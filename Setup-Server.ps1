<#
.SYNOPSIS
    Setup-Server.ps1 - Durcissement des ACL du share serveur PCPulse
.DESCRIPTION
    Script de hardening a executer en compte admin local sur le serveur
    hebergeant le share PCPulse$.

    Que fait ce script :
    1. Casse l'heritage NTFS sur la racine du share
    2. Cree les sous-dossiers \release\, \killed\, \logs\ avec ACLs EXPLICITES
       (au lieu d'heriter de la racine, ce qui donnait Domain Computers en
       Modify sur \release\ - faille majeure permettant a un PC compromis
       de pousser un Collector malveillant et obtenir SYSTEM sur tout le parc)
    3. Retire BUILTIN\Users (CreateFiles + AppendData) qui etait trop large
    4. Ajoute optionnellement un gMSA dedie (parametre -gMSAName)
    5. Domain Computers garde Modify SUR LES DOSSIERS NECESSAIRES uniquement
       (racine pour deposer JSON, killed/ pour rapport kill switch, logs/ pour
       les logs Collector). Sur \release\ : LECTURE SEULE.
    6. Owner = groupe Domain Admins (pas un compte personnel pour audit propre)

    Le script est IDEMPOTENT : peut etre relance sans risque pour re-aligner
    les ACLs si quelqu'un a fait une modif manuelle entre-temps.

    POST-CONDITIONS attendues apres execution :
    - \\<SERVER>\PCPulse$\           = Domain Computers Modify, gMSA Modify (si fourni)
    - \\<SERVER>\PCPulse$\release\   = Domain Computers READ ONLY,
                                       gMSA READ ONLY (si fourni),
                                       Admins du domaine FullControl
    - \\<SERVER>\PCPulse$\killed\    = Domain Computers Modify, gMSA Modify
    - \\<SERVER>\PCPulse$\logs\      = Domain Computers Modify, gMSA Modify

.NOTES
    Version  : 2.0
    Auteur   : Damien Gouhier
    Licence  : MIT

    PRE-REQUIS :
    - Compte admin local sur le serveur (admin local ou Domain Admin)
    - PCPulse$ deja cree (partage SMB + sous-dossiers ; voir docs/INSTALL.md,
      section "Creation du share" : New-SmbShare). Ce script NE cree PAS le
      share, il en durcit les ACL.
    - Optionnel : un gMSA cree au prealable dans AD si tu veux remplacer
      SYSTEM par un compte de service dedie (recommande)
    - Connexion AD fonctionnelle pour resoudre les SIDs

.PARAMETER ShareName
    Nom du share SMB heberge sur ce serveur. Defaut : 'PCPulse$'.

.PARAMETER gMSAName
    Nom du compte gMSA a ajouter aux ACLs. Format : 'nom-du-compte$' (avec le $ final).
    Exemple : 'svc-pcpulse$'.
    Si vide ou inexistant dans AD, le script skip cette partie sans erreur.
    Defaut : '' (pas de gMSA, classique SYSTEM).

.PARAMETER ExtraReadGroups
    Groupes AD existants a PRESERVER en ReadAndExecute sur tous les dossiers.
    Utile pour conserver les conventions infra de votre organisation
    (groupes admins T1, supervision, etc.) sans casser leurs workflows.
    Format : noms NetBIOS sans le domaine (ex: 'GR-ADMINS-IT').
    Defaut : @() (pas de groupes preserves)

.EXAMPLE
    # Configuration minimale (pas de gMSA, classique SYSTEM)
    .\Setup-Server.ps1

.EXAMPLE
    # Avec gMSA
    .\Setup-Server.ps1 -gMSAName 'svc-pcpulse$'

.EXAMPLE
    # Avec gMSA + preservation d'un groupe admin existant
    .\Setup-Server.ps1 -gMSAName 'svc-pcpulse$' -ExtraReadGroups @('GR-ADMINS-IT')

.EXAMPLE
    # Mode dry-run : affiche ce qui serait fait sans rien modifier
    .\Setup-Server.ps1 -WhatIf
#>

#Requires -Version 5.1
#Requires -RunAsAdministrator

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$ShareName = 'PCPulse$',

    # Nom du gMSA a ajouter aux ACLs (optionnel).
    # Format : 'nom-du-compte$' (avec le $ final, exemple 'svc-pcpulse$').
    # Si vide ou inexistant dans AD, cette partie est skippee.
    [string]$gMSAName  = '',

    # Groupes AD existants a PRESERVER en lecture sur tous les
    # dossiers du share. Permet de garder les conventions infra de
    # votre organisation (groupes admins, supervision, etc.) sans
    # casser leurs workflows existants.
    # Format : noms NetBIOS sans le domaine (ex: 'GR-ADMINS-IT').
    # Defaut : aucun groupe preserve.
    [string[]]$ExtraReadGroups = @()
)

# ============================================================
# HELPERS UI
# ============================================================
function Write-Step { Write-Host ("`n===== " + $args[0] + " =====") -ForegroundColor Cyan }
function Write-OK   { Write-Host ("    [OK]   " + $args[0]) -ForegroundColor Green }
function Write-Warn { Write-Host ("    [!!]   " + $args[0]) -ForegroundColor Yellow }
function Write-Err  { Write-Host ("    [KO]   " + $args[0]) -ForegroundColor Red }
function Write-Info { Write-Host ("    [..]   " + $args[0]) -ForegroundColor Gray }

Write-Host ""
Write-Host "======================================================" -ForegroundColor Cyan
Write-Host " PCPulse - Setup serveur (durcissement des ACL)" -ForegroundColor Cyan
Write-Host "======================================================" -ForegroundColor Cyan
if ($WhatIfPreference) {
    Write-Host " *** MODE DRY-RUN : aucune modification ne sera appliquee ***" -ForegroundColor Magenta
}

# ============================================================
# ETAPE 1 : Detection environnement et resolution des SIDs
# ============================================================
Write-Step "1/6 Detection environnement"

$cs = Get-CimInstance Win32_ComputerSystem
if (-not $cs.PartOfDomain) {
    Write-Err "Le serveur n'est pas joint a un domaine AD"
    exit 1
}
Write-OK "Serveur : $($env:COMPUTERNAME) sur domaine $($cs.Domain)"

# --- Resolution du SID du domaine via le compte machine local
try {
    $machineAccount = New-Object System.Security.Principal.NTAccount("$env:USERDNSDOMAIN\$env:COMPUTERNAME$")
    $machineSid     = $machineAccount.Translate([System.Security.Principal.SecurityIdentifier])
    $domainSid      = $machineSid.AccountDomainSid.Value
    Write-OK "SID du domaine : $domainSid"
} catch {
    Write-Err "Impossible de resoudre le SID du domaine : $_"
    exit 1
}

# --- Resolution des groupes AD via leurs SIDs well-known (RIDs Microsoft)
function Resolve-DomainGroup {
    param([string]$Rid, [string]$Label)
    try {
        $sid  = New-Object System.Security.Principal.SecurityIdentifier("$domainSid-$Rid")
        $name = $sid.Translate([System.Security.Principal.NTAccount]).Value
        Write-OK "$Label : $name"
        return @{ Sid = $sid; Name = $name }
    } catch {
        Write-Err "Echec resolution $Label (RID $Rid) : $_"
        return $null
    }
}

$grpAdmins    = Resolve-DomainGroup -Rid 512 -Label "Domain Admins"
$grpComputers = Resolve-DomainGroup -Rid 515 -Label "Domain Computers"

if (-not $grpAdmins -or -not $grpComputers) {
    Write-Err "Impossible de continuer sans les groupes AD de base"
    exit 1
}

# Resolution des groupes "extra" a preserver en lecture
# (groupes infra existants qu'on ne doit pas casser)
$extraGroupsResolved = @()
foreach ($grpName in $ExtraReadGroups) {
    if (-not $grpName) { continue }
    try {
        $acct = New-Object System.Security.Principal.NTAccount("$env:USERDOMAIN\$grpName")
        $sid  = $acct.Translate([System.Security.Principal.SecurityIdentifier])
        $extraGroupsResolved += @{ Sid = $sid; Name = "$env:USERDOMAIN\$grpName" }
        Write-OK "Groupe extra preserve : $env:USERDOMAIN\$grpName"
    } catch {
        Write-Warn "Groupe extra '$grpName' introuvable - sera ignore"
    }
}

# --- SIDs well-known universels (memes en FR/EN/DE/etc.)
$sidSystem        = New-Object System.Security.Principal.SecurityIdentifier('S-1-5-18')
$sidBuiltinAdmins = New-Object System.Security.Principal.SecurityIdentifier('S-1-5-32-544')
$sidCreatorOwner  = New-Object System.Security.Principal.SecurityIdentifier('S-1-3-0')
Write-OK "SIDs systeme resolus (SYSTEM, BUILTIN\Administrators, CREATOR OWNER)"

# --- Resolution du gMSA (si fourni)
$gmsaResolved = $null
if ([string]::IsNullOrWhiteSpace($gMSAName)) {
    Write-Info "Pas de gMSA fourni (-gMSAName), les ACLs seront definies sans"
} else {
    try {
        $gmsaAccount  = New-Object System.Security.Principal.NTAccount("$env:USERDOMAIN\$gMSAName")
        $gmsaSid      = $gmsaAccount.Translate([System.Security.Principal.SecurityIdentifier])
        $gmsaResolved = @{ Sid = $gmsaSid; Name = "$env:USERDOMAIN\$gMSAName" }
        Write-OK "gMSA $gMSAName : $($gmsaSid.Value)"
    } catch {
        Write-Warn "gMSA $gMSAName non trouve dans AD - les ACLs gMSA seront skippees"
        Write-Warn "  Verifie que le compte existe et que tu peux le resoudre"
    }
}

# ============================================================
# ETAPE 2 : Localisation du share et inventaire actuel
# ============================================================
Write-Step "2/6 Localisation share et inventaire"

$share = Get-SmbShare -Name $ShareName -ErrorAction SilentlyContinue
if (-not $share) {
    Write-Err "Le share $ShareName n'existe pas. Cree-le d'abord (voir docs/INSTALL.md, section 'Creation du share'), par ex. :"
    Write-Err "  New-SmbShare -Name '$ShareName' -Path 'D:\PCPulse' -FullAccess 'Administrateurs'"
    Write-Err "  puis cree les sous-dossiers release\ , killed\ , logs\"
    exit 1
}
$sharePath = $share.Path
Write-OK "Share trouve : $sharePath"

# Sous-dossiers attendus
$releaseDir = Join-Path $sharePath 'release'
$killedDir  = Join-Path $sharePath 'killed'
$logsDir    = Join-Path $sharePath 'logs'

# Inventaire actuel
foreach ($d in @($releaseDir, $killedDir, $logsDir)) {
    if (Test-Path $d) {
        $count = @(Get-ChildItem $d -Force -ErrorAction SilentlyContinue).Count
        Write-Info "  $d : existe ($count elements)"
    } else {
        Write-Warn "  $d : ABSENT (sera cree)"
    }
}

# ============================================================
# ETAPE 3 : Sauvegarde des ACLs actuelles (audit + rollback)
# ============================================================
Write-Step "3/6 Sauvegarde des ACLs actuelles"

$backupDir = Join-Path $sharePath '.acl-backup'
if ($PSCmdlet.ShouldProcess($backupDir, "Creer dossier backup ACLs")) {
    New-Item -Path $backupDir -ItemType Directory -Force | Out-Null
}

$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
foreach ($d in @($sharePath, $releaseDir, $killedDir, $logsDir)) {
    if (Test-Path $d) {
        $name = if ($d -eq $sharePath) { 'root' } else { Split-Path $d -Leaf }
        $backupFile = Join-Path $backupDir "acl-$name-$timestamp.txt"
        if ($PSCmdlet.ShouldProcess($backupFile, "Sauvegarder ACL")) {
            try {
                $aclData = @"
=== ACL backup $name - $(Get-Date) ===
Path: $d

Owner: $((Get-Acl $d).Owner)
SDDL: $((Get-Acl $d).Sddl)

Access rules:
$((Get-Acl $d).Access | Format-Table IdentityReference, FileSystemRights, AccessControlType, IsInherited -AutoSize | Out-String)
"@
                $aclData | Out-File -FilePath $backupFile -Encoding UTF8
                Write-OK "Backup ACL : $name -> acl-$name-$timestamp.txt"
            } catch {
                Write-Warn "Echec backup ACL $name : $_"
            }
        }
    }
}

# ============================================================
# ETAPE 4 : Creation des sous-dossiers manquants
# ============================================================
Write-Step "4/6 Creation des sous-dossiers"

foreach ($d in @($releaseDir, $killedDir, $logsDir)) {
    if (-not (Test-Path $d)) {
        if ($PSCmdlet.ShouldProcess($d, "Creer dossier")) {
            New-Item -Path $d -ItemType Directory -Force | Out-Null
            Write-OK "Cree : $d"
        }
    } else {
        Write-OK "Existe deja : $d"
    }
}

# ============================================================
# ETAPE 5 : ACLs explicites par dossier
# ============================================================
Write-Step "5/6 Application des ACLs hardenees"

# Helper : applique des ACLs explicites a un dossier (casse l'heritage,
# nettoie les regles existantes, applique les nouvelles).
function Set-HardenedAcl {
    param(
        [Parameter(Mandatory)] [string] $Path,
        [Parameter(Mandatory)] [array]  $Rules,
        [string] $OwnerSid
    )
    if (-not (Test-Path $Path)) {
        Write-Warn "Path absent : $Path"
        return
    }
    Write-Info "Hardening : $Path"

    if (-not $PSCmdlet.ShouldProcess($Path, "Apply hardened ACL")) {
        # En mode WhatIf, on liste juste les regles cibles
        foreach ($r in $Rules) {
            $name = try { $r.Identity.Translate([System.Security.Principal.NTAccount]).Value } catch { $r.Identity.Value }
            Write-Info "  WOULD: $name -> $($r.Right)"
        }
        return
    }

    # 1. Casser l'heritage : convertir les regles heritees en regles explicites
    $acl = Get-Acl $Path
    $acl.SetAccessRuleProtection($true, $true)  # protect=true, preserve=true
    Set-Acl -Path $Path -AclObject $acl

    # 2. Recuperer ACL fraiche apres protection, puis purger TOUTES les regles
    $acl = Get-Acl $Path
    @($acl.Access) | ForEach-Object {
        [void]$acl.RemoveAccessRule($_)
    }

    # 3. Re-appliquer les regles cibles
    $inheritFlags = [System.Security.AccessControl.InheritanceFlags]::ContainerInherit -bor `
                    [System.Security.AccessControl.InheritanceFlags]::ObjectInherit
    $propagation  = [System.Security.AccessControl.PropagationFlags]::None

    foreach ($r in $Rules) {
        # CREATOR OWNER ne doit s'appliquer qu'aux objets enfants (pas au dossier
        # lui-meme, sinon on cree une boucle bizarre)
        $iflags = $inheritFlags
        $pflags = $propagation
        if ($r.Identity.Value -eq 'S-1-3-0') {
            # CREATOR OWNER : InheritOnly + ContainerInherit + ObjectInherit
            $pflags = [System.Security.AccessControl.PropagationFlags]::InheritOnly
        }

        try {
            $rule = New-Object System.Security.AccessControl.FileSystemAccessRule(
                $r.Identity, $r.Right, $iflags, $pflags, "Allow"
            )
            $acl.AddAccessRule($rule)
            $name = try { $r.Identity.Translate([System.Security.Principal.NTAccount]).Value } catch { $r.Identity.Value }
            Write-OK "  $name : $($r.Right)"
        } catch {
            Write-Err "  Echec regle pour $($r.Identity) : $_"
        }
    }

    # 4. Owner si demande
    if ($OwnerSid) {
        try {
            $ownerSidObj = New-Object System.Security.Principal.SecurityIdentifier($OwnerSid)
            $acl.SetOwner($ownerSidObj)
            $ownerName = try { $ownerSidObj.Translate([System.Security.Principal.NTAccount]).Value } catch { $OwnerSid }
            Write-OK "  Owner : $ownerName"
        } catch {
            Write-Warn "  Echec set owner : $_"
        }
    }

    # 5. Apply
    try {
        Set-Acl -Path $Path -AclObject $acl -ErrorAction Stop
        Write-OK "  ACL appliquee"
    } catch {
        Write-Err "  Echec Set-Acl : $_"
    }
}

# ----- ACLs racine du share -----
# Racine : les PC doivent pouvoir y deposer leur JSON, mais pas modifier
# les sous-dossiers (qui auront leurs ACLs explicites).
$rulesRoot = @(
    @{ Identity = $sidSystem;        Right = 'FullControl' }
    @{ Identity = $sidBuiltinAdmins; Right = 'FullControl' }
    @{ Identity = $grpAdmins.Sid;    Right = 'FullControl' }
    @{ Identity = $grpComputers.Sid; Right = 'Modify' }     # depose JSON
    @{ Identity = $sidCreatorOwner;  Right = 'FullControl' } # owner sur ses fichiers
)
if ($gmsaResolved) {
    $rulesRoot += @{ Identity = $gmsaResolved.Sid; Right = 'Modify' }  # gMSA depose JSON
}
foreach ($grp in $extraGroupsResolved) {
    $rulesRoot += @{ Identity = $grp.Sid; Right = 'ReadAndExecute' }
}
Set-HardenedAcl -Path $sharePath -Rules $rulesRoot -OwnerSid $grpAdmins.Sid.Value

# ----- ACLs release\ (LE plus critique : LECTURE SEULE pour PC) -----
$rulesRelease = @(
    @{ Identity = $sidSystem;        Right = 'FullControl' }
    @{ Identity = $sidBuiltinAdmins; Right = 'FullControl' }
    @{ Identity = $grpAdmins.Sid;    Right = 'FullControl' }
    @{ Identity = $grpComputers.Sid; Right = 'ReadAndExecute' }  # LECTURE SEULE
)
if ($gmsaResolved) {
    $rulesRelease += @{ Identity = $gmsaResolved.Sid; Right = 'ReadAndExecute' }  # gMSA lit Collector
}
foreach ($grp in $extraGroupsResolved) {
    $rulesRelease += @{ Identity = $grp.Sid; Right = 'ReadAndExecute' }
}
Set-HardenedAcl -Path $releaseDir -Rules $rulesRelease -OwnerSid $grpAdmins.Sid.Value

# ----- ACLs killed\ -----
$rulesKilled = @(
    @{ Identity = $sidSystem;        Right = 'FullControl' }
    @{ Identity = $sidBuiltinAdmins; Right = 'FullControl' }
    @{ Identity = $grpAdmins.Sid;    Right = 'FullControl' }
    @{ Identity = $grpComputers.Sid; Right = 'Modify' }     # PC ecrivent leur rapport mort
    @{ Identity = $sidCreatorOwner;  Right = 'FullControl' }
)
if ($gmsaResolved) {
    $rulesKilled += @{ Identity = $gmsaResolved.Sid; Right = 'Modify' }
}
foreach ($grp in $extraGroupsResolved) {
    $rulesKilled += @{ Identity = $grp.Sid; Right = 'ReadAndExecute' }
}
Set-HardenedAcl -Path $killedDir -Rules $rulesKilled -OwnerSid $grpAdmins.Sid.Value

# ----- ACLs logs\ -----
$rulesLogs = @(
    @{ Identity = $sidSystem;        Right = 'FullControl' }
    @{ Identity = $sidBuiltinAdmins; Right = 'FullControl' }
    @{ Identity = $grpAdmins.Sid;    Right = 'FullControl' }
    @{ Identity = $grpComputers.Sid; Right = 'Modify' }     # PC ecrivent leurs logs
    @{ Identity = $sidCreatorOwner;  Right = 'FullControl' }
)
if ($gmsaResolved) {
    $rulesLogs += @{ Identity = $gmsaResolved.Sid; Right = 'Modify' }
}
foreach ($grp in $extraGroupsResolved) {
    $rulesLogs += @{ Identity = $grp.Sid; Right = 'ReadAndExecute' }
}
Set-HardenedAcl -Path $logsDir -Rules $rulesLogs -OwnerSid $grpAdmins.Sid.Value

# ============================================================
# ETAPE 6 : Verification post-application
# ============================================================
Write-Step "6/6 Verification post-application"

function Show-AclSummary {
    param([string]$Path, [string]$Label)
    if (-not (Test-Path $Path)) {
        Write-Warn "$Label : path absent"
        return
    }
    Write-Host "    $Label" -ForegroundColor Yellow
    $acl = Get-Acl $Path
    Write-Host "      Owner : $($acl.Owner)" -ForegroundColor Gray
    $acl.Access | ForEach-Object {
        $cls = if ($_.IsInherited) { 'Inherited' } else { 'Explicit' }
        $name = $_.IdentityReference.Value
        $rights = $_.FileSystemRights
        Write-Host "      [$cls] $name -> $rights" -ForegroundColor Gray
    }
    Write-Host ""
}

Show-AclSummary -Path $sharePath  -Label "Racine $sharePath"
Show-AclSummary -Path $releaseDir -Label "release\ (CRITIQUE - doit etre Read pour Domain Computers)"
Show-AclSummary -Path $killedDir  -Label "killed\"
Show-AclSummary -Path $logsDir    -Label "logs\"

# Tests d'acces simulé : verifier que Domain Computers ne peut PAS modifier release\
Write-Step "Tests securite (simulation acces)"

# Test 1 : Domain Computers peut-il toujours lire release\ ?
$canReadRelease = $false
try {
    $aclRelease = Get-Acl $releaseDir
    foreach ($rule in $aclRelease.Access) {
        if ($rule.IdentityReference.Translate([System.Security.Principal.SecurityIdentifier]).Value -eq $grpComputers.Sid.Value) {
            if ($rule.FileSystemRights -band [System.Security.AccessControl.FileSystemRights]::ReadAndExecute) {
                $canReadRelease = $true
                break
            }
        }
    }
} catch {}

if ($canReadRelease) {
    Write-OK "Domain Computers peut LIRE release\ (necessaire pour update Collector)"
} else {
    Write-Err "Domain Computers ne peut PAS lire release\ - le Updater va planter !"
}

# Test 2 : Domain Computers peut-il MODIFIER release\ ? (NE DOIT PAS pouvoir)
$canModifyRelease = $false
try {
    $aclRelease = Get-Acl $releaseDir
    foreach ($rule in $aclRelease.Access) {
        if ($rule.IdentityReference.Translate([System.Security.Principal.SecurityIdentifier]).Value -eq $grpComputers.Sid.Value) {
            if (($rule.FileSystemRights -band [System.Security.AccessControl.FileSystemRights]::Write) -or
                ($rule.FileSystemRights -band [System.Security.AccessControl.FileSystemRights]::Modify) -or
                ($rule.FileSystemRights -band [System.Security.AccessControl.FileSystemRights]::FullControl)) {
                # Verifier que c'est pas juste WriteAttributes / WriteExtendedAttributes
                $rightsStr = $rule.FileSystemRights.ToString()
                if ($rightsStr -match 'Modify|Write,|FullControl|^Write$' -and $rightsStr -notmatch 'Attributes') {
                    $canModifyRelease = $true
                    break
                }
            }
        }
    }
} catch {}

if ($canModifyRelease) {
    Write-Err "FAILLE : Domain Computers peut encore modifier release\ - hardening incomplet"
} else {
    Write-OK "Domain Computers ne peut PAS modifier release\ (exposition fermee)"
}

# Test 3 : gMSA si present
if ($gmsaResolved) {
    $canGmsaRead = $false
    try {
        $aclRelease = Get-Acl $releaseDir
        foreach ($rule in $aclRelease.Access) {
            if ($rule.IdentityReference.Translate([System.Security.Principal.SecurityIdentifier]).Value -eq $gmsaResolved.Sid.Value) {
                if ($rule.FileSystemRights -band [System.Security.AccessControl.FileSystemRights]::ReadAndExecute) {
                    $canGmsaRead = $true
                    break
                }
            }
        }
    } catch {}

    if ($canGmsaRead) {
        Write-OK "gMSA peut LIRE release\ (necessaire pour Updater)"
    } else {
        Write-Err "gMSA ne peut PAS lire release\ - migration va planter"
    }
}

Write-Host ""
Write-Host "======================================================" -ForegroundColor Green
Write-Host " Hardening termine" -ForegroundColor Green
Write-Host "======================================================" -ForegroundColor Green
Write-Host ""
Write-Host "Backup ACLs original : $backupDir" -ForegroundColor Cyan
Write-Host ""
Write-Host "PROCHAINES ETAPES :" -ForegroundColor Cyan
Write-Host "  1. Tester depuis 1 PC pilote :" -ForegroundColor White
Write-Host "       - L'Updater peut toujours lire release\ (download nouveau Collector)"
Write-Host "       - Le Collector peut deposer son JSON a la racine"
Write-Host "       - Le killswitch peut ecrire dans killed\"
Write-Host ""
Write-Host "  2. Tester en compte utilisateur que tu ne peux PAS modifier release\ :" -ForegroundColor White
Write-Host '       New-Item -Path \\<SERVER>\PCPulse$\release\test.txt -ItemType File'
Write-Host "       (doit echouer avec 'Acces refuse')"
Write-Host ""
Write-Host "  3. Apres validation, faire la meme verif depuis un compte machine compromis (simulation)" -ForegroundColor White
Write-Host "  4. Documenter les ACLs finales dans SECURITY.md du repo GitHub" -ForegroundColor White
Write-Host ""
