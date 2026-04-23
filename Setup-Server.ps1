<#
.SYNOPSIS
    Setup-Server.ps1 - Prepare le serveur pour heberger PCPulse
.DESCRIPTION
    Execute sur le serveur SRV-PCPULSE (Windows Server 2022 joint au
    domaine votre-domaine.fr) en session Administrateur local.

    Execute en sequence :
      1. Verification des prerequis (domaine, features SMB, firewall)
      2. Creation du dossier D:\PCPulse (ou C:\PCPulse si D: absent)
      3. Creation du partage SMB cache \\SRV-PCPULSE\PCPulse$
      4. Configuration des permissions NTFS explicites
      5. Depot de config.psd1 et ip-ranges.csv de base

    Compatible Windows Server 2022. Peut etre relance sans risque
    (idempotent sur les elements cles : partage, permissions).

.NOTES
    Version : 1.0
    Auteur  : Damien Gouhier
    Licence : MIT

.EXAMPLE
    # En session admin local du serveur :
    Set-ExecutionPolicy -Scope Process Bypass
    .\Setup-Server.ps1
#>

#Requires -Version 5.1
#Requires -RunAsAdministrator

# ============================================================
# CONFIGURATION (adapter si besoin)
# ============================================================
# NOTE : $DomainName est auto-detecte (etape 1). Si l'auto-detect
# echoue ou renvoie quelque chose d'incorrect, hardcode la valeur
# ici (ex: $DomainName = 'MONDOMAINE').
$DomainName      = $null                     # auto-detecte plus bas
$ShareName       = 'PCPulse$'                # nom du partage (avec $ = cache)
$PreferredDrive  = 'D:\PCPulse'              # chemin preferre
$FallbackDrive   = 'C:\PCPulse'              # fallback si D: absent

# Couleurs pour les messages
function Write-Step  { Write-Host ("`n===== " + $args[0] + " =====") -ForegroundColor Cyan }
function Write-OK    { Write-Host ("  [OK] " + $args[0])             -ForegroundColor Green }
function Write-Warn  { Write-Host ("  [!!] " + $args[0])             -ForegroundColor Yellow }
function Write-Err   { Write-Host ("  [KO] " + $args[0])             -ForegroundColor Red }
function Write-Info  { Write-Host ("  [..] " + $args[0])             -ForegroundColor Gray }

Write-Host ""
Write-Host "======================================================" -ForegroundColor Cyan
Write-Host "  PCPulse - Setup serveur (SRV-PCPULSE)"                  -ForegroundColor Cyan
Write-Host "======================================================" -ForegroundColor Cyan

# ============================================================
# ETAPE 1 : Verifications prerequis
# ============================================================
Write-Step "1/5 Verifications prerequis"

# 1.1 - Domaine + auto-detection du nom NetBIOS
$cs = Get-CimInstance Win32_ComputerSystem
if ($cs.Domain -eq 'WORKGROUP' -or -not $cs.PartOfDomain) {
    Write-Err "Le serveur n'est pas joint a un domaine AD."
    Write-Err "Joindre d'abord avec : Add-Computer -DomainName '<votre.domaine.fr>' -Credential (Get-Credential) -Restart"
    exit 1
}
Write-OK "Serveur joint au domaine FQDN : $($cs.Domain)"

# Auto-detection du nom NetBIOS (nt4-style, e.g. MONDOMAINE)
try {
    $ntDomain = (Get-CimInstance Win32_NTDomain | Where-Object { $_.DomainName -and $_.DnsForestName -eq $cs.Domain } | Select-Object -First 1).DomainName
    if (-not $ntDomain) {
        $ntDomain = ($cs.Domain -split '\.')[0].ToUpper()
    }
    $DomainName = $ntDomain
    Write-OK "Nom NetBIOS du domaine detecte : $DomainName"
} catch {
    $DomainName = ($cs.Domain -split '\.')[0].ToUpper()
    Write-Warn "Auto-detection partielle, utilisation de : $DomainName"
}

# ============================================================
# Resolution des groupes via leurs SIDs "well-known".
# Les SIDs sont UNIVERSELS (meme en AD francais ou les groupes
# s'appellent "Admins du domaine" / "Ordinateurs du domaine").
#
# Structure d'un SID de groupe built-in :
#   <SID du domaine>-<RID>
#
# RIDs standards Microsoft :
#   -512 = Domain Admins (ou "Admins du domaine" en FR)
#   -515 = Domain Computers (ou "Ordinateurs du domaine" en FR)
#
# Ref: https://learn.microsoft.com/en-us/windows/security/identity-protection/access-control/security-identifiers
# ============================================================

# 1. Extraire le SID du domaine via le compte machine local (qui a un SID de domaine)
try {
    $machineAccount = New-Object System.Security.Principal.NTAccount("$env:USERDNSDOMAIN\$env:COMPUTERNAME$")
    $machineSid = $machineAccount.Translate([System.Security.Principal.SecurityIdentifier])
    $domainSid = $machineSid.AccountDomainSid.Value
    Write-OK "SID du domaine : $domainSid"
} catch {
    Write-Err "Impossible de resoudre le SID du domaine via le compte machine : $_"
    Write-Err "Verifier que le serveur est authentifie aupres d'un DC."
    exit 1
}

# 2. Traduire les SIDs -512 et -515 en noms de groupes (selon la langue du Windows)
try {
    $sidAdmins = New-Object System.Security.Principal.SecurityIdentifier("$domainSid-512")
    $grpAdminsName = $sidAdmins.Translate([System.Security.Principal.NTAccount]).Value
    $grpAdmins = @{ Account = $grpAdminsName; Sid = $sidAdmins.Value }
    Write-OK "Groupe admin resolu : $grpAdminsName"
} catch {
    Write-Err "Impossible de resoudre le groupe Domain Admins (SID -512)"
    exit 1
}

try {
    $sidComputers = New-Object System.Security.Principal.SecurityIdentifier("$domainSid-515")
    $grpComputersName = $sidComputers.Translate([System.Security.Principal.NTAccount]).Value
    $grpComputers = @{ Account = $grpComputersName; Sid = $sidComputers.Value }
    Write-OK "Groupe ordinateurs resolu : $grpComputersName"
} catch {
    Write-Err "Impossible de resoudre le groupe Domain Computers (SID -515)"
    exit 1
}

Write-OK "Nom NetBIOS du serveur : $($env:COMPUTERNAME)"

# 1.2 - Feature File Server
$feature = Get-WindowsFeature FS-FileServer
if ($feature.InstallState -ne 'Installed') {
    Write-Info "Installation du role File Server..."
    Install-WindowsFeature FS-FileServer -IncludeManagementTools | Out-Null
    Write-OK "Role File Server installe"
} else {
    Write-OK "Role File Server deja present"
}

# 1.3 - Firewall SMB
$smbRules = Get-NetFirewallRule -DisplayGroup "File and Printer Sharing" -ErrorAction SilentlyContinue |
    Where-Object { $_.Enabled -eq $false -and $_.Profile -match 'Domain|Any' }
if ($smbRules) {
    Write-Info "Activation des regles firewall SMB (domaine)..."
    Enable-NetFirewallRule -DisplayGroup "File and Printer Sharing" -ErrorAction SilentlyContinue
    Write-OK "Regles firewall SMB activees"
} else {
    Write-OK "Firewall SMB deja ouvert"
}

# ============================================================
# ETAPE 2 : Creation du dossier
# ============================================================
Write-Step "2/5 Creation du dossier de stockage"

# Choisir D: ou C: selon disponibilite
$TargetPath = $FallbackDrive
if (Test-Path 'D:\') {
    $TargetPath = $PreferredDrive
    Write-OK "Disque D: detecte, utilisation de $TargetPath"
} else {
    Write-Warn "Disque D: non present, utilisation de $TargetPath"
}

# Creer le dossier principal + logs
if (-not (Test-Path $TargetPath)) {
    New-Item -Path $TargetPath -ItemType Directory -Force | Out-Null
    Write-OK "Cree : $TargetPath"
} else {
    Write-OK "Deja present : $TargetPath"
}

$logsPath = Join-Path $TargetPath 'logs'
if (-not (Test-Path $logsPath)) {
    New-Item -Path $logsPath -ItemType Directory -Force | Out-Null
    Write-OK "Cree : $logsPath"
} else {
    Write-OK "Deja present : $logsPath"
}

# ============================================================
# ETAPE 3 : Creation du partage SMB
# ============================================================
Write-Step "3/5 Creation du partage SMB"

$existingShare = Get-SmbShare -Name $ShareName -ErrorAction SilentlyContinue
if ($existingShare) {
    Write-Warn "Partage $ShareName existe deja, mise a jour des acces"
    # Supprimer les acces existants pour repartir propre
    $existingShare | Get-SmbShareAccess |
        ForEach-Object { Revoke-SmbShareAccess -Name $ShareName -AccountName $_.AccountName -Force -ErrorAction SilentlyContinue } |
        Out-Null
} else {
    New-SmbShare -Name $ShareName `
                 -Path $TargetPath `
                 -Description "PCPulse fleet monitoring share" `
                 -EncryptData $true | Out-Null
    Write-OK "Partage $ShareName cree (SMB3 encrypted)"
}

# Retirer "Everyone" par defaut AVANT d'ajouter les vraies permissions
$everyoneAccess = Get-SmbShareAccess -Name $ShareName -ErrorAction SilentlyContinue | Where-Object AccountName -match "Everyone|Tout le monde"
if ($everyoneAccess) {
    foreach ($e in $everyoneAccess) {
        Revoke-SmbShareAccess -Name $ShareName -AccountName $e.AccountName -Force -ErrorAction SilentlyContinue | Out-Null
    }
    Write-OK "Acces 'Everyone/Tout le monde' retire"
}

# Acces Full pour le groupe admin du domaine (Domain Admins / Admins du domaine)
try {
    Grant-SmbShareAccess -Name $ShareName -AccountName $grpAdmins.Account -AccessRight Full -Force -ErrorAction Stop | Out-Null
    Write-OK "$($grpAdmins.Account) : Full Access"
} catch {
    Write-Err "Echec Grant $($grpAdmins.Account) : $_"
}

# Acces Change (read+write) pour les comptes machine (Domain Computers / Ordinateurs du domaine)
try {
    Grant-SmbShareAccess -Name $ShareName -AccountName $grpComputers.Account -AccessRight Change -Force -ErrorAction Stop | Out-Null
    Write-OK "$($grpComputers.Account) : Change (read+write)"
} catch {
    Write-Err "Echec Grant $($grpComputers.Account) : $_"
}

Write-Info "Acces au partage configures :"
Get-SmbShareAccess -Name $ShareName | Format-Table AccountName, AccessRight, AccessControlType -AutoSize

# ============================================================
# ETAPE 4 : Permissions NTFS
# ============================================================
Write-Step "4/5 Permissions NTFS"

# Retirer l'heritage, preserver les droits actuels comme explicites
$acl = Get-Acl $TargetPath
$acl.SetAccessRuleProtection($true, $true)
Set-Acl $TargetPath $acl

# Rebuild les droits proprement
$acl = Get-Acl $TargetPath

# Retirer les regles "Users" et "Authenticated Users" (trop larges)
@($acl.Access) | Where-Object {
    $_.IdentityReference.Value -match 'Users$|Authenticated Users|Everyone'
} | ForEach-Object { $acl.RemoveAccessRule($_) | Out-Null }

# Ajouter les 4 permissions explicites. Toutes sont resolues via leurs SIDs
# "well-known" pour garantir que ca marche sur Windows FR/EN/DE/etc.
#
# SIDs universels (Microsoft) :
#   S-1-5-18     = NT AUTHORITY\SYSTEM
#   S-1-5-32-544 = BUILTIN\Administrators (Administrateurs en FR)
#   <domain>-512 = Domain Admins (Admins du domaine en FR)
#   <domain>-515 = Domain Computers (Ordinateurs du domaine en FR)
$inheritFlags = [System.Security.AccessControl.InheritanceFlags]::ContainerInherit -bor
                [System.Security.AccessControl.InheritanceFlags]::ObjectInherit
$propagation  = [System.Security.AccessControl.PropagationFlags]::None

# Resoudre SYSTEM et BUILTIN\Administrators via leurs SIDs universels
$sidSystem = New-Object System.Security.Principal.SecurityIdentifier('S-1-5-18')
$sidBuiltinAdmins = New-Object System.Security.Principal.SecurityIdentifier('S-1-5-32-544')

$rules = @(
    # SYSTEM : indispensable pour les services Windows
    @{ Identity = $sidSystem;                Right = "FullControl" }
    # Administrators locaux : full (SID -544, nom localise par Windows)
    @{ Identity = $sidBuiltinAdmins;         Right = "FullControl" }
    # Domain Admins / Admins du domaine : full (SID -512)
    @{ Identity = $sidAdmins;                Right = "FullControl" }
    # Domain Computers / Ordinateurs du domaine : Modify (SID -515)
    @{ Identity = $sidComputers;             Right = "Modify" }
)

foreach ($r in $rules) {
    try {
        # Utiliser le SID directement dans la rule (fonctionne sur tous les Windows)
        $rule = New-Object System.Security.AccessControl.FileSystemAccessRule(
            $r.Identity, $r.Right, $inheritFlags, $propagation, "Allow"
        )
        $acl.AddAccessRule($rule)
        # Pour l'affichage, on traduit en nom lisible
        try {
            $displayName = $r.Identity.Translate([System.Security.Principal.NTAccount]).Value
        } catch {
            $displayName = $r.Identity.Value
        }
        Write-OK "$displayName : $($r.Right)"
    } catch {
        Write-Err "Echec ACL pour $($r.Identity) : $($_.Exception.Message)"
    }
}

try {
    Set-Acl $TargetPath $acl -ErrorAction Stop
    Write-OK "Permissions NTFS appliquees"
} catch {
    Write-Err "Echec Set-Acl : $_"
}

# ============================================================
# ETAPE 5 : Fichiers de config par defaut
# ============================================================
Write-Step "5/5 Fichiers de config par defaut"

$configFile = Join-Path $TargetPath 'config.psd1'
if (Test-Path $configFile) {
    Write-Warn "config.psd1 existe deja, non modifie"
} else {
    $configContent = @'
@{
    HistoriqueJours      = 30
    SeuilBootLong        = 2
    SeuilBootMax         = 30
    SeuilDiskAlert       = 10
    SeuilCrashRecent     = 7
    SeuilDiskWarning     = 25
    SeuilOfflineJours    = 1
    DashboardTitle       = 'PCPulse'
    DashboardSubtitle    = 'Supervision du parc - Pilote'
    MaskHealthyByDefault = $false
    CsvRanges            = 'ip-ranges.csv'
    ScoreWeights         = @{
        BSOD          = 5
        WHEA          = 4
        Crash         = 3
        Thermal       = 3
        GPU_TDR       = 2
        DiskAlert     = 2
        BootLong      = 1
        WHEACorrected = 1
        BSODRecent    = 2
        Offline       = 3
        BatteryAlert  = 2
        BootPerfAlert = 2
        SMARTAlert    = 4
    }
}
'@
    $configContent | Out-File -FilePath $configFile -Encoding UTF8
    Write-OK "config.psd1 cree"
}

$csvFile = Join-Path $TargetPath 'ip-ranges.csv'
if (Test-Path $csvFile) {
    Write-Warn "ip-ranges.csv existe deja, non modifie"
} else {
    $csvContent = @'
"Field1","Pattern1","Field2","Pattern2","Entity"
"primary_local_ip","192.168.1.0/24","","","A-COMPLETER"
'@
    $csvContent | Out-File -FilePath $csvFile -Encoding UTF8
    Write-Warn "ip-ranges.csv cree avec 1 ligne placeholder - A COMPLETER avec les vraies plages IP du parc !"
}

# ============================================================
# RESUME
# ============================================================
$sharePathUNC       = "\\$($env:COMPUTERNAME)\$ShareName"
$sharePathUNCFull   = "\\$($env:COMPUTERNAME).$($cs.Domain)\$ShareName"
$csvPath            = Join-Path $TargetPath 'ip-ranges.csv'

Write-Host ""
Write-Host "======================================================" -ForegroundColor Green
Write-Host "  Setup termine !"                                     -ForegroundColor Green
Write-Host "======================================================" -ForegroundColor Green
Write-Host ""
Write-Host "Partage SMB accessible via :"
Write-Host "  $sharePathUNC" -ForegroundColor Yellow
Write-Host "  $sharePathUNCFull" -ForegroundColor Yellow
Write-Host ""
Write-Host "Contenu actuel :"
Get-ChildItem $TargetPath | Format-Table Name, Length, LastWriteTime -AutoSize

Write-Host ""
Write-Host "PROCHAINES ETAPES :" -ForegroundColor Cyan
Write-Host ""
Write-Host "  1. Sur un PC client, tester la connectivite :" -ForegroundColor White
Write-Host "     Test-NetConnection -ComputerName $($env:COMPUTERNAME) -Port 445"
Write-Host "     Get-ChildItem $sharePathUNC"
Write-Host ""
Write-Host "  2. Sur un PC pilote, installer PCPulse via :" -ForegroundColor White
Write-Host "     Install-Client.ps1 -ServerPath $sharePathUNC"
Write-Host ""
Write-Host "  3. Completer ${csvPath} avec les vraies plages IP" -ForegroundColor White
Write-Host ""
