<#
.SYNOPSIS
    Install-Client.ps1 - Installe PCPulse sur un PC pilote (avec auto-update)
.DESCRIPTION
    A executer sur chaque PC pilote en session Administrateur.
    Execute en sequence :
      1. Verifie la connectivite au share serveur
      2. Cree C:\ProgramData\PCPulse\
      3. Copie 01_Collector.ps1 + PCPulse-Updater.ps1 + version.txt
      4. Cree la tache planifiee PCPulse-Collector (pointe sur l'UPDATER)
      5. Lance immediatement pour generer le premier JSON
      6. Verifie la presence du JSON sur le serveur

    La tache planifiee pointe sur PCPulse-Updater.ps1, pas sur le Collector
    directement. Ainsi, a chaque cycle horaire, le PC verifie automatiquement
    si une nouvelle version du Collector est disponible sur le share
    (\\SERVER\PCPulse$\release\) et se met a jour le cas echeant.

    Voir SECURITY.md pour le trust model recommande (ACLs, gMSA, etc.).

.PARAMETER ServerPath
    Chemin UNC du partage serveur (obligatoire).
    Ex: \\SRV-PCPULSE\PCPulse$

.PARAMETER SourceDir
    Dossier local contenant 01_Collector.ps1, PCPulse-Updater.ps1, et
    optionnellement version.txt. Par defaut : dossier du script courant.

.PARAMETER InitialVersion
    Version a ecrire dans version.txt si aucun n'est fourni dans SourceDir.
    Par defaut : "2.0".

.PARAMETER gMSAName
    Nom du compte gMSA a utiliser pour la tache planifiee, au lieu de SYSTEM
    (recommande pour audit propre, voir SECURITY.md).
    Format : 'nom-du-compte$' (avec le $ final).
    Pre-requis : le PC doit etre membre du groupe AD ayant
    PrincipalsAllowedToRetrieveManagedPassword sur le gMSA. Test prealable :
        Test-ADServiceAccount -Identity 'svc-pcpulse'
    Si vide ou non fourni, la tache tourne en SYSTEM (comportement par defaut).
    Si fourni mais le compte n'est pas resolvable depuis ce PC, le script
    echoue avec un message clair (pas de fallback silencieux sur SYSTEM).

.PARAMETER SkipFirstRun
    Ne pas lancer la tache immediatement apres creation.

.EXAMPLE
    # Installation classique en SYSTEM
    .\Install-Client.ps1 -ServerPath "\\SRV-PCPULSE\PCPulse$"

.EXAMPLE
    # Installation avec gMSA dedie (recommande)
    .\Install-Client.ps1 -ServerPath "\\SRV-PCPULSE\PCPulse$" -gMSAName 'svc-pcpulse$'

.EXAMPLE
    # Avec dossier source alternatif
    .\Install-Client.ps1 -ServerPath "\\SRV-PCPULSE\PCPulse$" -SourceDir "C:\Temp\pcpulse"

.NOTES
    Version  : 2.1 (support gMSA optionnel, voir SECURITY.md)
    Auteur   : Damien Gouhier
    Licence  : MIT

    Fichiers attendus dans SourceDir :
      - 01_Collector.ps1     (obligatoire)
      - PCPulse-Updater.ps1  (obligatoire)
      - version.txt          (optionnel, "2.1" par defaut)

    CHANGELOG :
    v2.1 : Resolution robuste de SourceDir (cascade PSScriptRoot /
           PSCommandPath / MyInvocation au lieu d'un defaut de parametre
           fragile qui pouvait etre vide selon le contexte d'appel).
           Correction parsing du message d'aide gMSA (-replace sorti de
           la string interpolee).
           Bump version par defaut 2.0 -> 2.1 (aligne sur la release v2.1).
    v2.0 : Support gMSA optionnel via -gMSAName (au lieu de SYSTEM hardcode).
           Bump version par defaut 1.0 -> 2.0 pour aligner sur la release v2.0.
           Pointe vers SECURITY.md pour le trust model.
    v1.0 : Initiale - installation classique en SYSTEM.
#>

#Requires -Version 5.1
#Requires -RunAsAdministrator

param(
    [Parameter(Mandatory = $true)]
    [string]$ServerPath,

    [string]$SourceDir,

    [string]$InitialVersion = '2.1',

    # v2.0 : support gMSA optionnel pour la tache planifiee
    # Si vide, on garde l'ancien comportement (SYSTEM via ServiceAccount)
    [string]$gMSAName = '',

    [switch]$SkipFirstRun
)

# ============================================================
# RESOLUTION ROBUSTE DU DOSSIER SOURCE
# ============================================================
# $PSScriptRoot utilise comme valeur par defaut de parametre est fragile :
# selon le contexte d'invocation (-File, dot-source, hote tiers) il peut etre
# vide au moment du binding des parametres. On resout donc $SourceDir ici,
# dans le corps du script, via une cascade de methodes fiables. Si -SourceDir
# est fourni explicitement a l'appel, il est respecte tel quel.
if ([string]::IsNullOrWhiteSpace($SourceDir)) {
    if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
        $SourceDir = $PSScriptRoot
    } elseif (-not [string]::IsNullOrWhiteSpace($PSCommandPath)) {
        $SourceDir = Split-Path -Parent $PSCommandPath
    } elseif ($MyInvocation.MyCommand.Path) {
        $SourceDir = Split-Path -Parent $MyInvocation.MyCommand.Path
    } else {
        Write-Host " [KO] Impossible de determiner le dossier source automatiquement." -ForegroundColor Red
        Write-Host " [KO] Relancer en passant -SourceDir avec le chemin du dossier contenant" -ForegroundColor Red
        Write-Host " [KO] 01_Collector.ps1, PCPulse-Updater.ps1 et version.txt." -ForegroundColor Red
        exit 1
    }
}

# ============================================================
# CONFIGURATION
# ============================================================
$LocalPath    = "C:\ProgramData\PCPulse"
$CollectorSrc = Join-Path $SourceDir '01_Collector.ps1'
$UpdaterSrc   = Join-Path $SourceDir 'PCPulse-Updater.ps1'
$VersionSrc   = Join-Path $SourceDir 'version.txt'

$CollectorDst = Join-Path $LocalPath '01_Collector.ps1'
$UpdaterDst   = Join-Path $LocalPath 'PCPulse-Updater.ps1'
$VersionDst   = Join-Path $LocalPath 'version.txt'
$BackupDir    = Join-Path $LocalPath 'backup'

$TaskName     = 'PCPulse-Collector'

# Couleurs
function Write-Step { Write-Host ("`n===== " + $args[0] + " =====") -ForegroundColor Cyan }
function Write-OK   { Write-Host (" [OK] " + $args[0]) -ForegroundColor Green }
function Write-Warn { Write-Host (" [!!] " + $args[0]) -ForegroundColor Yellow }
function Write-Err  { Write-Host (" [KO] " + $args[0]) -ForegroundColor Red }
function Write-Info { Write-Host (" [..] " + $args[0]) -ForegroundColor Gray }

Write-Host ""
Write-Host "======================================================" -ForegroundColor Cyan
Write-Host " PCPulse - Install client sur $($env:COMPUTERNAME)" -ForegroundColor Cyan
if ($gMSAName) {
    Write-Host " Mode: avec auto-updater + gMSA ($gMSAName)" -ForegroundColor Cyan
} else {
    Write-Host " Mode: avec auto-updater (SYSTEM)" -ForegroundColor Cyan
}
Write-Host "======================================================" -ForegroundColor Cyan

# ============================================================
# ETAPE 1 : Verification des sources
# ============================================================
Write-Step "1/7 Verification des sources"

if (-not (Test-Path $CollectorSrc)) {
    Write-Err "Source du Collector introuvable : $CollectorSrc"
    Write-Err "Placer 01_Collector.ps1 dans le meme dossier que ce script,"
    Write-Err "ou passer -SourceDir avec le chemin complet."
    exit 1
}
Write-OK "Collector : $CollectorSrc ($([math]::Round((Get-Item $CollectorSrc).Length / 1KB, 1)) KB)"

if (-not (Test-Path $UpdaterSrc)) {
    Write-Err "Source de l'Updater introuvable : $UpdaterSrc"
    Write-Err "Placer PCPulse-Updater.ps1 dans le meme dossier que ce script."
    exit 1
}
Write-OK "Updater : $UpdaterSrc ($([math]::Round((Get-Item $UpdaterSrc).Length / 1KB, 1)) KB)"

$versionToWrite = $InitialVersion
if (Test-Path $VersionSrc) {
    $versionToWrite = (Get-Content $VersionSrc -Raw).Trim()
    Write-OK "Version (depuis version.txt source) : $versionToWrite"
} else {
    Write-Info "version.txt absent dans SourceDir, utilisation de : $versionToWrite"
}

# ============================================================
# ETAPE 1bis : Validation gMSA (si fourni)
# ============================================================
if ($gMSAName) {
    Write-Step "1bis/7 Validation gMSA"

    # Test-ADServiceAccount n'existe que si RSAT-AD est installe.
    # On essaye, sinon on skip avec un warning.
    if (Get-Command Test-ADServiceAccount -ErrorAction SilentlyContinue) {
        $gmsaShort = $gMSAName -replace '\$$',''
        try {
            $ok = Test-ADServiceAccount -Identity $gmsaShort -ErrorAction Stop
            if ($ok) {
                Write-OK "gMSA $gMSAName accessible depuis ce PC"
            } else {
                Write-Err "Test-ADServiceAccount a retourne False pour $gmsaShort"
                Write-Err "Verifier que ce PC est dans le groupe AD ayant"
                Write-Err "PrincipalsAllowedToRetrieveManagedPassword sur le gMSA."
                Write-Err "Note : un reboot est souvent necessaire apres ajout au groupe."
                exit 1
            }
        } catch {
            Write-Err "Test-ADServiceAccount a echoue : $_"
            Write-Err "Le gMSA $gMSAName ne sera probablement pas utilisable."
            exit 1
        }
    } else {
        Write-Warn "Test-ADServiceAccount non disponible (RSAT-AD non installe sur ce PC)"
        Write-Warn "On continue sans validation prealable - la tache echouera"
        Write-Warn "au demarrage si le gMSA n'est pas accessible."
    }
}

# ============================================================
# ETAPE 2 : Test connectivite au share serveur
# ============================================================
Write-Step "2/7 Test acces au partage serveur"
$serverName = ($ServerPath -split '\\')[2]

$ping = Test-NetConnection -ComputerName $serverName -Port 445 -WarningAction SilentlyContinue
if (-not $ping.TcpTestSucceeded) {
    Write-Err "Serveur $serverName injoignable sur le port 445 (SMB)"
    Write-Err "Verifier : firewall, DNS, joint au meme domaine"
    exit 1
}
Write-OK "Port 445 ouvert vers $serverName"

try {
    $accessible = Test-Path $ServerPath -ErrorAction Stop
    if ($accessible) {
        Write-OK "Partage accessible : $ServerPath"
    } else {
        Write-Warn "Partage $ServerPath inexistant"
    }
} catch {
    Write-Warn "Partage $ServerPath inaccessible depuis ce compte utilisateur"
    Write-Warn "(Note : la tache planifiee tournera en SYSTEM/gMSA, Kerberos devrait passer)"
    Write-Warn "On continue (sera teste a l'etape 6)"
}

# ============================================================
# ETAPE 3 : Deployer les scripts localement
# ============================================================
Write-Step "3/7 Deploiement local"

if (-not (Test-Path $LocalPath)) {
    New-Item -Path $LocalPath -ItemType Directory -Force | Out-Null
    Write-OK "Cree : $LocalPath"
}
if (-not (Test-Path $BackupDir)) {
    New-Item -Path $BackupDir -ItemType Directory -Force | Out-Null
    Write-OK "Cree : $BackupDir"
}

# Copier le Collector
Copy-Item -Path $CollectorSrc -Destination $CollectorDst -Force
$srcHash = (Get-FileHash $CollectorSrc -Algorithm SHA256).Hash
$dstHash = (Get-FileHash $CollectorDst -Algorithm SHA256).Hash
if ($srcHash -ne $dstHash) {
    Write-Err "Hash SHA256 different pour Collector : copie corrompue"
    exit 1
}
Write-OK "Collector copie : $CollectorDst (SHA256 OK)"

# Copier l'Updater
Copy-Item -Path $UpdaterSrc -Destination $UpdaterDst -Force
$srcHash = (Get-FileHash $UpdaterSrc -Algorithm SHA256).Hash
$dstHash = (Get-FileHash $UpdaterDst -Algorithm SHA256).Hash
if ($srcHash -ne $dstHash) {
    Write-Err "Hash SHA256 different pour Updater : copie corrompue"
    exit 1
}
Write-OK "Updater copie : $UpdaterDst (SHA256 OK)"

# Ecrire version.txt
Set-Content -Path $VersionDst -Value $versionToWrite -Encoding UTF8
Write-OK "version.txt ecrit : $versionToWrite"

# ============================================================
# ETAPE 4 : Creation de la tache planifiee
# ============================================================
Write-Step "4/7 Tache planifiee PCPulse-Collector"

$existingTask = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($existingTask) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
    Write-Info "Tache existante supprimee"
}

# IMPORTANT : la tache pointe sur l'UPDATER, pas sur le Collector direct.
$argList = '-ExecutionPolicy Bypass -WindowStyle Hidden -NonInteractive ' +
           "-File `"$UpdaterDst`" -SharePath `"$ServerPath`""
$action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $argList

$triggerStartup = New-ScheduledTaskTrigger -AtStartup
$triggerHourly  = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(5) `
                  -RepetitionInterval (New-TimeSpan -Hours 1) `
                  -RepetitionDuration (New-TimeSpan -Days 9999)

# v2.0 : Principal SYSTEM par defaut, ou gMSA si fourni
if ($gMSAName) {
    # gMSA : LogonType=Password est obligatoire (pas ServiceAccount)
    # Note : le compte doit etre au format "DOMAIN\nom$"
    $gmsaUserId = "$env:USERDOMAIN\$gMSAName"
    $principal = New-ScheduledTaskPrincipal -UserId $gmsaUserId `
                                             -LogonType Password `
                                             -RunLevel Highest
    Write-OK "Principal : $gmsaUserId (gMSA, LogonType=Password)"
} else {
    $principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' `
                                             -LogonType ServiceAccount `
                                             -RunLevel Highest
    Write-OK "Principal : SYSTEM (LogonType=ServiceAccount)"
}

$settings = New-ScheduledTaskSettingsSet `
              -AllowStartIfOnBatteries `
              -DontStopIfGoingOnBatteries `
              -StartWhenAvailable `
              -ExecutionTimeLimit (New-TimeSpan -Minutes 30) `
              -MultipleInstances IgnoreNew `
              -RestartCount 2 `
              -RestartInterval (New-TimeSpan -Minutes 10)

try {
    Register-ScheduledTask -TaskName $TaskName `
                          -Description "PCPulse : wrapper updater + collecte horaire" `
                          -Action $action `
                          -Trigger @($triggerStartup, $triggerHourly) `
                          -Principal $principal `
                          -Settings $settings -ErrorAction Stop | Out-Null
} catch {
    Write-Err "Echec creation de la tache : $_"
    if ($gMSAName) {
        Write-Err "Verifier que :"
        Write-Err "  1. Le PC est dans le groupe AD pour ce gMSA"
        Write-Err "  2. Un reboot a eu lieu apres l'ajout au groupe"
        Write-Err "  3. Le gMSA a 'Log on as a batch job' (via GPO)"
        $gMSAShort = $gMSAName -replace '\$$', ''
        Write-Err "Tester avec : Test-ADServiceAccount -Identity '$gMSAShort'"
    }
    exit 1
}
Write-OK "Tache $TaskName creee"
Write-OK "  Execute  : PCPulse-Updater.ps1 (wrapper de mise a jour)"
Write-OK "  Triggers : au demarrage + toutes les heures"

# ============================================================
# ETAPE 5 : Premier run
# ============================================================
if ($SkipFirstRun) {
    Write-Step "5/7 Premier run (SAUTE par -SkipFirstRun)"
    Write-Info "La tache demarrera au prochain boot ou dans 5 min"
} else {
    Write-Step "5/7 Premier run pour validation"
    Write-Info "Demarrage de la tache..."
    Start-ScheduledTask -TaskName $TaskName

    Write-Info "Attente de la fin d'execution (max 2 min)..."
    Write-Info "(Note: le Collector a un delai anti-collision aleatoire ~0-10 min,"
    Write-Info " il est normal de voir un timeout au 1er run, le JSON arrivera +tard)"

    $timeout = 120
    $elapsed = 0
    while ($elapsed -lt $timeout) {
        Start-Sleep -Seconds 5
        $elapsed += 5
        $taskInfo = Get-ScheduledTask -TaskName $TaskName
        if ($taskInfo.State -eq 'Ready') { break }
    }
    if ($elapsed -ge $timeout) {
        Write-Warn "Timeout : la tache tourne encore (normal, delai anti-collision)"
    } else {
        Write-OK "Execution terminee en ~$elapsed secondes"
    }
}

# ============================================================
# ETAPE 6 : Verification cote serveur (best effort)
# ============================================================
Write-Step "6/7 Verification sur le serveur"

$expectedJson = Join-Path $ServerPath "$($env:COMPUTERNAME).json"
$expectedLog  = Join-Path $ServerPath "logs\$($env:COMPUTERNAME).log"

Start-Sleep -Seconds 3

try {
    if (Test-Path $expectedJson -ErrorAction Stop) {
        $jsonInfo = Get-Item $expectedJson
        Write-OK "JSON present : $expectedJson"
        Write-OK "  Taille : $([math]::Round($jsonInfo.Length / 1KB, 1)) KB"
        Write-OK "  Modifie : $($jsonInfo.LastWriteTime)"
    } else {
        Write-Warn "JSON absent (peut etre normal si Collector en delai anti-collision)"
    }
} catch {
    Write-Warn "Impossible de verifier le JSON depuis ce compte (droits SMB)"
    Write-Warn "Verifier manuellement sur le serveur : Get-ChildItem $ServerPath"
}

try {
    if (Test-Path $expectedLog -ErrorAction Stop) {
        Write-OK "Log present : $expectedLog"
    } else {
        Write-Warn "Log absent (peut etre normal au tout 1er run)"
    }
} catch {
    Write-Warn "Impossible de verifier le log depuis ce compte (droits SMB)"
}

# ============================================================
# ETAPE 7 : Resume
# ============================================================
Write-Step "7/7 Resume"
Write-Host ""
Write-Host "======================================================" -ForegroundColor Green
Write-Host "  Installation terminee sur $($env:COMPUTERNAME)" -ForegroundColor Green
Write-Host "======================================================" -ForegroundColor Green
Write-Host ""
Write-Host "LAYOUT LOCAL :" -ForegroundColor Cyan
Write-Host "  $LocalPath\"
Write-Host "    |- 01_Collector.ps1       (version $versionToWrite)"
Write-Host "    |- PCPulse-Updater.ps1    (wrapper auto-update)"
Write-Host "    |- version.txt            ($versionToWrite)"
Write-Host "    |- backup\                (historique versions, rolling 5)"
Write-Host "    |- updater.log            (cree au 1er run)"
Write-Host ""
Write-Host "COMMANDES UTILES :" -ForegroundColor Cyan
Write-Host "  Etat tache    : Get-ScheduledTask -TaskName $TaskName | Format-List"
Write-Host "  Forcer run    : Start-ScheduledTask -TaskName $TaskName"
Write-Host "  Log updater   : Get-Content $LocalPath\updater.log -Tail 20"
Write-Host ""
Write-Host "POUR PUBLIER UNE NOUVELLE VERSION DU COLLECTOR :" -ForegroundColor Cyan
Write-Host "  1. Sur le serveur, mettre a jour $ServerPath\release\ :"
Write-Host "     - Copier le nouveau 01_Collector.ps1"
Write-Host "     - Ecrire version.txt avec le nouveau numero (ex: 2.1)"
Write-Host "  2. Les PC se mettront a jour dans l'heure qui suit."
Write-Host ""
Write-Host "POUR DESINSTALLER (proprement) :" -ForegroundColor Cyan
Write-Host "  Voir le killswitch (deposer un fichier sentinelle dans \release\)"
Write-Host "  Voir SECURITY.md > Killswitch pour la procedure complete."
Write-Host ""
Write-Host "POUR DESINSTALLER MANUELLEMENT :" -ForegroundColor Cyan
Write-Host "  Unregister-ScheduledTask -TaskName '$TaskName' -Confirm:`$false"
Write-Host "  Remove-Item '$LocalPath' -Recurse -Force"
Write-Host ""
