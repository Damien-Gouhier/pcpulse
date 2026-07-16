<#
.SYNOPSIS
    PCPulse-Updater.ps1 - Auto-update du Collector + KillSwitch
.DESCRIPTION
    Wrapper execute par la tache planifiee a la place du Collector direct.
    A chaque run :
      1. Verifie le killswitch (fichier sentinelle sur le share)
         -> si actif, auto-desinstalle PCPulse et exit
      2. Self-update : si l'Updater lui-meme differe de release\ (SHA256),
         se met a jour (effectif au prochain run)
      3. Verifie s'il existe une nouvelle version du Collector sur le share
      4. Si diff : backup, copie, verif SHA256, met a jour version locale
      5. Lance le Collector local avec le SharePath fourni

    Workflow killswitch :
      - L'admin depose un fichier sentinelle dans \release\ (defaut : KILLSWITCH.txt)
      - Le fichier doit contenir une phrase de confirmation exacte
        (defaut : CONFIRM-UNINSTALL-PCPULSE)
      - Le PC ecrit un rapport dans \killed\<HOSTNAME>.txt
      - Le PC supprime sa tache planifiee, son dossier local, et exit
      - L'admin retire le fichier sentinelle quand toutes les machines
        ont ete kill (cleanup manuel)

    En cas d'echec a n'importe quelle etape (sauf killswitch), le Collector
    deja present localement continue de tourner. AUCUN rollback auto.

.PARAMETER SharePath
    Chemin UNC du share serveur (ex: \\SRV-PCPULSE\PCPulse$).
    Transmis au Collector apres l'update.

.EXAMPLE
    # Utilise par la tache planifiee :
    PCPulse-Updater.ps1 -SharePath "\\SRV-PCPULSE\PCPulse$"

.NOTES
    Version : 1.3
    Auteur  : Damien Gouhier
    Licence : MIT

    CHANGELOG :
    v1.3 : updater.log deplace sur le share (logs\<HOST>.updater.log) : plus aucun
           journal sur le poste agent, purge de l'ancien updater.log local. Durcissement
           ACL du dossier runtime (SYSTEM + Administrateurs, SID en dur) via Set-PCPulseAcl.
    v1.2 : Housekeeping local (Invoke-LocalCleanup) : a chaque cycle, supprime
           le dossier de deploiement C:\Temp\PCPulse-deploy laisse par
           Deploy-PCPulse, borne updater.log a N jours, sweep les battreport
           orphelins. Ne touche pas PsExec. Non-bloquant, idempotent.
    v1.1 : Ajout du killswitch (auto-desinstallation a distance via
           fichier sentinelle). Configurable via config.psd1.
    v1.0 : Initiale - auto-update + verification SHA256.
    v1.1 : Self-update de l'Updater par SHA256 (resout le bootstrap : un Updater
           deploye one-shot par GPO peut desormais se mettre a jour seul, sans
           re-deploiement). Backups Updater purges par Remove-OldBackups.

    Layout attendu cote serveur :
        \\SRV-PCPULSE\PCPulse$\
            release\
                01_Collector.ps1     (derniere version officielle)
                PCPulse-Updater.ps1  (self-update de l'Updater par SHA256)
                version.txt          (ex: "1.1")
                KILLSWITCH.txt       (optionnel, declenche le kill)
            killed\                 (rapports de desinstallation)

    Layout genere cote client :
        C:\ProgramData\PCPulse\
            01_Collector.ps1        (version courante)
            PCPulse-Updater.ps1     (ce script)
            version.txt             (version active)
            backup\                 (5 derniers backups, rolling)
            (updater.log deplace sur le share : logs\<HOST>.updater.log)
            .update.lock            (lock file temporaire)

    Personnalisation killswitch (optionnel) via config.psd1 :
        KillSwitch = @{
            Enabled  = $true
            Filename = 'MON_NOM.txt'
            Phrase   = 'MA-PHRASE-SECRETE'
        }
#>

#Requires -Version 5.1

param(
    [Parameter(Mandatory = $true)]
    [string]$SharePath
)

# ============================================================
# CONFIGURATION
# ============================================================
$LocalDir       = "C:\ProgramData\PCPulse"
$CollectorLocal = Join-Path $LocalDir '01_Collector.ps1'
$UpdaterLocal   = Join-Path $LocalDir 'PCPulse-Updater.ps1'
$VersionLocal   = Join-Path $LocalDir 'version.txt'
$LockFile       = Join-Path $LocalDir '.update.lock'
$BackupDir      = Join-Path $LocalDir 'backup'
$UpdaterLog       = Join-Path $SharePath "logs\$($env:COMPUTERNAME).updater.log"  # v1.3 : sur le share, plus sur le poste agent
$LegacyUpdaterLog = Join-Path $LocalDir 'updater.log'                             # ancien emplacement local, purge par Invoke-LocalCleanup

$ReleaseDir   = Join-Path $SharePath 'release'
$KilledDir    = Join-Path $SharePath 'killed'
$CollectorSrv = Join-Path $ReleaseDir '01_Collector.ps1'
$VersionSrv   = Join-Path $ReleaseDir 'version.txt'
$UpdaterSrv   = Join-Path $ReleaseDir 'PCPulse-Updater.ps1'

$BackupRetention = 5

# v1.2 : nettoyage local des traces post-deploiement (housekeeping)
$DeployLeftoverDir       = 'C:\Temp\PCPulse-deploy'   # depose par Deploy-PCPulse, inutile apres install
$UpdaterLogRetentionDays = 7                          # borne updater.log (rotation par date)
$TempBattReportPattern   = 'battreport_*.xml'         # rapports batterie orphelins du Collector

# Defauts killswitch (si pas surcharge dans config.psd1)
$KillSwitchDefaults = @{
    Enabled  = $true
    Filename = 'KILLSWITCH.txt'
    Phrase   = 'CONFIRM-UNINSTALL-PCPULSE'
}

# Nom de la tache planifiee a supprimer lors du kill
$ScheduledTaskName = 'PCPulse-Collector'

# ============================================================
# FONCTIONS UTILITAIRES
# ============================================================
function Write-UpdaterLog {
    param([string]$Message, [string]$Level = 'INFO')
    try {
        $logDir = Split-Path $UpdaterLog
        if ($logDir -and -not (Test-Path -LiteralPath $logDir)) { $null = New-Item $logDir -ItemType Directory -Force -ErrorAction SilentlyContinue }
        $line = "{0} | {1} | {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
        Add-Content -Path $UpdaterLog -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue
    } catch {
        # Silent : on ne veut pas qu'un probleme de log casse l'updater
    }
}

function Get-KillSwitchConfig {
    # Charge la config killswitch depuis config.psd1 si dispo, sinon defauts
    $cfg = $KillSwitchDefaults.Clone()
    $configFile = Join-Path $SharePath 'config.psd1'

    if (Test-Path $configFile) {
        try {
            $userCfg = Import-PowerShellDataFile -Path $configFile -ErrorAction Stop
            if ($userCfg.ContainsKey('KillSwitch')) {
                if ($null -ne $userCfg.KillSwitch.Enabled)  { $cfg.Enabled  = [bool]$userCfg.KillSwitch.Enabled }
                if ($userCfg.KillSwitch.Filename)            { $cfg.Filename = [string]$userCfg.KillSwitch.Filename }
                if ($userCfg.KillSwitch.Phrase)              { $cfg.Phrase   = [string]$userCfg.KillSwitch.Phrase }
            }
        } catch {
            Write-UpdaterLog "Echec lecture config.psd1, utilisation des defauts killswitch : $_" 'WARN'
        }
    }
    return $cfg
}

function Invoke-KillSwitch {
    # Verifie le fichier sentinelle. Retourne $true si le killswitch
    # a ete declenche (auto-desinstallation effectuee, l'updater doit exit).
    $cfg = Get-KillSwitchConfig

    if (-not $cfg.Enabled) {
        return $false
    }

    $sentinelFile = Join-Path $ReleaseDir $cfg.Filename
    if (-not (Test-Path $sentinelFile)) {
        return $false
    }

    # Le fichier sentinelle existe : verifier le contenu
    try {
        $content = (Get-Content $sentinelFile -Raw -ErrorAction Stop).Trim()
    } catch {
        Write-UpdaterLog "Sentinelle detectee mais lecture echouee : $_" 'ERROR'
        return $false
    }

    if ($content -ne $cfg.Phrase) {
        Write-UpdaterLog "Sentinelle '$($cfg.Filename)' detectee mais phrase invalide, kill ignore" 'WARN'
        return $false
    }

    # Killswitch confirme : auto-desinstallation
    Write-UpdaterLog "=== KILLSWITCH ACTIF ===" 'WARN'
    Write-UpdaterLog "Fichier sentinelle : $sentinelFile"
    Write-UpdaterLog "Auto-desinstallation en cours..."

    # 1. Ecrire un rapport dans \killed\<HOSTNAME>.txt sur le share
    try {
        if (-not (Test-Path $KilledDir)) {
            New-Item -Path $KilledDir -ItemType Directory -Force | Out-Null
        }
        $report = @"
=== PCPulse Killswitch report ===
ComputerName : $env:COMPUTERNAME
Timestamp    : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Version      : $(if (Test-Path $VersionLocal) { (Get-Content $VersionLocal -Raw).Trim() } else { 'unknown' })
Sentinelle   : $($cfg.Filename)
"@
        $reportFile = Join-Path $KilledDir "$env:COMPUTERNAME.txt"
        $report | Out-File -FilePath $reportFile -Encoding UTF8 -Force -ErrorAction SilentlyContinue
        Write-UpdaterLog "Rapport killed ecrit : $reportFile"
    } catch {
        Write-UpdaterLog "Echec ecriture rapport killed (non-bloquant) : $_" 'WARN'
    }

    # 2. Supprimer la tache planifiee (avant de supprimer les fichiers,
    #    pour eviter qu'elle se relance pendant le cleanup)
    try {
        $task = Get-ScheduledTask -TaskName $ScheduledTaskName -ErrorAction SilentlyContinue
        if ($task) {
            Unregister-ScheduledTask -TaskName $ScheduledTaskName -Confirm:$false -ErrorAction Stop
            Write-UpdaterLog "Tache planifiee supprimee : $ScheduledTaskName"
        }
    } catch {
        Write-UpdaterLog "Echec suppression tache planifiee : $_" 'ERROR'
        # On continue quand meme : si la tache reste, elle re-tentera le kill au prochain cycle
    }

    # 3. Supprimer le dossier local PCPulse
    #    Note : on ne peut pas se supprimer soi-meme (ce script est en cours
    #    d'execution depuis $UpdaterLocal). On lance un job differe qui supprime
    #    apres notre exit.
    try {
        $cleanupCmd = @"
Start-Sleep -Seconds 5
Remove-Item -Path '$LocalDir' -Recurse -Force -ErrorAction SilentlyContinue
"@
        # Lance un PowerShell detache qui nettoiera apres notre exit
        Start-Process -FilePath 'powershell.exe' `
                      -ArgumentList '-NoProfile','-WindowStyle','Hidden','-Command',$cleanupCmd `
                      -WindowStyle Hidden | Out-Null
        Write-UpdaterLog "Cleanup differe lance (suppression de $LocalDir dans 5s)"
    } catch {
        Write-UpdaterLog "Echec lancement cleanup differe : $_" 'ERROR'
    }

    Write-UpdaterLog "=== Fin killswitch (exit) ==="
    return $true
}

function Invoke-Collector {
    param([string]$SharePath)
    if (-not (Test-Path $CollectorLocal)) {
        Write-UpdaterLog "Collector local introuvable : $CollectorLocal" 'ERROR'
        return
    }
    Write-UpdaterLog "Execution du Collector local (SharePath: $SharePath)"
    try {
        & $CollectorLocal -SharePath $SharePath
        Write-UpdaterLog "Collector termine"
    } catch {
        Write-UpdaterLog "Erreur execution Collector : $_" 'ERROR'
    }
}

function Remove-OldBackups {
    if (-not (Test-Path $BackupDir)) { return }
    # v1.2 : purge les deux familles de backups (Collector + Updater), retention chacune
    foreach ($pattern in @('01_Collector_v*.ps1', 'PCPulse-Updater_*.ps1')) {
        $backups = Get-ChildItem -Path $BackupDir -Filter $pattern -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTime -Descending
        if ($backups.Count -gt $BackupRetention) {
            $toDelete = $backups | Select-Object -Skip $BackupRetention
            foreach ($b in $toDelete) {
                Remove-Item $b.FullName -Force -ErrorAction SilentlyContinue
                Write-UpdaterLog "Backup ancien supprime : $($b.Name)"
            }
        }
    }
}

function Update-Self {
    # ============================================================
    # v1.2 : SELF-UPDATE de l'Updater par comparaison SHA256.
    #   - Compare le hash de CE script (local) a release\PCPulse-Updater.ps1
    #   - Si different : backup, copie, re-verif SHA256, restauration si echec
    #   - NON-BLOQUANT : tout echec est logge mais n'interrompt pas le cycle
    #   - La nouvelle version s'applique au PROCHAIN declenchement : ce process
    #     tourne deja sur l'ancien code charge en memoire, on ne re-execute pas.
    # ============================================================
    if (-not (Test-Path $UpdaterSrv)) {
        return   # Pas d'Updater publie dans release\ : rien a faire
    }

    try {
        $srvHash = (Get-FileHash -Path $UpdaterSrv -Algorithm SHA256 -ErrorAction Stop).Hash
    } catch {
        Write-UpdaterLog "Self-update : echec lecture SHA256 Updater serveur (non-bloquant) : $_" 'WARN'
        return
    }

    $localHash = $null
    if (Test-Path $UpdaterLocal) {
        try { $localHash = (Get-FileHash -Path $UpdaterLocal -Algorithm SHA256 -ErrorAction Stop).Hash } catch { $localHash = $null }
    }

    if ($localHash -and ($localHash -eq $srvHash)) {
        return   # Updater deja a jour
    }

    Write-UpdaterLog "Self-update : nouvelle version Updater detectee (local: $localHash | serveur: $srvHash)"

    # Backup de l'Updater courant
    $timestamp  = Get-Date -Format 'yyyyMMdd_HHmmss'
    $backupFile = Join-Path $BackupDir "PCPulse-Updater_${timestamp}.ps1"
    if (Test-Path $UpdaterLocal) {
        try {
            Copy-Item -Path $UpdaterLocal -Destination $backupFile -Force -ErrorAction Stop
            Write-UpdaterLog "Self-update : backup cree : $(Split-Path $backupFile -Leaf)"
        } catch {
            Write-UpdaterLog "Self-update : echec backup, abandon (non-bloquant) : $_" 'WARN'
            return
        }
    }

    # Copie du nouvel Updater depuis release\
    try {
        Copy-Item -Path $UpdaterSrv -Destination $UpdaterLocal -Force -ErrorAction Stop
    } catch {
        Write-UpdaterLog "Self-update : echec copie nouvel Updater : $_" 'ERROR'
        return
    }

    # Re-verif SHA256 apres copie (integrite)
    try {
        $newHash = (Get-FileHash -Path $UpdaterLocal -Algorithm SHA256 -ErrorAction Stop).Hash
        if ($newHash -ne $srvHash) {
            Write-UpdaterLog "Self-update : SHA256 mismatch apres copie (attendu $srvHash, obtenu $newHash)" 'ERROR'
            if (Test-Path $backupFile) {
                Copy-Item -Path $backupFile -Destination $UpdaterLocal -Force -ErrorAction SilentlyContinue
                Write-UpdaterLog "Self-update : backup restaure suite au mismatch SHA256"
            }
            return
        }
    } catch {
        Write-UpdaterLog "Self-update : echec re-verif SHA256 (non-bloquant) : $_" 'WARN'
        return
    }

    Write-UpdaterLog "Self-update : Updater mis a jour, effectif au prochain cycle" 'SUCCESS'
}

function Set-PCPulseAcl {
    # Verrouille le dossier runtime : SYSTEM + Administrateurs uniquement.
    # SID en dur (S-1-5-18 SYSTEM / S-1-5-32-544 Administrateurs) pour etre
    # independant de la langue de Windows (parc FR). Idempotent (ne fait rien
    # si l'heritage est deja coupe), non-bloquant.
    param([string]$Path)
    try {
        if (-not (Test-Path -LiteralPath $Path)) { return }
        $acl = Get-Acl -LiteralPath $Path
        if ($acl.AreAccessRulesProtected) { return }
        $sidSystem = New-Object System.Security.Principal.SecurityIdentifier('S-1-5-18')
        $sidAdmins = New-Object System.Security.Principal.SecurityIdentifier('S-1-5-32-544')
        $inherit   = [System.Security.AccessControl.InheritanceFlags]'ContainerInherit,ObjectInherit'
        $none      = [System.Security.AccessControl.PropagationFlags]::None
        $full      = [System.Security.AccessControl.FileSystemRights]::FullControl
        $allow     = [System.Security.AccessControl.AccessControlType]::Allow
        # Coupe l'heritage (ne recopie pas les ACE herites) puis retire les ACE explicites.
        $acl.SetAccessRuleProtection($true, $false)
        foreach ($ace in @($acl.Access | Where-Object { -not $_.IsInherited })) { [void]$acl.RemoveAccessRule($ace) }
        [void]$acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($sidSystem, $full, $inherit, $none, $allow)))
        [void]$acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($sidAdmins, $full, $inherit, $none, $allow)))
        Set-Acl -LiteralPath $Path -AclObject $acl
    } catch {
        # non-bloquant : un echec ACL ne casse jamais l'updater / l'install
    }
}

function Invoke-LocalCleanup {
    # v1.2 : HOUSEKEEPING LOCAL - supprime les traces post-deploiement qui ne
    # servent plus a rien et pourraient renseigner un attaquant sur un poste.
    # Tourne a chaque cycle. Idempotent, non-bloquant. Ne touche JAMAIS le
    # produit vivant (tache, scripts runtime, JSON buffer, backups) ni PsExec.

    # 1. Dossier de deploiement laisse par Deploy-PCPulse (Install-Client ne
    #    l'effacait pas). Revele share/serveur/killswitch/layout en clair.
    #    Garde stricte : uniquement ce chemin exact, jamais le dossier runtime.
    try {
        # Suppression INCONDITIONNELLE : la console de deploiement porte un nom
        # DIFFERENT (jamais C:\Temp\PCPulse-deploy), donc aucune regle de contenu.
        if ((Test-Path -LiteralPath $DeployLeftoverDir) -and ($DeployLeftoverDir -ine $LocalDir)) {
            Remove-Item -LiteralPath $DeployLeftoverDir -Recurse -Force -ErrorAction Stop
            Write-UpdaterLog "Cleanup : dossier de deploiement supprime ($DeployLeftoverDir)"
        }
    } catch {
        Write-UpdaterLog "Cleanup : echec suppression dossier deploiement (non-bloquant) : $_" 'WARN'
    }

    # 2. Rapports batterie orphelins du Collector dans le TEMP (SYSTEM -> C:\Windows\Temp).
    try {
        $tempDir = [System.IO.Path]::GetTempPath()
        Get-ChildItem -Path $tempDir -Filter $TempBattReportPattern -File -ErrorAction SilentlyContinue | ForEach-Object {
            Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue
            Write-UpdaterLog "Cleanup : rapport batterie orphelin supprime ($($_.Name))"
        }
    } catch {
        Write-UpdaterLog "Cleanup : echec sweep battreport (non-bloquant) : $_" 'WARN'
    }

    # 3. Rotation de updater.log : ne garder que les N derniers jours (borne la
    #    taille et limite la divulgation d'historique). Meme discipline CRLF que
    #    Invoke-LogCleanup du Collector (split tolerant, rejoin CRLF, newline final).
    try {
        if (Test-Path -LiteralPath $UpdaterLog) {
            $cutoff  = (Get-Date).AddDays(-$UpdaterLogRetentionDays).ToString('yyyy-MM-dd')
            $content = Get-Content -LiteralPath $UpdaterLog -Raw -ErrorAction Stop
            $lignes  = $content -split "`r?`n"
            $gardees = foreach ($l in $lignes) {
                $ligne = $l.TrimEnd("`r")
                if (($ligne.Length -ge 10) -and ($ligne.Substring(0, 10) -ge $cutoff)) { $ligne }
            }
            $finalText = if ($gardees) { ($gardees -join "`r`n") + "`r`n" } else { '' }
            [System.IO.File]::WriteAllText($UpdaterLog, $finalText, [System.Text.UTF8Encoding]::new($false))
        }
    } catch {
        Write-UpdaterLog "Cleanup : echec rotation updater.log (non-bloquant) : $_" 'WARN'
    }

    # 4. Purge de l'ancien updater.log LOCAL (v1.3 : le log est desormais sur le
    #    share). Retire la trace residuelle des postes deja deployes.
    try {
        if (Test-Path -LiteralPath $LegacyUpdaterLog) {
            Remove-Item -LiteralPath $LegacyUpdaterLog -Force -ErrorAction SilentlyContinue
        }
    } catch { }

    # 5. Durcissement ACL du dossier runtime : SYSTEM + Administrateurs uniquement
    #    (empeche un compte standard de lire les scripts = chemin share, killswitch).
    Set-PCPulseAcl -Path $LocalDir
}

# ============================================================
# INIT : creation du layout local si necessaire
# ============================================================
if (-not (Test-Path $LocalDir))  { New-Item -Path $LocalDir  -ItemType Directory -Force | Out-Null }
if (-not (Test-Path $BackupDir)) { New-Item -Path $BackupDir -ItemType Directory -Force | Out-Null }

Write-UpdaterLog "=== Debut cycle updater ==="
Write-UpdaterLog "SharePath: $SharePath"
Write-UpdaterLog "ComputerName: $env:COMPUTERNAME"

# ============================================================
# ETAPE 0 (NEW v1.1) : KILLSWITCH - verifier AVANT tout le reste
# ============================================================
if (Test-Path $ReleaseDir) {
    if (Invoke-KillSwitch) {
        # Killswitch declenche : on a auto-desinstalle, on exit immediatement
        Write-UpdaterLog "=== Fin cycle updater (killswitch) ==="
        exit 0
    }
}

# ============================================================
# ETAPE 0b (NEW v1.2) : HOUSEKEEPING LOCAL (traces post-deploiement)
# ============================================================
Invoke-LocalCleanup

# ============================================================
# ETAPE 1 : VERROU ANTI-COLLISION
# ============================================================
if (Test-Path $LockFile) {
    $lockAge = (Get-Date) - (Get-Item $LockFile).LastWriteTime
    if ($lockAge.TotalMinutes -lt 30) {
        Write-UpdaterLog "Lock actif (age $([int]$lockAge.TotalMinutes) min), abandon" 'WARN'
        exit 0
    }
    Write-UpdaterLog "Lock orphelin detecte (age $([int]$lockAge.TotalMinutes) min), suppression"
    Remove-Item $LockFile -Force -ErrorAction SilentlyContinue
}

try {
    New-Item -Path $LockFile -ItemType File -Force | Out-Null

    # ============================================================
    # ETAPE 2 : VERIFICATION CONNECTIVITE SHARE
    # ============================================================
    if (-not (Test-Path $ReleaseDir)) {
        Write-UpdaterLog "Share release inaccessible ($ReleaseDir), execution du Collector local sans update" 'WARN'
        Invoke-Collector -SharePath $SharePath
        return
    }

    # ============================================================
    # ETAPE 2b : SELF-UPDATE DE L'UPDATER (effectif au prochain cycle)
    # ============================================================
    Update-Self

    # ============================================================
    # ETAPE 3 : LECTURE DES VERSIONS
    # ============================================================
    $localVer = if (Test-Path $VersionLocal) {
        (Get-Content $VersionLocal -Raw -ErrorAction SilentlyContinue).Trim()
    } else {
        '0.0'
    }

    if (-not (Test-Path $VersionSrv)) {
        Write-UpdaterLog "Pas de version.txt serveur, execution du Collector local (version $localVer)" 'WARN'
        Invoke-Collector -SharePath $SharePath
        return
    }

    $srvVer = (Get-Content $VersionSrv -Raw -ErrorAction Stop).Trim()
    Write-UpdaterLog "Version locale: $localVer | Version serveur: $srvVer"

    # ============================================================
    # ETAPE 4 : COMPARAISON
    # ============================================================
    if ($localVer -eq $srvVer) {
        Write-UpdaterLog "Deja a jour"
        Invoke-Collector -SharePath $SharePath
        return
    }

    # ============================================================
    # ETAPE 5 : UPDATE
    # ============================================================
    Write-UpdaterLog "Update detectee : $localVer -> $srvVer"

    if (-not (Test-Path $CollectorSrv)) {
        Write-UpdaterLog "version.txt annonce $srvVer mais 01_Collector.ps1 absent du share" 'ERROR'
        Invoke-Collector -SharePath $SharePath
        return
    }

    # 5a - SHA256 source AVANT copie
    try {
        $srvHash = (Get-FileHash -Path $CollectorSrv -Algorithm SHA256 -ErrorAction Stop).Hash
        Write-UpdaterLog "SHA256 serveur : $srvHash"
    } catch {
        Write-UpdaterLog "Echec lecture SHA256 serveur : $_" 'ERROR'
        Invoke-Collector -SharePath $SharePath
        return
    }

    # 5b - Backup
    $backupFile = $null
    if (Test-Path $CollectorLocal) {
        $timestamp  = Get-Date -Format 'yyyyMMdd_HHmmss'
        $backupFile = Join-Path $BackupDir "01_Collector_v${localVer}_${timestamp}.ps1"
        try {
            Copy-Item -Path $CollectorLocal -Destination $backupFile -Force -ErrorAction Stop
            Write-UpdaterLog "Backup cree : $(Split-Path $backupFile -Leaf)"
        } catch {
            Write-UpdaterLog "Echec backup : $_" 'ERROR'
            Invoke-Collector -SharePath $SharePath
            return
        }
    }

    # 5c - Copie nouveau Collector
    try {
        Copy-Item -Path $CollectorSrv -Destination $CollectorLocal -Force -ErrorAction Stop
        Write-UpdaterLog "Copie du nouveau Collector OK"
    } catch {
        Write-UpdaterLog "Echec copie du nouveau Collector : $_" 'ERROR'
        if ($backupFile -and (Test-Path $backupFile)) {
            try {
                Copy-Item -Path $backupFile -Destination $CollectorLocal -Force -ErrorAction Stop
                Write-UpdaterLog "Backup restaure (continue sur $localVer)"
            } catch {
                Write-UpdaterLog "Echec restauration backup : $_" 'ERROR'
            }
        }
        Invoke-Collector -SharePath $SharePath
        return
    }

    # 5d - Verification SHA256 apres copie
    try {
        $localHash = (Get-FileHash -Path $CollectorLocal -Algorithm SHA256 -ErrorAction Stop).Hash
        if ($localHash -ne $srvHash) {
            Write-UpdaterLog "SHA256 mismatch apres copie (attendu $srvHash, obtenu $localHash)" 'ERROR'
            if ($backupFile -and (Test-Path $backupFile)) {
                Copy-Item -Path $backupFile -Destination $CollectorLocal -Force
                Write-UpdaterLog "Backup restaure suite au mismatch SHA256"
            }
            Invoke-Collector -SharePath $SharePath
            return
        }
        Write-UpdaterLog "SHA256 valide : $localHash"
    } catch {
        Write-UpdaterLog "Erreur verification SHA256 : $_" 'ERROR'
        Invoke-Collector -SharePath $SharePath
        return
    }

    # 5e - Mise a jour version.txt local
    try {
        Set-Content -Path $VersionLocal -Value $srvVer -Encoding UTF8 -ErrorAction Stop
        Write-UpdaterLog "version.txt local mis a jour : $srvVer" 'SUCCESS'
    } catch {
        Write-UpdaterLog "Echec mise a jour version.txt local : $_" 'ERROR'
    }

    # 5f - Cleanup backups
    Remove-OldBackups

    # ============================================================
    # ETAPE 6 : LANCER LE COLLECTOR
    # ============================================================
    Invoke-Collector -SharePath $SharePath
}
catch {
    Write-UpdaterLog "Exception non geree : $_" 'ERROR'
    Invoke-Collector -SharePath $SharePath
}
finally {
    if (Test-Path $LockFile) {
        Remove-Item $LockFile -Force -ErrorAction SilentlyContinue
        Write-UpdaterLog "Lock libere"
    }
    Write-UpdaterLog "=== Fin cycle updater ==="
}

