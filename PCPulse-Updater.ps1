<#
.SYNOPSIS
    PCPulse-Updater.ps1 - Auto-update du Collector depuis le share serveur
.DESCRIPTION
    Wrapper execute par la tache planifiee a la place du Collector direct.
    Verifie a chaque run s'il existe une nouvelle version sur le share,
    l'applique le cas echeant, puis lance le Collector local.

    Workflow :
      1. Acquiert un verrou (evite double execution concurrente)
      2. Compare version locale vs version serveur
      3. Si diff : backup, copie, verif SHA256, met a jour version locale
      4. Lance le Collector local avec le SharePath fourni
      5. Libere le verrou

    En cas d'echec a n'importe quelle etape, le Collector deja present
    localement continue de tourner. AUCUN rollback auto (evite les boucles
    infernales). Les erreurs sont logees dans updater.log pour diagnostic.

.PARAMETER SharePath
    Chemin UNC du share serveur (ex: \\SRV-PCPULSE\PCPulse$).
    Transmis au Collector apres l'update.

.EXAMPLE
    # Utilise par la tache planifiee :
    PCPulse-Updater.ps1 -SharePath "\\SRV-PCPULSE\PCPulse$"

.NOTES
    Version : 1.0
    Auteur  : Damien Gouhier
    Licence : MIT

    Layout attendu cote serveur :
      \\SRV-PCPULSE\PCPulse$\
        release\
          01_Collector.ps1   (derniere version officielle)
          version.txt        (ex: "1.1")

    Layout genere cote client :
      C:\ProgramData\PCPulse\
        01_Collector.ps1     (version courante)
        PCPulse-Updater.ps1  (ce script)
        version.txt          (version active)
        backup\              (5 derniers backups, rolling)
        updater.log          (log dedie)
        .update.lock         (lock file temporaire)
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
$VersionLocal   = Join-Path $LocalDir 'version.txt'
$LockFile       = Join-Path $LocalDir '.update.lock'
$BackupDir      = Join-Path $LocalDir 'backup'
$UpdaterLog     = Join-Path $LocalDir 'updater.log'

$ReleaseDir     = Join-Path $SharePath 'release'
$CollectorSrv   = Join-Path $ReleaseDir '01_Collector.ps1'
$VersionSrv     = Join-Path $ReleaseDir 'version.txt'

$BackupRetention = 5   # nombre de versions a conserver

# ============================================================
# FONCTIONS UTILITAIRES
# ============================================================
function Write-UpdaterLog {
    param([string]$Message, [string]$Level = 'INFO')
    try {
        $line = "{0} | {1} | {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
        Add-Content -Path $UpdaterLog -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue
    } catch {
        # Silent : on ne veut pas qu'un probleme de log casse l'updater
    }
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
    $backups = Get-ChildItem -Path $BackupDir -Filter '01_Collector_v*.ps1' -ErrorAction SilentlyContinue |
               Sort-Object LastWriteTime -Descending
    if ($backups.Count -gt $BackupRetention) {
        $toDelete = $backups | Select-Object -Skip $BackupRetention
        foreach ($b in $toDelete) {
            Remove-Item $b.FullName -Force -ErrorAction SilentlyContinue
            Write-UpdaterLog "Backup ancien supprime : $($b.Name)"
        }
    }
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
# ETAPE 1 : VERROU ANTI-COLLISION
# ============================================================
if (Test-Path $LockFile) {
    # Lock existe : soit une autre instance tourne, soit un crash precedent
    # a laisse un lock orphelin. On le considere orphelin apres 30 min.
    $lockAge = (Get-Date) - (Get-Item $LockFile).LastWriteTime
    if ($lockAge.TotalMinutes -lt 30) {
        Write-UpdaterLog "Lock actif (age $([int]$lockAge.TotalMinutes) min), abandon" 'WARN'
        exit 0
    }
    Write-UpdaterLog "Lock orphelin detecte (age $([int]$lockAge.TotalMinutes) min), suppression"
    Remove-Item $LockFile -Force -ErrorAction SilentlyContinue
}

try {
    # Creer le lock file avec timestamp courant
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

    # Verifier que le nouveau Collector existe sur le share
    if (-not (Test-Path $CollectorSrv)) {
        Write-UpdaterLog "version.txt annonce $srvVer mais 01_Collector.ps1 absent du share" 'ERROR'
        Invoke-Collector -SharePath $SharePath
        return
    }

    # 5a - Calcul du SHA256 source AVANT copie (source de verite)
    try {
        $srvHash = (Get-FileHash -Path $CollectorSrv -Algorithm SHA256 -ErrorAction Stop).Hash
        Write-UpdaterLog "SHA256 serveur : $srvHash"
    } catch {
        Write-UpdaterLog "Echec lecture SHA256 serveur : $_" 'ERROR'
        Invoke-Collector -SharePath $SharePath
        return
    }

    # 5b - Backup de la version actuelle (si elle existe)
    $backupFile = $null
    if (Test-Path $CollectorLocal) {
        $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
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

    # 5c - Copie du nouveau Collector
    try {
        Copy-Item -Path $CollectorSrv -Destination $CollectorLocal -Force -ErrorAction Stop
        Write-UpdaterLog "Copie du nouveau Collector OK"
    } catch {
        Write-UpdaterLog "Echec copie du nouveau Collector : $_" 'ERROR'
        # Tentative de restauration du backup
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
            # Restaurer le backup
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

    # 5e - Mise a jour du version.txt local (seulement si tout le reste a reussi)
    try {
        Set-Content -Path $VersionLocal -Value $srvVer -Encoding UTF8 -ErrorAction Stop
        Write-UpdaterLog "version.txt local mis a jour : $srvVer" 'SUCCESS'
    } catch {
        Write-UpdaterLog "Echec mise a jour version.txt local : $_" 'ERROR'
        # Le Collector est deja copie, on continue quand meme
    }

    # 5f - Nettoyage des vieux backups (rolling 5)
    Remove-OldBackups

    # ============================================================
    # ETAPE 6 : LANCER LE COLLECTOR (nouvelle version ou ancienne en cas d'echec)
    # ============================================================
    Invoke-Collector -SharePath $SharePath
}
catch {
    Write-UpdaterLog "Exception non geree : $_" 'ERROR'
    # Tenter quand meme de lancer le Collector local
    Invoke-Collector -SharePath $SharePath
}
finally {
    # ============================================================
    # ETAPE 7 : LIBERATION DU VERROU
    # ============================================================
    if (Test-Path $LockFile) {
        Remove-Item $LockFile -Force -ErrorAction SilentlyContinue
        Write-UpdaterLog "Lock libere"
    }
    Write-UpdaterLog "=== Fin cycle updater ==="
}
