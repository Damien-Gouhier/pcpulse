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
      4. Si diff : verif signature + anti-downgrade, puis installation ATOMIQUE
         verifiee (copie tmp -> SHA -> renommage), sans backup ; maj version locale
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
    Version : 1.9
    Auteur  : Damien Gouhier
    Licence : MIT

    CHANGELOG :
    v1.9 : [SECURITE] Suite audit. (1) TOCTOU ferme : la signature est desormais
           verifiee sur le fichier TMP LOCAL (les octets reellement installes puis
           executes) dans Install-VerifiedUpdate, et non plus sur le fichier du share
           (re-lisible/permutable par un attaquant entre la verif et la copie).
           (2) Anti-downgrade de l'UPDATER : le self-update lit la version dans le
           fichier signe a installer (marqueur $UpdaterVersion) et refuse toute
           version < locale -> un ancien Updater signe ne peut plus etre re-injecte.
           (3) Anti-downgrade Collector durci : (a) une version serveur non parseable
           en [version] est REFUSEE ; (b) la version du Collector est lue dans le
           FICHIER SIGNE (marqueur $CollectorVersion) et non plus dans version.txt du
           share (non signe, falsifiable) -> plus de reinjection d'un vieux Collector
           signe via un version.txt gonfle. Fail-closed si marqueur illisible.
    v1.8 : [FIX] Nettoyage des backups deplace HORS du bloc update -> execute a
           CHAQUE cycle. Avant, Remove-OldBackups n'etait appele qu'en cas de MAJ
           du Collector ; un poste "Deja a jour" sortait avant et ne purgeait
           jamais ses backups legacy (pre-1.6) -> 0 poste nettoye observe sur le
           parc. Desormais auto-guerison systematique, meme deja a jour.
    v1.7 : [FIX CRITIQUE] Verification de signature RELACHEE sur le pin. L'ancien
           check exigeait Status='Valid' (confiance de chaine OS) ; le certificat
           DSI n'etant pas encore de confiance sur le parc, tout etait refuse
           (UnknownError) -> parc fige. Desormais : on exige la signature par le
           thumbprint EPINGLE + integrite (rejet si NotSigned/HashMismatch), mais
           on TOLERE UnknownError/NotTrusted (chaine non deployee). Le pin est
           l'ancre de confiance ; la confiance de chaine OS devient un bonus.
    v1.6 : [DURCISSEMENT] ZERO backup local. Le schema "backup -> copie -> restaure
           si KO" est remplace par une installation ATOMIQUE VERIFIEE
           (Install-VerifiedUpdate : copie vers tmp -> controle SHA -> renommage
           atomique). Le fichier vivant n'est jamais remplace par une copie
           corrompue -> plus besoin de backup. Remove-OldBackups supprime desormais
           le dossier backup (nettoie les backups laisses par les versions
           anterieures sur les postes). Moins de vieux code exploitable, plus robuste.
    v1.5 : [SECURITE] Authenticite de la chaine de mise a jour. Le SHA256 ne
           verifiait que l'integrite (calcule sur le meme share) : write sur
           release\ = RCE SYSTEM sur le parc. Ajout de la VERIFICATION DE
           SIGNATURE Authenticode + thumbprint EPINGLE ($PinnedThumbprints) :
           - appliquee d'ABORD au self-update de l'Updater (maillon critique),
             puis au Collector, AVANT toute copie/adoption ;
           - liste de thumbprints vide = verif DESACTIVEE (deploiement progressif
             sans casser le parc, on active en renseignant le thumbprint) ;
           - log du STATUT EXACT au refus (NotSigned/HashMismatch/NotTrusted...).
           [SECURITE] ANTI-DOWNGRADE : refus d'une version serveur < locale (empeche
           le rejeu d'un ancien Collector legitimement signe + son vieux version.txt).
    v1.4 : [FIX] Rotation de updater.log : cap de taille. Get-Content -Raw chargeait
           tout le fichier -> OutOfMemoryException sur un log enorme (observe 2 Go)
           -> rotation en echec a chaque cycle, log jamais borne. Au-dela de
           UpdaterLogMaxMB, lecture par -Tail (memoire bornee). Meme classe de bug
           que le Collector 2.3.2, cote Updater.
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
            (plus de dossier backup\ depuis v1.6 : installation atomique verifiee)
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
# v1.9 : version machine-lisible de CET Updater. Sert de plancher anti-downgrade au
# self-update (Install-VerifiedUpdate lit ce marqueur dans le fichier a installer).
# NE PAS renommer/reformater cette ligne : elle est parsee par regex.
$UpdaterVersion = [version]'1.9'
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

# v1.6 : plus de $BackupRetention -> ZERO backup local (installation atomique verifiee).

# v1.2 : nettoyage local des traces post-deploiement (housekeeping)
$DeployLeftoverDir       = 'C:\Temp\PCPulse-deploy'   # depose par Deploy-PCPulse, inutile apres install
$UpdaterLogRetentionDays = 7                          # borne updater.log (rotation par date)
$UpdaterLogMaxMB         = 20                          # v1.4 : au-dela, rotation via -Tail (evite l'OOM de Get-Content -Raw sur un log enorme)
$UpdaterLogTailLines     = 5000                        # v1.4 : lignes conservees quand le log depasse le cap
$TempBattReportPattern   = 'battreport_*.xml'         # rapports batterie orphelins du Collector

# ============================================================
# v1.5 : SIGNATURE DE CODE (authenticite de la chaine de mise a jour)
# ============================================================
# Le SHA256 ne verifie que l'INTEGRITE (le fichier n'a pas ete corrompu depuis le
# share), PAS l'AUTHENTICITE. Un attaquant avec write sur release\ peut deposer un
# Collector/Updater malveillant + son hash => RCE SYSTEM sur le parc. La signature
# Authenticode + thumbprint EPINGLE ferme ce trou : on n'execute QUE des fichiers
# signes par le certificat de la DSI.
#
# DEPLOIEMENT PROGRESSIF (evite le "chicken-and-egg") :
#   - Liste VIDE -> verification DESACTIVEE (log WARN, comportement inchange).
#     On peut ainsi deployer d'abord cet Updater "qui sait verifier" sans casser
#     le parc, PUIS signer les scripts + renseigner le thumbprint ci-dessous, PUIS
#     redeployer -> la barriere s'active.
#   - Renseigner le(s) thumbprint(s) SHA1 du cert de signature (SVR19RDS/ADCS).
#     Prevoir DEUX valeurs pendant une rotation de certificat (ancien + nouveau).
# Recuperer un thumbprint : (Get-AuthenticodeSignature .\01_Collector.ps1).SignerCertificate.Thumbprint
$PinnedThumbprints = @(
    # 'AAAA1111BBBB2222CCCC3333DDDD4444EEEE5555'   # <-- cert DSI (a renseigner)
)

# Defauts killswitch (si pas surcharge dans config.psd1)
# v1.5 : Enabled = $false par defaut (OPT-IN). Un killswitch actif par defaut avec
# une phrase/nom PUBLICS (visibles dans le repo) = risque de DoS parc si les ACL
# sautent. Le killswitch ne s'active QUE si config.psd1 declare explicitement
# KillSwitch.Enabled = $true (avec une phrase + un nom PERSONNALISES, non publics).
$KillSwitchDefaults = @{
    Enabled  = $false
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
    # v1.5 : config.psd1 en priorite dans release\ (lecture seule postes), repli racine.
    $configFile = Join-Path $SharePath 'release\config.psd1'
    if (-not (Test-Path $configFile)) { $configFile = Join-Path $SharePath 'config.psd1' }

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
    # v1.6 : PLUS AUCUN backup local (installation atomique verifiee, cf.
    # Install-VerifiedUpdate). On supprime le dossier backup s'il existe -> nettoie
    # les backups laisses par les versions anterieures sur les postes deja deployes
    # (moins de vieux code exploitable). Non-bloquant.
    if (Test-Path $BackupDir) {
        try {
            Remove-Item -LiteralPath $BackupDir -Recurse -Force -ErrorAction Stop
            Write-UpdaterLog "Backups locaux supprimes (dossier backup nettoye)"
        } catch {
            Write-UpdaterLog "Nettoyage backups : echec suppression dossier (non-bloquant) : $_" 'WARN'
        }
    }
}

function Test-PCPulseSignature {
    # v1.5 : verifie la signature Authenticode d'un fichier contre les thumbprints
    # EPINGLES ($PinnedThumbprints). Retourne $true si OK (ou si verification
    # desactivee = liste vide). Logge le STATUT EXACT au refus (attaque vs cert
    # expire vs CRL injoignable -> diagnostic en 30 s dans 6 mois).
    param([string]$Path)

    if (-not $PinnedThumbprints -or $PinnedThumbprints.Count -eq 0) {
        Write-UpdaterLog "Signature : verification DESACTIVEE (aucun thumbprint epingle) - a activer une fois les scripts signes" 'WARN'
        return $true
    }
    if (-not (Test-Path -LiteralPath $Path)) {
        Write-UpdaterLog "Signature : fichier absent ($Path)" 'ERROR'
        return $false
    }
    try {
        $sig = Get-AuthenticodeSignature -LiteralPath $Path -ErrorAction Stop
    } catch {
        Write-UpdaterLog "Signature : echec Get-AuthenticodeSignature sur $(Split-Path $Path -Leaf) : $_" 'ERROR'
        return $false
    }
    # v1.7 : le PIN est l'ANCRE DE CONFIANCE. On ne depend PAS de la confiance de
    # chaine de l'OS (Trusted Root/Publishers) ni de la CRL : sur le parc, tant que
    # le certificat de la DSI n'est pas deploye dans les magasins (AllSigned repousse),
    # Get-AuthenticodeSignature renvoie 'UnknownError'/'NotTrusted' pour un fichier
    # POURTANT correctement signe -> l'ancien check 'Status -eq Valid' refusait tout
    # et FIGEAIT le parc (plus aucune MAJ). On REJETTE seulement les etats qui prouvent
    # une absence de signature ou une ALTERATION, puis on exige le thumbprint epingle.
    # Securite : nul ne peut signer avec la cle privee de la DSI ; embarquer le cert
    # public + signer avec une autre cle => HashMismatch (rejete). Le pin + hash
    # intact suffit a garantir l'authenticite, sans confiance de chaine OS.
    if ($sig.Status -eq 'NotSigned' -or $sig.Status -eq 'HashMismatch' -or $sig.Status -eq 'NotSupportedFileFormat') {
        Write-UpdaterLog "Signature REFUSEE ($(Split-Path $Path -Leaf)) : statut='$($sig.Status)'" 'ERROR'
        return $false
    }
    $thumb = if ($sig.SignerCertificate) { $sig.SignerCertificate.Thumbprint } else { $null }
    if (-not $thumb -or ($PinnedThumbprints -notcontains $thumb)) {
        Write-UpdaterLog "Signature REFUSEE ($(Split-Path $Path -Leaf)) : thumbprint non epingle ('$thumb') statut='$($sig.Status)'" 'ERROR'
        return $false
    }
    if ($sig.Status -ne 'Valid') {
        # Signe par le bon cert, hash intact, mais chaine non de confiance sur CE poste
        # (cert DSI pas encore dans les magasins). Accepte sur le pin, trace en WARN.
        Write-UpdaterLog "Signature ACCEPTEE sur pin ($(Split-Path $Path -Leaf)) : thumbprint OK, statut='$($sig.Status)' (chaine non de confiance sur ce poste)" 'WARN'
    }
    return $true
}

function Install-VerifiedUpdate {
    # v1.6 : remplace le schema "backup -> copie -> re-verif -> restaure si KO" par
    # une INSTALLATION SURE SANS BACKUP : la signature est deja verifiee sur Source
    # en amont ; ici on copie vers un tmp, on controle que la copie est FIDELE (SHA
    # == source), puis renommage ATOMIQUE (local, pas de piege UNC). Le fichier
    # vivant n'est JAMAIS remplace par une copie partielle/corrompue -> aucun backup
    # necessaire (moins de vieux code exploitable sur les postes). Retourne $true si
    # Dest a bien ete mis a jour ; sinon Dest reste INCHANGE.
    param([string]$Source, [string]$Dest, [string]$ExpectedSha, [version]$MinVersionInFile)
    # v1.9 : le tmp DOIT finir en .ps1 -> Get-AuthenticodeSignature lit la signature
    # selon l'EXTENSION (bloc de signature PowerShell) ; un .tmp n'est pas reconnu
    # comme script signable -> SignerCertificate null -> refus a tort. (Bug attrape
    # au test V0044 : "thumbprint non epingle ('')" sur un .tmp pourtant signe.)
    $tmp = "$Dest.$PID.new.ps1"
    try {
        Copy-Item -LiteralPath $Source -Destination $tmp -Force -ErrorAction Stop

        # v1.9 : VERIF SIGNATURE SUR LE TMP (les octets reellement installes puis
        # executes), et NON sur le fichier du share. Ferme le TOCTOU : meme si le
        # share est permute apres coup, on n'installe QUE ce qu'on a verifie ici.
        if (-not (Test-PCPulseSignature -Path $tmp)) {
            Write-UpdaterLog "Install : signature du fichier copie invalide -> abandon, Dest inchange" 'ERROR'
            Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
            return $false
        }

        # SHA : simple detection de copie fidele (la securite = la signature ci-dessus).
        $tmpSha = (Get-FileHash -LiteralPath $tmp -Algorithm SHA256 -ErrorAction Stop).Hash
        if ($ExpectedSha -and ($tmpSha -ne $ExpectedSha)) {
            Write-UpdaterLog "Install : SHA de la copie != source ($tmpSha vs $ExpectedSha) -> abandon, Dest inchange" 'ERROR'
            Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
            return $false
        }

        # v1.9 : ANTI-DOWNGRADE sur le tmp verifie (self-update de l'Updater). On lit
        # la version dans le fichier signe qu'on s'apprete a installer -> un ancien
        # Updater (meme legitimement signe) ne peut plus etre re-injecte.
        if ($MinVersionInFile) {
            $fileVer = $null
            $vm = Select-String -LiteralPath $tmp -Pattern "\`$(?:Updater|Collector)Version\s*=\s*\[version\]'([0-9.]+)'" | Select-Object -First 1
            if ($vm) { $fileVer = $vm.Matches[0].Groups[1].Value -as [version] }
            if (-not $fileVer) {
                Write-UpdaterLog "Install : version illisible dans le fichier signe -> abandon (anti-downgrade)" 'ERROR'
                Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
                return $false
            }
            if ($fileVer -lt $MinVersionInFile) {
                Write-UpdaterLog "Install : DOWNGRADE Updater REFUSE ($fileVer < $MinVersionInFile) -> Dest inchange" 'ERROR'
                Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
                return $false
            }
        }

        Move-Item -LiteralPath $tmp -Destination $Dest -Force -ErrorAction Stop   # renommage atomique (disque local)
        return $true
    } catch {
        Write-UpdaterLog "Install : echec copie verifiee ($_) -> Dest inchange" 'ERROR'
        Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
        return $false
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

    # v1.9 : la verif de signature ET l'anti-downgrade se font desormais DANS
    # Install-VerifiedUpdate, SUR LE TMP (octets reellement installes) -> ferme le
    # TOCTOU (maillon critique : remplacer l'Updater par un binaire non verifie ferait
    # sauter toute la barriere). -MinVersionInFile = version de CET Updater : un ancien
    # Updater, meme legitimement signe, ne peut plus etre re-injecte via le share.
    if (Install-VerifiedUpdate -Source $UpdaterSrv -Dest $UpdaterLocal -ExpectedSha $srvHash -MinVersionInFile $UpdaterVersion) {
        Write-UpdaterLog "Self-update : Updater mis a jour, effectif au prochain cycle" 'SUCCESS'
    }
    # (echec deja logge par Install-VerifiedUpdate ; l'Updater actuel reste en place)
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
    #    v1.4 : CAP DE TAILLE. Avant, Get-Content -Raw chargeait TOUT le fichier en
    #    memoire -> OutOfMemoryException sur un log devenu enorme (observe : 2 Go) ->
    #    la rotation echouait a chaque cycle et le log ne redescendait JAMAIS. Au-dela
    #    du cap, on ne lit que la FIN via -Tail (memoire bornee, lecture depuis la fin
    #    du fichier), ce qui suffit a re-borner puis purger par date. Meme classe de
    #    correctif que le Collector 2.3.2.
    try {
        if (Test-Path -LiteralPath $UpdaterLog) {
            $cutoff = (Get-Date).AddDays(-$UpdaterLogRetentionDays).ToString('yyyy-MM-dd')
            $sizeMB = ((Get-Item -LiteralPath $UpdaterLog).Length) / 1MB
            if ($sizeMB -gt $UpdaterLogMaxMB) {
                # Trop gros pour -Raw (OOM) : on ne recupere que les dernieres lignes.
                $lignes = @(Get-Content -LiteralPath $UpdaterLog -Tail $UpdaterLogTailLines -ErrorAction Stop)
            } else {
                $lignes = (Get-Content -LiteralPath $UpdaterLog -Raw -ErrorAction Stop) -split "`r?`n"
            }
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

    # 5. Sweep des fichiers .new.ps1 d'installation atomique ORPHELINS (v1.8). Install-VerifiedUpdate
    #    ecrit "<fichier>.<PID>.new.ps1" puis le renomme ; il le nettoie sur succes ET
    #    sur echec, mais un process TUE entre les deux peut en laisser un. Ce cleanup
    #    tourne en DEBUT de cycle -> tout .new.ps1 present ici vient d'un cycle anterieur
    #    = orphelin -> on le degage. Borne le bruit dans le dossier runtime.
    try {
        Get-ChildItem -LiteralPath $LocalDir -Filter '*.new.ps1' -File -ErrorAction SilentlyContinue | ForEach-Object {
            Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue
            Write-UpdaterLog "Cleanup : tmp d'install orphelin supprime ($($_.Name))"
        }
    } catch { }

    # 6. Durcissement ACL du dossier runtime : SYSTEM + Administrateurs uniquement
    #    (empeche un compte standard de lire les scripts = chemin share, killswitch).
    Set-PCPulseAcl -Path $LocalDir
}

# ============================================================
# INIT : creation du layout local si necessaire
# ============================================================
if (-not (Test-Path $LocalDir))  { New-Item -Path $LocalDir  -ItemType Directory -Force | Out-Null }
# v1.6 : plus de creation de dossier backup (zero backup local). Remove-OldBackups
# supprime meme le dossier s'il subsiste d'une version anterieure.

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
    # ETAPE 3b : ANTI-DOWNGRADE (v1.5)
    # ============================================================
    # Un attaquant avec write sur release\ peut redeposer un ancien Collector
    # LEGITIMEMENT signe + son vieux version.txt : la signature serait valide et le
    # parc regresserait vers une version vulnerable. On refuse toute version serveur
    # STRICTEMENT INFERIEURE a la locale. (Si un numero n'est pas parseable en
    # [version], on n'applique pas la garde pour ne pas bloquer un cas legitime.)
    $lv = $localVer -as [version]
    $sv = $srvVer   -as [version]
    # v1.9 : si la version LOCALE est saine mais la version SERVEUR n'est pas parseable
    # en [version] (ex. un '0.0-x' injecte pour contourner la garde ci-dessous), c'est
    # suspect -> on REFUSE (on ne se laisse pas downgrader via un version.txt malforme).
    if ($lv -and -not $sv) {
        Write-UpdaterLog "Version serveur '$srvVer' non parseable -> UPDATE REFUSE (anti-downgrade) -> on garde $localVer" 'ERROR'
        Invoke-Collector -SharePath $SharePath
        return
    }
    if ($lv -and $sv -and ($sv -lt $lv)) {
        Write-UpdaterLog "DOWNGRADE REFUSE : version serveur $srvVer < locale $localVer -> on garde $localVer" 'ERROR'
        Invoke-Collector -SharePath $SharePath
        return
    }

    # ============================================================
    # NETTOYAGE BACKUPS (v1.8) : a CHAQUE cycle, hors du bloc update.
    # Avant, Remove-OldBackups etait dans l'ETAPE 5 -> jamais atteint quand le
    # poste etait "Deja a jour" -> les backups legacy (pre-1.6) ne partaient
    # jamais (0 poste nettoye observe). Ici on nettoie systematiquement ->
    # auto-guerison sur tout le parc, meme deja a jour.
    # ============================================================
    Remove-OldBackups

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

    # 5a-bis - (v1.9) La verif de signature se fait desormais DANS Install-VerifiedUpdate,
    # SUR LE TMP (les octets reellement installes/executes) et non sur le fichier du
    # share -> ferme le TOCTOU. Plus de controle de signature sur la source ici :
    # Install-VerifiedUpdate refuse et laisse le Collector local inchange si la copie
    # n'est pas signee par le cert epingle.

    # 5b - Installation SURE du nouveau Collector, SANS backup (v1.6) :
    # copie -> verif SHA (copie fidele) -> renommage atomique. Le Collector vivant
    # n'est jamais remplace par une copie partielle/corrompue. En cas d'echec, il
    # reste INCHANGE et on continue de tourner dessus.
    # v1.9 : -MinVersionInFile = version LOCALE (source fiable, sur l'endpoint) comme
    # plancher. Install-VerifiedUpdate lit la version dans le Collector SIGNE ($CollectorVersion)
    # et refuse un downgrade -> on ne fait plus confiance a version.txt du share (non signe,
    # qu'un attaquant pourrait gonfler pour reinjecter un vieux Collector signe). Fail-closed
    # si le marqueur est illisible (vieux Collector sans marqueur = refuse).
    if (-not (Install-VerifiedUpdate -Source $CollectorSrv -Dest $CollectorLocal -ExpectedSha $srvHash -MinVersionInFile $lv)) {
        Write-UpdaterLog "Update ECHEC : Collector non remplace -> on garde $localVer" 'ERROR'
        Invoke-Collector -SharePath $SharePath
        return
    }
    Write-UpdaterLog "Nouveau Collector installe (SHA verifie) : $srvHash" 'SUCCESS'

    # 5e - Mise a jour version.txt local
    try {
        Set-Content -Path $VersionLocal -Value $srvVer -Encoding UTF8 -ErrorAction Stop
        Write-UpdaterLog "version.txt local mis a jour : $srvVer" 'SUCCESS'
    } catch {
        Write-UpdaterLog "Echec mise a jour version.txt local : $_" 'ERROR'
    }

    # (nettoyage backups : deplace en debut de cycle en v1.8, cf. ci-dessus)

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

