#Requires -Version 5.1
<#
.SYNOPSIS
    PCPulse - Decommission-PC.ps1 v2.3 - marquage du cycle de vie des postes
.DESCRIPTION
    Outil interactif (double-clic via un .bat lanceur) pour gerer la mise en
    decommission d'un poste, cote DONNEE uniquement : il ne touche JAMAIS la
    machine cible. Il ecrit un registre separe :
        <RegistryPath>\decommissioning.json

    ARCHITECTURE (v2.0) : le registre vit dans un dossier ou les techs ont deja
    le droit d'ecrire avec leur COMPTE DE SESSION NORMAL (ex: un dossier metier
    sur un serveur non durci). L'outil n'accede PLUS au share durci PCPulse : pas
    de credentials, pas de compte a privileges, pas de RunAs. Le tech lance, coche, c'est tout.
    Le Dashboard (genere par un poste/serveur qui a le droit de lire ce dossier)
    lit ce registre pour afficher les badges "A decommissionner / Fait / En
    retard". Le registre ne pilote qu'un badge : aucun code, aucun deploiement.

    Workflow : Marquer (A faire) -> Cloturer (Fait). Repousser l'echeance et
    Retirer une marque sont possibles tant que l'entree existe (reserves aux
    comptes DecommissionAdmins).

    Anti-faute de frappe : si un motif (regex) est defini dans decom-config.psd1
    (cle PcNameRegex), le nom saisi est valide dessus ; un nom hors-format demande
    une confirmation. Sans motif configure : aucun controle (tout nom accepte).
    Le motif reste dans TA config (non publiee) : aucune nomenclature en dur ici.

    Attribution : l'operateur (qui marque / qui cloture) est capte via le compte
    qui lance l'outil ($env:USERNAME). AssignedTo (le tech CENSE s'en occuper)
    est choisi dans la liste DecommissionTechs.

    Multi-ecriture : plusieurs techs peuvent modifier le registre en meme temps
    -> verrou par fichier lock + retry, et ecriture ATOMIQUE SMB-safe (fichier tmp
    puis Move-Item -Force ; JAMAIS [IO.File]::Replace qui echoue sur un chemin UNC).
    Le verrou est tenu par un HANDLE OUVERT pendant toute l'action : un technicien
    qui reste longtemps sur un prompt ne peut pas se le faire voler (c'est l'OS qui
    refuse la suppression, pas une comparaison de dates), et un verrou reellement
    orphelin reste recuperable puisqu'un processus mort libere ses handles. Le lock
    porte en plus un GUID que Save-Registry revalide avant toute ecriture : si le
    verrou a disparu, l'ecriture est REFUSEE au lieu d'ecraser le registre.
.PARAMETER RegistryPath
    Dossier OU vivent le registre (decommissioning.json), le lock et le fichier
    de reglages (decom-config.psd1). C'est l'emplacement ou les techs ont le
    droit d'ecrire (ex: \\SERVER\SHARE\...\decom). Obligatoire.
.PARAMETER GraceDays
    Delai par defaut avant echeance (TargetDate = aujourd'hui + N jours). Defaut
    lu dans decom-config.psd1 (DecommissionGraceDays) sinon 30.
.NOTES
    Auteur     : Damien Gouhier
    Repository : https://github.com/Damien-Gouhier/pcpulse
    Licence    : MIT
    Version    : 2.3
    Runtime    : PowerShell 5.1+ (lance depuis le poste d'un tech, compte normal)
.CHANGELOG
    v2.3 : [FIX CRITIQUE] Vol de verrou -> perte de donnee SILENCIEUSE. Le lock
           etait cree puis referme aussitot, et la detection d'abandon reposait sur
           le seul LastWriteTime, jamais retouche apres l'acquisition : un tech
           reste plus de 5 min sur un prompt (Raison, Select-Tech) se faisait
           RECUPERER son verrou. Les deux operateurs appelaient alors Save-Registry,
           qui reecrit le tableau ENTIER -> le dernier ecrasait l'entree de l'autre,
           sans erreur ni trace. Deux barrieres independantes :
           (1) le handle du lock reste OUVERT en FileShare::Read pendant toute
               l'action -> la suppression concurrente echoue au niveau de l'OS
               (Read n'accorde pas FILE_SHARE_DELETE), quel que soit l'age du
               verrou. Un operateur lent n'est plus depossede.
           (2) le lock porte un GUID ; Save-Registry appelle Test-RegistryLockHeld
               en tout premier et THROW si le token a disparu ou change -> refus
               d'ecrire bruyant au lieu d'un ecrasement muet.
           La detection de verrou abandonne n'est plus temporelle : on SONDE le
           fichier en FileShare::None ; si l'ouverture reussit, plus aucun handle ne
           le tient donc le detenteur est mort (crash, Ctrl+C, fenetre fermee) et le
           residu est recupere IMMEDIATEMENT, sans attendre 5 min. Le controle d'age
           subsiste en repli pour le cas ou la sonde echoue sans detenteur vivant
           (ACL NTFS refusant l'ouverture du fichier cree par un autre technicien).
           Remove-RegistryLock ne supprime plus le lock que s'il porte notre token
           (sinon il volerait a son tour le verrou de celui qui l'a re-pris), et un
           echec de liberation est desormais SIGNALE au lieu d'etre avale.
           [FIX] Lock orphelin si l'ecriture du token echouait apres CreateNew
           (share plein, session SMB coupee) : le fichier restait avec un token
           inconnu de tous -> registre bloque, message trompeur "verrouille par un
           autre operateur". Le fichier est maintenant supprime dans ce cas.
           La recuperation rapide ne s'applique qu'aux locks PORTANT UN TOKEN (3 champs) :
           un lock ecrit par un v2.2 (2 champs, handle referme) retombe sur le controle
           d'age, sinon on lui volerait son verrou pendant la fenetre de deploiement
           mixte -- en reintroduisant le bug qu'on corrige.
           [FIX] -LiteralPath sur tous les acces fichier (lock, registre, config) :
           un RegistryPath contenant des crochets ("...\Parc [ancien]\decom") etait
           interprete comme un MOTIF par -Path -> Test-Path $false sur un dossier
           existant, et l'outil devenait inutilisable de facon inexplicable.
           [SECURITE] Fail-open sur les actions reservees. "config ABSENTE" (aucune
           restriction, voulu) et "config PRESENTE ET ILLISIBLE" (panne) tombaient sur
           le MEME test ($Admins vide) : une faute de frappe dans decom-config.psd1 --
           ou le bug de crochets ci-dessus -- ouvrait Repousser [3] et Retirer [4] a
           tout le monde. Une config presente et illisible bloque desormais [3]/[4].
           Et toute cle INCONNUE du .psd1 est nommee en console : une cle mal
           orthographiee (DecomissionAdmins, un M) parse sans erreur et etait ignoree
           en silence, avec le meme fail-open a la clef.
           [FIX] Chemin relatif : -RegistryPath est resolu en ABSOLU au demarrage. Le
           script melange cmdlets PowerShell et API .NET (pour la litteralite et le
           controle des share modes), et les deux ne resolvent PAS les chemins relatifs
           pareil (emplacement PowerShell vs [Environment]::CurrentDirectory) -> on
           aurait lu un dossier et ecrit dans un autre.
           [FIX] Move-Item -Destination echappe : la destination n'a pas de variante
           litterale et etait traitee comme un motif -> sur un chemin a crochets,
           CHAQUE sauvegarde echouait (le .tmp restait sur le share).
    v2.2 : [FIX] Lecture du registre forcee en UTF-8 (Get-Registry). Save-Registry
           ecrit en UTF-8 mais Get-Content SANS -Encoding relit en ANSI (defaut
           PS 5.1) -> le "c cedille" de "Francois" se re-encodait a chaque cycle,
           mojibake compose qui explosait (observe : 1 nom -> 281 751 caracteres,
           fichier a 1,7 Mo). Force -Encoding UTF8, coherent avec l'ecriture.
           [FIX] Enumeration explicite apres ConvertFrom-Json : en PS 5.1 un tableau
           JSON est emis NON enumere -> "@($raw | ConvertFrom-Json)" donnait Count=1
           avec toutes les entrees empilees dans [0] (symptome "System.Object[]" a la
           cloture). On enumere en liste plate (aplatit aussi tout sous-tableau).
    v2.1 : Validation de nom pilotee par config (cle PcNameRegex de decom-config.psd1)
           au lieu d'un motif code en dur -> aucune nomenclature d'environnement
           dans le code publie. Sans motif : aucun controle.
    v2.0 : Re-architecture. Le registre vit dans un dossier ecrivable au compte
           NORMAL du tech (hors share durci) : suppression totale de la
           machinerie de credentials (WNet / Get-Credential / compte a privileges / -SharePath
           / -AskCredential). Reglages lus depuis decom-config.psd1 a cote du
           registre (DecommissionTechs / DecommissionAdmins / DecommissionGraceDays,
           rien de sensible). Anti-typo par motif (regex) au lieu d'une lecture
           des JSON du parc. Sonde d'ecriture ciblee sur RegistryPath.
           Cloture/Repousse : reconstruction de l'entree (muter en place un objet
           ConvertFrom-Json pouvait echouer "propriete introuvable"). Cloture a
           un seul poste = confirmation directe o/N (pas de numero a saisir).
    v1.0 : Version initiale (registre sur le share durci via un compte a privileges).
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$RegistryPath,

    [int]$GraceDays = 0   # 0 = prendre la valeur config (DecommissionGraceDays) ou 30
)

$ErrorActionPreference = 'Stop'

# Normalisation du chemin en ABSOLU avant tout Join-Path.
# Ce script melange volontairement cmdlets PowerShell (Test-Path, Remove-Item) et API
# .NET (File::Open, File::WriteAllText, Directory::CreateDirectory) -- les .NET pour
# leur litteralite et pour le controle des share modes du verrou. Or les deux ne
# resolvent PAS les chemins relatifs de la meme facon : les cmdlets partent de
# l'emplacement PowerShell courant, les API .NET de [Environment]::CurrentDirectory,
# qui n'est pas synchronise avec lui. Un -RegistryPath relatif (rien ne l'interdit)
# ferait donc lire un dossier et ecrire dans un autre : la sonde d'ecriture passerait,
# Get-Registry lirait un dossier vide, et Save-Registry ecrirait le .tmp ailleurs avant
# d'echouer sur le Move-Item. GetUnresolvedProviderPathFromPSPath resout sans exiger
# que la cible existe (le dossier peut etre a creer) et sans developper les jokers.
$RegistryPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($RegistryPath)

$RegistryFile = Join-Path $RegistryPath 'decommissioning.json'
$LockFile     = Join-Path $RegistryPath 'decommissioning.lock'
$ConfigFile   = Join-Path $RegistryPath 'decom-config.psd1'
$Utf8NoBom    = [System.Text.UTF8Encoding]::new($false)
$Operator     = $env:USERNAME

# ============================================================
# SONDE D'ECRITURE : echec ici = probleme de DROITS, pas de verrou.
# ============================================================
try {
    # -LiteralPath partout sur les chemins de config : un RegistryPath contenant des
    # crochets (ex. "\\SRV\SHARE\Parc [ancien]\decom") est interprete comme un MOTIF
    # par -Path -> Test-Path renvoie $false sur un dossier qui existe. Meme classe de
    # bug que l'encodage : silencieux, et il ne se manifeste que chez le client qui a
    # un chemin exotique.
    # [System.IO.Directory]::CreateDirectory et non New-Item : New-Item n'a pas de
    # -LiteralPath, son -Path resout les jokers. Les API .NET sont litterales.
    if (-not (Test-Path -LiteralPath $RegistryPath)) { [void][System.IO.Directory]::CreateDirectory($RegistryPath) }
    $probe = Join-Path $RegistryPath (".pcpulse.write.test.$PID.tmp")
    [System.IO.File]::WriteAllText($probe, 'ok', $Utf8NoBom)
    Remove-Item -LiteralPath $probe -Force -ErrorAction SilentlyContinue
} catch {
    Write-Host "[X] Ecriture IMPOSSIBLE dans le dossier du registre (compte : $env:USERNAME)." -ForegroundColor Red
    Write-Host "    Ce n'est PAS un verrou : ce compte n'a pas les droits d'ecriture sur" -ForegroundColor Yellow
    Write-Host "    $RegistryPath" -ForegroundColor Yellow
    Write-Host "    Verifiez le chemin et vos droits, puis relancez." -ForegroundColor Yellow
    exit 1
}

# ============================================================
# REGLAGES : liste des techs + admins + delai (decom-config.psd1, non sensible)
# ============================================================
$Techs       = @()
$Admins      = @()
$graceCfg    = 30
$PcNameRegex = ''
$configBroken = $false
if (Test-Path -LiteralPath $ConfigFile) {
    try {
        # -LiteralPath : Import-PowerShellDataFile resout les jokers sur -Path. Avec un
        # RegistryPath a crochets, l'import echouait -> $Admins vide -> voir plus bas,
        # c'etait un fail-OPEN sur les actions reservees.
        $cfg = Import-PowerShellDataFile -LiteralPath $ConfigFile -ErrorAction Stop
        if ($cfg.DecommissionTechs)      { $Techs       = @($cfg.DecommissionTechs) }
        if ($cfg.DecommissionAdmins)     { $Admins      = @($cfg.DecommissionAdmins) }
        if ($cfg.DecommissionGraceDays)  { $graceCfg    = [int]$cfg.DecommissionGraceDays }
        if ($cfg.PcNameRegex)            { $PcNameRegex = [string]$cfg.PcNameRegex }
        # Meme esprit que le coverage-check du Dashboard : une cle mal orthographiee
        # (DecomissionAdmins, un M) parse SANS erreur et est ignoree en silence. Sur
        # DecommissionAdmins la consequence est un fail-open (liste vide = tout le monde
        # admin), donc on nomme toute cle inconnue au lieu de la laisser passer.
        $knownCfgKeys = @('DecommissionTechs','DecommissionAdmins','DecommissionGraceDays','PcNameRegex')
        $unknownCfg   = @($cfg.Keys | Where-Object { $_ -notin $knownCfgKeys })
        if ($unknownCfg.Count -gt 0) {
            Write-Host "[!] decom-config.psd1 : cle(s) INCONNUE(S) ignoree(s) -> $($unknownCfg -join ', ')" -ForegroundColor Yellow
            Write-Host "    Faute de frappe ? Cles attendues : $($knownCfgKeys -join ', ')" -ForegroundColor Yellow
        }
    } catch {
        $configBroken = $true
        Write-Host "[!] decom-config.psd1 PRESENT mais illisible ($_)." -ForegroundColor Red
        Write-Host "    Liste techs vide (saisie libre) et actions [3]/[4] BLOQUEES par precaution." -ForegroundColor Yellow
    }
}
if ($GraceDays -le 0) { $GraceDays = $graceCfg }

# Repousser [3] / Retirer [4] reserves aux comptes listes dans DecommissionAdmins
# (comparaison sur le compte de SESSION $env:USERNAME). Liste vide/absente =>
# aucune restriction (tout le monde a acces) : c'est VOULU, un deploiement sans
# config n'est pas un deploiement restreint.
# v2.3 : mais "config ABSENTE" et "config PRESENTE ET ILLISIBLE" ne sont pas la meme
# chose. Le second cas est une PANNE, et il tombait sur le meme test ($Admins vide)
# -> une simple faute de frappe dans le .psd1, ou un chemin a crochets, ouvrait [3]
# et [4] a tout le monde. Fail-open sur un controle d'acces : on ferme.
$IsAdmin = (-not $configBroken) -and (($Admins.Count -eq 0) -or ($Admins -contains $env:USERNAME))

# ============================================================
# VERROU (multi-ecriture) : lock-file + retry + recuperation auto
# ============================================================
# PIEGE CORRIGE EN v2.3 -- "VOL DE VERROU" ET PERTE DE DONNEE SILENCIEUSE.
#
# Jusqu'en v2.2, le lock etait un fichier cree puis IMMEDIATEMENT REFERME, et la
# detection de verrou abandonne reposait sur le seul LastWriteTime, qui n'etait
# jamais retouche apres l'acquisition. Consequence : un technicien reste plus de
# StaleMinutes sur un prompt (Read-NonEmpty "Raison", Select-Tech -- un appel
# telephonique suffit) voyait son verrou juge "abandonne" et RECUPERE par un
# second operateur. Les deux detenaient alors leur propre copie de $entries en
# memoire, les deux appelaient Save-Registry -- qui reecrit le tableau ENTIER --
# et le dernier a ecrire ECRASAIT l'entree de l'autre. Sans aucune trace : aucune
# erreur, aucun log, l'entree disparaissait simplement du registre.
#
# Le correctif tient sur deux barrieres independantes :
#
#  1. EMPECHER LE VOL (barriere OS, pas barriere horloge). Le handle du lock reste
#     OUVERT pendant toute la duree de l'action, en FileShare::Read : les autres
#     peuvent LIRE le fichier (afficher qui tient le verrou, verifier le token)
#     mais ni l'ecrire ni le SUPPRIMER -- la suppression exige que tous les
#     handles ouverts aient concede FILE_SHARE_DELETE, ce qui n'est pas le cas.
#     Un Remove-Item concurrent ECHOUE donc, quel que soit l'age du verrou. C'est
#     Windows (et le serveur SMB, qui honore les share modes) qui arbitre, plus
#     une comparaison de dates : un operateur lent n'est plus jamais depossede.
#     Corollaire : un verrou REELLEMENT abandonne reste recuperable, car un
#     processus mort libere ses handles -- la recuperation cesse d'etre une
#     heuristique temporelle pour devenir un fait ("plus personne ne le tient").
#     Le controle d'age est CONSERVE, mais il ne peut plus faire de degat : si un
#     vivant tient le handle, la suppression echoue de toute facon. Il ne sert
#     plus qu'au cas residuel du handle SMB orphelin (crash machine brutal /
#     coupure reseau : le serveur garde le handle jusqu'au timeout de session).
#
#  2. EMPECHER LA CORRUPTION SI LE VERROU DISPARAIT QUAND MEME (defense en
#     profondeur : suppression manuelle du .lock, share remonte, timeout SMB).
#     Le lock porte un GUID, memorise dans $Script:LockToken. Save-Registry
#     appelle Test-RegistryLockHeld en TOUT PREMIER et throw si le token a
#     disparu ou change. Le catch du menu affiche alors "Action interrompue
#     (registre inchange)" : on refuse d'ecrire au lieu d'ecraser en silence.
#     Un refus bruyant vaut infiniment mieux qu'une perte de donnee muette.
$Script:LockToken  = $null   # GUID de NOTRE verrou (null = on ne tient rien)
$Script:LockHandle = $null   # handle maintenu ouvert = barriere anti-suppression

function Get-RegistryLock {
    param([int]$TimeoutSec = 20, [int]$StaleMinutes = 5)
    # Re-entrance : si on tient DEJA le verrou, ne pas retenter un CreateNew -- il
    # echouerait sur notre propre fichier et la sonde d'orphelin buterait sur notre
    # propre handle (les regles de partage sont par handle, pas par processus), pour
    # finir sur un faux "verrouille par un autre operateur" apres 20 s d'attente.
    # Le flux actuel ne peut pas y tomber (le finally du menu libere a chaque tour) ;
    # c'est un garde-fou pour un futur appel imbrique.
    if ($Script:LockHandle) { return $true }
    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    while ((Get-Date) -lt $deadline) {
        try {
            # CreateNew echoue si le fichier existe deja = verrou tenu par un autre.
            # FileShare::Read (et NON None) : les autres processus doivent pouvoir
            # RELIRE le token pour verifier qu'ils ne tiennent pas un verrou perime.
            # Read n'inclut pas Delete -> le fichier reste indestructible tant que
            # ce handle vit.
            $token = [guid]::NewGuid().ToString('N')
            $fs = [System.IO.File]::Open($LockFile, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::Write, [System.IO.FileShare]::Read)
            try {
                $w = New-Object System.IO.StreamWriter($fs, $Utf8NoBom)
                $w.WriteLine("$Operator | $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | $token")
                # INVARIANT A NE PAS RETIRER : ce Flush() est le seul chose qui pousse
                # le token sur le disque. On ne dispose PAS $w/$fs (le handle ouvert EST
                # le verrou) et StreamWriter n'a pas de finaliseur en .NET Framework :
                # sans ce Flush, le fichier de lock resterait VIDE, Test-RegistryLockHeld
                # echouerait toujours et Save-Registry refuserait toute ecriture.
                $w.Flush()
            } catch {
                # L'ecriture du token a echoue APRES que CreateNew ait cree le fichier
                # (share plein, session SMB coupee entre l'Open et le WriteLine). Sans
                # ce Remove-Item on laisserait un lock ORPHELIN portant un token que
                # personne ne connait : registre bloque pour tout le monde, avec le
                # message trompeur "verrouille par un autre operateur".
                # Dispose SOUS try/catch : dans le scenario meme qu'on traite (share
                # plein), Dispose re-flushe le buffer du FileStream et releve la MEME
                # IOException -> sans ce garde, l'exception sortirait avant le
                # Remove-Item et le lock orphelin survivrait quand meme.
                try { $fs.Dispose() } catch {}
                Remove-Item -LiteralPath $LockFile -Force -ErrorAction SilentlyContinue
                throw
            }
            $Script:LockHandle = $fs
            $Script:LockToken  = $token
            return $true
        } catch {
            # Verrou tenu par un autre, OU residu orphelin. On distingue les deux par
            # une SONDE DETERMINISTE plutot que par une heuristique de date.
            #
            # Ouvrir le lock en exclusivite TOTALE (FileShare::None) : si ca REUSSIT,
            # c'est qu'aucun autre handle ne le tient -> le detenteur est mort (crash,
            # Ctrl+C, fenetre fermee) et le fichier n'est qu'un residu, quel que soit
            # son age. Si un detenteur est VIVANT, son acces Write n'est pas autorise
            # par le FileShare::None qu'on demande -> la sonde echoue et on boucle.
            # C'est l'OS qui repond "est-ce que quelqu'un le tient", au lieu de deduire
            # l'abandon du temps ecoule : plus de fenetre de 5 min a attendre apres un
            # Ctrl+C, et plus jamais de vol d'un verrou legitimement tenu.
            $orphan = $false
            try {
                $lockProbe = [System.IO.File]::Open($LockFile, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
                try {
                    # COMPATIBILITE v2.2 -> v2.3 (fenetre de deploiement mixte).
                    # La sonde ne prouve "personne ne le tient" que face a un pair qui
                    # TIENT un handle, c'est-a-dire un pair v2.3. Un lock ecrit par un
                    # Decommission-PC.ps1 v2.2 (fenetre restee ouverte pendant la mise a
                    # jour de l'outil) etait cree puis REFERME aussitot : la sonde
                    # reussirait et on lui volerait son verrou -- en reintroduisant
                    # exactement la perte de donnee qu'on corrige.
                    # On distingue les deux par la FORME du lock : v2.3 ecrit trois
                    # champs ("operateur | date | token"), v2.2 en ecrivait deux. Sans
                    # token, on ne se croit pas autorise a recuperer vite : on retombe
                    # sur le controle d'age, qui est le contrat de la v2.2.
                    $sr        = New-Object System.IO.StreamReader($lockProbe, [System.Text.Encoding]::UTF8)
                    $lockLine  = $sr.ReadLine()
                    $orphan    = (($lockLine -split '\|').Count -ge 3)
                } finally { $lockProbe.Dispose() }
            } catch {}

            if ($orphan) {
                try {
                    # -ErrorAction Stop : un echec silencieux ici relancerait la boucle
                    # sans temporisation (sonde OK -> Remove KO -> continue -> ...), soit
                    # une boucle serree a 100 % de CPU pendant tout le TimeoutSec, en
                    # affichant des milliers de "verrou recupere" mensongers. Cas reel :
                    # ACL du dossier partage accordant l'ecriture mais pas la SUPPRESSION
                    # d'un fichier cree par un autre technicien.
                    Remove-Item -LiteralPath $LockFile -Force -ErrorAction Stop
                    Write-Host "  [i] Verrou orphelin (plus aucun detenteur) -> recupere." -ForegroundColor Yellow
                    continue    # retente CreateNew immediatement
                } catch {
                    Write-Host "  [!] Residu de verrou impossible a supprimer ($_)." -ForegroundColor Red
                    Start-Sleep -Milliseconds 400
                    continue
                }
            }

            # REPLI sur l'age : la sonde peut echouer pour une raison qui n'est PAS un
            # detenteur vivant -- typiquement une ACL NTFS qui nous refuse l'ouverture
            # en ecriture du fichier cree par un autre technicien (les techs ecrivent
            # avec leur compte de session dans un dossier partage). Sans ce repli, un
            # tel residu bloquerait le registre indefiniment. Ce chemin est sans risque :
            # si un detenteur est vivant, le Remove-Item echoue de toute facon (son
            # FileShare::Read n'accorde pas DELETE).
            try {
                $lockAge = (Get-Date) - (Get-Item -LiteralPath $LockFile -ErrorAction Stop).LastWriteTime
                if ($lockAge.TotalMinutes -ge $StaleMinutes) {
                    Remove-Item -LiteralPath $LockFile -Force -ErrorAction Stop
                    Write-Host ("  [i] Verrou residuel ({0:N0} min) -> recupere." -f $lockAge.TotalMinutes) -ForegroundColor Yellow
                }
            } catch {}
            Start-Sleep -Milliseconds 400
        }
    }
    return $false
}

function Test-RegistryLockHeld {
    # Retourne $true seulement si le lock existe ENCORE et porte NOTRE token.
    if (-not $Script:LockToken) { return $false }
    try {
        if (-not (Test-Path -LiteralPath $LockFile)) { return $false }
        # PIEGE PARTAGE DE FICHIER : on ne peut PAS relire avec Get-Content ici.
        # Notre propre handle de verrou est ouvert en FileAccess::Write ; un
        # lecteur classique (Get-Content, File::ReadAllText, StreamReader) demande
        # implicitement FileShare::Read, ce qui INTERDIT l'acces Write deja detenu
        # -> violation de partage, meme depuis notre processus (les regles de
        # partage s'appliquent par HANDLE, pas par processus). Il faut donc ouvrir
        # explicitement en concedant FileShare::ReadWrite.
        $rs = [System.IO.File]::Open($LockFile, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        try {
            $sr   = New-Object System.IO.StreamReader($rs, [System.Text.Encoding]::UTF8)
            $line = $sr.ReadLine()
        } finally { $rs.Dispose() }
        return ($line -like "*$($Script:LockToken)*")
    } catch {
        # Lock illisible : on ne peut pas PROUVER qu'on le tient -> on considere
        # que non. Cote Save-Registry cela produit un refus d'ecrire, jamais un
        # ecrasement : c'est le sens de la faute qu'on veut.
        return $false
    }
}

function Remove-RegistryLock {
    # Liberer le handle AVANT de supprimer, sinon notre propre barriere nous bloque.
    try { if ($Script:LockHandle) { $Script:LockHandle.Dispose() } } catch {}
    $Script:LockHandle = $null
    # Ne supprimer que si le fichier est bien LE NOTRE : si notre verrou a ete
    # perdu et re-pris par un autre operateur entre-temps, supprimer ici lui
    # volerait son verrou a son tour (propagation du bug qu'on corrige).
    try {
        if ((Test-Path -LiteralPath $LockFile) -and (Test-RegistryLockHeld)) {
            # Pas de -ErrorAction SilentlyContinue muet : un echec de suppression
            # laisse notre lock sur le disque sans detenteur, ce qui bloque les autres
            # techniciens jusqu'au chemin de recuperation. Autant le dire tout de suite.
            Remove-Item -LiteralPath $LockFile -Force -ErrorAction Stop
        }
    } catch {
        Write-Host "  [!] Le verrou n'a pas pu etre libere ($_). Il sera recupere automatiquement au prochain lancement." -ForegroundColor Yellow
    }
    $Script:LockToken = $null
}

# ============================================================
# LECTURE / ECRITURE du registre
# ============================================================
function Get-Registry {
    if (-not (Test-Path -LiteralPath $RegistryFile)) { return @() }
    try {
        # PIEGE ENCODAGE PS 5.1 : Save-Registry ecrit en UTF-8 (sans BOM), mais
        # Get-Content SANS -Encoding relit en ANSI par defaut -> le "c cedille" de
        # "Francois" (0xC3 0xA7 en UTF-8) devient "Ã§", re-sauve en UTF-8, relu ANSI...
        # le mojibake se COMPOSE a chaque cycle et explose (observe : 1 nom -> 281 751
        # caracteres, fichier a 1,7 Mo). On force la lecture UTF-8 : coherent avec
        # l'ecriture, plus aucune recomposition.
        $raw = Get-Content -LiteralPath $RegistryFile -Raw -Encoding UTF8 -ErrorAction Stop
        if ([string]::IsNullOrWhiteSpace($raw)) { return @() }
        # PIEGE PowerShell 5.1 : "$raw | ConvertFrom-Json" emet un tableau JSON
        # comme UN SEUL objet non-enumere. Un "@(...)" direct donnerait alors
        # $entries.Count = 1 avec les N entrees empilees dans $entries[0] (tableau
        # imbrique) -> Where-Object teste .Statut sur le tableau et fait tout passer
        # d'un bloc (symptome "System.Object[]"). On enumere explicitement pour
        # obtenir une liste PLATE. Le foreach aplatit aussi tout sous-tableau
        # eventuel (auto-reparation : la prochaine sauvegarde reecrit propre).
        $data = $raw | ConvertFrom-Json
        $out  = @()
        foreach ($item in $data) { $out += $item }
        return $out
    } catch {
        Write-Host "[!] Registre illisible ($_). Repartir d'une liste vide serait DANGEREUX -> abandon." -ForegroundColor Red
        throw
    }
}
function Save-Registry {
    param([object[]]$Entries)
    # GARDE-FOU EN TOUT PREMIER (v2.3) : ne JAMAIS reecrire le registre sans la
    # preuve qu'on tient toujours le verrou. Save-Registry reecrit le tableau
    # ENTIER a partir d'un $entries lu en memoire il y a peut-etre plusieurs
    # minutes : ecrire sans verrou, c'est ecraser en silence tout ce qu'un autre
    # operateur a fait entre-temps. Le throw remonte au catch du menu, qui affiche
    # "Action interrompue (registre inchange)" -- message exact, l'action est
    # perdue mais AUCUNE donnee d'autrui ne l'est.
    if (-not (Test-RegistryLockHeld)) {
        throw "verrou perdu (lock absent, expire ou repris par un autre operateur) - ecriture refusee pour ne pas ecraser le travail d'un autre. Relance l'action."
    }
    # Ecriture ATOMIQUE SMB-safe : tmp unique -> Move-Item -Force (rename cote
    # serveur). Pas de [IO.File]::Replace (throw sur UNC).
    # Piege PS 5.1 : ConvertTo-Json d'un tableau a 1 element emet un OBJET, pas un
    # tableau -> on force les crochets a la main pour 0 et 1 element.
    $arr = @($Entries)
    if     ($arr.Count -eq 0) { $json = '[]' }
    elseif ($arr.Count -eq 1) { $json = '[' + ($arr[0] | ConvertTo-Json -Depth 6) + ']' }
    else                      { $json = $arr | ConvertTo-Json -Depth 6 }
    $tmp = "$RegistryFile.$PID.$([guid]::NewGuid().ToString('N')).tmp"
    [System.IO.File]::WriteAllText($tmp, $json, $Utf8NoBom)
    # -Destination n'a pas de variante litterale et est traite comme un MOTIF par le
    # provider FileSystem : sur un chemin a crochets, la destination ne se resout pas
    # et CHAQUE sauvegarde echoue (le .tmp reste sur le share, action perdue). On
    # echappe les metacaracteres de joker ; sans effet sur un chemin normal.
    Move-Item -LiteralPath $tmp -Force `
        -Destination ([System.Management.Automation.WildcardPattern]::Escape($RegistryFile))
}

# ============================================================
# HELPERS UI
# ============================================================
function Test-PcNameFormat {
    # Valide le nom sur le motif PcNameRegex de decom-config.psd1 (optionnel).
    # Sans motif configure -> aucun controle (tout nom non vide accepte).
    param([string]$Pc)
    if (-not $PcNameRegex) { return $true }
    return ($Pc -match $PcNameRegex)
}

function Read-NonEmpty {
    param([string]$Prompt)
    do { $v = (Read-Host $Prompt).Trim() } while (-not $v)
    return $v
}

function Select-Tech {
    if ($Techs.Count -eq 0) {
        return (Read-NonEmpty "  Assigner a (nom du tech)")
    }
    Write-Host "  Assigner a :"
    for ($i = 0; $i -lt $Techs.Count; $i++) { Write-Host ("    [{0}] {1}" -f ($i+1), $Techs[$i]) }
    do {
        $sel = (Read-Host "  Numero du tech").Trim()
        $ok  = ($sel -match '^\d+$') -and ([int]$sel -ge 1) -and ([int]$sel -le $Techs.Count)
        if (-not $ok) { Write-Host "  Choix invalide." -ForegroundColor Yellow }
    } while (-not $ok)
    return $Techs[[int]$sel - 1]
}

function New-HistoryEntry {
    param([string]$Action, [string]$Detail = '')
    [PSCustomObject]@{
        Date   = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        Action = $Action
        By     = $Operator
        Detail = $Detail
    }
}

# ============================================================
# ACTIONS
# ============================================================
function Invoke-Mark {
    $entries = Get-Registry
    $pc = (Read-NonEmpty "  Nom du PC a decommissionner").ToUpper()

    $existing = $entries | Where-Object { $_.PC -eq $pc }
    if ($existing) {
        Write-Host "  '$pc' est deja dans le registre (statut: $($existing.Statut)). Utilise Repousser/Cloturer." -ForegroundColor Yellow
        return
    }

    if (-not (Test-PcNameFormat $pc)) {
        Write-Host "  (!) '$pc' ne correspond pas au format attendu ($PcNameRegex)." -ForegroundColor Yellow
        if ((Read-Host "  Confirmer ce nom malgre tout ? (o/N)").Trim().ToLower() -ne 'o') {
            Write-Host "  Annule." -ForegroundColor Yellow
            return
        }
    }

    if ((Read-Host "  Confirmer le marquage de '$pc' ? (o/N)").Trim().ToLower() -ne 'o') {
        Write-Host "  Annule." -ForegroundColor Yellow
        return
    }

    $reason   = Read-NonEmpty "  Raison (ex: remplace, HS, vol, fin de vie)"
    $assigned = Select-Tech
    $target   = (Get-Date).AddDays($GraceDays).ToString('yyyy-MM-dd')

    $entry = [PSCustomObject]@{
        PC         = $pc
        Statut     = 'A faire'
        Reason     = $reason
        AssignedTo = $assigned
        MarkedAt   = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        Operator   = $Operator
        TargetDate = $target
        DoneAt     = $null
        DoneBy     = $null
        History    = @( New-HistoryEntry -Action 'Marque' -Detail "assigne a $assigned, echeance $target" )
    }
    Save-Registry -Entries (@($entries) + $entry)
    Write-Host "  [OK] '$pc' marque 'A faire', assigne a $assigned, echeance $target." -ForegroundColor Green
}

function Invoke-Close {
    $entries = @(Get-Registry)
    $todo = @($entries | Where-Object { $_.Statut -eq 'A faire' })
    if ($todo.Count -eq 0) { Write-Host "  Rien a cloturer." -ForegroundColor Yellow; return }

    if ($todo.Count -eq 1) {
        # Un seul poste a cloturer -> confirmation directe (pas de numero a saisir).
        $pc = $todo[0].PC
        Write-Host ("  A cloturer : {0}  (assigne {1}, {2})" -f $todo[0].PC, $todo[0].AssignedTo, $todo[0].Reason)
        if ((Read-Host "  Marquer '$pc' FAIT ? (o/N)").Trim().ToLower() -ne 'o') { Write-Host "  Annule." -ForegroundColor Yellow; return }
    } else {
        Write-Host "  A faire :"
        for ($i = 0; $i -lt $todo.Count; $i++) {
            Write-Host ("    [{0}] {1,-14} assigne: {2,-10} echeance: {3}  ({4})" -f ($i+1), $todo[$i].PC, $todo[$i].AssignedTo, $todo[$i].TargetDate, $todo[$i].Reason)
        }
        $sel = (Read-Host "  Numero a marquer FAIT (Entree = annuler)").Trim()
        if (-not ($sel -match '^\d+$') -or [int]$sel -lt 1 -or [int]$sel -gt $todo.Count) { Write-Host "  Annule." -ForegroundColor Yellow; return }
        $pc = $todo[[int]$sel - 1].PC
    }

    # On RECONSTRUIT l'entree (muter en place un objet ConvertFrom-Json peut
    # echouer "propriete introuvable") : objet neuf = toujours modifiable.
    $new = @()
    foreach ($e in $entries) {
        if ($e.PC -eq $pc) {
            $e = [PSCustomObject]@{
                PC         = $e.PC
                Statut     = 'Fait'
                Reason     = $e.Reason
                AssignedTo = $e.AssignedTo
                MarkedAt   = $e.MarkedAt
                Operator   = $e.Operator
                TargetDate = $e.TargetDate
                DoneAt     = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
                DoneBy     = $Operator
                History    = @($e.History) + (New-HistoryEntry -Action 'Cloture' -Detail 'decommission realisee')
            }
        }
        $new += $e
    }
    Save-Registry -Entries $new
    Write-Host "  [OK] '$pc' marque FAIT par $Operator." -ForegroundColor Green
}

function Invoke-Postpone {
    $entries = @(Get-Registry)
    $todo = @($entries | Where-Object { $_.Statut -eq 'A faire' })
    if ($todo.Count -eq 0) { Write-Host "  Rien a repousser." -ForegroundColor Yellow; return }
    for ($i = 0; $i -lt $todo.Count; $i++) {
        Write-Host ("    [{0}] {1,-14} echeance actuelle: {2}" -f ($i+1), $todo[$i].PC, $todo[$i].TargetDate)
    }
    $sel = (Read-Host "  Numero a repousser").Trim()
    if (-not ($sel -match '^\d+$') -or [int]$sel -lt 1 -or [int]$sel -gt $todo.Count) { Write-Host "  Annule." -ForegroundColor Yellow; return }
    $pc = $todo[[int]$sel - 1].PC
    $days = (Read-Host "  Repousser de combien de jours ? (defaut $GraceDays)").Trim()
    if (-not ($days -match '^\d+$')) { $days = $GraceDays }
    $newTarget = (Get-Date).AddDays([int]$days).ToString('yyyy-MM-dd')
    # Reconstruction (cf. Invoke-Close) plutot que mutation en place.
    $new = @()
    foreach ($e in $entries) {
        if ($e.PC -eq $pc) {
            $old = $e.TargetDate
            $e = [PSCustomObject]@{
                PC         = $e.PC
                Statut     = $e.Statut
                Reason     = $e.Reason
                AssignedTo = $e.AssignedTo
                MarkedAt   = $e.MarkedAt
                Operator   = $e.Operator
                TargetDate = $newTarget
                DoneAt     = $e.DoneAt
                DoneBy     = $e.DoneBy
                History    = @($e.History) + (New-HistoryEntry -Action 'Repousse' -Detail "$old -> $newTarget")
            }
        }
        $new += $e
    }
    Save-Registry -Entries $new
    Write-Host "  [OK] '$pc' repousse au $newTarget." -ForegroundColor Green
}

function Invoke-Remove {
    $entries = @(Get-Registry)
    if ($entries.Count -eq 0) { Write-Host "  Registre vide." -ForegroundColor Yellow; return }
    for ($i = 0; $i -lt $entries.Count; $i++) {
        Write-Host ("    [{0}] {1,-14} [{2}] assigne: {3}" -f ($i+1), $entries[$i].PC, $entries[$i].Statut, $entries[$i].AssignedTo)
    }
    $sel = (Read-Host "  Numero de la marque a RETIRER").Trim()
    if (-not ($sel -match '^\d+$') -or [int]$sel -lt 1 -or [int]$sel -gt $entries.Count) { Write-Host "  Annule." -ForegroundColor Yellow; return }
    $pc = $entries[[int]$sel - 1].PC
    if ((Read-Host "  Retirer definitivement la marque de '$pc' ? (o/N)").Trim().ToLower() -ne 'o') { Write-Host "  Annule." -ForegroundColor Yellow; return }
    $kept = @($entries | Where-Object { $_.PC -ne $pc })
    Save-Registry -Entries $kept
    Write-Host "  [OK] Marque de '$pc' retiree." -ForegroundColor Green
}

function Invoke-List {
    $entries = @(Get-Registry)
    if ($entries.Count -eq 0) { Write-Host "  Registre vide." -ForegroundColor Cyan; return }
    Write-Host ("  {0,-14} {1,-8} {2,-10} {3,-12} {4}" -f 'PC','STATUT','ASSIGNE','ECHEANCE','RAISON') -ForegroundColor Cyan
    foreach ($e in ($entries | Sort-Object Statut, TargetDate)) {
        $late = ($e.Statut -eq 'A faire' -and $e.TargetDate -lt (Get-Date -Format 'yyyy-MM-dd'))
        $col  = if ($e.Statut -eq 'Fait') { 'Green' } elseif ($late) { 'Red' } else { 'White' }
        $flag = if ($late) { ' (EN RETARD)' } else { '' }
        Write-Host ("  {0,-14} {1,-8} {2,-10} {3,-12} {4}{5}" -f $e.PC, $e.Statut, $e.AssignedTo, $e.TargetDate, $e.Reason, $flag) -ForegroundColor $col
    }
}

# ============================================================
# BOUCLE PRINCIPALE
# ============================================================
Write-Host ""
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host " PCPulse - Decommission" -ForegroundColor Cyan
Write-Host " Registre  : $RegistryFile" -ForegroundColor Cyan
Write-Host " Operateur : $Operator" -ForegroundColor Cyan
Write-Host "==================================================" -ForegroundColor Cyan

do {
    Write-Host ""
    Write-Host " [1] Marquer une machine   [2] Cloturer (Fait)"
    if ($IsAdmin) {
        Write-Host " [3] Repousser l'echeance  [4] Retirer une marque"
    }
    Write-Host " [5] Lister                [Q] Quitter"
    $choice = (Read-Host " Choix").Trim().ToUpper()

    if ($choice -eq 'Q') { break }
    if ($choice -notin @('1','2','3','4','5')) { Write-Host " Choix invalide." -ForegroundColor Yellow; continue }
    if (($choice -in @('3','4')) -and -not $IsAdmin) {
        Write-Host " [!] Action reservee a l'administrateur PCPulse." -ForegroundColor Yellow; continue
    }

    # Lister ne modifie rien -> pas besoin de verrou
    if ($choice -eq '5') {
        try { Invoke-List } catch { Write-Host " [!] Lecture du registre impossible : $_" -ForegroundColor Red }
        continue
    }

    if (-not (Get-RegistryLock)) {
        Write-Host " [!] Registre verrouille par un autre operateur, reessaie dans un instant." -ForegroundColor Red
        continue
    }
    try {
        switch ($choice) {
            '1' { Invoke-Mark }
            '2' { Invoke-Close }
            '3' { Invoke-Postpone }
            '4' { Invoke-Remove }
        }
    } catch {
        Write-Host " [!] Action interrompue (registre inchange) : $_" -ForegroundColor Red
    } finally {
        Remove-RegistryLock
    }
} while ($true)

Write-Host ""
Write-Host "A bientot." -ForegroundColor Cyan
