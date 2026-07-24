#Requires -Version 5.1
<#
.SYNOPSIS
    PCPulse - Decommission-PC.ps1 v2.1 - marquage du cycle de vie des postes
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
    -> verrou par fichier lock + retry (recuperation auto d'un lock abandonne),
    et ecriture ATOMIQUE SMB-safe (fichier tmp puis Move-Item -Force ; JAMAIS
    [IO.File]::Replace qui echoue sur un chemin UNC).
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
    Version    : 2.2
    Runtime    : PowerShell 5.1+ (lance depuis le poste d'un tech, compte normal)
.CHANGELOG
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

$RegistryFile = Join-Path $RegistryPath 'decommissioning.json'
$LockFile     = Join-Path $RegistryPath 'decommissioning.lock'
$ConfigFile   = Join-Path $RegistryPath 'decom-config.psd1'
$Utf8NoBom    = [System.Text.UTF8Encoding]::new($false)
$Operator     = $env:USERNAME

# ============================================================
# SONDE D'ECRITURE : echec ici = probleme de DROITS, pas de verrou.
# ============================================================
try {
    if (-not (Test-Path $RegistryPath)) { New-Item -ItemType Directory -Path $RegistryPath -Force | Out-Null }
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
if (Test-Path $ConfigFile) {
    try {
        $cfg = Import-PowerShellDataFile -Path $ConfigFile -ErrorAction Stop
        if ($cfg.DecommissionTechs)      { $Techs       = @($cfg.DecommissionTechs) }
        if ($cfg.DecommissionAdmins)     { $Admins      = @($cfg.DecommissionAdmins) }
        if ($cfg.DecommissionGraceDays)  { $graceCfg    = [int]$cfg.DecommissionGraceDays }
        if ($cfg.PcNameRegex)            { $PcNameRegex = [string]$cfg.PcNameRegex }
    } catch {
        Write-Host "[!] decom-config.psd1 illisible ($_). Liste techs vide, saisie libre." -ForegroundColor Yellow
    }
}
if ($GraceDays -le 0) { $GraceDays = $graceCfg }

# Repousser [3] / Retirer [4] reserves aux comptes listes dans DecommissionAdmins
# (comparaison sur le compte de SESSION $env:USERNAME). Liste vide/absente =>
# aucune restriction (tout le monde a acces).
$IsAdmin = ($Admins.Count -eq 0) -or ($Admins -contains $env:USERNAME)

# ============================================================
# VERROU (multi-ecriture) : lock-file + retry + recuperation auto
# ============================================================
function Get-RegistryLock {
    param([int]$TimeoutSec = 20, [int]$StaleMinutes = 5)
    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    while ((Get-Date) -lt $deadline) {
        try {
            # CreateNew echoue si le fichier existe deja = verrou tenu par un autre
            $fs = [System.IO.File]::Open($LockFile, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
            $w = New-Object System.IO.StreamWriter($fs)
            $w.WriteLine("$Operator | $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
            $w.Flush(); $w.Dispose(); $fs.Dispose()
            return $true
        } catch {
            # Verrou tenu. S'il est plus vieux que StaleMinutes, c'est un verrou
            # ABANDONNE (fenetre restee ouverte / run interrompu) -> on le recupere.
            try {
                $lockAge = (Get-Date) - (Get-Item $LockFile -ErrorAction Stop).LastWriteTime
                if ($lockAge.TotalMinutes -ge $StaleMinutes) {
                    Write-Host ("  [i] Verrou abandonne ({0:N0} min) -> recupere." -f $lockAge.TotalMinutes) -ForegroundColor Yellow
                    Remove-Item $LockFile -Force -ErrorAction SilentlyContinue
                    continue
                }
            } catch {}
            Start-Sleep -Milliseconds 400
        }
    }
    return $false
}
function Remove-RegistryLock {
    try { if (Test-Path $LockFile) { Remove-Item $LockFile -Force -ErrorAction SilentlyContinue } } catch {}
}

# ============================================================
# LECTURE / ECRITURE du registre
# ============================================================
function Get-Registry {
    if (-not (Test-Path $RegistryFile)) { return @() }
    try {
        # PIEGE ENCODAGE PS 5.1 : Save-Registry ecrit en UTF-8 (sans BOM), mais
        # Get-Content SANS -Encoding relit en ANSI par defaut -> le "c cedille" de
        # "Francois" (0xC3 0xA7 en UTF-8) devient "Ã§", re-sauve en UTF-8, relu ANSI...
        # le mojibake se COMPOSE a chaque cycle et explose (observe : 1 nom -> 281 751
        # caracteres, fichier a 1,7 Mo). On force la lecture UTF-8 : coherent avec
        # l'ecriture, plus aucune recomposition.
        $raw = Get-Content -Path $RegistryFile -Raw -Encoding UTF8 -ErrorAction Stop
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
    Move-Item -LiteralPath $tmp -Destination $RegistryFile -Force
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
