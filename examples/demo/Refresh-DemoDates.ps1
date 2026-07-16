#Requires -Version 5.1
<#
.SYNOPSIS
    Refresh-DemoDates.ps1 - Re-ancre les dates des JSON de demo sur "maintenant".

.DESCRIPTION
    Les JSON de demo contiennent des dates figees. Or le Dashboard calcule le
    statut "hors-ligne" par rapport a l'heure courante (SeuilOfflineJours, defaut
    1 jour). Des dates figees finissent donc par faire apparaitre TOUS les PC de
    demo hors-ligne, ce qui casse le rendu de la demo.

    Ce script decale, dans chaque .json du dossier, toutes les dates d'ACTIVITE
    par un delta calcule fichier par fichier, de sorte que :
      - un PC "online"  ait CollectedAt ramene a maintenant (moins 2h)
      - un PC "offline" ait CollectedAt ramene a maintenant moins N jours
    Les ecarts internes (boots, crashs, throttling, WHEA...) sont preserves : on
    applique le meme delta a toutes les dates du fichier.

    Les dates MATERIEL ne sont PAS touchees (DriverDate, YearOfManufacture,
    WeekOfManufacture) : ce ne sont pas des dates d'activite.

    La modification se fait par SUBSTITUTION TEXTE (regex sur les paires
    "cle": "date"), pas via ConvertTo-Json. Raison : en PS 5.1, ConvertTo-Json
    casse les tableaux a un seul element (les transforme en objet) et reformate
    tout le fichier. La substitution texte ne touche QUE les dates ciblees.

.PARAMETER Path
    Dossier contenant les .json de demo. Defaut : le dossier du script.

.EXAMPLE
    pwsh .\examples\demo\Refresh-DemoDates.ps1

.EXAMPLE
    pwsh .\Refresh-DemoDates.ps1 -Path "C:\chemin\vers\demos"

.NOTES
    Version : 1.0
    Auteur  : Damien Gouhier
    Licence : MIT

    Comportement idempotent : on peut le relancer autant de fois qu'on veut,
    il re-ancre toujours sur l'heure courante.
#>

param(
    [string]$Path
)

# ------------------------------------------------------------
# Resolution du dossier cible
# ------------------------------------------------------------
if ([string]::IsNullOrWhiteSpace($Path)) {
    if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
        $Path = $PSScriptRoot
    } else {
        $Path = (Get-Location).Path
    }
}

# ------------------------------------------------------------
# Configuration
# ------------------------------------------------------------
# PC a considerer comme HORS-LIGNE -> nb de jours dans le passe pour CollectedAt.
# Tout PC absent de cette table est considere ONLINE (collecte il y a 2h).
$OfflineDays = @{
    'OFFLINE-005' = 13
}
$OnlineOffsetHours = 2

$ci = [System.Globalization.CultureInfo]::InvariantCulture

# Cles dont la valeur est une date d'ACTIVITE a decaler.
# (Volontairement absentes : DriverDate, YearOfManufacture, WeekOfManufacture.)
$dateKeys = @(
    'CollectedAt','LastRealColdBoot','LastBoot','DerniereActivite',
    'Timestamp','DateBoot','FirstSeen','LastSeen','BootStartTime','Date','Day'
)
$keyAlt  = ($dateKeys | ForEach-Object { [regex]::Escape($_) }) -join '|'
$pattern = '"(' + $keyAlt + ')"\s*:\s*"([^"]+)"'

# ------------------------------------------------------------
# Traitement
# ------------------------------------------------------------
$files = Get-ChildItem -Path $Path -Filter '*.json' -File -ErrorAction Stop
if (-not $files) {
    Write-Warning "Aucun fichier .json trouve dans : $Path"
    return
}

$now   = Get-Date
$count = 0

Write-Host ""
Write-Host "Rafraichissement des dates de demo dans : $Path" -ForegroundColor Cyan
Write-Host ("-" * 64)

foreach ($f in $files) {
    $text = [System.IO.File]::ReadAllText($f.FullName)

    # Extraire le CollectedAt d'origine (sert a calculer le delta du fichier)
    $mCA = [regex]::Match($text, '"CollectedAt"\s*:\s*"([^"]+)"')
    if (-not $mCA.Success) {
        Write-Host ("  [skip] {0} (pas de CollectedAt)" -f $f.Name) -ForegroundColor DarkGray
        continue
    }
    $caStr = $mCA.Groups[1].Value

    # Nom du PC (pour savoir s'il est online/offline)
    $mPC = [regex]::Match($text, '"PC"\s*:\s*"([^"]+)"')
    $pc  = if ($mPC.Success) { $mPC.Groups[1].Value } else { $f.BaseName }

    try {
        $caOrig = [datetime]::ParseExact($caStr, 'yyyy-MM-dd HH:mm:ss', $ci)
    } catch {
        Write-Warning ("  CollectedAt illisible ('{0}') dans {1}, ignore" -f $caStr, $f.Name)
        continue
    }

    if ($OfflineDays.ContainsKey($pc)) {
        $target   = $now.AddDays(-1 * $OfflineDays[$pc])
        $stateLbl = "offline (-{0}j)" -f $OfflineDays[$pc]
    } else {
        $target   = $now.AddHours(-1 * $OnlineOffsetHours)
        $stateLbl = "online"
    }

    $delta = $target - $caOrig
    # v2.3.1 (#16) : les dates SANS heure (ex CPUThrottling.Day) sont decalees du
    # MEME delta que les FirstSeen/LastSeen du meme objet, puis tronquees a la date.
    # Avant : delta ARRONDI a l'entier de jours -> pouvait desaligner Day d'un jour
    # calendaire par rapport a ses FirstSeen/LastSeen (delta different).

    # Evaluateur regex : detecte le format, parse, decale, reformate a l'identique.
    $eval = {
        param($m)
        $key = $m.Groups[1].Value
        $val = $m.Groups[2].Value

        if ($val -match '^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$') {
            $dt  = [datetime]::ParseExact($val, 'yyyy-MM-dd HH:mm:ss', $ci)
            $new = ($dt + $delta).ToString('yyyy-MM-dd HH:mm:ss')
        }
        elseif ($val -match '^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d+Z$') {
            # v2.3.1 (#16) : parse en UTC (AssumeUniversal) pour que le suffixe Z
            # reste semantiquement correct apres decalage.
            $dt  = [datetime]::ParseExact($val.Substring(0, 19), "yyyy-MM-dd'T'HH:mm:ss", $ci, `
                       [System.Globalization.DateTimeStyles]::AssumeUniversal -bor [System.Globalization.DateTimeStyles]::AdjustToUniversal)
            $new = ($dt + $delta).ToString("yyyy-MM-dd'T'HH:mm:ss") + '.0000000Z'
        }
        elseif ($val -match '^\d{4}-\d{2}-\d{2}$') {
            $dt  = [datetime]::ParseExact($val, 'yyyy-MM-dd', $ci)
            $new = ($dt + $delta).ToString('yyyy-MM-dd')
        }
        else {
            return $m.Value   # format non reconnu : on ne touche pas
        }

        return '"' + $key + '": "' + $new + '"'
    }.GetNewClosure()

    $newText = [regex]::Replace($text, $pattern, [System.Text.RegularExpressions.MatchEvaluator]$eval)

    # Ecriture UTF-8 SANS BOM (JSON standard)
    [System.IO.File]::WriteAllText($f.FullName, $newText, [System.Text.UTF8Encoding]::new($false))

    Write-Host ("  {0,-16} {1,-16} CollectedAt -> {2}" -f $pc, $stateLbl, $target.ToString('yyyy-MM-dd HH:mm:ss')) -ForegroundColor Green
    $count++
}

Write-Host ("-" * 64)
Write-Host ("[OK] {0} fichier(s) rafraichi(s)." -f $count) -ForegroundColor Cyan
Write-Host ""
