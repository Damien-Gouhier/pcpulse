#Requires -Version 7.0
<#
.SYNOPSIS
    PCPulse Dashboard v1.8
.DESCRIPTION
    Lit les JSON produits par le Collector sur tous les PC du parc et
    genere un tableau de bord HTML autonome avec KPIs, filtres, tri,
    drill-down par PC, dark/light mode.
.NOTES
    Auteur       : Damien Gouhier
    Repository   : https://github.com/Damien-Gouhier/pcpulse
    Licence      : MIT
    Version      : 1.8
    Runtime      : PowerShell 7+ (pwsh.exe)
.CHANGELOG
    v1.8 : Exposition des donnees hardware v1.8 du Collector
           - Panel Materiel enrichi de 3 nouveaux blocs :
             * RAM : installee / slots / capacite max + detail barrettes
                     (Slot, Type DDR, Vitesse, Fabricant).
             * GPU : nom + version + date driver pour chaque adapteur.
             * CPU Throttling : agregation quotidienne des events 35/55
                    Kernel-Processor-Power, avec badge alerte >=3 jours.
           - Backward-compat totale : si JSON v1.7 ou anterieur, les
             nouveaux blocs affichent discretement "Donnees non
             disponibles (Collector < v1.8)".
    v1.7 : Debruitage des Top Crashers (Signal vs Bruit)
           - Scoring automatique de chaque processus crasheur :
             score = crashs_par_PC / sqrt(PC_impactes).
             Penalise la dispersion (un crash reparti sur 10 PC est
             noye, un crash concentre sur 1 PC ressort).
           - 3 sections dans le panel Top Crashers parc global :
             * Signaux locaux (score >=3) : a investiguer prioritairement
             * Problemes repartis (2<=score<3) : bug applicatif probable
             * Bruit ambient (score <2) : replie par defaut
           - Blacklist Hard : processus totalement ignores (ex: le
             fameux microsoftsearchbing.exe qui pourrissait deja NXT).
           - Blacklist Soft : processus forces en section Bruit meme
             si leur score les aurait remontes (dellosd, shellexpe,
             gamebar, ASUS utilitaires, Dell techhub).
           - Drill-down par PC : meme logique locale, bruiteurs grises
             avec badge "bruit" au lieu d'etre cache.
    v1.6 : Intelligence de diagnostic (Wave 1)
           - Bandeau verdict global par PC (4 niveaux : Sain / A surveiller
             / Incident probable / Critique) calcule selon seuils sur
             crashs, WHEA, batterie, CPU age, SMART, bursts I/O.
           - Section "Signaux croises" (5 patterns temporels a fenetre
             10 min) : burst I/O -> hard crash, WHEA PCIe -> crash, etc.
           - Nom CPU ajoute dans le panel Materiel.
           - Libelles detailles des Hard crashes (Coupure alim / Reprise
             veille / User bouton power) exploitant CrashCause v1.6.
           - Schema accepte elargi a ('1.4','1.5','1.6').
    v1.5 : Rendu des clusters Event 51 (Disk slow / I/O timeout)
           - Propagation des champs Count/IsBurst/FirstSeen/LastSeen
             produits par le Collector v1.5 jusqu'au rendu JS.
           - Affichage enrichi : "xN events en 1s" + badge BURST si
             Count >= 50 (signal materiel fort).
           - Backward-compat JSON v1.4- : affichage inchange.
           - Schema accepte elargi a ('1.4','1.5').
    Voir CHANGELOG.md du repo pour l'historique des versions.
.EXAMPLE
    .\02_Dashboard.ps1
    # Utilise le SharePath par defaut (C:\PCPulse)
.EXAMPLE
    .\02_Dashboard.ps1 -SharePath "D:\Data\pcpulse"
    # Utilise un chemin custom
.EXAMPLE
    .\02_Dashboard.ps1 -FiltrePC "LAPTOP-*"
    # Ne charge que les PC dont le nom commence par LAPTOP-
#>

param(
    [string]$SharePath = 'C:\PCPulse',
    [string]$FiltrePC  = '*'
)

# ============================================================
# CONSTANTES
# ============================================================
$SchemaCompatible = @('1.4','1.5','1.6','1.7','1.8')   # historique jusqu'à v1.7 + v1.8 (RAM/GPU/CPU Throttling)
$ConfigFile       = Join-Path $SharePath 'config.psd1'
$OutputHTML       = Join-Path $env:TEMP ("PCPulse-Dashboard-" + (Get-Date -Format 'yyyyMMdd-HHmm') + '.html')

# Valeurs par defaut si config.psd1 est absent / corrompu
$DefaultConfig = @{
    SeuilBootLong        = 2
    SeuilCrashRecent     = 7
    SeuilDiskAlert       = 10
    SeuilDiskWarning     = 25
    SeuilOfflineJours    = 1
    DashboardTitle       = 'PCPulse'
    DashboardSubtitle    = 'Supervision du parc'
    CsvRanges            = 'ip-ranges.csv'
    MaskHealthyByDefault = $false
    ScoreWeights         = @{
        BSOD          = 5
        WHEA          = 4
        Crash         = 3
        Thermal       = 3
        GPU_TDR       = 2
        DiskAlert     = 2
        BootLong      = 1
        Offline       = 5
        # v5.3 : EDR down = critique, batterie usee = mineur
        SentinelDown  = 5
        Battery       = 1
        # v5.4 : BootPerf et SMART
        BootPerfSlow  = 1
        DiskHealth    = 3
    }
}

# ============================================================
# FONCTIONS UTILITAIRES
# ============================================================

function Import-MonitorConfig {
    param([string]$Path, [hashtable]$Defaults)
    if (-not (Test-Path $Path)) {
        Write-Host "[!] Config absente ($Path), utilisation des defaults" -ForegroundColor Yellow
        return $Defaults.Clone()
    }
    try {
        $loaded = Import-PowerShellDataFile -Path $Path -ErrorAction Stop
        $merged = @{}
        foreach ($key in $Defaults.Keys) {
            $merged[$key] = if ($loaded.ContainsKey($key)) { $loaded[$key] } else { $Defaults[$key] }
        }
        Write-Host "[+] Config chargee : $Path" -ForegroundColor Green
        return $merged
    } catch {
        Write-Host "[!] Erreur lecture config ($_), utilisation des defaults" -ForegroundColor Yellow
        return $Defaults.Clone()
    }
}

function ConvertTo-HtmlSafe {
    param($Text)
    if ($null -eq $Text) { return '' }
    return [System.Net.WebUtility]::HtmlEncode([string]$Text)
}

function Build-CIDRRange {
    param([string]$CIDR, [string]$Site)
    try {
        $parts   = $CIDR.Split('/')
        $prefix  = [int]$parts[1]
        $netRaw  = [System.Net.IPAddress]::Parse($parts[0]).GetAddressBytes()

        $maskBytes = [byte[]]::new(4)
        for ($i = 0; $i -lt 4; $i++) {
            $bitsInByte = [math]::Min(8, [math]::Max(0, $prefix - ($i * 8)))
            if     ($bitsInByte -eq 0) { $maskBytes[$i] = 0 }
            elseif ($bitsInByte -eq 8) { $maskBytes[$i] = 255 }
            else                       { $maskBytes[$i] = [byte](256 - [math]::Pow(2, 8 - $bitsInByte)) }
        }

        $netBytes = [byte[]]::new(4)
        for ($i = 0; $i -lt 4; $i++) {
            $netBytes[$i] = $netRaw[$i] -band $maskBytes[$i]
        }

        return [PSCustomObject]@{
            CIDR      = $CIDR
            Site      = $Site
            NetBytes  = $netBytes
            MaskBytes = $maskBytes
        }
    } catch {
        Write-Warning "Range ignoree ($CIDR) : $_"
        return $null
    }
}

function Get-SiteFromIP {
    param([string]$IP, [array]$Ranges)
    if (-not $IP -or $IP -eq 'N/A' -or $IP -like '169.254.*' -or $IP -like '127.*') {
        return 'Inconnu'
    }
    try {
        $ipBytes = [System.Net.IPAddress]::Parse($IP).GetAddressBytes()
    } catch {
        return 'Inconnu'
    }
    foreach ($r in $Ranges) {
        $match = $true
        for ($i = 0; $i -lt 4; $i++) {
            if (($ipBytes[$i] -band $r.MaskBytes[$i]) -ne $r.NetBytes[$i]) {
                $match = $false
                break
            }
        }
        if ($match) { return $r.Site }
    }
    return 'Inconnu'
}

function ConvertTo-Array {
    param($Data)
    if ($null -eq $Data) { return @() }
    if ($Data -is [array]) { return $Data }
    return @($Data)
}

# ============================================================
# CHARGEMENT DE LA CONFIG
# ============================================================
$cfg = Import-MonitorConfig -Path $ConfigFile -Defaults $DefaultConfig

$SeuilBootLong        = [int]$cfg.SeuilBootLong
$SeuilCrashRecent     = [int]$cfg.SeuilCrashRecent
$SeuilDiskAlert       = [int]$cfg.SeuilDiskAlert
$SeuilDiskWarning     = [int]$cfg.SeuilDiskWarning
$SeuilOfflineJours    = [int]$cfg.SeuilOfflineJours
$DashboardTitle       = [string]$cfg.DashboardTitle
$DashboardSubtitle    = [string]$cfg.DashboardSubtitle
$MaskHealthyByDefault = [bool]$cfg.MaskHealthyByDefault
$ScoreWeights         = $cfg.ScoreWeights

# Resolution du chemin CSV
$csvRel = [string]$cfg.CsvRanges
if ([System.IO.Path]::IsPathRooted($csvRel)) {
    $CsvRanges = $csvRel
} else {
    $CsvRanges = Join-Path $SharePath $csvRel
}

# ============================================================
# LECTURE DU CSV DE RANGES IP
# ============================================================
$rangesIP = @()
if (Test-Path $CsvRanges) {
    Write-Host "[*] Chargement des ranges IP depuis $CsvRanges" -ForegroundColor Cyan
    $raw = Import-Csv -Path $CsvRanges -Delimiter ',' |
        Where-Object { $_.Pattern1 -like '*/*' }

    foreach ($r in $raw) {
        $built = Build-CIDRRange -CIDR $r.Pattern1.Trim() -Site $r.Entity.Trim()
        if ($built) { $rangesIP += $built }
    }
    Write-Host "[+] $($rangesIP.Count) range(s) chargee(s) et pre-compilee(s)" -ForegroundColor Green
} else {
    Write-Host "[!] CSV de ranges non trouve : $CsvRanges (colonne Site desactivee)" -ForegroundColor Yellow
}
$showSite = ($rangesIP.Count -gt 0)

# ============================================================
# LECTURE DES JSON (avec verification SchemaVersion)
# ============================================================
Write-Host "[*] Lecture des donnees depuis $SharePath ..." -ForegroundColor Cyan

$jsonFiles = Get-ChildItem -Path $SharePath -Filter "$FiltrePC.json" -ErrorAction Stop
$allData   = [System.Collections.Generic.List[PSCustomObject]]::new()
$schemaWarnings = [System.Collections.Generic.List[string]]::new()

foreach ($file in $jsonFiles) {
    try {
        $content = Get-Content -Path $file.FullName -Raw -Encoding UTF8
        $parsed  = $content | ConvertFrom-Json

        $schema = $parsed.SchemaVersion
        if (-not $schema) {
            $schemaWarnings.Add("$($file.Name) : SchemaVersion absent (Collector < v5)")
        } elseif ($schema -notin $SchemaCompatible) {
            $schemaWarnings.Add("$($file.Name) : schema $schema (attendu $($SchemaCompatible -join '/'))")
        }

        $allData.Add($parsed)
    } catch {
        Write-Warning "Impossible de lire $($file.Name) : $_"
    }
}

if ($schemaWarnings.Count -gt 0) {
    Write-Host "[!] $($schemaWarnings.Count) fichier(s) avec schema non conforme :" -ForegroundColor Yellow
    $schemaWarnings | Select-Object -First 5 | ForEach-Object {
        Write-Host "    $_" -ForegroundColor Yellow
    }
    if ($schemaWarnings.Count -gt 5) {
        Write-Host "    ... et $($schemaWarnings.Count - 5) autre(s)" -ForegroundColor Yellow
    }
    Write-Host "    (le Dashboard reste tolerant, mais certaines donnees peuvent manquer)" -ForegroundColor DarkGray
}

if ($allData.Count -eq 0) {
    Write-Host "[-] Aucune donnee trouvee." -ForegroundColor Red
    exit 1
}
Write-Host "[+] $($allData.Count) PC(s) charge(s)" -ForegroundColor Green

# ============================================================
# CONSTRUCTION DU PAYLOAD JS
# ============================================================
$now       = Get-Date
$embedData = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($pc in $allData) {
    $site = Get-SiteFromIP -IP $pc.Machine.IP -Ranges $rangesIP

    $collectedAt = [datetime]$pc.Machine.CollectedAt
    $hoursAgo    = [math]::Round(($now - $collectedAt).TotalHours, 1)
    $isOffline   = ($hoursAgo -gt (24 * $SeuilOfflineJours))

    $crashList = @($pc.Events | Where-Object { $_.EventId -eq 41 } |
        Sort-Object Timestamp -Descending | ForEach-Object {
            # v1.6 : propager CrashCause (null sur JSON v1.4/1.5, rempli sur v1.6).
            # Le JS gere les deux cas avec un fallback sur Type/Detail.
            [PSCustomObject]@{
                Timestamp  = $_.Timestamp
                Type       = if ($_.Type) { $_.Type } else { 'Freeze' }
                Detail     = ConvertTo-HtmlSafe $_.Detail
                CrashCause = $_.CrashCause
                Message    = ConvertTo-HtmlSafe $_.Message
            }
        })

    $bootList = @($pc.BootDurations | Sort-Object { [datetime]$_.DateBoot } | ForEach-Object {
        [PSCustomObject]@{
            DateBoot      = $_.DateBoot
            DurationMin   = $_.DurationMin
            EstBootLong   = $_.EstBootLong
            PrecedentType = $_.PrecedentType
            Method        = if ($_.Method) { $_.Method } else { 'unknown' }
            # v6.0 : fix BootType manquant dans l'embed (le Collector l'ecrivait
            # bien dans le JSON mais le Dashboard ne le recopiait pas dans le
            # payload JS, d'ou le panneau "Repartition demarrages" qui
            # affichait tout en "Inconnu").
            BootType      = if ($_.BootType) { $_.BootType } else { 'Unknown' }
        }
    })

    $bsodList = @(ConvertTo-Array $pc.BSODs | Where-Object { $_.Date } | ForEach-Object {
        [PSCustomObject]@{ Date = $_.Date; Nom = $_.Nom }
    })

    $warningList = @(ConvertTo-Array $pc.ResourceWarnings | Where-Object { $_.Timestamp } | ForEach-Object {
        # v1.5 : propager les champs de clustering Event 51 (Count/IsBurst/FirstSeen/LastSeen).
        # Les JSON v1.4 n'ont pas ces champs : $_.Count retournera $null et sera serialise
        # en null dans le JSON, ce que le JS gere avec 'typeof w.Count === "number"'.
        [PSCustomObject]@{
            Timestamp = $_.Timestamp
            Type      = $_.Type
            Detail    = ConvertTo-HtmlSafe $_.Detail
            Count     = $_.Count
            IsBurst   = $_.IsBurst
            FirstSeen = $_.FirstSeen
            LastSeen  = $_.LastSeen
        }
    })

    $topRAMList = @(ConvertTo-Array $pc.TopRAM | Where-Object { $_.Name } | ForEach-Object {
        [PSCustomObject]@{
            Name         = ConvertTo-HtmlSafe $_.Name
            WorkingSetMB = $_.WorkingSetMB
            CPUSeconds   = $_.CPUSeconds
        }
    })

    $diskList = @(ConvertTo-Array $pc.DiskInfo | Where-Object { $_.Drive } | ForEach-Object {
        [PSCustomObject]@{
            Drive   = $_.Drive
            Label   = if ($_.Label) { ConvertTo-HtmlSafe $_.Label } else { 'Sans nom' }
            TotalGB = $_.TotalGB
            UsedGB  = $_.UsedGB
            FreeGB  = $_.FreeGB
            PctUsed = $_.PctUsed
            PctFree = $_.PctFree
            IsAlert = [bool]$_.IsAlert
        }
    })

    $crasherList = @(ConvertTo-Array $pc.TopCrashers | Where-Object { $_.AppName } | ForEach-Object {
        [PSCustomObject]@{
            AppName    = ConvertTo-HtmlSafe $_.AppName
            CrashCount = $_.CrashCount
        }
    })

    # ================================================================
    # HARDWARE HEALTH : deux schemas supportes
    #   v5.2 : WHEA_Fatal (par occurrence) + WHEA_Corrected (agrege par signature)
    #   v5.0 : WHEA_CPU / WHEA_RAM / WHEA_PCIe (par ID d'event, imprecise)
    # On normalise tout vers le schema v5.2 en memoire.
    # ================================================================
    $hwHealth = [PSCustomObject]@{
        WHEA_Fatal     = @()
        WHEA_Corrected = @()
        GPU_TDR        = @()
        Thermal        = @()
        # v1.8 : CPU throttling detaille (agregation quotidienne Event 35/55).
        # Reste vide si Collector < v1.8 : pas de regression, le JS gere.
        CPUThrottling  = @()
    }

    if ($pc.HardwareHealth) {
        $hh = $pc.HardwareHealth

        # --- Detection du schema ---
        $hasV52 = ($null -ne $hh.PSObject.Properties['WHEA_Fatal']) -or ($null -ne $hh.PSObject.Properties['WHEA_Corrected'])
        $hasV50 = ($null -ne $hh.PSObject.Properties['WHEA_CPU'])   -or ($null -ne $hh.PSObject.Properties['WHEA_RAM']) -or ($null -ne $hh.PSObject.Properties['WHEA_PCIe'])

        if ($hasV52) {
            # Schema v5.2 : lecture directe
            $hwHealth.WHEA_Fatal = @(ConvertTo-Array $hh.WHEA_Fatal | Where-Object { $_.Timestamp } | ForEach-Object {
                [PSCustomObject]@{
                    Timestamp   = $_.Timestamp
                    EventId     = $_.EventId
                    Severity    = $_.Severity
                    Component   = ConvertTo-HtmlSafe $_.Component
                    ErrorSource = ConvertTo-HtmlSafe $_.ErrorSource
                    BDF         = ConvertTo-HtmlSafe $_.BDF
                    Detail      = ConvertTo-HtmlSafe $_.Detail
                }
            })
            $hwHealth.WHEA_Corrected = @(ConvertTo-Array $hh.WHEA_Corrected | Where-Object { $_.LastSeen } | ForEach-Object {
                [PSCustomObject]@{
                    Component   = ConvertTo-HtmlSafe $_.Component
                    EventId     = $_.EventId
                    ErrorSource = ConvertTo-HtmlSafe $_.ErrorSource
                    BDF         = ConvertTo-HtmlSafe $_.BDF
                    Count       = $_.Count
                    FirstSeen   = $_.FirstSeen
                    LastSeen    = $_.LastSeen
                    Detail      = ConvertTo-HtmlSafe $_.Detail
                }
            })
        }
        elseif ($hasV50) {
            # Schema v5.0 legacy : on convertit les trois listes en WHEA_Fatal
            # (on est conservateur : tout etait mis en "erreur" donc on considere Fatal)
            $legacyFatal = @()
            foreach ($k in @('WHEA_CPU', 'WHEA_RAM', 'WHEA_PCIe')) {
                $comp = switch ($k) { 'WHEA_CPU' {'CPU'}; 'WHEA_RAM' {'RAM'}; 'WHEA_PCIe' {'PCIe'} }
                $legacyFatal += @(ConvertTo-Array $hh.$k | Where-Object { $_.Timestamp } | ForEach-Object {
                    [PSCustomObject]@{
                        Timestamp   = $_.Timestamp
                        EventId     = if ($_.EventId) { $_.EventId } else { 0 }
                        Severity    = 2
                        Component   = $comp
                        ErrorSource = "(legacy v5.0)"
                        BDF         = ""
                        Detail      = ConvertTo-HtmlSafe $_.Detail
                    }
                })
            }
            $hwHealth.WHEA_Fatal = $legacyFatal
            # WHEA_Corrected reste vide : l'info n'existe pas en v5.0
        }

        $hwHealth.GPU_TDR = @(ConvertTo-Array $hh.GPU_TDR | Where-Object { $_.Timestamp } | ForEach-Object {
            [PSCustomObject]@{
                Timestamp = $_.Timestamp
                Driver    = ConvertTo-HtmlSafe $_.Driver
                Detail    = ConvertTo-HtmlSafe $_.Detail
            }
        })
        $hwHealth.Thermal = @(ConvertTo-Array $hh.Thermal | Where-Object { $_.Timestamp } | ForEach-Object {
            [PSCustomObject]@{
                Timestamp   = $_.Timestamp
                AlertType   = ConvertTo-HtmlSafe $_.AlertType
                Temperature = ConvertTo-HtmlSafe $_.Temperature
                Detail      = ConvertTo-HtmlSafe $_.Detail
            }
        })
        # v1.8 : CPU throttling (present uniquement sur JSON >= 1.8)
        if ($hh.PSObject.Properties['CPUThrottling']) {
            $hwHealth.CPUThrottling = @(ConvertTo-Array $hh.CPUThrottling | Where-Object { $_.Day } | ForEach-Object {
                [PSCustomObject]@{
                    Day       = $_.Day
                    EventId   = $_.EventId
                    Type      = ConvertTo-HtmlSafe $_.Type
                    Count     = $_.Count
                    FirstSeen = $_.FirstSeen
                    LastSeen  = $_.LastSeen
                    Detail    = ConvertTo-HtmlSafe $_.Detail
                }
            })
        }
    }

    # Stats v5.2 si presentes, fallback vers stats v5.0
    $statsObj          = $pc.Stats
    $totalWHEAFatal    = if ($statsObj.TotalWHEAFatal -ne $null)    { $statsObj.TotalWHEAFatal }    else { @($hwHealth.WHEA_Fatal).Count }
    $totalWHEACorr     = if ($statsObj.TotalWHEACorrected -ne $null){ $statsObj.TotalWHEACorrected }else { 0 }
    $totalGPU          = if ($statsObj.TotalGPU)                    { $statsObj.TotalGPU }          else { @($hwHealth.GPU_TDR).Count }
    $totalThermal      = if ($statsObj.TotalThermal)                { $statsObj.TotalThermal }      else { @($hwHealth.Thermal).Count }
    $totalHWAlerting   = $totalWHEAFatal + $totalGPU + $totalThermal   # les "vraies" alertes (hors corrigees)

    # BootsByType : disponible en v5.2 uniquement
    $bootsByType = $null
    if ($statsObj.BootsByType) {
        $bootsByType = [PSCustomObject]@{
            ColdBoot    = if ($statsObj.BootsByType.ColdBoot)    { $statsObj.BootsByType.ColdBoot }    else { 0 }
            FastStartup = if ($statsObj.BootsByType.FastStartup) { $statsObj.BootsByType.FastStartup } else { 0 }
            Resume      = if ($statsObj.BootsByType.Resume)      { $statsObj.BootsByType.Resume }      else { 0 }
            Unknown     = if ($statsObj.BootsByType.Unknown)     { $statsObj.BootsByType.Unknown }     else { 0 }
        }
    }

    $cpuName = if ($pc.Machine.CPUName) { ([string]$pc.Machine.CPUName).Trim() } else { '' }

    # ================================================================
    # BATTERIE (v5.3+) : null si absent (JSON v5.2 ou desktop)
    # ================================================================
    $batteryEmbed = $null
    if ($pc.BatteryInfo) {
        $batteryEmbed = [PSCustomObject]@{
            HasBattery         = [bool]$pc.BatteryInfo.HasBattery
            Manufacturer       = ConvertTo-HtmlSafe ([string]$pc.BatteryInfo.Manufacturer)
            Chemistry          = ConvertTo-HtmlSafe ([string]$pc.BatteryInfo.Chemistry)
            DesignCapacity     = if ($null -ne $pc.BatteryInfo.DesignCapacity)     { [int64]$pc.BatteryInfo.DesignCapacity }     else { 0 }
            FullChargeCapacity = if ($null -ne $pc.BatteryInfo.FullChargeCapacity) { [int64]$pc.BatteryInfo.FullChargeCapacity } else { 0 }
            HealthPercent      = if ($null -ne $pc.BatteryInfo.HealthPercent)      { [double]$pc.BatteryInfo.HealthPercent }     else { 0 }
            HealthCategory     = ConvertTo-HtmlSafe ([string]$pc.BatteryInfo.HealthCategory)
            CycleCount         = $pc.BatteryInfo.CycleCount   # peut etre $null
            CurrentChargePct   = if ($null -ne $pc.BatteryInfo.CurrentChargePct)   { [int]$pc.BatteryInfo.CurrentChargePct }     else { 0 }
            Status             = ConvertTo-HtmlSafe ([string]$pc.BatteryInfo.Status)
            IsAlert            = [bool]$pc.BatteryInfo.IsAlert
        }
    }

    # ================================================================
    # SERVICES HEALTH (v5.3+) : null si absent
    # On structure en tableau de services pour faciliter l'extensibilite
    # (on pourra rajouter FortiClient/Intune/Dell sans casser le dashboard)
    # ================================================================
    $servicesEmbed = $null
    if ($pc.ServicesHealth -and $pc.ServicesHealth.SentinelAgent) {
        $sa = $pc.ServicesHealth.SentinelAgent
        $servicesEmbed = [PSCustomObject]@{
            SentinelAgent = [PSCustomObject]@{
                DisplayName = ConvertTo-HtmlSafe ([string]$sa.DisplayName)
                ServiceName = ConvertTo-HtmlSafe ([string]$sa.ServiceName)
                Installed   = [bool]$sa.Installed
                Status      = ConvertTo-HtmlSafe ([string]$sa.Status)
                StartType   = ConvertTo-HtmlSafe ([string]$sa.StartType)
                IsAlert     = [bool]$sa.IsAlert
            }
        }
    }

    # ================================================================
    # BOOT PERFORMANCE (v5.4+) : null si absent
    # ================================================================
    $bootPerfEmbed = $null
    if ($pc.BootPerformance) {
        $bp = $pc.BootPerformance
        $lastBootEmbed = $null
        if ($bp.LastBoot) {
            $lastBootEmbed = [PSCustomObject]@{
                Timestamp                   = $bp.LastBoot.Timestamp
                Level                       = ConvertTo-HtmlSafe ([string]$bp.LastBoot.Level)
                BootTimeMs                  = [int64]$bp.LastBoot.BootTimeMs
                MainPathBootTimeMs          = [int64]$bp.LastBoot.MainPathBootTimeMs
                BootPostBootTimeMs          = [int64]$bp.LastBoot.BootPostBootTimeMs
                UserProfileProcessingTimeMs = [int64]$bp.LastBoot.UserProfileProcessingTimeMs
                ExplorerInitTimeMs          = [int64]$bp.LastBoot.ExplorerInitTimeMs
                NumStartupApps              = [int]$bp.LastBoot.NumStartupApps
                IsRebootAfterInstall        = [bool]$bp.LastBoot.IsRebootAfterInstall
                IsSlow                      = [bool]$bp.LastBoot.IsSlow
            }
        }
        $historyEmbed = @()
        if ($bp.History) {
            $historyEmbed = @(ConvertTo-Array $bp.History | ForEach-Object {
                [PSCustomObject]@{
                    Timestamp                   = $_.Timestamp
                    Level                       = ConvertTo-HtmlSafe ([string]$_.Level)
                    BootTimeMs                  = [int64]$_.BootTimeMs
                    MainPathBootTimeMs          = [int64]$_.MainPathBootTimeMs
                    BootPostBootTimeMs          = [int64]$_.BootPostBootTimeMs
                    UserProfileProcessingTimeMs = [int64]$_.UserProfileProcessingTimeMs
                    ExplorerInitTimeMs          = [int64]$_.ExplorerInitTimeMs
                    NumStartupApps              = [int]$_.NumStartupApps
                    IsRebootAfterInstall        = [bool]$_.IsRebootAfterInstall
                    IsSlow                      = [bool]$_.IsSlow
                }
            })
        }
        $bootPerfEmbed = [PSCustomObject]@{
            LastBoot = $lastBootEmbed
            History  = $historyEmbed
            Stats    = if ($bp.Stats) {
                [PSCustomObject]@{
                    BootsAnalyzed  = [int]$bp.Stats.BootsAnalyzed
                    AvgBootTimeMs  = [int]$bp.Stats.AvgBootTimeMs
                    AvgMainPathMs  = [int]$bp.Stats.AvgMainPathMs
                    AvgPostBootMs  = [int]$bp.Stats.AvgPostBootMs
                    MaxBootTimeMs  = [int]$bp.Stats.MaxBootTimeMs
                    SlowBootsCount = [int]$bp.Stats.SlowBootsCount
                }
            } else { $null }
            IsAlert = [bool]$bp.IsAlert
        }
    }

    # ================================================================
    # DISK HEALTH SMART (v5.4+) : array, vide si absent
    # ================================================================
    $diskHealthEmbed = @()
    if ($pc.DiskHealth) {
        $diskHealthEmbed = @(ConvertTo-Array $pc.DiskHealth | ForEach-Object {
            [PSCustomObject]@{
                FriendlyName           = ConvertTo-HtmlSafe ([string]$_.FriendlyName)
                MediaType              = ConvertTo-HtmlSafe ([string]$_.MediaType)
                BusType                = ConvertTo-HtmlSafe ([string]$_.BusType)
                SizeGB                 = $_.SizeGB
                OperationalStatus      = ConvertTo-HtmlSafe ([string]$_.OperationalStatus)
                HealthStatus           = ConvertTo-HtmlSafe ([string]$_.HealthStatus)
                TemperatureC           = $_.TemperatureC
                TemperatureMaxC        = $_.TemperatureMaxC
                WearPct                = $_.WearPct
                PowerOnHours           = $_.PowerOnHours
                ReadErrorsTotal        = $_.ReadErrorsTotal
                ReadErrorsUncorrected  = $_.ReadErrorsUncorrected
                WriteErrorsTotal       = $_.WriteErrorsTotal
                WriteErrorsUncorrected = $_.WriteErrorsUncorrected
                IsAlert                = [bool]$_.IsAlert
                AlertReasons           = @(ConvertTo-Array $_.AlertReasons | ForEach-Object { ConvertTo-HtmlSafe ([string]$_) })
            }
        })
    }

    # ================================================================
    # MONITORS (v5.5+) : array, vide si absent ou si que des ecrans internes
    # ================================================================
    $monitorsEmbed = @()
    if ($pc.Monitors) {
        $monitorsEmbed = @(ConvertTo-Array $pc.Monitors | ForEach-Object {
            [PSCustomObject]@{
                ManufacturerCode  = ConvertTo-HtmlSafe ([string]$_.ManufacturerCode)
                Manufacturer      = ConvertTo-HtmlSafe ([string]$_.Manufacturer)
                Model             = ConvertTo-HtmlSafe ([string]$_.Model)
                SerialNumber      = ConvertTo-HtmlSafe ([string]$_.SerialNumber)
                ProductCode       = ConvertTo-HtmlSafe ([string]$_.ProductCode)
                YearOfManufacture = $_.YearOfManufacture
                WeekOfManufacture = $_.WeekOfManufacture
                AgeYears          = $_.AgeYears
                Active            = [bool]$_.Active
                VideoOutputTech   = $_.VideoOutputTech
            }
        })
    }

    # ================================================================
    # MEMORY INVENTORY (v1.8) : null si Collector < v1.8
    # ================================================================
    $memoryEmbed = $null
    if ($pc.MemoryInventory) {
        $memInv = $pc.MemoryInventory
        $memoryEmbed = [PSCustomObject]@{
            TotalInstalledGB = if ($null -ne $memInv.TotalInstalledGB) { [int]$memInv.TotalInstalledGB } else { 0 }
            MaxCapacityGB    = if ($null -ne $memInv.MaxCapacityGB)    { [int]$memInv.MaxCapacityGB }    else { 0 }
            TotalSlots       = if ($null -ne $memInv.TotalSlots)       { [int]$memInv.TotalSlots }       else { 0 }
            OccupiedSlots    = if ($null -ne $memInv.OccupiedSlots)    { [int]$memInv.OccupiedSlots }    else { 0 }
            FreeSlots        = if ($null -ne $memInv.FreeSlots)        { [int]$memInv.FreeSlots }        else { 0 }
            CanUpgrade       = [bool]$memInv.CanUpgrade
            Modules          = @(ConvertTo-Array $memInv.Modules | ForEach-Object {
                [PSCustomObject]@{
                    Slot         = ConvertTo-HtmlSafe ([string]$_.Slot)
                    Bank         = ConvertTo-HtmlSafe ([string]$_.Bank)
                    CapacityGB   = if ($null -ne $_.CapacityGB) { [int]$_.CapacityGB } else { 0 }
                    Type         = ConvertTo-HtmlSafe ([string]$_.Type)
                    SpeedMHz     = if ($null -ne $_.SpeedMHz) { [int]$_.SpeedMHz } else { 0 }
                    Manufacturer = ConvertTo-HtmlSafe ([string]$_.Manufacturer)
                    PartNumber   = ConvertTo-HtmlSafe ([string]$_.PartNumber)
                }
            })
        }
    }

    # ================================================================
    # GPU INVENTORY (v1.8) : tableau vide si Collector < v1.8
    # ================================================================
    $gpuEmbed = @()
    if ($pc.GPUInventory) {
        $gpuEmbed = @(ConvertTo-Array $pc.GPUInventory | ForEach-Object {
            [PSCustomObject]@{
                Name          = ConvertTo-HtmlSafe ([string]$_.Name)
                DriverVersion = ConvertTo-HtmlSafe ([string]$_.DriverVersion)
                DriverDate    = ConvertTo-HtmlSafe ([string]$_.DriverDate)
            }
        })
    }

    $embedData.Add([PSCustomObject]@{
        PC               = ConvertTo-HtmlSafe ([string]$pc.Machine.PC).Trim()
        IP               = ConvertTo-HtmlSafe $pc.Machine.IP
        Site             = ConvertTo-HtmlSafe $site
        CurrentUser      = ConvertTo-HtmlSafe $pc.Machine.CurrentUser
        LastBoot         = $pc.Machine.LastBoot
        UptimeDays       = $pc.Machine.UptimeDays
        CollectedAt      = $pc.Machine.CollectedAt
        IsOffline        = $isOffline
        CPUName          = ConvertTo-HtmlSafe $cpuName
        CPUVendor        = ConvertTo-HtmlSafe $pc.Machine.CPUVendor
        CPUGen           = $pc.Machine.CPUGen
        CPUYear          = $pc.Machine.CPUYear
        CPUAge           = $pc.Machine.CPUAge
        CPUAgeCategory   = if ($pc.Machine.CPUAgeCategory) { $pc.Machine.CPUAgeCategory } else { 'Inconnu' }
        ConnectionType   = if ($pc.Machine.ConnectionType) { $pc.Machine.ConnectionType } else { 'Inconnu' }
        # v5.6 : chassis info (peut etre null si Collector < v5.6)
        ChassisInfo      = if ($pc.Machine.ChassisInfo) {
            [PSCustomObject]@{
                ChassisType  = $pc.Machine.ChassisInfo.ChassisType
                ChassisLabel = ConvertTo-HtmlSafe ([string]$pc.Machine.ChassisInfo.ChassisLabel)
                IsLaptop     = [bool]$pc.Machine.ChassisInfo.IsLaptop
                IsDesktop    = [bool]$pc.Machine.ChassisInfo.IsDesktop
                IsAIO        = [bool]$pc.Machine.ChassisInfo.IsAIO
            }
        } else { $null }
        Crashes          = $crashList
        Boots            = $bootList
        BSODs            = $bsodList
        ResourceWarnings = $warningList
        TopRAM           = $topRAMList
        DiskInfo         = $diskList
        TopCrashers      = $crasherList
        HardwareHealth   = $hwHealth
        BootsByType      = $bootsByType
        TotalWHEAFatal   = $totalWHEAFatal
        TotalWHEACorr    = $totalWHEACorr
        TotalGPU         = $totalGPU
        TotalThermal     = $totalThermal
        TotalHardware    = $totalHWAlerting
        # v1.6 : total des hard crashs filtres (BSODSilent+SleepResumeFailed+PowerLoss).
        # null sur JSON v1.4/1.5, int sur v1.6. Le JS gere ce cas dans computeVerdict.
        TotalHardCrash   = if ($null -ne $pc.Stats.TotalHardCrash) { [int]$pc.Stats.TotalHardCrash } else { $null }
        # v5.3 / v5.4 additions (peuvent etre $null selon le schema du JSON source)
        BatteryInfo      = $batteryEmbed
        ServicesHealth   = $servicesEmbed
        BootPerformance  = $bootPerfEmbed
        DiskHealth       = $diskHealthEmbed
        # v5.5 additions
        Monitors         = $monitorsEmbed
        # v1.8 additions
        MemoryInventory  = $memoryEmbed
        GPUInventory     = $gpuEmbed
    })
}

# Force l'array meme avec 1 seul element (bug ConvertTo-Json connu)
$jsonEmbed    = ConvertTo-Json -InputObject @($embedData) -Depth 6
$weightsEmbed = ConvertTo-Json -InputObject $ScoreWeights -Compress

$titleHtml    = ConvertTo-HtmlSafe $DashboardTitle
$subtitleHtml = ConvertTo-HtmlSafe $DashboardSubtitle
$maskHealthyJs = if ($MaskHealthyByDefault) { 'true' } else { 'false' }
$showSiteJs    = if ($showSite) { 'true' } else { 'false' }

# Favicon SVG minimaliste (moniteur avec barres) encode en data URI
$faviconSvg = @'
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64"><rect x="6" y="10" width="52" height="36" rx="4" fill="#6c63ff"/><rect x="10" y="14" width="44" height="28" rx="2" fill="#1a1a2e"/><rect x="14" y="30" width="4" height="8" fill="#4ecca3"/><rect x="22" y="26" width="4" height="12" fill="#ffa502"/><rect x="30" y="22" width="4" height="16" fill="#ff6b6b"/><rect x="38" y="28" width="4" height="10" fill="#4ecca3"/><rect x="46" y="24" width="4" height="14" fill="#a855f7"/><rect x="24" y="48" width="16" height="4" fill="#6c63ff"/><rect x="18" y="52" width="28" height="4" rx="1" fill="#6c63ff"/></svg>
'@
$faviconBytes = [System.Text.Encoding]::UTF8.GetBytes($faviconSvg)
$faviconB64   = [Convert]::ToBase64String($faviconBytes)

# ============================================================
# HTML COMPLET
# ============================================================
$html = @"
<!DOCTYPE html>
<html lang="fr" data-theme="dark">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<link rel="icon" type="image/svg+xml" href="data:image/svg+xml;base64,$faviconB64">
<title>$titleHtml Dashboard</title>
<style>
    /* ===== VARIABLES THEME =====
       v5.6.1 : palette dark retravaillee pour ameliorer le contraste.
       Ancienne palette : text-dim #aaa text-muted #888 text-faint #666
       -> ratio de contraste insuffisant (en dessous du seuil WCAG AA
       pour les textes secondaires). Nouvelle palette : gris plus clairs
       pour matcher la lisibilite du theme clair.
    */
    :root {
        --bg-main:      #1a1a2e;
        --bg-panel:     #22223d;
        --bg-elevated:  #121a2e;
        --border:       #3a3a5c;    /* un peu plus clair pour qu'on voie les traits */
        --text-main:    #f0f2f8;    /* etait #eee -> plus blanc pur */
        --text-dim:     #c8cbd6;    /* etait #aaa -> +20 en lumiere */
        --text-muted:   #9ea3b3;    /* etait #888 -> +22 en lumiere (lisible en labels) */
        --text-faint:   #7a8192;    /* etait #666 -> +28 (WCAG AA atteint) */
        --text-ghost:   #5e6478;    /* etait #555 -> placeholders */
        --accent:       #8d85ff;    /* etait #6c63ff -> accent plus lumineux sur fond sombre */
        --green:        #5dd4a8;    /* etait #4ecca3 -> plus vif */
        --orange:       #ffb84d;    /* etait #ffa502 -> plus vif */
        --red:          #ff7a7a;    /* etait #ff6b6b -> plus vif */
        --pink:         #ffa8f0;
        --cyan:         #38bcd2;    /* etait #17a2b8 -> plus lumineux */
        --purple:       #b570f7;    /* etait #a855f7 -> plus lumineux */
        --yellow:       #ffc555;    /* etait #ffaa44 -> plus lumineux */
        --bg-danger:    #2a1a1a;
        --bg-hw:        #2a1a3a;
        --bg-warning:   #2a2a1a;
        --bg-success:   #1a2a22;
    }
    [data-theme="light"] {
        --bg-main:      #f4f5fa;
        --bg-panel:     #ffffff;
        --bg-elevated:  #eef0f7;
        --border:       #d8dbe8;
        --text-main:    #1a1a2e;
        --text-dim:     #495266;
        --text-muted:   #6b7380;
        --text-faint:   #8890a0;
        --text-ghost:   #a8b0c0;
        --accent:       #5b52e0;
        --green:        #16a085;
        --orange:       #e67e22;
        --red:          #d63446;
        --pink:         #c941a0;
        --cyan:         #138496;
        --purple:       #8e44ad;
        --yellow:       #d17200;
        --bg-danger:    #fce8eb;
        --bg-hw:        #f3e9fb;
        --bg-warning:   #fdf3dc;
        --bg-success:   #e0f4ec;
    }

    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        background: var(--bg-main);
        color: var(--text-main);
        padding: 24px;
        min-height: 100vh;
        transition: background 0.2s, color 0.2s;
    }
    .header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 24px;
        border-bottom: 1px solid var(--border);
        padding-bottom: 16px;
    }
    .header h1 { font-size: 24px; font-weight: 400; color: var(--accent); }
    .header h1 span { font-weight: 700; color: var(--text-main); }
    .header-right { display: flex; align-items: center; gap: 16px; }
    .timestamp { color: var(--text-muted); font-size: 12px; }

    .theme-toggle {
        background: var(--bg-panel);
        border: 1px solid var(--border);
        color: var(--text-dim);
        padding: 6px 10px;
        border-radius: 6px;
        cursor: pointer;
        font-size: 14px;
        transition: all 0.2s;
    }
    .theme-toggle:hover { border-color: var(--accent); color: var(--accent); }

    /* ===== BARRE DE FILTRES ===== */
    .toolbar {
        display: flex;
        align-items: center;
        gap: 8px;
        padding: 10px 20px;
        background: var(--bg-panel);
        border-radius: 10px 10px 0 0;
        flex-wrap: wrap;
    }
    .toolbar .label { color: var(--text-muted); font-size: 12px; margin-right: 6px; }
    .range-btn, .filter-btn {
        padding: 5px 14px;
        border-radius: 20px;
        border: 1px solid var(--border);
        background: transparent;
        color: var(--text-dim);
        cursor: pointer;
        font-size: 12px;
        transition: all 0.15s;
    }
    .range-btn:hover, .filter-btn:hover { border-color: var(--accent); color: var(--accent); }
    .range-btn.active { background: var(--accent); color: white; border-color: var(--accent); }
    .filter-btn.active { background: var(--accent); color: white; border-color: var(--accent); }

    .divider {
        width: 1px;
        height: 20px;
        background: var(--border);
        margin: 0 4px;
    }

    .toolbar select, .toolbar .search-input {
        padding: 5px 10px;
        border-radius: 6px;
        border: 1px solid var(--border);
        background: var(--bg-main);
        color: var(--text-main);
        font-size: 12px;
        cursor: pointer;
    }
    .toolbar select:focus, .toolbar .search-input:focus {
        outline: none;
        border-color: var(--accent);
    }

    .search-wrap { margin-left: auto; position: relative; }
    .search-input {
        padding: 6px 12px 6px 32px !important;
        border-radius: 20px !important;
        width: 220px;
    }
    .search-icon {
        position: absolute;
        left: 10px;
        top: 50%;
        transform: translateY(-50%);
        color: var(--text-faint);
        font-size: 14px;
        pointer-events: none;
    }

    .export-btn {
        background: var(--green);
        color: #0a1a12;
        border: none;
        padding: 6px 14px;
        border-radius: 6px;
        cursor: pointer;
        font-size: 12px;
        font-weight: 600;
        transition: all 0.15s;
    }
    .export-btn:hover { transform: translateY(-1px); opacity: 0.9; }

    /* ===== BOOT TYPE CHIPS (v5.2) ===== */
    .boot-type-chip {
        display: inline-flex;
        align-items: center;
        gap: 4px;
        padding: 2px 8px;
        border-radius: 10px;
        font-size: 10px;
        font-weight: 600;
        white-space: nowrap;
    }
    .boot-cold    { background: #1a2a3d; color: #66aaff; }
    .boot-fast    { background: #3d3d1a; color: var(--yellow); }
    .boot-resume  { background: #2a1a3d; color: var(--purple); }
    .boot-unknown { background: #2a2a3a; color: var(--text-muted); }
    [data-theme="light"] .boot-cold    { background: #dde6f7; color: #3a6cbf; }
    [data-theme="light"] .boot-fast    { background: #fdf3dc; color: var(--yellow); }
    [data-theme="light"] .boot-resume  { background: #f3e9fb; color: var(--purple); }
    [data-theme="light"] .boot-unknown { background: #e6e9f2; color: var(--text-muted); }

    /* ===== KPI CARDS ===== */
    .kpi-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(160px, 1fr));
        gap: 14px;
        margin: 16px 0 20px;
    }
    .kpi-card {
        background: linear-gradient(135deg, var(--bg-panel), var(--bg-main));
        padding: 16px 18px;
        border-radius: 10px;
        border: 1px solid var(--border);
        cursor: pointer;
        transition: all 0.2s;
        position: relative;
    }
    .kpi-card:hover {
        transform: translateY(-2px);
        border-color: var(--accent);
    }
    .kpi-card.active {
        border-color: var(--accent);
        box-shadow: 0 0 0 2px var(--accent);
    }
    .kpi-card.card-warning   { border-left: 3px solid var(--orange); }
    .kpi-card.card-danger    { border-left: 3px solid var(--red); }
    .kpi-card.card-info      { border-left: 3px solid var(--cyan); }
    .kpi-card.card-hardware  { border-left: 3px solid var(--purple); }
    .kpi-card.card-success   { border-left: 3px solid var(--green); }
    .kpi-value { font-size: 26px; font-weight: 700; margin-bottom: 4px; }
    /* v5.6 : label KPI plus contraste pour la lisibilite
       (avant : text-muted qui se perdait sur fond sombre) */
    .kpi-label {
        font-size: 11px;
        color: var(--text-dim);
        text-transform: uppercase;
        letter-spacing: 0.5px;
        font-weight: 600;
    }

    /* ===== SUMMARY BAR (v5.5) : chiffres essentiels en haut =====
       Un seul coup d'oeil pour savoir "combien de PC au total, combien
       en ligne, combien sains". Les autres KPIs sont regroupes plus bas. */
    .summary-bar {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
        gap: 12px;
        margin: 16px 0 22px;
    }
    .summary-tile {
        background: var(--bg-panel);
        border: 1px solid var(--border);
        border-radius: 10px;
        padding: 14px 18px;
        display: flex;
        align-items: center;
        gap: 14px;
    }
    .summary-icon {
        width: 42px; height: 42px;
        border-radius: 10px;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 20px;
        flex-shrink: 0;
    }
    .summary-icon.blue   { background: rgba(108, 99, 255, 0.15); color: var(--accent); }
    .summary-icon.green  { background: rgba(78, 204, 163, 0.15); color: var(--green); }
    .summary-icon.red    { background: rgba(255, 107, 107, 0.15); color: var(--red); }
    .summary-content { flex: 1; min-width: 0; }
    .summary-value { font-size: 24px; font-weight: 700; color: var(--text); line-height: 1.1; }
    /* v5.6 : label summary plus contraste aussi */
    .summary-label { font-size: 11px; color: var(--text-dim); text-transform: uppercase; letter-spacing: 0.5px; margin-top: 2px; font-weight: 600; }
    .summary-sub   { font-size: 10px; color: var(--text-muted); margin-top: 2px; }

    /* ===== KPI GROUPS (v5.5) : 4 familles thematiques =====
       Ameliore la lisibilite quand il y a beaucoup d'indicateurs :
       l'oeil trouve direct la famille qui l'interesse au lieu de
       scanner 15 cartes eparpillees. */
    .kpi-group {
        margin: 0 0 16px 0;
    }
    /* v5.6 : header de groupe plus visible (text-dim au lieu de text-muted) */
    .kpi-group-header {
        display: flex;
        align-items: center;
        gap: 8px;
        margin-bottom: 8px;
        font-size: 11px;
        text-transform: uppercase;
        letter-spacing: 0.6px;
        color: var(--text-dim);
        font-weight: 700;
    }
    .kpi-group-icon { font-size: 14px; opacity: 0.9; }
    .kpi-group-line {
        flex: 1;
        height: 1px;
        background: var(--border);
    }
    .kpi-group .kpi-grid {
        margin: 0;
        /* Les cartes dans un groupe sont un peu plus compactes */
        grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
        gap: 10px;
    }
    /* Dans un groupe, les cartes sont legerement moins hautes */
    .kpi-group .kpi-card { padding: 12px 14px; }
    .kpi-group .kpi-value { font-size: 22px; margin-bottom: 2px; }
    .kpi-group .kpi-label { font-size: 10.5px; }

    /* v5.6 : cartes "vides" attenuees mais LABEL reste lisible
       (seule la valeur s'eclaircit, le label garde son poids).
       v5.6.1 : valeur attenuee passe de text-faint a text-muted pour
       etre lisible meme en dark mode (le "0" etait trop pale avant). */
    .kpi-card.kpi-quiet .kpi-value { color: var(--text-muted) !important; font-weight: 500; }
    .kpi-card.kpi-quiet { opacity: 0.92; }
    .kpi-card.kpi-quiet:hover { opacity: 1; }
    /* v5.6 : quand la carte est calme, la barre coloree a gauche est
       neutralisee pour ne pas creer de faux signal visuel */
    .kpi-card.kpi-quiet.card-warning,
    .kpi-card.kpi-quiet.card-danger,
    .kpi-card.kpi-quiet.card-info,
    .kpi-card.kpi-quiet.card-hardware {
        border-left-color: var(--border);
    }

    /* Cartes "critiques" : value > 0 et theme danger -> plus d'emphase */
    .kpi-card.kpi-loud {
        background: linear-gradient(135deg, rgba(255, 107, 107, 0.08), var(--bg-panel));
    }
    .kpi-card.kpi-loud .kpi-value { font-weight: 800; }

    .color-green  { color: var(--green); }
    .color-red    { color: var(--red); }
    .color-orange { color: var(--orange); }
    .color-blue   { color: var(--accent); }
    .color-cyan   { color: var(--cyan); }
    .color-purple { color: var(--purple); }
    .color-pink   { color: var(--pink); }
    .color-yellow { color: var(--yellow); }

    /* ===== KPI GROUP BUTTONS (v5.6) =====
       Remplace les 4 lignes pleine largeur par une barre horizontale
       de 4 boutons compacts. Gain d'espace vertical important, plus
       facile a scanner en 1 coup d'oeil.
       - Bouton "calme" (aucune alerte)  -> discret (gris)
       - Bouton "chaud"  (alertes > 0)    -> colore + badge
       - Clic = filtre tableau sur cette famille
       - Survol = popover avec les sous-KPIs detailles (chacun cliquable) */
    .kpi-group-bar {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
        gap: 10px;
        margin: 16px 0 22px;
    }
    .kpi-group-btn {
        position: relative;
        background: var(--bg-panel);
        border: 1px solid var(--border);
        border-left: 3px solid var(--border);
        border-radius: 8px;
        padding: 12px 14px;
        cursor: pointer;
        display: flex;
        align-items: center;
        gap: 10px;
        transition: border-color 0.15s, transform 0.15s, background 0.15s;
        text-align: left;
        font-family: inherit;
        color: inherit;
    }
    .kpi-group-btn:hover {
        border-color: var(--accent);
        transform: translateY(-1px);
    }
    .kpi-group-btn.active {
        border-color: var(--accent);
        box-shadow: 0 0 0 2px var(--accent);
    }
    .kpi-group-btn-icon {
        width: 34px; height: 34px;
        border-radius: 8px;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 16px;
        flex-shrink: 0;
        background: var(--bg-main);
    }
    .kpi-group-btn-content { flex: 1; min-width: 0; }
    .kpi-group-btn-title {
        font-size: 12px;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        color: var(--text-dim);
        line-height: 1.2;
    }
    .kpi-group-btn-sub {
        font-size: 10.5px;
        color: var(--text-muted);
        margin-top: 2px;
    }
    .kpi-group-btn-count {
        font-size: 22px;
        font-weight: 800;
        color: var(--text-muted);   /* v5.6.1 : etait text-faint, trop pale en dark */
        margin-left: 8px;
    }

    /* Etats colores selon le nombre d'alertes */
    .kpi-group-btn.warn {
        border-left-color: var(--orange);
        background: linear-gradient(135deg, rgba(255, 165, 2, 0.06), var(--bg-panel));
    }
    .kpi-group-btn.warn .kpi-group-btn-icon { background: rgba(255, 165, 2, 0.15); color: var(--orange); }
    .kpi-group-btn.warn .kpi-group-btn-count { color: var(--orange); }
    .kpi-group-btn.danger {
        border-left-color: var(--red);
        background: linear-gradient(135deg, rgba(255, 107, 107, 0.08), var(--bg-panel));
    }
    .kpi-group-btn.danger .kpi-group-btn-icon { background: rgba(255, 107, 107, 0.15); color: var(--red); }
    .kpi-group-btn.danger .kpi-group-btn-count { color: var(--red); }

    /* Teinte discrete sur l'icone selon la famille (visible quand le
       bouton est calme, sinon overridee par warn/danger) */
    .kpi-group-btn.calm.family-security .kpi-group-btn-icon { color: var(--purple); }
    .kpi-group-btn.calm.family-stability .kpi-group-btn-icon { color: var(--pink); }
    .kpi-group-btn.calm.family-performance .kpi-group-btn-icon { color: var(--yellow); }
    .kpi-group-btn.calm.family-material .kpi-group-btn-icon { color: var(--cyan); }

    /* Popover au survol : detail des sous-KPIs de la famille */
    .kpi-group-popover {
        position: absolute;
        top: calc(100% + 8px);
        left: -1px;
        right: -1px;
        z-index: 50;
        background: var(--bg-panel);
        border: 1px solid var(--border);
        border-radius: 8px;
        padding: 10px;
        box-shadow: 0 8px 24px rgba(0,0,0,0.35);
        display: none;
        min-width: 240px;
    }
    [data-theme="light"] .kpi-group-popover {
        box-shadow: 0 8px 24px rgba(0,0,0,0.12);
    }
    .kpi-group-btn:hover .kpi-group-popover {
        display: block;
    }
    /* Pont invisible pour que la souris puisse transiter du bouton
       au popover sans fermer celui-ci */
    .kpi-group-btn:hover::after {
        content: '';
        position: absolute;
        top: 100%;
        left: 0; right: 0;
        height: 10px;
    }
    .kpi-popover-items {
        display: flex;
        flex-direction: column;
        gap: 3px;
    }
    .kpi-popover-item {
        display: flex;
        justify-content: space-between;
        align-items: center;
        padding: 7px 10px;
        border-radius: 5px;
        cursor: pointer;
        transition: background 0.1s;
        font-size: 12px;
    }
    .kpi-popover-item:hover { background: var(--bg-main); }
    .kpi-popover-item.active {
        background: var(--bg-main);
        outline: 1px solid var(--accent);
    }
    .kpi-popover-item .pop-label { color: var(--text-dim); font-weight: 500; }
    .kpi-popover-item .pop-value {
        font-weight: 700;
        font-size: 14px;
        min-width: 28px;
        text-align: right;
        color: var(--text-muted);   /* v5.6.1 : etait text-faint */
    }
    .kpi-popover-item .pop-value.warn   { color: var(--orange); }
    .kpi-popover-item .pop-value.danger { color: var(--red); }

    /* v5.6 : indicateurs tableau en mode "outline" quand tout va bien
       -> seul ce qui est critique (rouge) ressort visuellement */
    .indicator-dot.ok {
        background: transparent;
        color: var(--green);
        border: 1.5px solid var(--green);
        line-height: 11px;
    }

    /* ===== TABLE ===== */
    .table-container {
        background: var(--bg-panel);
        border-radius: 0 0 10px 10px;
        overflow: hidden;
    }
    .table-header {
        padding: 14px 20px;
        border-bottom: 1px solid var(--border);
        display: flex;
        justify-content: space-between;
        align-items: center;
    }
    .table-header h2 {
        font-size: 14px;
        font-weight: 600;
        color: var(--text-dim);
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    .count { color: var(--text-muted); font-size: 12px; }
    .active-filters {
        display: flex;
        gap: 6px;
        flex-wrap: wrap;
        margin-top: 6px;
    }
    .active-filter-chip {
        background: var(--accent);
        color: white;
        padding: 2px 8px 2px 10px;
        border-radius: 12px;
        font-size: 10px;
        display: inline-flex;
        align-items: center;
        gap: 6px;
    }
    .active-filter-chip .close {
        cursor: pointer;
        font-weight: 700;
        opacity: 0.7;
    }
    .active-filter-chip .close:hover { opacity: 1; }

    /* ===== v1.5 : Pagination ===== */
    .pagination-bar {
        display: flex;
        justify-content: space-between;
        align-items: center;
        gap: 16px;
        padding: 12px 16px;
        background: var(--bg-panel);
        border-top: 1px solid var(--border);
        flex-wrap: wrap;
        font-size: 12px;
    }
    .pagination-info {
        color: var(--text-muted);
        font-size: 12px;
        white-space: nowrap;
    }
    .pagination-controls {
        display: flex;
        align-items: center;
        gap: 4px;
        flex-wrap: wrap;
    }
    .pagination-label {
        color: var(--text-muted);
        font-size: 12px;
        margin-right: 4px;
    }
    .pagination-per, .pagination-page, .pagination-nav, .pagination-showmore, .pagination-showall {
        background: var(--bg-main);
        border: 1px solid var(--border);
        color: var(--text-main);
        padding: 5px 10px;
        border-radius: 4px;
        cursor: pointer;
        font-size: 12px;
        font-weight: 500;
        transition: all 0.15s;
    }
    .pagination-per:hover:not(.active),
    .pagination-page:hover:not(.active),
    .pagination-nav:hover:not([disabled]),
    .pagination-showmore:hover,
    .pagination-showall:hover {
        background: var(--bg-hover);
        border-color: var(--accent);
    }
    .pagination-per.active, .pagination-page.active {
        background: var(--accent);
        color: white;
        border-color: var(--accent);
    }
    .pagination-nav[disabled] {
        opacity: 0.4;
        cursor: not-allowed;
    }
    .pagination-sep {
        display: inline-block;
        width: 1px;
        height: 20px;
        background: var(--border);
        margin: 0 6px;
    }
    .pagination-ellipsis {
        color: var(--text-muted);
        padding: 0 4px;
    }
    .pagination-showmore {
        margin-left: 12px;
        background: var(--bg-main);
        color: var(--accent);
        border-color: var(--accent);
    }
    .pagination-showall {
        margin-left: 4px;
        background: transparent;
        color: var(--text-muted);
        font-style: italic;
    }
    /* Responsive : empiler info + controls sur mobile */
    @media (max-width: 900px) {
        .pagination-bar {
            flex-direction: column;
            align-items: flex-start;
        }
        .pagination-controls {
            width: 100%;
        }
    }

    .table-scroll { overflow-x: auto; }
    table { width: 100%; border-collapse: collapse; min-width: 1200px; }
    th, td { padding: 10px 16px; text-align: left; font-size: 13px; }
    th {
        background: var(--bg-main);
        color: var(--text-muted);
        font-weight: 600;
        text-transform: uppercase;
        font-size: 11px;
        letter-spacing: 0.5px;
        border-bottom: 1px solid var(--border);
        user-select: none;
        white-space: nowrap;
    }
    th.sortable { cursor: pointer; transition: color 0.15s; }
    th.sortable:hover { color: var(--accent); }
    th.sort-active { color: var(--accent); }
    th .sort-arrow { margin-left: 4px; font-size: 9px; color: var(--text-ghost); }
    th.sort-active .sort-arrow { color: var(--accent); }

    tbody tr.row-main    { border-bottom: 1px solid var(--border); transition: background 0.1s; }
    tbody tr.row-main:hover { background: var(--bg-main); }
    tbody tr.row-danger   { background: var(--bg-danger) !important; }
    tbody tr.row-hardware { background: var(--bg-hw) !important; }
    tbody tr.row-warning  { background: var(--bg-warning) !important; }

    .badge {
        padding: 2px 10px;
        border-radius: 12px;
        font-size: 11px;
        font-weight: 600;
        display: inline-block;
    }
    .badge-online  { background: #1a3d1a; color: var(--green); }
    .badge-offline { background: #3d1a1a; color: var(--red); }
    [data-theme="light"] .badge-online  { background: #d5f2e9; color: var(--green); }
    [data-theme="light"] .badge-offline { background: #fcdfe3; color: var(--red); }

    .site-badge {
        padding: 2px 8px;
        border-radius: 4px;
        font-size: 11px;
        background: var(--border);
        color: var(--text-dim);
    }
    .site-badge.inconnu { opacity: 0.5; font-style: italic; }

    /* ===== SCORE BADGE ===== */
    .score-badge {
        display: inline-block;
        padding: 3px 9px;
        border-radius: 12px;
        font-size: 12px;
        font-weight: 700;
        min-width: 32px;
        text-align: center;
    }
    .score-ok      { background: #1a3d1a; color: var(--green); }
    .score-warn    { background: #3d3d1a; color: var(--orange); }
    .score-danger  { background: #3d1a1a; color: var(--red); }
    .score-critic  { background: var(--red); color: white; }
    [data-theme="light"] .score-ok      { background: #d5f2e9; color: var(--green); }
    [data-theme="light"] .score-warn    { background: #fdf3dc; color: var(--orange); }
    [data-theme="light"] .score-danger  { background: #fcdfe3; color: var(--red); }
    [data-theme="light"] .score-critic  { background: var(--red); color: white; }

    .kpi-crash    { color: var(--red); font-weight: 700; font-size: 15px; }
    .kpi-bsod     { color: var(--pink); font-weight: 700; font-size: 15px; }
    .kpi-hardware { color: var(--purple); font-weight: 700; font-size: 15px; }

    .crash-badge, .hw-badge, .warn-badge {
        display: inline-block;
        padding: 2px 7px;
        border-radius: 4px;
        font-size: 10px;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        margin-right: 6px;
        vertical-align: middle;
    }
    .crash-badge.bsod        { background: #3d1a3d; color: var(--pink); }
    .crash-badge.freeze-app  { background: #1a2a3d; color: #66aaff; }
    .crash-badge.hard-reset  { background: #2e1a1a; color: #ff9966; }
    .crash-badge.freeze      { background: #1a2535; color: #6699cc; }
    .crash-stopcode          { color: var(--red); font-size: 11px; font-weight: 600; margin-right: 8px; }
    .crash-bugname           { color: var(--yellow); font-size: 11px; font-style: italic; margin-right: 8px; }
    .crash-app               { color: var(--yellow); font-size: 11px; margin-right: 8px; }

    .hw-badge.whea-cpu  { background: #3d1a2a; color: #ff6b9d; }
    .hw-badge.whea-ram  { background: #3d2a1a; color: #ffaa66; }
    .hw-badge.whea-pcie { background: #2a1a3d; color: #aa99ff; }
    .hw-badge.gpu-tdr   { background: #1a2a3d; color: #66aaff; }
    .hw-badge.thermal   { background: #3d2a1a; color: #ff9944; }

    .warn-badge.ram-exhaustion { background: #3d1a1a; color: var(--red); }
    .warn-badge.cpu-throttling { background: #3d2e1a; color: var(--yellow); }
    .warn-badge.disk-full      { background: #3d2a1a; color: #ff9966; }
    .warn-badge.disk-slow      { background: #2a2a3d; color: #99aaff; }
    /* v1.5 : badge burst pour Event 51 cluster massif (>50 events) */
    .burst-badge {
        display: inline-block;
        margin-left: 6px;
        padding: 1px 6px;
        border-radius: 3px;
        font-size: 10px;
        font-weight: 700;
        background: var(--red);
        color: white;
        letter-spacing: 0.5px;
        vertical-align: middle;
    }
    .warn-detail               { color: var(--text-muted); font-size: 11px; margin-right: 8px; }
    .kpi-warning               { color: var(--yellow); font-weight: 700; }

    /* v1.6 : bloc CPU dans panel Materiel */
    .cpu-info-block { padding: 4px 0; }
    .cpu-info-block .cpu-name {
        font-family: ui-monospace, 'SF Mono', Consolas, monospace;
        font-size: 12px;
        color: var(--text);
        margin-bottom: 4px;
        word-break: break-word;
    }
    .cpu-info-block .cpu-meta {
        font-size: 11px;
        color: var(--text-muted);
        display: flex;
        align-items: center;
        gap: 8px;
        flex-wrap: wrap;
    }
    .cpu-badge-inline {
        display: inline-block;
        padding: 1px 6px;
        border-radius: 3px;
        font-size: 10px;
        font-weight: 600;
    }
    .cpu-badge-inline.ok      { background: rgba(34,197,94,0.15);  color: var(--green); }
    .cpu-badge-inline.warning { background: rgba(234,179,8,0.15);  color: var(--yellow); }
    .cpu-badge-inline.danger  { background: rgba(239,68,68,0.15);  color: var(--red); }

    /* v1.6 : bandeau verdict global (4 niveaux) */
    .verdict-banner {
        display: flex;
        align-items: center;
        gap: 10px;
        padding: 10px 14px;
        margin: 0 0 12px 0;
        border-radius: 6px;
        border-left: 4px solid;
        font-size: 13px;
    }
    .verdict-banner.sain     { background: rgba(34,197,94,0.08);  border-color: var(--green);  color: var(--text); }
    .verdict-banner.watch    { background: rgba(234,179,8,0.08);  border-color: var(--yellow); color: var(--text); }
    .verdict-banner.incident { background: rgba(251,146,60,0.10); border-color: #fb923c;       color: var(--text); }
    .verdict-banner.critical { background: rgba(239,68,68,0.10);  border-color: var(--red);    color: var(--text); }
    .verdict-banner .v-icon    { font-size: 18px; }
    .verdict-banner .v-label   { font-weight: 700; margin-right: 6px; }
    .verdict-banner .v-reasons { color: var(--text-muted); font-size: 12px; }

    /* v1.6 : section Signaux croises (correlations temporelles) */
    .correlations-block {
        margin: 10px 0;
        padding: 10px 12px;
        background: rgba(139,92,246,0.06);
        border-left: 3px solid #8b5cf6;
        border-radius: 4px;
    }
    .correlations-block h5 {
        margin: 0 0 6px 0;
        font-size: 11px;
        color: #8b5cf6;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        font-weight: 700;
    }
    .correlation-item {
        display: flex;
        align-items: flex-start;
        gap: 8px;
        padding: 4px 0;
        font-size: 12px;
        color: var(--text);
    }
    .correlation-item .c-icon { font-size: 14px; flex-shrink: 0; }
    .correlation-item .c-severity { font-weight: 700; margin-right: 4px; }
    .correlation-item.crit  .c-severity { color: var(--red); }
    .correlation-item.warn  .c-severity { color: var(--yellow); }
    .correlation-item .c-detail { color: var(--text-muted); font-size: 11px; }

    .uptime-badge, .cpu-badge, .conn-badge, .chassis-badge {
        display: inline-block;
        padding: 2px 6px;
        border-radius: 4px;
        font-size: 10px;
        font-weight: 600;
    }
    .uptime-ok      { background: #1a3d1a; color: var(--green); }
    .uptime-warning { background: #3d3d1a; color: var(--orange); }
    .uptime-danger  { background: #3d1a1a; color: var(--red); }
    .cpu-recent       { background: #1a3d1a; color: var(--green); }
    .cpu-vieillissant { background: #3d3d1a; color: var(--orange); }
    .cpu-ancien       { background: #3d1a1a; color: var(--red); }
    .cpu-inconnu      { background: #2a2a3a; color: var(--text-muted); }
    .conn-ethernet   { background: #1a3d2a; color: var(--green); }
    .conn-wifi       { background: #1a2a3d; color: #6699ff; }
    .conn-autre      { background: #2a2a3a; color: var(--text-dim); }
    .conn-deconnecte { background: #3d1a1a; color: var(--red); }
    /* v5.8 : chassis badges */
    .chassis-laptop  { background: #1a2a3d; color: #6699ff; }
    .chassis-desktop { background: #2a1a3d; color: var(--purple); }
    .chassis-aio     { background: #3d2a1a; color: var(--orange); }
    .chassis-autre   { background: #2a2a3a; color: var(--text-muted); }

    [data-theme="light"] .uptime-ok      { background: #d5f2e9; color: var(--green); }
    [data-theme="light"] .uptime-warning { background: #fdf3dc; color: var(--orange); }
    [data-theme="light"] .uptime-danger  { background: #fcdfe3; color: var(--red); }
    [data-theme="light"] .cpu-recent       { background: #d5f2e9; color: var(--green); }
    [data-theme="light"] .cpu-vieillissant { background: #fdf3dc; color: var(--orange); }
    [data-theme="light"] .cpu-ancien       { background: #fcdfe3; color: var(--red); }
    [data-theme="light"] .cpu-inconnu      { background: #e6e9f2; color: var(--text-muted); }
    [data-theme="light"] .conn-ethernet   { background: #d5f2e9; color: var(--green); }
    [data-theme="light"] .conn-wifi       { background: #dde6f7; color: #3a6cbf; }
    [data-theme="light"] .conn-autre      { background: #e6e9f2; color: var(--text-dim); }
    [data-theme="light"] .conn-deconnecte { background: #fcdfe3; color: var(--red); }
    [data-theme="light"] .chassis-laptop  { background: #dde6f7; color: #3a6cbf; }
    [data-theme="light"] .chassis-desktop { background: #f3e9fb; color: var(--purple); }
    [data-theme="light"] .chassis-aio     { background: #fdf0dc; color: var(--orange); }
    [data-theme="light"] .chassis-autre   { background: #e6e9f2; color: var(--text-muted); }

    /* ===== DISK BARS ===== */
    .disk-bar-wrap { display: flex; align-items: center; gap: 6px; margin: 4px 0; }
    .disk-drive    { font-weight: 600; color: var(--text-dim); min-width: 24px; }
    .disk-bar      { flex: 1; height: 8px; background: var(--border); border-radius: 4px; overflow: hidden; }
    .disk-bar-fill { height: 100%; border-radius: 4px; transition: width 0.3s; }
    .disk-bar-fill.ok      { background: var(--green); }
    .disk-bar-fill.warning { background: var(--orange); }
    .disk-bar-fill.danger  { background: var(--red); }
    .disk-info-text { font-size: 10px; color: var(--text-muted); min-width: 80px; text-align: right; }

    /* ===== FRESHNESS (derniere activite) ===== */
    .freshness-ok       { color: var(--green);  font-size: 11px; }
    .freshness-warning  { color: var(--orange); font-size: 11px; }
    .freshness-danger   { color: var(--red);    font-size: 11px; }

    /* ===== BOOT ===== */
    .boot-badge { display: inline-block; padding: 2px 6px; border-radius: 4px; font-size: 11px; font-weight: 700; }
    .boot-ok      { background: #1a3d1a; color: var(--green); }
    .boot-warning { background: #3d3d1a; color: var(--orange); }
    .boot-danger  { background: #3d1a1a; color: var(--red); }
    .boot-date    { font-size: 11px; color: var(--text-faint); margin-top: 2px; }
    [data-theme="light"] .boot-ok      { background: #d5f2e9; color: var(--green); }
    [data-theme="light"] .boot-warning { background: #fdf3dc; color: var(--orange); }
    [data-theme="light"] .boot-danger  { background: #fcdfe3; color: var(--red); }

    /* ===== CRASHERS (item) ===== */
    .crasher-item {
        display: flex;
        align-items: center;
        gap: 6px;
        padding: 3px 0;
        border-bottom: 1px solid var(--bg-main);
    }
    .crasher-item:last-child { border-bottom: none; }
    .crasher-name  { color: var(--yellow); font-size: 11px; }
    .crasher-count { color: var(--red); font-weight: 700; font-size: 11px; }

    /* ===== DRILL-DOWN ===== */
    .row-main td:first-child { cursor: pointer; user-select: none; }
    .row-main td:first-child:hover { color: var(--accent); }
    .toggle-icon { display: inline-block; margin-right: 6px; font-size: 10px; color: var(--text-muted); transition: transform 0.2s; }
    .row-main.open .toggle-icon { transform: rotate(90deg); }
    .row-detail { display: none; background: var(--bg-elevated) !important; }
    .row-detail.visible { display: table-row; }
    .row-detail td { padding: 0 !important; border-bottom: 1px solid var(--border) !important; white-space: normal !important; }
    .detail-box { padding: 14px 20px 16px 36px; display: flex; gap: 24px; flex-wrap: wrap; }
    .detail-section { flex: 1; min-width: 180px; padding-left: 12px; border-left: 3px solid var(--border); }
    .detail-section.sec-crash   { border-left-color: var(--red); }
    .detail-section.sec-boot    { border-left-color: var(--orange); }
    .detail-section.sec-disk    { border-left-color: var(--green); }
    .detail-section.sec-crasher { border-left-color: var(--yellow); }
    .detail-section.sec-hw      { border-left-color: var(--purple); }
    .detail-section.sec-perf    { border-left-color: var(--cyan); }
    /* v5.3 / v5.4 */
    .detail-section.sec-battery { border-left-color: var(--green); }
    .detail-section.sec-edr     { border-left-color: var(--purple); }
    .detail-section.sec-bootperf{ border-left-color: var(--orange); }
    .detail-section.sec-smart   { border-left-color: var(--cyan); }
    .detail-section.sec-monitors { border-left-color: var(--purple); }
    /* v1.8 */
    .detail-section.sec-ram      { border-left-color: var(--cyan); }
    .detail-section.sec-gpu      { border-left-color: var(--orange); }
    .detail-section.sec-throttle { border-left-color: var(--red); }

    /* v1.8 : RAM section */
    .ram-summary {
        padding: 8px 0 10px 0;
        border-bottom: 1px solid var(--border);
        margin-bottom: 8px;
    }
    .ram-summary-total {
        display: flex;
        align-items: baseline;
        gap: 6px;
        margin-bottom: 6px;
    }
    .ram-big {
        font-size: 24px;
        font-weight: 700;
        color: var(--cyan);
    }
    .ram-unit { font-size: 12px; color: var(--text-muted); }
    .ram-summary-slots {
        font-size: 11px;
        color: var(--text-muted);
        margin-left: 8px;
    }
    .ram-summary-meta {
        display: flex;
        gap: 10px;
        align-items: center;
        font-size: 11px;
        color: var(--text-muted);
        flex-wrap: wrap;
    }
    .ram-meta-ok { color: var(--green); font-weight: 600; }
    .ram-badge {
        display: inline-block;
        padding: 2px 8px;
        border-radius: 10px;
        font-size: 10px;
        font-weight: 600;
    }
    .ram-badge.upgrade-ok { background: rgba(34,197,94,0.15);  color: var(--green); }
    .ram-badge.upgrade-no { background: rgba(239,68,68,0.12);  color: var(--red); }
    .ram-modules {
        display: flex;
        flex-direction: column;
        gap: 3px;
    }
    .ram-module-row {
        display: flex;
        gap: 10px;
        align-items: center;
        padding: 4px 6px;
        background: var(--bg-main);
        border-radius: 4px;
        font-size: 11px;
    }
    .ram-slot-name {
        font-family: ui-monospace, 'SF Mono', Consolas, monospace;
        color: var(--text-muted);
        min-width: 70px;
    }
    .ram-capacity {
        font-weight: 700;
        color: var(--text);
        min-width: 50px;
    }
    .ram-module-meta {
        color: var(--text-muted);
        flex: 1;
        font-size: 10px;
    }

    /* v1.8 : GPU section */
    .gpu-row {
        padding: 6px 0;
        border-bottom: 1px solid var(--border);
    }
    .gpu-row:last-child { border-bottom: none; }
    .gpu-name {
        font-family: ui-monospace, 'SF Mono', Consolas, monospace;
        font-size: 12px;
        color: var(--text);
        margin-bottom: 2px;
    }
    .gpu-meta {
        font-size: 11px;
        display: flex;
        gap: 6px;
        align-items: center;
    }
    .gpu-driver-ok  { color: var(--text-muted); }
    .gpu-driver-old { color: var(--yellow); }
    .gpu-old-tag {
        background: rgba(234,179,8,0.15);
        color: var(--yellow);
        padding: 1px 6px;
        border-radius: 3px;
        font-size: 9px;
        font-weight: 600;
        text-transform: uppercase;
    }

    /* v1.8 : CPU Throttling section */
    .throttle-summary {
        padding: 6px 10px;
        border-radius: 4px;
        font-size: 11px;
        font-weight: 600;
        margin-bottom: 8px;
    }
    .throttle-sev-info     { background: rgba(59,130,246,0.10); color: var(--cyan);   }
    .throttle-sev-watch    { background: rgba(234,179,8,0.12);  color: var(--yellow); border-left: 3px solid var(--yellow); padding-left: 8px; }
    .throttle-sev-critical { background: rgba(239,68,68,0.12);  color: var(--red);    border-left: 3px solid var(--red);    padding-left: 8px; }
    .throttle-row {
        display: flex;
        gap: 10px;
        align-items: center;
        padding: 4px 0;
        font-size: 11px;
    }
    .throttle-day {
        font-family: ui-monospace, 'SF Mono', Consolas, monospace;
        color: var(--text);
        min-width: 90px;
    }
    .throttle-type {
        color: var(--text-muted);
        flex: 1;
    }
    .throttle-count {
        color: var(--red);
        font-weight: 600;
        font-size: 10px;
    }
    .throttle-more {
        font-size: 10px;
        color: var(--text-faint);
        font-style: italic;
        padding: 4px 0 0 0;
    }

    /* Pastilles indicateurs compacts (colonne tableau)
       v5.7 : gap 5px (etait 3px) pour mieux separer les pastilles.
       Avant elles se collaient et donnaient un bloc illisible. */
    .indicator-row {
        display: inline-flex;
        gap: 5px;
        align-items: center;
        padding: 2px 4px;
    }
    .indicator-dot {
        display: inline-block; width: 15px; height: 15px; border-radius: 50%;
        font-size: 9px; font-weight: 700; text-align: center; line-height: 15px;
        color: #fff; cursor: help; user-select: none;
        flex-shrink: 0;
    }
    .indicator-dot.ok   { background: var(--green); }
    .indicator-dot.warn { background: var(--orange); }
    .indicator-dot.ko   { background: var(--red); }
    .indicator-dot.na   { background: var(--border); color: var(--text-muted); }

    /* Drill-down batterie : barre de sante */
    .batt-bar-wrap { display: flex; align-items: center; gap: 8px; margin: 8px 0; }
    .batt-bar     { flex: 1; height: 12px; background: var(--bg-main); border-radius: 6px; overflow: hidden; border: 1px solid var(--border); }
    .batt-bar-fill { height: 100%; transition: width 0.3s; }
    .batt-bar-fill.good    { background: linear-gradient(90deg, var(--green), #3bb891); }
    .batt-bar-fill.warning { background: linear-gradient(90deg, var(--orange), #ff8c00); }
    .batt-bar-fill.danger  { background: linear-gradient(90deg, var(--red), #d03030); }
    .batt-meta    { font-size: 11px; color: var(--text-muted); }
    .batt-meta strong { color: var(--text-dim); }

    /* Boot Perf : breakdown des phases */
    .bootperf-phases { display: flex; flex-direction: column; gap: 6px; margin: 8px 0; }
    .bootperf-phase  { display: flex; align-items: center; gap: 8px; font-size: 11px; }
    .bootperf-label  { flex: 0 0 140px; color: var(--text-muted); }
    .bootperf-bar    { flex: 1; height: 8px; background: var(--bg-main); border-radius: 4px; overflow: hidden; border: 1px solid var(--border); }
    .bootperf-bar-fill { height: 100%; background: var(--orange); }
    .bootperf-bar-fill.slow  { background: var(--red); }
    .bootperf-bar-fill.fast  { background: var(--green); }
    .bootperf-value  { flex: 0 0 75px; text-align: right; font-weight: 700; color: var(--text-dim); }
    .bootperf-meta   { font-size: 10px; color: var(--text-faint); margin-top: 6px; }
    .bootperf-row    { display: flex; gap: 12px; padding: 4px 0; border-bottom: 1px solid var(--bg-main); font-size: 11px; }
    .bootperf-row:last-child { border-bottom: none; }
    .bootperf-row .date { color: var(--text-faint); flex: 1; }
    .bootperf-row .total { color: var(--text-dim); font-weight: 700; }
    .bootperf-row.slow .total { color: var(--red); }

    /* SMART : carte par disque */
    .smart-disk { border: 1px solid var(--border); border-radius: 6px; padding: 8px 10px; margin-bottom: 8px; background: var(--bg-main); }
    .smart-disk-head { display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 6px; }
    .smart-disk-name { font-weight: 700; color: var(--text-dim); font-size: 12px; }
    .smart-disk-sub  { font-size: 10px; color: var(--text-faint); margin-top: 2px; }
    .smart-grid      { display: grid; grid-template-columns: 1fr 1fr; gap: 4px 12px; font-size: 10.5px; }
    .smart-kv        { display: flex; justify-content: space-between; }
    .smart-kv .k     { color: var(--text-muted); }
    .smart-kv .v     { color: var(--text-dim); font-weight: 600; }
    .smart-kv .v.ok   { color: var(--green); }
    .smart-kv .v.warn { color: var(--orange); }
    .smart-kv .v.ko   { color: var(--red); }
    .smart-kv .v.na   { color: var(--text-ghost); font-style: italic; font-weight: normal; }
    .smart-health-badge { font-size: 10px; padding: 2px 6px; border-radius: 3px; font-weight: 700; }
    .smart-health-badge.healthy   { background: var(--bg-success); color: var(--green); }
    .smart-health-badge.warning   { background: var(--bg-warning); color: var(--orange); }
    .smart-health-badge.unhealthy { background: var(--bg-danger); color: var(--red); }
    .smart-alerts    { margin-top: 6px; display: flex; gap: 4px; flex-wrap: wrap; }
    .smart-alert-chip { font-size: 9.5px; padding: 2px 5px; border-radius: 3px; background: var(--bg-danger); color: var(--red); font-weight: 600; }

    /* ===== MONITORS (v5.7) : inventaire ecrans externes =====
       Chaque moniteur = une "carte" fine avec nom/fab/serial/age.
       Un badge vert "actif" quand branche, gris quand plus branche.
       Un badge orange "age X ans" quand ecran ancien. */
    .monitor-card {
        border: 1px solid var(--border);
        border-radius: 6px;
        padding: 8px 10px;
        margin-bottom: 6px;
        background: var(--bg-main);
    }
    .monitor-card.old { border-left: 3px solid var(--orange); }
    .monitor-head {
        display: flex;
        justify-content: space-between;
        align-items: baseline;
        gap: 10px;
        margin-bottom: 4px;
    }
    .monitor-name {
        font-weight: 700;
        color: var(--text-dim);
        font-size: 12px;
        flex: 1;
        min-width: 0;
    }
    .monitor-manuf {
        font-size: 11px;
        color: var(--text-muted);
        font-weight: 500;
    }
    .monitor-badges { display: flex; gap: 4px; flex-shrink: 0; }
    .mon-badge {
        font-size: 9.5px;
        padding: 2px 6px;
        border-radius: 3px;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.3px;
    }
    .mon-badge.active   { background: var(--bg-success); color: var(--green); }
    .mon-badge.inactive { background: var(--border);     color: var(--text-muted); }
    .mon-badge.old      { background: var(--bg-warning); color: var(--orange); }
    .monitor-meta {
        font-size: 10.5px;
        color: var(--text-muted);
        display: flex;
        gap: 12px;
        flex-wrap: wrap;
    }
    .monitor-meta strong { color: var(--text-dim); }

    /* ===== INVENTAIRE MONITORS GLOBAL (panneau bas de page) ===== */
    .monitor-inventory {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
        gap: 10px;
    }
    .monitor-inv-tile {
        background: var(--bg-main);
        border: 1px solid var(--border);
        border-radius: 6px;
        padding: 10px 12px;
    }
    .monitor-inv-value {
        font-size: 22px;
        font-weight: 700;
        color: var(--text);
    }
    .monitor-inv-label {
        font-size: 10.5px;
        color: var(--text-muted);
        text-transform: uppercase;
        letter-spacing: 0.5px;
        margin-top: 2px;
    }
    .monitor-inv-sub {
        font-size: 10.5px;
        color: var(--text-dim);
        margin-top: 4px;
        line-height: 1.5;
    }
    .monitor-inv-sub .chip {
        display: inline-block;
        background: var(--bg-panel);
        padding: 2px 6px;
        border-radius: 3px;
        margin: 2px 3px 2px 0;
        font-size: 10px;
    }

    .detail-section h4 { font-size: 11px; text-transform: uppercase; letter-spacing: 0.5px; color: var(--text-muted); margin-bottom: 8px; }
    .detail-item { font-size: 12px; padding: 4px 0; border-bottom: 1px solid var(--bg-main); display: flex; justify-content: space-between; gap: 16px; align-items: flex-start; flex-wrap: wrap; }
    .detail-item:last-child { border-bottom: none; }
    .detail-date  { color: var(--text-faint); white-space: nowrap; font-size: 11px; }
    .detail-ago   { color: var(--text-muted); font-size: 10px; font-style: italic; }
    .detail-info  { color: var(--text-dim); }
    .detail-dur   { color: var(--orange); font-weight: 700; white-space: nowrap; }
    .detail-empty { color: var(--text-ghost); font-style: italic; font-size: 12px; }

    /* ===== ONGLETS DRILL-DOWN (v5.5) =====
       Les 10 sections de detail sont regroupees en 5 onglets thematiques :
       Vue d'ensemble / Stabilite / Demarrage / Materiel / Securite.
       Evite le "mur de cartes" ou tout avait le meme poids visuel. */
    .detail-tabs {
        display: flex;
        gap: 4px;
        border-bottom: 1px solid var(--border);
        padding: 0 20px 0 36px;
        margin: 8px 0 0 0;
        overflow-x: auto;
        scrollbar-width: thin;
    }
    .detail-tab {
        padding: 8px 14px;
        cursor: pointer;
        font-size: 11px;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        color: var(--text-muted);
        border-bottom: 2px solid transparent;
        margin-bottom: -1px;
        transition: color 0.15s, border-color 0.15s;
        white-space: nowrap;
        display: flex;
        align-items: center;
        gap: 6px;
        background: none;
        border-left: none;
        border-right: none;
        border-top: none;
    }
    .detail-tab:hover { color: var(--text-dim); }
    .detail-tab.active {
        color: var(--accent);
        border-bottom-color: var(--accent);
    }
    /* Badge de compteur sur les onglets (nb d'alertes dans l'onglet) */
    .detail-tab .tab-badge {
        font-size: 10px;
        background: var(--red);
        color: #fff;
        padding: 1px 5px;
        border-radius: 8px;
        font-weight: 700;
        min-width: 14px;
        text-align: center;
    }
    .detail-tab .tab-badge.quiet {
        background: var(--border);
        color: var(--text-muted);
    }
    .detail-tab-panel {
        display: none;
        padding: 14px 20px 16px 36px;
    }
    .detail-tab-panel.active { display: block; }
    .detail-tab-panel .detail-box {
        padding: 0;
    }

    /* Onglet "Vue d'ensemble" : liste compacte des alertes actives */
    .overview-list { display: flex; flex-direction: column; gap: 6px; }
    .overview-alert {
        display: flex;
        align-items: center;
        gap: 10px;
        padding: 8px 12px;
        border-radius: 6px;
        background: var(--bg-main);
        border-left: 3px solid var(--orange);
        font-size: 12px;
    }
    .overview-alert.critical { border-left-color: var(--red); background: rgba(255, 107, 107, 0.08); }
    .overview-alert.warning  { border-left-color: var(--orange); }
    .overview-alert.info     { border-left-color: var(--cyan); }
    .overview-alert .ov-icon { font-size: 16px; width: 24px; text-align: center; flex-shrink: 0; }
    .overview-alert .ov-text { flex: 1; color: var(--text-dim); }
    .overview-alert .ov-text strong { color: var(--text); }
    .overview-alert .ov-meta { color: var(--text-faint); font-size: 11px; white-space: nowrap; }
    .overview-empty {
        padding: 18px 12px;
        text-align: center;
        color: var(--green);
        font-size: 12.5px;
        background: rgba(78, 204, 163, 0.06);
        border-radius: 6px;
        border: 1px dashed rgba(78, 204, 163, 0.3);
    }
    .overview-empty strong { display: block; font-size: 14px; margin-bottom: 4px; }

    /* ===== MODE COMPACT DU TABLEAU (v5.5) =====
       Par defaut on cache les colonnes "avancees" (Crash/BSOD/HW/Disque/Perf)
       pour alleger le tableau. Un bouton toolbar permet de tout afficher. */
    body:not(.advanced-cols) .col-advanced { display: none !important; }

    /* ===== PANNEAU TOP CRASHERS PARC ===== */
    .global-panel {
        margin-top: 24px;
        background: var(--bg-panel);
        border-radius: 10px;
        padding: 20px;
        border: 1px solid var(--border);
    }
    .global-panel h3 {
        font-size: 14px;
        font-weight: 600;
        color: var(--text-dim);
        text-transform: uppercase;
        letter-spacing: 0.5px;
        margin-bottom: 14px;
    }
    .global-crashers-grid {
        display: grid;
        grid-template-columns: repeat(auto-fill, minmax(280px, 1fr));
        gap: 10px;
    }
    .global-crasher-row {
        display: flex;
        align-items: center;
        gap: 10px;
        padding: 8px 12px;
        background: var(--bg-main);
        border-radius: 6px;
        border-left: 3px solid var(--yellow);
    }
    .global-crasher-name { color: var(--yellow); font-size: 12px; font-weight: 600; flex: 1; overflow: hidden; text-overflow: ellipsis; }
    .global-crasher-stats { font-size: 11px; color: var(--text-muted); white-space: nowrap; }
    .global-crasher-total { color: var(--red); font-weight: 700; }

    /* v1.7 : 3 sections Top Crashers global (Local / Reparti / Bruit) */
    .crasher-section { margin-bottom: 18px; }
    .crasher-section-title {
        margin: 0 0 10px 0;
        font-size: 13px;
        color: var(--text);
        display: flex;
        align-items: center;
        gap: 8px;
        font-weight: 700;
    }
    .crasher-section-count {
        background: var(--bg-alt);
        color: var(--text-muted);
        padding: 1px 8px;
        border-radius: 10px;
        font-size: 11px;
        font-weight: 600;
    }
    .crasher-section-hint {
        font-size: 11px;
        color: var(--text-faint);
        font-weight: 400;
        margin-left: 4px;
    }
    .crasher-section-empty {
        padding: 8px 12px;
        background: var(--bg-main);
        border-radius: 6px;
        font-size: 11px;
        color: var(--text-faint);
        font-style: italic;
    }
    .crasher-collapsible { cursor: pointer; user-select: none; }
    .crasher-toggle { font-size: 11px; color: var(--text-muted); margin-right: 2px; }
    .crasher-noise-body { margin-top: 4px; }

    /* Distinction visuelle par niveau (couleur bordure gauche) */
    .global-crasher-row.local  { border-left-color: var(--red); }
    .global-crasher-row.spread { border-left-color: var(--yellow); }
    .global-crasher-row.noise  { border-left-color: var(--text-faint); opacity: 0.75; }

    /* Badge score */
    .crasher-score {
        display: inline-block;
        min-width: 32px;
        text-align: center;
        padding: 2px 6px;
        border-radius: 3px;
        font-family: ui-monospace, 'SF Mono', Consolas, monospace;
        font-size: 11px;
        font-weight: 700;
    }
    .crasher-score-local  { background: rgba(239,68,68,0.18); color: var(--red); }
    .crasher-score-spread { background: rgba(234,179,8,0.18); color: var(--yellow); }
    .crasher-score-noise  { background: var(--bg-alt);       color: var(--text-muted); }

    /* Cadenas "forced to noise" via blacklist soft */
    .crasher-forced { font-size: 11px; opacity: 0.6; margin-right: -4px; }

    /* v1.7 : drill-down PC - bruiteurs grises */
    .crasher-soft-noise { opacity: 0.55; }
    .crasher-noise-tag {
        display: inline-block;
        font-size: 9px;
        background: var(--bg-alt);
        color: var(--text-faint);
        padding: 1px 5px;
        border-radius: 3px;
        margin-left: 4px;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }

    /* ===== BOOT BREAKDOWN (v5.2) ===== */
    .boot-breakdown { display: flex; flex-direction: column; gap: 10px; }
    .boot-breakdown-bar {
        display: flex;
        height: 28px;
        background: var(--bg-main);
        border-radius: 6px;
        overflow: hidden;
        border: 1px solid var(--border);
    }
    .boot-breakdown-segment {
        display: flex;
        align-items: center;
        justify-content: center;
        color: white;
        font-weight: 700;
        font-size: 11px;
        min-width: 0;
        overflow: hidden;
        white-space: nowrap;
        transition: all 0.3s;
    }
    .boot-breakdown-segment.cold    { background: #3a6cbf; }
    .boot-breakdown-segment.fast    { background: var(--yellow); color: #2a1a00; }
    .boot-breakdown-segment.resume  { background: var(--purple); }
    .boot-breakdown-segment.unknown { background: var(--text-faint); }
    .boot-breakdown-legend {
        display: flex;
        gap: 16px;
        font-size: 11px;
        color: var(--text-dim);
        flex-wrap: wrap;
    }
    .boot-breakdown-legend-item { display: flex; align-items: center; gap: 6px; }
    .boot-breakdown-legend-dot { width: 10px; height: 10px; border-radius: 2px; }
    .boot-breakdown-empty { color: var(--text-ghost); font-style: italic; padding: 10px; text-align: center; }
    .boot-breakdown-note {
        font-size: 11px;
        color: var(--text-muted);
        background: var(--bg-main);
        border-left: 3px solid var(--accent);
        padding: 8px 12px;
        border-radius: 4px;
    }

    /* ===== LEGEND ===== */
    .legend { display: flex; gap: 20px; padding: 12px 20px; font-size: 11px; color: var(--text-faint); flex-wrap: wrap; border-top: 1px solid var(--border); }
    .legend-item { display: flex; align-items: center; gap: 6px; }
    .legend-dot { width: 10px; height: 10px; border-radius: 3px; }
    .dot-danger   { background: #ff6b6b33; border: 1px solid var(--red); }
    .dot-warning  { background: #ffa50233; border: 1px solid var(--orange); }
    .dot-hardware { background: #a855f733; border: 1px solid var(--purple); }
    .dot-ok       { background: #4ecca333; border: 1px solid var(--green); }
</style>
</head>
<body>

<div class="header">
    <h1><span>$titleHtml</span> &mdash; $subtitleHtml</h1>
    <div class="header-right">
        <button class="theme-toggle" id="themeToggle" onclick="toggleTheme()" title="Basculer clair/sombre">&#9788;</button>
        <div class="timestamp">G&eacute;n&eacute;r&eacute; le : $($now.ToString('dd/MM/yyyy HH:mm'))</div>
    </div>
</div>

<div class="toolbar">
    <span class="label">P&eacute;riode :</span>
    <button class="range-btn" onclick="setDays(1, this)">24h</button>
    <button class="range-btn" onclick="setDays(7, this)">7 jours</button>
    <button class="range-btn" onclick="setDays(15, this)">15 jours</button>
    <button class="range-btn active" onclick="setDays(30, this)">30 jours</button>

    <div class="divider"></div>

    <button class="filter-btn" id="maskHealthyBtn" onclick="toggleMaskHealthy()" title="Masquer les PC sans probleme">Masquer sains</button>
    <button class="filter-btn" id="advancedColsBtn" onclick="toggleAdvancedCols()" title="Afficher les colonnes techniques (Crash, BSOD, HW, Disque, Perf)">Vue d&eacute;taill&eacute;e</button>

    <div class="divider"></div>

    <label style="color: var(--text-muted); font-size: 12px;">Site :</label>
    <select id="siteFilter" onchange="filterChanged()"><option value="">Tous</option></select>
    <label style="color: var(--text-muted); font-size: 12px;">CPU :</label>
    <select id="cpuFilter" onchange="filterChanged()">
        <option value="">Tous</option>
        <option value="Recent">Recent</option>
        <option value="Vieillissant">Vieillissant</option>
        <option value="Ancien">Ancien</option>
        <option value="Inconnu">Inconnu</option>
    </select>

    <div class="search-wrap">
        <span class="search-icon">&#128269;</span>
        <input class="search-input" type="text" id="searchInput" placeholder="PC ou utilisateur..." oninput="filterChanged()">
    </div>
    <button class="export-btn" onclick="exportCSV()" title="Exporter la vue courante en CSV">&#8681; CSV</button>
</div>

<!-- v5.5 : Summary bar avec les 3 chiffres essentiels -->
<div class="summary-bar" id="summaryBar"></div>

<!-- v5.6 : barre horizontale de 4 boutons de groupes (compact)
     remplace les 4 lignes pleine largeur de v5.5 -->
<div class="kpi-group-bar" id="kpiGroupBar"></div>

<!-- Garde pour retrocompat : kpiGrid devient cache, les groupes ci-dessus le remplacent -->
<div class="kpi-grid" id="kpiGrid" style="display:none"></div>
<div id="kpiGroups" style="display:none"></div>

<div class="table-container">
    <div class="table-header">
        <div>
            <h2>D&eacute;tails par appareil</h2>
            <div class="active-filters" id="activeFilters"></div>
        </div>
        <span class="count" id="pcCount"></span>
    </div>
    <!-- v1.5 : pagination haut -->
    <div id="paginationTop"></div>
    <div class="table-scroll" id="deviceTable">
        <table>
            <thead id="tableHead"></thead>
            <tbody id="tableBody"></tbody>
        </table>
    </div>
    <!-- v1.5 : pagination bas -->
    <div id="paginationBottom"></div>
    <div class="legend">
        <div class="legend-item"><div class="legend-dot dot-danger"></div> Crash ou BSOD</div>
        <div class="legend-item"><div class="legend-dot dot-hardware"></div> Erreur mat&eacute;rielle</div>
        <div class="legend-item"><div class="legend-dot dot-warning"></div> Boot long / disque critique</div>
        <div class="legend-item"><div class="legend-dot dot-ok"></div> Pas de probl&egrave;me</div>
    </div>
</div>

<div class="global-panel">
    <h3>R&eacute;partition des d&eacute;marrages (parc)</h3>
    <div id="bootBreakdown" class="boot-breakdown"></div>
</div>

<div class="global-panel">
    <h3>Top Crashers parc global</h3>
    <div class="global-crashers-grid" id="globalCrashers"></div>
</div>

<!-- v5.7 : inventaire des moniteurs externes du parc -->
<div class="global-panel" id="monitorPanel" style="display:none">
    <h3>&Eacute;crans secondaires branch&eacute;s</h3>
    <div class="monitor-inventory" id="monitorInventory"></div>
</div>

<script>
// ===== DONNEES =====
var pcData        = $jsonEmbed;
var scoreWeights  = $weightsEmbed;
var showSite      = $showSiteJs;
var seuilBootLong    = $SeuilBootLong;
var seuilCrashRecent = $SeuilCrashRecent;
var seuilDiskAlert   = $SeuilDiskAlert;
var seuilDiskWarning = $SeuilDiskWarning;
var generatedAt   = new Date('$($now.ToString("yyyy-MM-ddTHH:mm:ss"))');

// ===== ETAT UI =====
var state = {
    days: 30,
    daysBtn: null,
    maskHealthy: $maskHealthyJs,
    siteFilter: '',
    cpuFilter: '',
    kpiFilter: null,     // 'offline', 'crash', 'bsod', 'hw', 'bootLong', 'diskAlert', 'oldCpu', 'crashRecent'
    sort: { col: 'score', dir: 'desc' },
    // Pagination : itemsPerPage peut valoir 20, 50, 100, ou 0 (tous)
    // Persistance via localStorage pour se souvenir du choix utilisateur
    itemsPerPage: (function() {
        try {
            var saved = localStorage.getItem('pcpulse_itemsPerPage');
            if (saved !== null) return parseInt(saved, 10);
        } catch (e) {}
        return 50;  // defaut
    })(),
    currentPage: 1
};

// ===== BUGCHECK MAPPING =====
// Stop codes BSOD les plus courants pour affichage symbolique
// Source : https://learn.microsoft.com/windows-hardware/drivers/debugger/bug-check-code-reference2
var BUGCHECKS = {
    '0x1':   'APC_INDEX_MISMATCH',
    '0xA':   'IRQL_NOT_LESS_OR_EQUAL',
    '0x1A':  'MEMORY_MANAGEMENT',
    '0x1E':  'KMODE_EXCEPTION_NOT_HANDLED',
    '0x3B':  'SYSTEM_SERVICE_EXCEPTION',
    '0x4E':  'PFN_LIST_CORRUPT',
    '0x50':  'PAGE_FAULT_IN_NONPAGED_AREA',
    '0x7A':  'KERNEL_DATA_INPAGE_ERROR',
    '0x7B':  'INACCESSIBLE_BOOT_DEVICE',
    '0x7E':  'SYSTEM_THREAD_EXCEPTION_NOT_HANDLED',
    '0x7F':  'UNEXPECTED_KERNEL_MODE_TRAP',
    '0x9F':  'DRIVER_POWER_STATE_FAILURE',
    '0xC1':  'SPECIAL_POOL_DETECTED_MEMORY_CORRUPTION',
    '0xC2':  'BAD_POOL_CALLER',
    '0xC4':  'DRIVER_VERIFIER_DETECTED_VIOLATION',
    '0xC5':  'DRIVER_CORRUPTED_EXPOOL',
    '0xD1':  'DRIVER_IRQL_NOT_LESS_OR_EQUAL',
    '0xDE':  'POOL_CORRUPTION_IN_FILE_AREA',
    '0xEF':  'CRITICAL_PROCESS_DIED',
    '0xF4':  'CRITICAL_OBJECT_TERMINATION',
    '0xF7':  'DRIVER_OVERRAN_STACK_BUFFER',
    '0x109': 'CRITICAL_STRUCTURE_CORRUPTION',
    '0x116': 'VIDEO_TDR_FAILURE',
    '0x124': 'WHEA_UNCORRECTABLE_ERROR',
    '0x133': 'DPC_WATCHDOG_VIOLATION',
    '0x139': 'KERNEL_SECURITY_CHECK_FAILURE',
    '0x154': 'UNEXPECTED_STORE_EXCEPTION',
    '0x1E1': 'VIDEO_DXGKRNL_FATAL_ERROR',
    '0xEF':  'CRITICAL_PROCESS_DIED'
};
function bugCheckName(stopCode) {
    if (!stopCode) return '';
    var key = stopCode.toUpperCase().replace(/^0X/, '0x');
    // Normalise : 0x0000007E -> 0x7E
    var num = key.replace(/^0x0+/, '0x');
    if (num === '0x') num = '0x0';
    return BUGCHECKS[num] || '';
}

// ===== UTILS =====
function parseDate(s) { return new Date(s.replace(' ', 'T')); }
function timeAgo(dateStr) {
    var d = parseDate(dateStr);
    var diffMs = generatedAt - d;
    var diffMin = Math.floor(diffMs / 60000);
    var diffH = Math.floor(diffMin / 60);
    var diffJ = Math.floor(diffH / 24);
    if (diffJ > 0) return 'il y a ' + diffJ + 'j';
    if (diffH > 0) return 'il y a ' + diffH + 'h';
    if (diffMin > 0) return 'il y a ' + diffMin + 'min';
    return 'a l instant';
}

// ============================================================
// v1.6 : VERDICT GLOBAL (4 niveaux)
// Regle : on remonte tous les criteres observes et on prend le
// niveau le plus grave. Chaque critere contribue a la liste de
// raisons affichee en metadata.
// ============================================================
// ============================================================
// v1.7 : DEBRUITAGE DES TOP CRASHERS (Signal vs Bruit)
// ============================================================
// Blacklist HARD : processus ecartes completement, invisibles partout.
// Reserve aux crashers "inactionnables par construction" dont le volume
// pourrit la vue sans apporter aucune info exploitable.
var CRASHER_BLACKLIST_HARD = [
    // Le vrai nom du binaire Microsoft a une typo interne : "searchINbing"
    // (et non "searchBing" comme on pourrait le croire). On garde les 2 variantes
    // par securite au cas ou Microsoft corrigerait la typo un jour.
    'microsoftsearchinbing.exe',  // typo reelle du binaire (cas terrain)
    'microsoftsearchbing.exe'     // orthographe logique (si jamais corrige)
];

// Blacklist SOFT : processus toujours affiches, mais forces en section
// "Bruit ambient" meme si leur score les aurait remontes plus haut.
// Garde un oeil dessus sans les laisser polluer le top.
var CRASHER_BLACKLIST_SOFT = [
    'dellosd.exe',                       // Dell On-Screen Display
    'shellexperiencehost.exe',           // UI shell Windows, redemarre seul
    'gamebar.exe',                       // Xbox Game Bar
    'asussystemanalysis.exe',            // utilitaire ASUS
    'asusverifyjwt.exe',                 // utilitaire ASUS
    'dell.techhub.diagnostics.subagent'  // Dell Tech Hub
];

function isBlacklistedHard(name) {
    if (!name) return false;
    var n = String(name).toLowerCase();
    for (var i = 0; i < CRASHER_BLACKLIST_HARD.length; i++) {
        if (n.indexOf(CRASHER_BLACKLIST_HARD[i]) >= 0) return true;
    }
    return false;
}
function isBlacklistedSoft(name) {
    if (!name) return false;
    var n = String(name).toLowerCase();
    for (var i = 0; i < CRASHER_BLACKLIST_SOFT.length; i++) {
        if (n.indexOf(CRASHER_BLACKLIST_SOFT[i]) >= 0) return true;
    }
    return false;
}

// Scoring : penalise la dispersion entre PC.
// score = (total/pcCount) / sqrt(pcCount) = moyenne_par_PC / sqrt(PC_impactes)
// Un processus concentre (8 crashs sur 1 PC) sort plus haut qu'un
// processus reparti (81 crashs sur 10 PC), qui est typiquement du bruit.
function computeCrasherScore(total, pcCount) {
    if (!pcCount || pcCount <= 0) return 0;
    var avgPerPc = total / pcCount;
    return avgPerPc / Math.sqrt(pcCount);
}

// Classification en 3 niveaux :
//   'local'   (score >= 3)  : concentre sur peu de PC, actionnable
//   'spread'  (2 <= score < 3) : reparti, possible bug app
//   'noise'   (score < 2)   : bruit ambient, pas actionnable
// Les entrees en blacklist SOFT sont forcees en 'noise' peu importe leur score.
function classifyCrasher(name, total, pcCount) {
    var score = computeCrasherScore(total, pcCount);
    if (isBlacklistedSoft(name)) return { level: 'noise', score: score, forced: true };
    if (score >= 3)              return { level: 'local',  score: score, forced: false };
    if (score >= 2)              return { level: 'spread', score: score, forced: false };
    return                              { level: 'noise',  score: score, forced: false };
}

function computeVerdict(p) {
    // Donnees sur lesquelles on raisonne
    var crashCount    = p.crashCount        || 0;
    var bsodCount     = p.bsodCount         || 0;
    var hardCrash     = (p.pc && p.pc.TotalHardCrash) ? p.pc.TotalHardCrash : 0;  // v1.6
    var wheaFatal     = (p.wheaFatal    || []).length;
    var wheaCorrUniq  = (p.wheaCorrected || []).length;
    var wheaCorrTot   = p.wheaCorrectedTotal || 0;
    var cpuCat        = (p.pc && p.pc.CPUAgeCategory) || 'Inconnu';
    var battPct       = (p.battery && p.battery.HasBattery === true && typeof p.battery.HealthPercent === 'number') ? p.battery.HealthPercent : null;
    var smartWearMax  = 0;
    if (p.pc && Array.isArray(p.pc.DiskHealth)) {
        p.pc.DiskHealth.forEach(function(d) {
            if (typeof d.WearPct === 'number' && d.WearPct > smartWearMax) smartWearMax = d.WearPct;
        });
    }
    // Burst I/O massif (>= 50 events en un cluster)
    var hasBurstMassive = (p.warnings || []).some(function(w) { return w.IsBurst === true; });
    var hasBurstAny     = (p.warnings || []).some(function(w) { return typeof w.Count === 'number' && w.Count > 1; });

    var reasons = [];
    var level   = 'sain';

    // --- NIVEAU CRITIQUE ---
    if (wheaFatal > 0) {
        reasons.push(wheaFatal + ' erreur(s) WHEA fatale(s)');
        level = 'critical';
    }
    if (hardCrash >= 5) {
        reasons.push(hardCrash + ' hard crashs/30j');
        level = 'critical';
    }
    if (battPct !== null && battPct < 50) {
        reasons.push('batterie ' + battPct + '%');
        level = 'critical';
    }
    if (smartWearMax > 80) {
        reasons.push('SSD wear ' + smartWearMax + '%');
        level = 'critical';
    }
    if (cpuCat === 'Ancien' && crashCount >= 1) {
        reasons.push('CPU ancien + crashs');
        level = 'critical';
    }

    // --- NIVEAU INCIDENT PROBABLE ---
    if (level !== 'critical') {
        if (hasBurstMassive) {
            reasons.push('burst I/O massif (>=50 events)');
            level = 'incident';
        }
        if (wheaCorrUniq > 20 || wheaCorrTot > 50) {
            reasons.push(wheaCorrUniq + ' signatures WHEA corrigees');
            level = 'incident';
        }
        if (bsodCount >= 1) {
            reasons.push(bsodCount + ' BSOD/30j');
            level = 'incident';
        }
        if (hardCrash >= 2 && hardCrash < 5) {
            reasons.push(hardCrash + ' hard crashs/30j');
            level = 'incident';
        }
    }

    // --- NIVEAU A SURVEILLER ---
    if (level !== 'critical' && level !== 'incident') {
        if (cpuCat === 'Vieillissant') {
            reasons.push('CPU vieillissant');
            level = 'watch';
        }
        if (battPct !== null && battPct >= 50 && battPct < 70) {
            reasons.push('batterie ' + battPct + '%');
            level = 'watch';
        }
        if (smartWearMax >= 50 && smartWearMax <= 80) {
            reasons.push('SSD wear ' + smartWearMax + '%');
            level = 'watch';
        }
        if (hasBurstAny && !hasBurstMassive) {
            reasons.push('burst I/O isole');
            level = 'watch';
        }
    }

    // Libelles pour affichage
    var labels = {
        sain:     { cls: 'sain',     icon: '&#129001;', label: 'Sain' },              // green circle
        watch:    { cls: 'watch',    icon: '&#129000;', label: 'A surveiller' },      // yellow circle
        incident: { cls: 'incident', icon: '&#128992;', label: 'Incident probable' }, // orange circle
        critical: { cls: 'critical', icon: '&#128308;', label: 'Critique' }           // red circle
    };

    var info = labels[level];
    return {
        level:   level,
        cls:     info.cls,
        icon:    info.icon,
        label:   info.label,
        reasons: reasons
    };
}

// ============================================================
// v1.6 : DETECTION DE PATTERNS TEMPORELS (fenetre 10 min)
// 5 patterns qui croisent differentes sources pour donner un
// diagnostic qu'une seule source ne revele pas.
// ============================================================
function detectCorrelations(p) {
    var CORR_WINDOW_MS = 10 * 60 * 1000;  // 10 min
    var findings = [];

    function tsOf(s) { return parseDate(s).getTime(); }

    // -- Pattern 1 : Burst I/O (>=50 events) -> Hard crash dans les 10 min
    var bursts = (p.warnings || []).filter(function(w) { return w.IsBurst === true; });
    var crashTimestamps = (p.crashes || []).map(function(c) { return { ts: tsOf(c.Timestamp), cause: c.CrashCause || '', type: c.Type }; });
    bursts.forEach(function(b) {
        var bTs = tsOf(b.Timestamp);
        var match = crashTimestamps.filter(function(c) {
            return c.ts >= bTs && (c.ts - bTs) <= CORR_WINDOW_MS;
        });
        if (match.length > 0) {
            findings.push({
                severity: 'crit',
                icon: '&#128190;',
                title: 'Probable panne disque',
                detail: 'Burst I/O massif (' + b.Count + ' events) le ' + b.Timestamp + ' suivi d un crash materiel dans les 10 min'
            });
        }
    });

    // -- Pattern 2 : WHEA Corrected PCIe -> crash systeme dans les 10 min
    var wheaPCIe = (p.wheaCorrected || []).filter(function(h) {
        return (h.ErrorSource || '').toUpperCase().indexOf('PCI') >= 0;
    });
    wheaPCIe.forEach(function(h) {
        var hTs = tsOf(h.LastSeen);
        var match = crashTimestamps.filter(function(c) {
            return Math.abs(c.ts - hTs) <= CORR_WINDOW_MS;
        });
        if (match.length > 0) {
            findings.push({
                severity: 'warn',
                icon: '&#128268;',
                title: 'Slot PCIe suspect',
                detail: 'WHEA PCIe (' + h.Count + ' occurrences) corrélée avec un crash dans les 10 min'
            });
        }
    });

    // -- Pattern 3 : Thermal event + BootTime > 2x moyenne historique
    var hasThermal = (p.thermal || []).length > 0;
    var avgBoot = (p.bootPerf && p.bootPerf.Stats && p.bootPerf.Stats.AvgBootTimeMs) ? p.bootPerf.Stats.AvgBootTimeMs : 0;
    var maxBoot = (p.bootPerf && p.bootPerf.Stats && p.bootPerf.Stats.MaxBootTimeMs) ? p.bootPerf.Stats.MaxBootTimeMs : 0;
    if (hasThermal && avgBoot > 0 && maxBoot > 2 * avgBoot) {
        findings.push({
            severity: 'warn',
            icon: '&#127777;',
            title: 'Refroidissement degrade',
            detail: 'Thermal event + pic de boot a ' + Math.round(maxBoot/1000) + 's (moyenne ' + Math.round(avgBoot/1000) + 's) : ventilo ou pate thermique a verifier'
        });
    }

    // -- Pattern 4 : BSOD >=2 avec meme BugCheck sur 7j
    var SEVEN_DAYS_MS = 7 * 24 * 3600 * 1000;
    var bsodByCode = {};
    (p.crashes || []).filter(function(c) { return c.Type === 'BSOD' && c.Detail; }).forEach(function(c) {
        var code = c.Detail;
        var ts = tsOf(c.Timestamp);
        if (!bsodByCode[code]) bsodByCode[code] = [];
        bsodByCode[code].push(ts);
    });
    Object.keys(bsodByCode).forEach(function(code) {
        var arr = bsodByCode[code].sort(function(a,b){ return b-a; });
        if (arr.length >= 2 && (arr[0] - arr[arr.length-1]) <= SEVEN_DAYS_MS) {
            var bugName = bugCheckName(code);
            findings.push({
                severity: 'warn',
                icon: '&#128165;',
                title: 'Crash recurrent',
                detail: arr.length + ' BSOD ' + code + (bugName ? ' (' + bugName + ')' : '') + ' en 7 jours'
            });
        }
    });

    // -- Pattern 5 : Hard crash >=2 en 24h
    var ONE_DAY_MS = 24 * 3600 * 1000;
    var hardCrashes = (p.crashes || []).filter(function(c) {
        return c.Type === 'Hard reset' && c.CrashCause && c.CrashCause !== 'UserForcedReset';
    }).map(function(c) { return tsOf(c.Timestamp); }).sort(function(a,b){ return b-a; });
    for (var i = 0; i < hardCrashes.length - 1; i++) {
        if ((hardCrashes[i] - hardCrashes[i+1]) <= ONE_DAY_MS) {
            findings.push({
                severity: 'warn',
                icon: '&#9889;',
                title: 'Instabilite marquee',
                detail: '>=2 hard crashs en moins de 24h : surveiller alim / thermique / drivers'
            });
            break;  // un seul match suffit
        }
    }

    return findings;
}
// Format humain de l'uptime a partir d'une valeur en jours (float).
// Regles :
//   < 1h       -> "Xmin"
//   1h a 24h   -> "Xh"
//   1j a 7j    -> "Xj Yh"   (ex: "1j 8h")
//   > 7j       -> "Xj"      (jours entiers, sans les heures)
function formatUptime(uptimeDays) {
    if (uptimeDays === null || uptimeDays === undefined) return 'N/A';
    var totalMinutes = Math.floor(uptimeDays * 24 * 60);
    if (totalMinutes < 60) return totalMinutes + 'min';
    var totalHours = Math.floor(totalMinutes / 60);
    if (totalHours < 24) return totalHours + 'h';
    var days = Math.floor(totalHours / 24);
    var remainingHours = totalHours - (days * 24);
    if (days < 7) return days + 'j ' + remainingHours + 'h';
    return days + 'j';
}
function freshnessClass(hoursAgo) {
    if (hoursAgo <= 2) return 'freshness-ok';
    if (hoursAgo <= 24) return 'freshness-warning';
    return 'freshness-danger';
}
function hoursSince(dateStr) {
    return (generatedAt - parseDate(dateStr)) / 3600000;
}

// ===== CALCUL SCORE SANTE =====
// Score compose a partir des poids de config. Plus c'est haut, plus c'est grave.
// v5.2 : on compte separement WHEA_Fatal, Thermal et GPU_TDR (plus de double
// comptage). Les WHEA corrected ne sont jamais comptees (c'est de la telemetrie).
function computeScore(p) {
    var s = 0;
    s += p.bsodCount          * scoreWeights.BSOD;
    s += p.wheaFatal.length   * scoreWeights.WHEA;
    s += p.crashCount         * scoreWeights.Crash;
    s += p.thermal.length     * scoreWeights.Thermal;
    s += p.gpuTDR.length      * scoreWeights.GPU_TDR;
    s += p.diskAlertCount     * scoreWeights.DiskAlert;
    s += p.bootLongCount      * scoreWeights.BootLong;
    if (p.pc.IsOffline)       s += scoreWeights.Offline;
    // v5.3 / v5.4 : nouvelles penalites (defensif si weights absents du config ancien)
    if (p.sentinelAlert)      s += (scoreWeights.SentinelDown || 5);
    if (p.batteryAlert)       s += (scoreWeights.Battery      || 1);
    if (p.bootPerfAlert)      s += (scoreWeights.BootPerfSlow || 1);
    if (p.diskSmartAlert)     s += (scoreWeights.DiskHealth   || 3);
    return s;
}
function scoreClass(s) {
    if (s === 0) return 'score-ok';
    if (s <= 5)  return 'score-warn';
    if (s <= 15) return 'score-danger';
    return 'score-critic';
}

// ===== INIT SITES DROPDOWN =====
function initSiteDropdown() {
    if (!showSite) {
        document.getElementById('siteFilter').style.display = 'none';
        // masquer aussi le label du site
        var labels = document.querySelectorAll('.toolbar label');
        for (var i = 0; i < labels.length; i++) {
            if (labels[i].textContent.indexOf('Site') !== -1) labels[i].style.display = 'none';
        }
        return;
    }
    var sites = {};
    pcData.forEach(function(pc) { if (pc.Site) sites[pc.Site] = true; });
    var sorted = Object.keys(sites).sort();
    var sel = document.getElementById('siteFilter');
    sorted.forEach(function(s) {
        var opt = document.createElement('option');
        opt.value = s;
        opt.textContent = s;
        sel.appendChild(opt);
    });
}

// ===== THEME =====
function toggleTheme() {
    var html = document.documentElement;
    var now = html.getAttribute('data-theme');
    var next = (now === 'dark') ? 'light' : 'dark';
    html.setAttribute('data-theme', next);
    document.getElementById('themeToggle').innerHTML = next === 'dark' ? '&#9788;' : '&#9789;';
    try { localStorage.setItem('pcmon-theme', next); } catch(e) {}
}
(function() {
    try {
        var saved = localStorage.getItem('pcmon-theme');
        if (saved) {
            document.documentElement.setAttribute('data-theme', saved);
            document.getElementById('themeToggle').innerHTML = saved === 'dark' ? '&#9788;' : '&#9789;';
        }
    } catch(e) {}
})();

// ===== FILTRES =====
// v1.5 : wrapper appele par search/site/cpu qui reset la page avant render
// (evite de se retrouver sur une page qui n'existe plus apres filtrage)
function filterChanged() {
    state.currentPage = 1;
    render();
}

function setDays(days, btn) {
    state.days = days;
    state.daysBtn = btn;
    state.currentPage = 1;  // v1.5 : retour page 1 sur changement de filtre
    document.querySelectorAll('.range-btn').forEach(function(b) { b.classList.remove('active'); });
    if (btn) btn.classList.add('active');
    render();
}
function toggleMaskHealthy() {
    state.maskHealthy = !state.maskHealthy;
    state.currentPage = 1;  // v1.5
    document.getElementById('maskHealthyBtn').classList.toggle('active', state.maskHealthy);
    render();
}

// v5.5 : bascule entre la vue compacte et la vue detaillee du tableau.
// Persiste le choix dans localStorage pour que l'utilisateur ne subisse
// pas un reset a chaque regeneration du dashboard.
function toggleAdvancedCols() {
    var on = document.body.classList.toggle('advanced-cols');
    document.getElementById('advancedColsBtn').classList.toggle('active', on);
    try { localStorage.setItem('pcmon-advcols', on ? '1' : '0'); } catch(e) {}
}

// v5.5 : changement d'onglet dans un drill-down PC. Le state est stocke
// en data-attribute sur le conteneur pour etre isole par ligne.
function selectDetailTab(idx, tab) {
    var root = document.getElementById('detail-' + idx);
    if (!root) return;
    var tabs = root.querySelectorAll('.detail-tab');
    var panels = root.querySelectorAll('.detail-tab-panel');
    for (var i = 0; i < tabs.length; i++) {
        tabs[i].classList.toggle('active', tabs[i].getAttribute('data-tab') === tab);
    }
    for (var j = 0; j < panels.length; j++) {
        panels[j].classList.toggle('active', panels[j].getAttribute('data-panel') === tab);
    }
}
function toggleKpiFilter(key) {
    state.kpiFilter = (state.kpiFilter === key) ? null : key;
    state.currentPage = 1;  // v1.5
    render();
}
function clearKpiFilter() {
    state.kpiFilter = null;
    state.currentPage = 1;  // v1.5
    render();
}

// ===== ENRICHISSEMENT DES PC =====
// Calcule pour chaque PC les compteurs filtres par la periode + le score
function enrichPC(pc, cutoff) {
    var crashes     = (pc.Crashes || []).filter(function(c) { return parseDate(c.Timestamp) >= cutoff; });
    // v5.7 : on separe la LISTE filtree par periode (pour les compteurs et bootsLongs)
    // de la DERNIERE info boot connue (toujours affichee, meme si hors periode).
    // Avant : si aucun boot dans la periode 24h -> colonne "Boot" = N/A,
    // alors que le JSON contenait bien un dernier boot connu plus ancien.
    var allBoots    = pc.Boots || [];
    var boots       = allBoots.filter(function(b) { return parseDate(b.DateBoot) >= cutoff; });
    var bsods       = (pc.BSODs   || []).filter(function(b) { return parseDate(b.Date) >= cutoff; });
    var warnings    = (pc.ResourceWarnings || []).filter(function(w) { return parseDate(w.Timestamp) >= cutoff; });
    var bootsLongs  = boots.filter(function(b) { return b.EstBootLong; });
    // Dernier boot connu : on prend dans allBoots, peu importe la periode
    var dernierBoot = allBoots.length > 0 ? allBoots[allBoots.length - 1] : null;
    var hw          = pc.HardwareHealth || { WHEA_Fatal: [], WHEA_Corrected: [], GPU_TDR: [], Thermal: [] };

    // WHEA_Fatal : filtre par periode
    var wheaFatal = (hw.WHEA_Fatal || []).filter(function(h) { return parseDate(h.Timestamp) >= cutoff; });
    // WHEA_Corrected : deja agrege, on filtre par LastSeen
    var wheaCorrected = (hw.WHEA_Corrected || []).filter(function(h) { return parseDate(h.LastSeen) >= cutoff; });
    // Somme des occurrences corrigees sur la periode
    var wheaCorrectedTotal = wheaCorrected.reduce(function(s, c) { return s + c.Count; }, 0);

    var gpuTDR  = (hw.GPU_TDR || []).filter(function(h) { return parseDate(h.Timestamp) >= cutoff; });
    var thermal = (hw.Thermal || []).filter(function(h) { return parseDate(h.Timestamp) >= cutoff; });

    // Decompose WHEA_Fatal par composant (pour affichage dans le drill-down)
    var fatalByComponent = { CPU: [], RAM: [], PCIe: [], GPU: [], Autre: [] };
    wheaFatal.forEach(function(h) {
        var c = h.Component || 'Autre';
        if (!fatalByComponent[c]) fatalByComponent[c] = [];
        fatalByComponent[c].push(h);
    });

    var diskAlerts = (pc.DiskInfo || []).filter(function(d) { return d.IsAlert; });

    // Boots par type (sur la periode filtree)
    var bootsByType = { ColdBoot: 0, FastStartup: 0, Resume: 0, Unknown: 0 };
    boots.forEach(function(b) {
        var t = b.BootType || 'Unknown';
        if (bootsByType[t] === undefined) bootsByType[t] = 0;
        bootsByType[t]++;
    });

    // ==============================================================
    // v5.3 : Batterie + Services EDR
    // v5.4 : Boot Performance + Disk Health SMART
    // Tous peuvent etre null/absents (retrocompat JSON v5.2 et anciens)
    // ==============================================================
    var battery = pc.BatteryInfo || null;
    var batteryAlert = !!(battery && battery.IsAlert);

    var sentinel = (pc.ServicesHealth && pc.ServicesHealth.SentinelAgent) ? pc.ServicesHealth.SentinelAgent : null;
    var sentinelAlert = !!(sentinel && sentinel.IsAlert);

    var bootPerf = pc.BootPerformance || null;
    var bootPerfAlert = !!(bootPerf && bootPerf.IsAlert);

    var diskHealth = pc.DiskHealth || [];
    var diskSmartAlerts = diskHealth.filter(function(d) { return d.IsAlert; });
    var diskSmartAlert = diskSmartAlerts.length > 0;
    // WearPct max parmi les disques (null si tous null)
    var diskWorstWear = null;
    diskHealth.forEach(function(d) {
        if (d.WearPct !== null && d.WearPct !== undefined) {
            if (diskWorstWear === null || d.WearPct > diskWorstWear) diskWorstWear = d.WearPct;
        }
    });

    // v5.7 : moniteurs externes (peut etre absent si Collector < v5.5)
    var monitors = pc.Monitors || [];
    var SEUIL_MONITOR_OLD_YEARS = 7;
    var oldMonitors = monitors.filter(function(m) {
        return m.AgeYears !== null && m.AgeYears !== undefined && m.AgeYears >= SEUIL_MONITOR_OLD_YEARS;
    });
    var oldMonitorAlert = oldMonitors.length > 0;

    var result = {
        pc: pc,
        crashes: crashes, boots: boots, bsods: bsods, warnings: warnings,
        topRAM: pc.TopRAM || [], diskInfo: pc.DiskInfo || [], diskAlerts: diskAlerts,
        topCrashers: pc.TopCrashers || [],
        bootsLongs: bootsLongs, dernierBoot: dernierBoot,
        bootsByType: bootsByType,
        crashCount: crashes.length,
        bsodCount: bsods.length,
        bsodClassifieCount: crashes.filter(function(c) { return c.Type === 'BSOD'; }).length,
        warningCount: warnings.length,
        bootLongCount: bootsLongs.length,
        diskAlertCount: diskAlerts.length,
        // v5.2 : WHEA separe Fatal vs Corrected
        wheaFatal: wheaFatal,
        wheaCorrected: wheaCorrected,
        wheaCorrectedTotal: wheaCorrectedTotal,
        fatalByComponent: fatalByComponent,
        gpuTDR: gpuTDR,
        thermal: thermal,
        // hwCount = seulement les alertes qui doivent impacter le score
        // (les corrected ne plombent pas un PC)
        hwCount: wheaFatal.length + gpuTDR.length + thermal.length,
        collectedHoursAgo: hoursSince(pc.CollectedAt),
        // v5.3 / v5.4
        battery: battery,
        batteryAlert: batteryAlert,
        sentinel: sentinel,
        sentinelAlert: sentinelAlert,
        bootPerf: bootPerf,
        bootPerfAlert: bootPerfAlert,
        diskHealth: diskHealth,
        diskSmartAlerts: diskSmartAlerts,
        diskSmartAlert: diskSmartAlert,
        diskWorstWear: diskWorstWear,
        // v5.7 : moniteurs externes
        monitors: monitors,
        oldMonitors: oldMonitors,
        oldMonitorAlert: oldMonitorAlert,
        // v1.8 : RAM/GPU/throttling (null ou [] si JSON < 1.8)
        memory: pc.MemoryInventory || null,
        gpuInventory: pc.GPUInventory || [],
        hardwareHealth: pc.HardwareHealth || {}
    };
    result.score = computeScore(result);
    return result;
}

// ===== TRI =====
function sortPCs(pcs) {
    var dir = state.sort.dir === 'asc' ? 1 : -1;
    pcs.sort(function(a, b) {
        switch (state.sort.col) {
            case 'name':   return dir * a.pc.PC.localeCompare(b.pc.PC);
            case 'site':   return dir * (a.pc.Site || '').localeCompare(b.pc.Site || '');
            case 'user':   return dir * (a.pc.CurrentUser || '').localeCompare(b.pc.CurrentUser || '');
            case 'uptime':
                var ua = a.pc.UptimeDays || 0, ub = b.pc.UptimeDays || 0;
                return dir * (ua - ub);
            case 'fresh':  return dir * (a.collectedHoursAgo - b.collectedHoursAgo);
            case 'boot':
                var ba = a.dernierBoot ? a.dernierBoot.DurationMin : 0;
                var bb = b.dernierBoot ? b.dernierBoot.DurationMin : 0;
                return dir * (ba - bb);
            case 'crash':  return dir * (a.crashCount - b.crashCount);
            case 'bsod':   return dir * (a.bsodCount - b.bsodCount);
            case 'hw':     return dir * (a.hwCount - b.hwCount);
            case 'disk':
                var da = a.diskInfo.length > 0 ? a.diskInfo.reduce(function(m, d) { return d.PctFree < m ? d.PctFree : m; }, 100) : 100;
                var db = b.diskInfo.length > 0 ? b.diskInfo.reduce(function(m, d) { return d.PctFree < m ? d.PctFree : m; }, 100) : 100;
                return dir * (da - db);
            case 'perf':   return dir * (a.warningCount - b.warningCount);
            case 'score':
            default:       return dir * (a.score - b.score);
        }
    });
}

// ===== FILTRE KPI =====
function matchKpiFilter(p, filter) {
    if (!filter) return true;
    var crashRecentCutoff = new Date(generatedAt);
    crashRecentCutoff.setDate(crashRecentCutoff.getDate() - seuilCrashRecent);
    switch (filter) {
        case 'offline':     return p.pc.IsOffline;
        case 'online':      return !p.pc.IsOffline;
        case 'crash':       return p.crashCount > 0;
        case 'bsod':        return p.bsodCount > 0;
        case 'hw':          return p.hwCount > 0;
        case 'wheaCorr':    return p.wheaCorrected.length > 0;
        case 'bootLong':    return p.bootLongCount > 0;
        case 'diskAlert':   return p.diskAlertCount > 0;
        case 'oldCpu':      return p.pc.CPUAgeCategory === 'Ancien';
        case 'crashRecent': return (p.pc.Crashes || []).some(function(c) { return parseDate(c.Timestamp) >= crashRecentCutoff; });
        // v5.3 / v5.4
        case 'edrDown':     return p.sentinelAlert;
        case 'battery':     return p.batteryAlert;
        case 'bootPerf':    return p.bootPerfAlert;
        case 'smart':       return p.diskSmartAlert;
        // v5.7
        case 'oldMonitor':  return p.oldMonitorAlert;
        default: return true;
    }
}

// ===== RENDER =====
function render() {
    var cutoff = new Date(generatedAt);
    cutoff.setDate(cutoff.getDate() - state.days);

    var searchTerm = document.getElementById('searchInput').value.trim().toLowerCase();
    state.siteFilter = document.getElementById('siteFilter').value;
    state.cpuFilter  = document.getElementById('cpuFilter').value;

    // Enrichir tous les PCs
    var allEnriched = pcData.map(function(pc) { return enrichPC(pc, cutoff); });

    // KPIs : calcules sur l'ensemble (filtres de type "search/site" mais pas "maskHealthy/kpiFilter")
    // Pour que clicker un KPI filtre le tableau sans cacher le KPI lui-meme
    var baseFiltered = allEnriched.filter(function(p) {
        if (state.siteFilter && p.pc.Site !== state.siteFilter) return false;
        if (state.cpuFilter  && p.pc.CPUAgeCategory !== state.cpuFilter) return false;
        if (searchTerm) {
            return p.pc.PC.toLowerCase().indexOf(searchTerm) !== -1 ||
                   (p.pc.CurrentUser || '').toLowerCase().indexOf(searchTerm) !== -1;
        }
        return true;
    });

    renderKpis(baseFiltered);

    // Vue tableau : on applique kpiFilter + maskHealthy
    var visible = baseFiltered.filter(function(p) {
        if (!matchKpiFilter(p, state.kpiFilter)) return false;
        if (state.maskHealthy && p.score === 0) return false;
        return true;
    });

    sortPCs(visible);

    // v1.5 : Pagination
    // On slice "visible" selon la page courante et itemsPerPage.
    // Si un filtre change, il est possible que currentPage pointe sur une
    // page qui n'existe plus (ex: j'etais page 5, je filtre -> 40 PC -> 1 page).
    // On borne currentPage pour eviter un tableau vide.
    var totalPages = (state.itemsPerPage === 0) ? 1 : Math.max(1, Math.ceil(visible.length / state.itemsPerPage));
    if (state.currentPage > totalPages) state.currentPage = totalPages;
    if (state.currentPage < 1) state.currentPage = 1;

    var visibleSlice;
    if (state.itemsPerPage === 0) {
        visibleSlice = visible;  // Tout afficher
    } else {
        var startIdx = (state.currentPage - 1) * state.itemsPerPage;
        var endIdx   = startIdx + state.itemsPerPage;
        visibleSlice = visible.slice(startIdx, endIdx);
    }

    renderTable(visibleSlice);
    renderPagination(visible.length, totalPages);
    renderActiveFilters();
    renderGlobalCrashers(baseFiltered);
    renderBootBreakdown(baseFiltered);
    renderMonitorInventory(baseFiltered);

    document.getElementById('pcCount').textContent = visible.length + ' / ' + allEnriched.length + ' appareil(s)';
}

// ===== v1.5 : PAGINATION =====
// Rendu du bloc de pagination (affiche en haut et en bas du tableau).
// totalItems = nombre total de PC apres tous les filtres (pour les compteurs)
// totalPages = calcule par render() pour coherence
function renderPagination(totalItems, totalPages) {
    var containers = ['paginationTop', 'paginationBottom'];
    var per = state.itemsPerPage;
    var cur = state.currentPage;

    // Etat actif/inactif pour les selecteurs
    function perActive(val) { return per === val ? ' active' : ''; }

    // Calculer les bornes d'affichage
    var startNum, endNum;
    if (per === 0) {
        startNum = totalItems > 0 ? 1 : 0;
        endNum   = totalItems;
    } else {
        startNum = totalItems > 0 ? (cur - 1) * per + 1 : 0;
        endNum   = Math.min(cur * per, totalItems);
    }

    // Construction des numeros de page (logique "1 ... 3 4 [5] 6 7 ... 20")
    var pagesHtml = '';
    if (per > 0 && totalPages > 1) {
        var range = [];
        if (totalPages <= 7) {
            // Peu de pages : on les affiche toutes
            for (var i = 1; i <= totalPages; i++) range.push(i);
        } else {
            // Beaucoup de pages : ellipses intelligentes
            range.push(1);
            if (cur > 3) range.push('...');
            for (var j = Math.max(2, cur - 1); j <= Math.min(totalPages - 1, cur + 1); j++) {
                range.push(j);
            }
            if (cur < totalPages - 2) range.push('...');
            range.push(totalPages);
        }
        range.forEach(function(p) {
            if (p === '...') {
                pagesHtml += '<span class="pagination-ellipsis">...</span>';
            } else {
                var active = p === cur ? ' active' : '';
                pagesHtml += '<button class="pagination-page' + active + '" onclick="goToPage(' + p + ')">' + p + '</button>';
            }
        });
    }

    // Bouton Precedent / Suivant
    var prevDisabled = (per === 0 || cur === 1) ? ' disabled' : '';
    var nextDisabled = (per === 0 || cur >= totalPages) ? ' disabled' : '';

    // Boutons "afficher plus" et "tout afficher"
    var canShowMore = (per > 0 && cur < totalPages);
    var showMoreBtn = canShowMore
        ? '<button class="pagination-showmore" onclick="showMorePage()" title="Afficher les ' + Math.min(per, totalItems - endNum) + ' PC suivants">+ Afficher ' + Math.min(per, totalItems - endNum) + ' de plus</button>'
        : '';
    var showAllBtn = (per !== 0 && totalItems > per)
        ? '<button class="pagination-showall" onclick="setItemsPerPage(0)" title="Desactiver la pagination et tout afficher">Tout afficher</button>'
        : '';

    var html =
        '<div class="pagination-bar">' +
            '<div class="pagination-info">' +
                'Affichage ' + startNum + '-' + endNum + ' sur ' + totalItems + ' PC' +
            '</div>' +
            '<div class="pagination-controls">' +
                '<span class="pagination-label">Par page :</span>' +
                '<button class="pagination-per' + perActive(20)  + '" onclick="setItemsPerPage(20)">20</button>' +
                '<button class="pagination-per' + perActive(50)  + '" onclick="setItemsPerPage(50)">50</button>' +
                '<button class="pagination-per' + perActive(100) + '" onclick="setItemsPerPage(100)">100</button>' +
                '<button class="pagination-per' + perActive(0)   + '" onclick="setItemsPerPage(0)">Tous</button>' +
                '<span class="pagination-sep"></span>' +
                '<button class="pagination-nav' + prevDisabled + '" onclick="goToPage(' + (cur - 1) + ')"' + prevDisabled + '>&larr; Prec.</button>' +
                pagesHtml +
                '<button class="pagination-nav' + nextDisabled + '" onclick="goToPage(' + (cur + 1) + ')"' + nextDisabled + '>Suiv. &rarr;</button>' +
                showMoreBtn +
                showAllBtn +
            '</div>' +
        '</div>';

    containers.forEach(function(id) {
        var el = document.getElementById(id);
        if (el) el.innerHTML = html;
    });
}

function setItemsPerPage(n) {
    state.itemsPerPage = n;
    state.currentPage = 1;
    try { localStorage.setItem('pcpulse_itemsPerPage', n); } catch (e) {}
    render();
}

function goToPage(n) {
    state.currentPage = n;
    render();
    // Scroll vers le haut du tableau pour le confort utilisateur
    var tableEl = document.getElementById('deviceTable');
    if (tableEl) tableEl.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

// "Afficher + de N" : passe a la page suivante sans changer itemsPerPage
// (difference avec goToPage : ne scroll pas en haut, reste dans la continuite)
function showMorePage() {
    state.currentPage++;
    render();
}

// Reset currentPage a 1 des qu'un filtre change (sinon on peut se retrouver
// sur une page qui n'existe plus et voir un tableau vide le temps d'un render)
function resetPaginationToFirstPage() {
    state.currentPage = 1;
}

function renderActiveFilters() {
    var container = document.getElementById('activeFilters');
    var chips = [];
    if (state.kpiFilter) {
        var labels = {
            'offline': 'Offline', 'online': 'En ligne', 'crash': 'Crash/Freeze',
            'bsod': 'BSOD', 'hw': 'Erreurs fatales HW', 'wheaCorr': 'WHEA corrigees',
            'bootLong': 'Boots longs',
            'diskAlert': 'Disques critiques', 'oldCpu': 'CPU anciens',
            'crashRecent': 'Crash recent (' + seuilCrashRecent + 'j)',
            // v5.3 / v5.4
            'edrDown': 'EDR en panne',
            'battery': 'Batterie usee',
            'bootPerf': 'Boots lents',
            'smart': 'SMART alerte',
            'oldMonitor': 'Ecrans anciens (>=7 ans)'
        };
        chips.push('<span class="active-filter-chip">' + labels[state.kpiFilter] + ' <span class="close" onclick="clearKpiFilter()">&times;</span></span>');
    }
    if (state.maskHealthy) {
        chips.push('<span class="active-filter-chip">PC sains masques <span class="close" onclick="toggleMaskHealthy()">&times;</span></span>');
    }
    container.innerHTML = chips.join('');
}

function kpiCard(key, colorClass, value, label, cardClass) {
    var active = state.kpiFilter === key ? ' active' : '';
    // v5.5 : attenuer les cards "vides" (value=0) et emphasiser les critiques
    var quietCls = '';
    if ((value === 0 || value === '0') && key !== '') {
        quietCls = ' kpi-quiet';
    } else if (typeof value === 'number' && value > 0 && (cardClass === 'card-danger' || cardClass === 'card-hardware')) {
        quietCls = ' kpi-loud';
    }
    return '<div class="kpi-card ' + cardClass + quietCls + active + '" onclick="toggleKpiFilter(\'' + key + '\')">' +
           '<div class="kpi-value ' + colorClass + '">' + value + '</div>' +
           '<div class="kpi-label">' + label + '</div></div>';
}

function renderKpis(pcs) {
    var totalPC = pcs.length;
    var totalOffline = pcs.filter(function(p) { return p.pc.IsOffline; }).length;
    var totalCrash = pcs.reduce(function(s, p) { return s + p.crashCount; }, 0);
    var totalBSOD = pcs.reduce(function(s, p) { return s + p.bsodCount; }, 0);
    var totalBootsLongs = pcs.reduce(function(s, p) { return s + p.bootLongCount; }, 0);
    var totalDiskAlerts = pcs.reduce(function(s, p) { return s + p.diskAlertCount; }, 0);
    var totalHardware = pcs.reduce(function(s, p) { return s + p.hwCount; }, 0);
    var pcWithCorrected = pcs.filter(function(p) { return p.wheaCorrected.length > 0; }).length;
    var totalOldCPU = pcs.filter(function(p) { return p.pc.CPUAgeCategory === 'Ancien'; }).length;

    var totalEdrDown = pcs.filter(function(p) { return p.sentinelAlert; }).length;
    var totalBatteryWorn = pcs.filter(function(p) { return p.batteryAlert; }).length;
    var totalBootPerfSlow = pcs.filter(function(p) { return p.bootPerfAlert; }).length;
    var totalSmartAlert = pcs.filter(function(p) { return p.diskSmartAlert; }).length;

    // v5.7 : inventaire moniteurs externes
    var totalMonitors = pcs.reduce(function(s, p) { return s + (p.monitors ? p.monitors.length : 0); }, 0);
    var pcWithOldMonitor = pcs.filter(function(p) { return p.oldMonitorAlert; }).length;
    var pcWithMonitor = pcs.filter(function(p) { return p.monitors && p.monitors.length > 0; }).length;

    var crashRecentCutoff = new Date(generatedAt);
    crashRecentCutoff.setDate(crashRecentCutoff.getDate() - seuilCrashRecent);
    var pcCrashRecents = pcs.filter(function(p) {
        return (p.pc.Crashes || []).some(function(c) { return parseDate(c.Timestamp) >= crashRecentCutoff; });
    }).length;

    var pcHealthy = pcs.filter(function(p) { return p.score === 0; }).length;
    var pcWithIssues = totalPC - pcHealthy;
    var pctOnline = totalPC > 0 ? Math.round(((totalPC - totalOffline) / totalPC) * 100) : 0;
    var pctHealthy = totalPC > 0 ? Math.round((pcHealthy / totalPC) * 100) : 0;

    // ========================================================
    // SUMMARY BAR : 3 chiffres essentiels en haut de page
    // ========================================================
    var summary = document.getElementById('summaryBar');
    summary.innerHTML =
        '<div class="summary-tile">' +
          '<div class="summary-icon blue">&#128187;</div>' +
          '<div class="summary-content">' +
            '<div class="summary-value">' + totalPC + '</div>' +
            '<div class="summary-label">PC monitor&eacute;s</div>' +
            '<div class="summary-sub">dans la p&eacute;riode s&eacute;lectionn&eacute;e</div>' +
          '</div>' +
        '</div>' +
        '<div class="summary-tile">' +
          '<div class="summary-icon green">&#10003;</div>' +
          '<div class="summary-content">' +
            '<div class="summary-value">' + (totalPC - totalOffline) + ' <span style="font-size:14px;color:var(--text-muted);font-weight:500">/ ' + totalPC + '</span></div>' +
            '<div class="summary-label">En ligne (24h)</div>' +
            '<div class="summary-sub">' + pctOnline + '% du parc joignable</div>' +
          '</div>' +
        '</div>' +
        '<div class="summary-tile">' +
          '<div class="summary-icon ' + (pcWithIssues === 0 ? 'green' : 'red') + '">' + (pcWithIssues === 0 ? '&#10003;' : '&#9888;') + '</div>' +
          '<div class="summary-content">' +
            '<div class="summary-value">' + pcHealthy + ' <span style="font-size:14px;color:var(--text-muted);font-weight:500">/ ' + totalPC + '</span></div>' +
            '<div class="summary-label">PC sans alerte</div>' +
            '<div class="summary-sub">' + pctHealthy + '% sains &middot; ' + pcWithIssues + ' &agrave; examiner</div>' +
          '</div>' +
        '</div>';

    // ========================================================
    // v5.6 : BARRE DE 4 BOUTONS DE GROUPES
    // ========================================================
    // Definition des 4 familles et de leurs sous-KPIs
    // Chaque sous-KPI = { filterKey, label, value, level('danger'|'warn'|'neutral') }
    var familySecurity = {
        key: 'security', family: 'family-security', icon: '&#128274;', title: 'S&eacute;curit&eacute;',
        subs: [
            { k: 'offline', label: 'Offline',      v: totalOffline,  level: (totalOffline > 0 ? 'warn' : 'neutral') },
            { k: 'edrDown', label: 'EDR en panne', v: totalEdrDown,  level: (totalEdrDown > 0 ? 'danger' : 'neutral') }
        ]
    };
    var familyStability = {
        key: 'stability', family: 'family-stability', icon: '&#128165;', title: 'Stabilit&eacute;',
        subs: [
            { k: 'crashRecent', label: 'Crash r&eacute;cent (' + seuilCrashRecent + 'j)', v: pcCrashRecents, level: (pcCrashRecents > 0 ? 'danger' : 'neutral') },
            { k: 'crash',       label: 'Crash/Freeze (' + state.days + 'j)',              v: totalCrash,     level: (totalCrash > 0 ? 'warn' : 'neutral') },
            { k: 'bsod',        label: 'BSOD (' + state.days + 'j)',                      v: totalBSOD,      level: (totalBSOD > 0 ? 'danger' : 'neutral') },
            { k: 'hw',          label: 'Erreurs fatales HW',                              v: totalHardware,  level: (totalHardware > 0 ? 'danger' : 'neutral') },
            { k: 'wheaCorr',    label: 'WHEA corrig&eacute;es',                           v: pcWithCorrected,level: 'neutral' }
        ]
    };
    var familyPerformance = {
        key: 'performance', family: 'family-performance', icon: '&#9889;', title: 'Performance',
        subs: [
            { k: 'bootLong', label: 'Boots longs',          v: totalBootsLongs,   level: (totalBootsLongs > 0 ? 'warn' : 'neutral') },
            { k: 'bootPerf', label: 'Boots lents (&gt;90s)', v: totalBootPerfSlow, level: (totalBootPerfSlow > 0 ? 'warn' : 'neutral') }
        ]
    };
    var familyMaterial = {
        key: 'material', family: 'family-material', icon: '&#128295;', title: 'Usure mat&eacute;rielle',
        subs: [
            { k: 'diskAlert',  label: 'Disques satur&eacute;s',           v: totalDiskAlerts,   level: (totalDiskAlerts > 0 ? 'warn' : 'neutral') },
            { k: 'smart',      label: 'SMART alerte',                     v: totalSmartAlert,   level: (totalSmartAlert > 0 ? 'warn' : 'neutral') },
            { k: 'battery',    label: 'Batterie us&eacute;e (&lt;60%)',   v: totalBatteryWorn,  level: (totalBatteryWorn > 0 ? 'warn' : 'neutral') },
            { k: 'oldCpu',     label: 'CPU anciens',                      v: totalOldCPU,       level: (totalOldCPU > 0 ? 'danger' : 'neutral') },
            { k: 'oldMonitor', label: 'Ecrans &acirc;g&eacute;s (&ge;7 ans)', v: pcWithOldMonitor, level: (pcWithOldMonitor > 0 ? 'warn' : 'neutral') }
        ]
    };

    var families = [familySecurity, familyStability, familyPerformance, familyMaterial];

    // Quels filtres KPI appartiennent a quelle famille (pour marquer
    // un bouton "active" si le filtre courant pointe vers cette famille)
    var keysOf = function(fam) { return fam.subs.map(function(s) { return s.k; }); };

    // Rendu des boutons
    var bar = document.getElementById('kpiGroupBar');
    var html = '';
    families.forEach(function(fam) {
        // Compter les alertes "chaudes" de la famille (warn + danger)
        // et celles qui sont "critiques" (danger)
        var hotCount = 0, criticalCount = 0;
        fam.subs.forEach(function(s) {
            if (s.level === 'danger' && s.v > 0) { hotCount += s.v; criticalCount += s.v; }
            else if (s.level === 'warn' && s.v > 0) { hotCount += s.v; }
        });

        var stateCls = 'calm';
        if (criticalCount > 0)   stateCls = 'danger';
        else if (hotCount > 0)   stateCls = 'warn';

        // "Active" = le filtre KPI courant est l'un des sous-KPIs de la famille
        var active = (state.kpiFilter && keysOf(fam).indexOf(state.kpiFilter) >= 0) ? ' active' : '';

        // Resume textuel sous le titre
        var sub;
        if (hotCount === 0) {
            sub = 'Aucune alerte';
        } else {
            var parts = [];
            fam.subs.forEach(function(s) {
                if (s.v > 0 && s.level !== 'neutral') parts.push(s.v + ' ' + s.label.replace(/&nbsp;/g,' '));
            });
            sub = parts.slice(0, 2).join(' &middot; ');
            if (parts.length > 2) sub += ' &hellip;';
        }

        // Popover : items cliquables pour filtrer sur un sous-KPI precis
        // onclick avec event.stopPropagation pour ne pas declencher le clic
        // du bouton parent (qui filtre sur la famille entiere)
        var itemsHtml = '';
        fam.subs.forEach(function(s) {
            var valCls = (s.v === 0) ? '' : (s.level === 'danger' ? 'danger' : (s.level === 'warn' ? 'warn' : ''));
            var itemActive = (state.kpiFilter === s.k) ? ' active' : '';
            itemsHtml += '<div class="kpi-popover-item' + itemActive + '" onclick="event.stopPropagation();toggleKpiFilter(\'' + s.k + '\')">' +
                         '<span class="pop-label">' + s.label + '</span>' +
                         '<span class="pop-value ' + valCls + '">' + s.v + '</span>' +
                         '</div>';
        });

        // Le clic du bouton filtre sur "crashRecent" pour stab, "edrDown" pour secu,
        // "bootPerf" pour perf, "diskAlert" pour usure : c'est le sous-KPI le plus
        // "important" de chaque famille. Ou alors on passe un filter meta "family"
        // -> plus simple : le clic toggle le filtre sur LE sous-KPI le plus critique
        // qui est > 0, sinon sur le premier de la famille.
        var primaryKey = fam.subs[0].k;
        for (var i = 0; i < fam.subs.length; i++) {
            if (fam.subs[i].v > 0 && fam.subs[i].level === 'danger') { primaryKey = fam.subs[i].k; break; }
        }
        if (primaryKey === fam.subs[0].k) {
            for (var j = 0; j < fam.subs.length; j++) {
                if (fam.subs[j].v > 0 && fam.subs[j].level === 'warn') { primaryKey = fam.subs[j].k; break; }
            }
        }

        html += '<button type="button" class="kpi-group-btn ' + stateCls + ' ' + fam.family + active + '" ' +
                'onclick="toggleKpiFilter(\'' + primaryKey + '\')">' +
                '<div class="kpi-group-btn-icon">' + fam.icon + '</div>' +
                '<div class="kpi-group-btn-content">' +
                  '<div class="kpi-group-btn-title">' + fam.title + '</div>' +
                  '<div class="kpi-group-btn-sub">' + sub + '</div>' +
                '</div>' +
                '<div class="kpi-group-btn-count">' + hotCount + '</div>' +
                '<div class="kpi-group-popover">' +
                  '<div class="kpi-popover-items">' + itemsHtml + '</div>' +
                '</div>' +
                '</button>';
    });
    bar.innerHTML = html;
}

// ===== AGREGATION PARC : REPARTITION DES DEMARRAGES (v5.2) =====
function renderBootBreakdown(pcs) {
    var container = document.getElementById('bootBreakdown');
    var totals = { ColdBoot: 0, FastStartup: 0, Resume: 0, Unknown: 0 };
    var pcCountWithBootData = 0;

    pcs.forEach(function(p) {
        var bt = p.bootsByType || {};
        var hasData = (bt.ColdBoot || 0) + (bt.FastStartup || 0) + (bt.Resume || 0) + (bt.Unknown || 0);
        if (hasData > 0) pcCountWithBootData++;
        totals.ColdBoot    += bt.ColdBoot    || 0;
        totals.FastStartup += bt.FastStartup || 0;
        totals.Resume      += bt.Resume      || 0;
        totals.Unknown     += bt.Unknown     || 0;
    });

    var grandTotal = totals.ColdBoot + totals.FastStartup + totals.Resume + totals.Unknown;

    if (grandTotal === 0) {
        container.innerHTML = '<div class="boot-breakdown-empty">Aucun demarrage detecte sur la periode ' +
            '(ou Collector en version anterieure a v5.2 : l\'Event 27 Kernel-Boot n\'etait pas encore collecte)</div>';
        return;
    }

    var pctCold    = (totals.ColdBoot    / grandTotal) * 100;
    var pctFast    = (totals.FastStartup / grandTotal) * 100;
    var pctResume  = (totals.Resume      / grandTotal) * 100;
    var pctUnknown = (totals.Unknown     / grandTotal) * 100;

    var barHtml = '<div class="boot-breakdown-bar">';
    if (pctCold > 0)    barHtml += '<div class="boot-breakdown-segment cold"    style="width:' + pctCold.toFixed(1)    + '%" title="ColdBoot : ' + totals.ColdBoot + '">'    + (pctCold    >= 5 ? totals.ColdBoot    : '') + '</div>';
    if (pctFast > 0)    barHtml += '<div class="boot-breakdown-segment fast"    style="width:' + pctFast.toFixed(1)    + '%" title="Fast Startup : ' + totals.FastStartup + '">' + (pctFast    >= 5 ? totals.FastStartup : '') + '</div>';
    if (pctResume > 0)  barHtml += '<div class="boot-breakdown-segment resume"  style="width:' + pctResume.toFixed(1)  + '%" title="Resume hibernation : ' + totals.Resume + '">' + (pctResume  >= 5 ? totals.Resume      : '') + '</div>';
    if (pctUnknown > 0) barHtml += '<div class="boot-breakdown-segment unknown" style="width:' + pctUnknown.toFixed(1) + '%" title="Inconnu : ' + totals.Unknown + '">' + (pctUnknown >= 5 ? totals.Unknown     : '') + '</div>';
    barHtml += '</div>';

    var legendHtml = '<div class="boot-breakdown-legend">' +
        '<div class="boot-breakdown-legend-item"><div class="boot-breakdown-legend-dot" style="background:#3a6cbf"></div>&#10052; ColdBoot : ' + totals.ColdBoot + ' (' + pctCold.toFixed(1) + '%)</div>' +
        '<div class="boot-breakdown-legend-item"><div class="boot-breakdown-legend-dot" style="background:var(--yellow)"></div>&#9889; Fast Startup : ' + totals.FastStartup + ' (' + pctFast.toFixed(1) + '%)</div>' +
        '<div class="boot-breakdown-legend-item"><div class="boot-breakdown-legend-dot" style="background:var(--purple)"></div>&#128164; Resume : ' + totals.Resume + ' (' + pctResume.toFixed(1) + '%)</div>';
    if (totals.Unknown > 0) {
        legendHtml += '<div class="boot-breakdown-legend-item"><div class="boot-breakdown-legend-dot" style="background:var(--text-faint)"></div>? Inconnu : ' + totals.Unknown + '</div>';
    }
    legendHtml += '</div>';

    // Note metier : expliquer ce que signifient ces chiffres
    var noteHtml = '';
    var fastDomine = (pctFast > 50);
    var coldDomine = (pctCold > 50);
    if (fastDomine) {
        noteHtml = '<div class="boot-breakdown-note">Le Fast Startup domine sur le parc (' + pctFast.toFixed(0) + '%). ' +
            'PCPulse detecte correctement les reprises d\'activite (Fast Startup et wake Modern Standby) : ' +
            'l\'uptime affiche reflete le temps depuis la derniere reprise utilisateur, pas depuis le dernier vrai cold boot.</div>';
    } else if (coldDomine && pctFast < 10) {
        noteHtml = '<div class="boot-breakdown-note">Le parc est majoritairement en ColdBoot (' + pctCold.toFixed(0) + '%). ' +
            'Fast Startup semble desactive sur la plupart des postes.</div>';
    }

    container.innerHTML = barHtml + legendHtml + noteHtml +
        '<div style="font-size:11px;color:var(--text-muted)">' + pcCountWithBootData + ' PC remontent des donnees de demarrage sur ' + pcs.length + ' (' + grandTotal + ' d&eacute;marrages au total)</div>';
}

// ===== AGREGATION PARC : TOP CRASHERS GLOBAL (v1.7 : 3 sections) =====
function renderGlobalCrashers(pcs) {
    var agg = {};
    pcs.forEach(function(p) {
        (p.topCrashers || []).forEach(function(c) {
            var key = c.AppName;
            // v1.7 : blacklist HARD - on zappe completement a l'agregation
            if (isBlacklistedHard(key)) return;
            if (!agg[key]) agg[key] = { name: key, total: 0, pcCount: 0, pcSet: {} };
            agg[key].total += c.CrashCount;
            if (!agg[key].pcSet[p.pc.PC]) {
                agg[key].pcCount++;
                agg[key].pcSet[p.pc.PC] = true;
            }
        });
    });

    var arr = Object.keys(agg).map(function(k) {
        var a = agg[k];
        var cls = classifyCrasher(a.name, a.total, a.pcCount);
        a.score  = cls.score;
        a.level  = cls.level;
        a.forced = cls.forced;
        return a;
    });

    // Tri principal : score decroissant dans chaque section
    arr.sort(function(a, b) { return b.score - a.score; });

    var container = document.getElementById('globalCrashers');
    if (arr.length === 0) {
        container.innerHTML = '<div class="detail-empty">Aucun crash d application agrege</div>';
        return;
    }

    var local  = arr.filter(function(a) { return a.level === 'local'; });
    var spread = arr.filter(function(a) { return a.level === 'spread'; });
    var noise  = arr.filter(function(a) { return a.level === 'noise'; });

    // Helper de rendu d'une ligne crasheur
    function renderRow(c) {
        var badgeCls = 'crasher-score-' + c.level;
        var forcedTag = c.forced ? '<span class="crasher-forced" title="Force en bruit (blacklist soft)">&#128274;</span>' : '';
        var scoreTxt = c.score.toFixed(1);
        return '<div class="global-crasher-row ' + c.level + '">' +
                  '<span class="global-crasher-name" title="' + c.name + '">' + c.name + '</span>' +
                  forcedTag +
                  '<span class="crasher-score ' + badgeCls + '" title="Score signal (plus haut = plus concentre sur peu de PC)">' + scoreTxt + '</span>' +
                  '<span class="global-crasher-stats">' +
                    '<span class="global-crasher-total">' + c.total + '</span>' +
                    ' crashs / ' + c.pcCount + ' PC' +
                  '</span>' +
                '</div>';
    }

    var html = '';

    // --- Section 1 : Signaux locaux ---
    html += '<div class="crasher-section crasher-section-local">' +
            '<h4 class="crasher-section-title">&#127919; Signaux locaux ' +
              '<span class="crasher-section-count">' + local.length + '</span>' +
              '<span class="crasher-section-hint">Crashs concentres sur peu de PC, a investiguer</span>' +
            '</h4>';
    if (local.length > 0) {
        html += '<div class="global-crashers-grid">' + local.map(renderRow).join('') + '</div>';
    } else {
        html += '<div class="crasher-section-empty">Aucun signal local detecte</div>';
    }
    html += '</div>';

    // --- Section 2 : Problemes repartis ---
    html += '<div class="crasher-section crasher-section-spread">' +
            '<h4 class="crasher-section-title">&#9888;&#65039; Problemes repartis ' +
              '<span class="crasher-section-count">' + spread.length + '</span>' +
              '<span class="crasher-section-hint">Possibles bugs applicatifs touchant plusieurs PC</span>' +
            '</h4>';
    if (spread.length > 0) {
        html += '<div class="global-crashers-grid">' + spread.map(renderRow).join('') + '</div>';
    } else {
        html += '<div class="crasher-section-empty">Aucun probleme reparti detecte</div>';
    }
    html += '</div>';

    // --- Section 3 : Bruit ambient (repliable, fermee par defaut) ---
    html += '<div class="crasher-section crasher-section-noise">' +
            '<h4 class="crasher-section-title crasher-collapsible" onclick="toggleNoise(this)">' +
              '<span class="crasher-toggle">&#9656;</span> ' +
              '&#128266; Bruit ambient ' +
              '<span class="crasher-section-count">' + noise.length + '</span>' +
              '<span class="crasher-section-hint">Bruit de fond parc (repli&eacute; par d&eacute;faut)</span>' +
            '</h4>' +
            '<div class="crasher-noise-body" style="display:none">';
    if (noise.length > 0) {
        html += '<div class="global-crashers-grid">' + noise.map(renderRow).join('') + '</div>';
    } else {
        html += '<div class="crasher-section-empty">Pas de bruit residuel</div>';
    }
    html += '</div></div>';

    container.innerHTML = html;
}

// v1.7 : toggle visuel section bruit
function toggleNoise(el) {
    var body = el.parentElement.querySelector('.crasher-noise-body');
    var toggle = el.querySelector('.crasher-toggle');
    if (!body) return;
    if (body.style.display === 'none') {
        body.style.display = '';
        if (toggle) toggle.innerHTML = '&#9662;';  // fleche bas
    } else {
        body.style.display = 'none';
        if (toggle) toggle.innerHTML = '&#9656;';  // fleche droite
    }
}

// ===== AGREGATION PARC : INVENTAIRE MONITORS (v5.7) =====
// Produit les stats d'inventaire des ecrans externes :
// - Nb total d'ecrans branches
// - Nb de PC avec au moins 1 ecran externe
// - Top fabricants avec compteur
// - Ecrans anciens (>= 7 ans)
function renderMonitorInventory(pcs) {
    var panel = document.getElementById('monitorPanel');
    var container = document.getElementById('monitorInventory');

    var allMonitors = [];
    var pcWithMonitor = 0;
    pcs.forEach(function(p) {
        if (p.monitors && p.monitors.length > 0) {
            pcWithMonitor++;
            p.monitors.forEach(function(m) { allMonitors.push({ mon: m, pc: p.pc.PC }); });
        }
    });

    // Si aucun moniteur remonte, on masque le panneau (ex: parc 100% Collector v5.4)
    if (allMonitors.length === 0) {
        panel.style.display = 'none';
        return;
    }
    panel.style.display = '';

    // Top fabricants
    var manufCount = {};
    var oldCount = 0;
    var ageSum = 0;
    var ageCount = 0;
    allMonitors.forEach(function(item) {
        var m = item.mon;
        var key = m.Manufacturer || m.ManufacturerCode || '?';
        manufCount[key] = (manufCount[key] || 0) + 1;
        if (m.AgeYears !== null && m.AgeYears !== undefined) {
            if (m.AgeYears >= 7) oldCount++;
            ageSum += m.AgeYears;
            ageCount++;
        }
    });
    var topManufs = Object.keys(manufCount).map(function(k) {
        return { name: k, count: manufCount[k] };
    }).sort(function(a, b) { return b.count - a.count; }).slice(0, 5);

    var avgAge = ageCount > 0 ? (ageSum / ageCount).toFixed(1) : 'N/A';

    // Construction du HTML
    var html = '';

    html += '<div class="monitor-inv-tile">' +
            '<div class="monitor-inv-value">' + allMonitors.length + '</div>' +
            '<div class="monitor-inv-label">&Eacute;crans secondaires</div>' +
            '<div class="monitor-inv-sub">sur ' + pcWithMonitor + ' PC du parc</div>' +
            '</div>';

    html += '<div class="monitor-inv-tile">' +
            '<div class="monitor-inv-value">' + avgAge + (avgAge !== 'N/A' ? ' ans' : '') + '</div>' +
            '<div class="monitor-inv-label">&Acirc;ge moyen</div>' +
            '<div class="monitor-inv-sub">calcul sur ' + ageCount + ' &eacute;cran(s)</div>' +
            '</div>';

    html += '<div class="monitor-inv-tile">' +
            '<div class="monitor-inv-value" style="color:' + (oldCount > 0 ? 'var(--orange)' : 'var(--text-muted)') + '">' + oldCount + '</div>' +
            '<div class="monitor-inv-label">&Eacute;crans &ge; 7 ans</div>' +
            '<div class="monitor-inv-sub">candidats au renouvellement</div>' +
            '</div>';

    var manufHtml = topManufs.map(function(m) {
        return '<span class="chip">' + m.name + ' : <strong>' + m.count + '</strong></span>';
    }).join('');
    html += '<div class="monitor-inv-tile" style="grid-column:span 2">' +
            '<div class="monitor-inv-label" style="margin-top:0">Top fabricants</div>' +
            '<div class="monitor-inv-sub">' + (manufHtml || 'Aucun') + '</div>' +
            '</div>';

    container.innerHTML = html;
}

// ===== TABLEAU =====
function sortArrow(col) {
    if (state.sort.col !== col) return '<span class="sort-arrow">&#9650;&#9660;</span>';
    return state.sort.dir === 'asc' ? '<span class="sort-arrow">&#9650;</span>' : '<span class="sort-arrow">&#9660;</span>';
}
function sortClass(col) {
    return 'sortable' + (state.sort.col === col ? ' sort-active' : '');
}
function sortColumn(col) {
    if (state.sort.col === col) {
        state.sort.dir = state.sort.dir === 'desc' ? 'asc' : 'desc';
    } else {
        state.sort.col = col;
        state.sort.dir = (col === 'name' || col === 'site' || col === 'user') ? 'asc' : 'desc';
    }
    render();
}

function renderTable(pcs) {
    var siteHeader = showSite ? '<th class="' + sortClass('site') + '" onclick="sortColumn(\'site\')">Site ' + sortArrow('site') + '</th>' : '';
    // v5.5 : les colonnes "techniques" (crash/bsod/hw/disque/perf) sont
    // masquees par defaut via la classe .col-advanced. Le bouton toolbar
    // "Vue detaillee" bascule le body.advanced-cols pour les reveler.
    document.getElementById('tableHead').innerHTML =
        '<tr>' +
        '<th class="' + sortClass('score') + '" onclick="sortColumn(\'score\')">Sant&eacute; ' + sortArrow('score') + '</th>' +
        '<th class="' + sortClass('name')  + '" onclick="sortColumn(\'name\')">Nom PC '  + sortArrow('name')  + '</th>' +
        '<th>Statut</th>' + siteHeader +
        '<th>IP</th>' +
        '<th class="' + sortClass('user')  + '" onclick="sortColumn(\'user\')">Utilisateur ' + sortArrow('user') + '</th>' +
        '<th>Conn.</th><th>CPU</th>' +
        '<th class="' + sortClass('uptime') + '" onclick="sortColumn(\'uptime\')">Uptime ' + sortArrow('uptime') + '</th>' +
        '<th class="' + sortClass('fresh')  + '" onclick="sortColumn(\'fresh\')">Vu '  + sortArrow('fresh')  + '</th>' +
        '<th class="' + sortClass('boot')   + '" onclick="sortColumn(\'boot\')">Boot ' + sortArrow('boot')   + '</th>' +
        '<th class="col-advanced ' + sortClass('crash')  + '" onclick="sortColumn(\'crash\')">Crash ' + sortArrow('crash') + '</th>' +
        '<th class="col-advanced ' + sortClass('bsod')   + '" onclick="sortColumn(\'bsod\')">BSOD '  + sortArrow('bsod')   + '</th>' +
        '<th class="col-advanced ' + sortClass('hw')     + '" onclick="sortColumn(\'hw\')">HW '      + sortArrow('hw')     + '</th>' +
        '<th class="col-advanced ' + sortClass('disk')   + '" onclick="sortColumn(\'disk\')">Disque ' + sortArrow('disk')  + '</th>' +
        '<th class="col-advanced ' + sortClass('perf')   + '" onclick="sortColumn(\'perf\')">Perf '  + sortArrow('perf')   + '</th>' +
        '<th title="Indicateurs : Batterie / EDR / BootPerf / SMART">Ind.</th>' +
        '</tr>';

    var tbody = '';
    var colspan = showSite ? 17 : 16;
    pcs.forEach(function(p, idx) {
        var pc = p.pc;
        var rowClass = 'row-ok';
        if (p.crashCount > 0 || p.bsodCount > 0) rowClass = 'row-danger';
        else if (p.hwCount > 0 || p.sentinelAlert || p.diskSmartAlert) rowClass = 'row-hardware';
        else if (p.bootLongCount > 0 || p.diskAlertCount > 0 || p.batteryAlert || p.bootPerfAlert) rowClass = 'row-warning';

        var scoreCell = '<span class="score-badge ' + scoreClass(p.score) + '" title="Score = BSOD:' + p.bsodCount + ' Crash:' + p.crashCount + ' WHEAFatal:' + p.wheaFatal.length + ' GPU:' + p.gpuTDR.length + ' Thermal:' + p.thermal.length + ' BootLong:' + p.bootLongCount + ' Disk:' + p.diskAlertCount + (p.pc.IsOffline ? ' +Offline' : '') + (p.sentinelAlert ? ' +EDR' : '') + (p.batteryAlert ? ' +Batt' : '') + (p.bootPerfAlert ? ' +BootSlow' : '') + (p.diskSmartAlert ? ' +SMART' : '') + '">' + p.score + '</span>';

        var statusBadge = pc.IsOffline
            ? '<span class="badge badge-offline">OFFLINE</span>'
            : '<span class="badge badge-online">OK</span>';

        var siteCell = '';
        if (showSite) {
            var siteClass = pc.Site === 'Inconnu' ? 'site-badge inconnu' : 'site-badge';
            siteCell = '<td><span class="' + siteClass + '">' + pc.Site + '</span></td>';
        }

        var uptimeCell = 'N/A';
        if (pc.UptimeDays !== null && pc.UptimeDays !== undefined) {
            var uptimeClass = 'uptime-ok';
            if (pc.UptimeDays > 30) uptimeClass = 'uptime-danger';
            else if (pc.UptimeDays > 14) uptimeClass = 'uptime-warning';
            uptimeCell = '<span class="uptime-badge ' + uptimeClass + '">' + formatUptime(pc.UptimeDays) + '</span>';
        }

        var freshCell = '<span class="' + freshnessClass(p.collectedHoursAgo) + '">' + timeAgo(pc.CollectedAt) + '</span>';

        var connType = pc.ConnectionType || 'Inconnu';
        var connClass = 'conn-autre';
        if (connType === 'Ethernet') connClass = 'conn-ethernet';
        else if (connType === 'WiFi') connClass = 'conn-wifi';
        else if (connType === 'Deconnecte') connClass = 'conn-deconnecte';
        var connCell = '<span class="conn-badge ' + connClass + '">' + connType + '</span>';

        var cpuAgeCategory = pc.CPUAgeCategory || 'Inconnu';
        var cpuClass = 'cpu-inconnu';
        if (cpuAgeCategory === 'Recent') cpuClass = 'cpu-recent';
        else if (cpuAgeCategory === 'Vieillissant') cpuClass = 'cpu-vieillissant';
        else if (cpuAgeCategory === 'Ancien') cpuClass = 'cpu-ancien';
        var cpuLabel = cpuAgeCategory;
        if (pc.CPUYear) cpuLabel += ' (' + pc.CPUYear + ')';
        var cpuCell = '<span class="cpu-badge ' + cpuClass + '" title="' + (pc.CPUName || '') + '">' + cpuLabel + '</span>';

        // v5.8 : badge chassis sous le CPU (laptop/desktop/AIO)
        if (pc.ChassisInfo && pc.ChassisInfo.ChassisLabel && pc.ChassisInfo.ChassisLabel !== 'Inconnu') {
            var chCls = 'chassis-autre', chIcon = '&#128221;';
            if (pc.ChassisInfo.IsLaptop)       { chCls = 'chassis-laptop';  chIcon = '&#128187;'; }
            else if (pc.ChassisInfo.IsAIO)     { chCls = 'chassis-aio';     chIcon = '&#128444;'; }
            else if (pc.ChassisInfo.IsDesktop) { chCls = 'chassis-desktop'; chIcon = '&#128421;'; }
            cpuCell += '<div style="margin-top:3px"><span class="chassis-badge ' + chCls + '" title="Type de chassis">' +
                       chIcon + ' ' + pc.ChassisInfo.ChassisLabel + '</span></div>';
        }

        var bootCell = 'N/A';
        if (p.dernierBoot) {
            var bootClass = 'boot-ok';
            if (p.dernierBoot.DurationMin > 5) bootClass = 'boot-danger';
            else if (p.dernierBoot.EstBootLong) bootClass = 'boot-warning';
            bootCell = '<span class="boot-badge ' + bootClass + '">' + p.dernierBoot.DurationMin + ' min</span><div class="boot-date">' + p.dernierBoot.DateBoot + '</div>';
        }

        var diskCell = '';
        if (p.diskInfo.length > 0) {
            var worstDisk = p.diskInfo.reduce(function(a, b) { return a.PctFree < b.PctFree ? a : b; });
            var diskColor = 'color-green';
            if (worstDisk.IsAlert) diskColor = 'color-red';
            else if (worstDisk.PctFree < seuilDiskWarning) diskColor = 'color-orange';
            diskCell = '<span class="' + diskColor + '" style="font-weight:700">' + worstDisk.PctFree + '% libre</span><div style="font-size:10px;color:var(--text-faint)">' + worstDisk.FreeGB + ' GB</div>';
        } else {
            diskCell = '<span style="color:var(--text-ghost)">N/A</span>';
        }

        var hwCell = p.hwCount > 0
            ? '<span class="kpi-hardware">' + p.hwCount + '</span>'
            : '<span style="color:var(--text-ghost)">-</span>';

        // ==============================================================
        // COLONNE INDICATEURS (v5.3 / v5.4)
        // 4 pastilles compactes : Batt / EDR / Boot / SMART
        // Etats : ok (vert) / warn (orange) / ko (rouge) / na (gris)
        // ==============================================================
        var indB, indE, indBP, indS;
        // Batterie
        if (!p.battery || !p.battery.HasBattery) {
            indB = '<span class="indicator-dot na" title="Pas de batterie (desktop/VM)">B</span>';
        } else if (p.batteryAlert) {
            indB = '<span class="indicator-dot ko" title="Batterie us&eacute;e : ' + p.battery.HealthPercent + '% (' + p.battery.HealthCategory + ')">B</span>';
        } else {
            indB = '<span class="indicator-dot ok" title="Batterie OK : ' + p.battery.HealthPercent + '% (' + p.battery.HealthCategory + ')">B</span>';
        }
        // EDR SentinelOne
        if (!p.sentinel) {
            indE = '<span class="indicator-dot na" title="Donn&eacute;e EDR absente (Collector &lt; v5.3)">E</span>';
        } else if (!p.sentinel.Installed) {
            indE = '<span class="indicator-dot ko" title="SentinelOne NON INSTALLE">E</span>';
        } else if (p.sentinelAlert) {
            indE = '<span class="indicator-dot ko" title="SentinelAgent : ' + p.sentinel.Status + '">E</span>';
        } else {
            indE = '<span class="indicator-dot ok" title="SentinelAgent : Running">E</span>';
        }
        // Boot Performance
        if (!p.bootPerf || !p.bootPerf.LastBoot) {
            indBP = '<span class="indicator-dot na" title="Pas de donn&eacute;e boot perf (aucun cold boot r&eacute;cent ou Collector &lt; v5.4)">P</span>';
        } else if (p.bootPerfAlert) {
            var lb = p.bootPerf.LastBoot;
            indBP = '<span class="indicator-dot ko" title="Boot lent : ' + (lb.BootTimeMs/1000).toFixed(1) + 's (Post-boot ' + (lb.BootPostBootTimeMs/1000).toFixed(1) + 's)">P</span>';
        } else {
            indBP = '<span class="indicator-dot ok" title="Boot OK : ' + (p.bootPerf.LastBoot.BootTimeMs/1000).toFixed(1) + 's">P</span>';
        }
        // SMART
        if (p.diskHealth.length === 0) {
            indS = '<span class="indicator-dot na" title="Donn&eacute;es SMART absentes (Collector &lt; v5.4)">S</span>';
        } else if (p.diskSmartAlert) {
            var reasons = [];
            p.diskSmartAlerts.forEach(function(d) { reasons.push(d.FriendlyName + ' (' + d.AlertReasons.join(', ') + ')'); });
            indS = '<span class="indicator-dot ko" title="SMART alerte : ' + reasons.join(' | ') + '">S</span>';
        } else {
            indS = '<span class="indicator-dot ok" title="SMART OK (' + p.diskHealth.length + ' disque' + (p.diskHealth.length > 1 ? 's' : '') + ')">S</span>';
        }
        var indicatorsCell = '<div class="indicator-row">' + indB + indE + indBP + indS + '</div>';

        tbody += '<tr class="' + rowClass + ' row-main" onclick="toggleDetail(\'detail-' + idx + '\')">' +
            '<td>' + scoreCell + '</td>' +
            '<td><span class="toggle-icon">&#9654;</span>' + pc.PC + '</td>' +
            '<td>' + statusBadge + '</td>' +
            siteCell +
            '<td>' + pc.IP + '</td>' +
            '<td>' + (pc.CurrentUser || '') + '</td>' +
            '<td>' + connCell + '</td>' +
            '<td>' + cpuCell + '</td>' +
            '<td>' + uptimeCell + '</td>' +
            '<td>' + freshCell + '</td>' +
            '<td>' + bootCell + '</td>' +
            '<td class="col-advanced"><span class="kpi-crash">' + p.crashCount + '</span>' + (p.bsodClassifieCount > 0 ? ' <span class="kpi-bsod" style="font-size:11px">(' + p.bsodClassifieCount + ' BSOD)</span>' : '') + '</td>' +
            '<td class="col-advanced"><span class="kpi-bsod">' + p.bsodCount + '</span></td>' +
            '<td class="col-advanced">' + hwCell + '</td>' +
            '<td class="col-advanced">' + diskCell + '</td>' +
            '<td class="col-advanced">' + (p.warningCount > 0 ? '<span class="kpi-warning">' + p.warningCount + '</span>' : '-') + '</td>' +
            '<td>' + indicatorsCell + '</td>' +
            '</tr>';

        tbody += renderDetailRow(p, idx, colspan);
    });

    if (pcs.length === 0) {
        tbody = '<tr><td colspan="' + colspan + '" style="text-align:center;padding:30px;color:var(--text-ghost);font-style:italic">Aucun PC ne correspond aux filtres actifs</td></tr>';
    }
    document.getElementById('tableBody').innerHTML = tbody;
}

function renderDetailRow(p, idx, colspan) {
    // v1.6 : libelles fins par CrashCause pour les Events 41.
    // Fallback sur les anciens libelles si CrashCause absent (JSON v1.4/1.5).
    function crashCauseLabel(cause) {
        if (cause === 'BSODSilent')        return 'BSOD silencieux';
        if (cause === 'SleepResumeFailed') return 'Reprise veille rat&eacute;e';
        if (cause === 'UserForcedReset')   return 'User bouton power';
        if (cause === 'PowerLoss')         return 'Coupure alim / thermal';
        if (cause === 'FreezeApp')         return 'Freeze applicatif';
        if (cause === 'FreezeUnknown')     return 'Freeze (cause inconnue)';
        return null;
    }

    // Crashes avec BugCheck symbolique + CrashCause v1.6
    var crashHTML = '';
    if (p.crashes.length > 0) {
        p.crashes.forEach(function(c) {
            var badge = '', detailSpan = '';
            var causeLbl = crashCauseLabel(c.CrashCause);

            if (c.Type === 'BSOD') {
                badge = '<span class="crash-badge bsod">BSOD</span>';
                detailSpan = '<span class="crash-stopcode">' + (c.Detail || '') + '</span>';
                var bugName = bugCheckName(c.Detail);
                if (bugName) detailSpan += '<span class="crash-bugname">' + bugName + '</span>';
            } else if (c.Type === 'Hard reset') {
                badge = '<span class="crash-badge hard-reset">Hard reset</span>';
                // v1.6 : preciser la cause si disponible
                if (causeLbl) detailSpan = '<span class="crash-app">' + causeLbl + '</span>';
            } else if (c.Detail) {
                badge = '<span class="crash-badge freeze-app">Freeze</span>';
                detailSpan = '<span class="crash-app">' + c.Detail + '</span>';
            } else {
                badge = '<span class="crash-badge freeze">Freeze</span>';
                if (causeLbl) detailSpan = '<span class="crash-app">' + causeLbl + '</span>';
            }
            crashHTML += '<div class="detail-item">' + badge + detailSpan + '<span class="detail-date">' + c.Timestamp + ' <span class="detail-ago">(' + timeAgo(c.Timestamp) + ')</span></span></div>';
        });
    } else {
        crashHTML = '<div class="detail-empty">Aucun crash sur la periode</div>';
    }

    // ========================================================
    // DRILL-DOWN DEMARRAGES (v5.2)
    //   - Mini-repartition par type (ColdBoot / FastStartup / Resume)
    //   - Liste des 8 derniers demarrages avec badge du type
    //   - Mise en evidence des boots longs (>= seuilBootLong)
    // ========================================================
    var bootLongHTML = '';

    // Ligne de repartition par type (seulement si on a des Cold/Fast/Resume)
    var bt = p.bootsByType || {};
    var nbCold = bt.ColdBoot || 0, nbFast = bt.FastStartup || 0, nbResume = bt.Resume || 0;
    var totalByType = nbCold + nbFast + nbResume;
    if (totalByType > 0) {
        bootLongHTML += '<div style="display:flex;gap:6px;flex-wrap:wrap;margin-bottom:10px;font-size:10px">';
        if (nbCold > 0)   bootLongHTML += '<span class="boot-type-chip boot-cold" title="Arret complet + demarrage">&#10052; Cold &times;' + nbCold + '</span>';
        if (nbFast > 0)   bootLongHTML += '<span class="boot-type-chip boot-fast" title="Fast Startup (hybrid boot)">&#9889; Fast &times;' + nbFast + '</span>';
        if (nbResume > 0) bootLongHTML += '<span class="boot-type-chip boot-resume" title="Reprise depuis hibernation">&#128164; Resume &times;' + nbResume + '</span>';
        bootLongHTML += '</div>';
    }

    // Liste des 8 derniers demarrages (tous types)
    if (p.boots.length > 0) {
        var recentBoots = p.boots.slice().reverse().slice(0, 8);
        recentBoots.forEach(function(b) {
            var typeBadge = '';
            var t = b.BootType || 'Unknown';
            if      (t === 'ColdBoot')    typeBadge = '<span class="boot-type-chip boot-cold" title="Cold boot">&#10052;</span>';
            else if (t === 'FastStartup') typeBadge = '<span class="boot-type-chip boot-fast" title="Fast Startup">&#9889;</span>';
            else if (t === 'Resume')      typeBadge = '<span class="boot-type-chip boot-resume" title="Resume hibernation">&#128164;</span>';
            else                          typeBadge = '<span class="boot-type-chip boot-unknown" title="Type inconnu">?</span>';

            var durClass = b.EstBootLong ? 'detail-dur' : 'detail-info';
            bootLongHTML += '<div class="detail-item">' +
                typeBadge +
                '<span class="detail-date" style="flex:1">' + b.DateBoot + '</span>' +
                '<span class="' + durClass + '">' + b.DurationMin + ' min</span>' +
                '</div>';
        });
        if (p.boots.length > 8) {
            bootLongHTML += '<div class="detail-empty">... et ' + (p.boots.length - 8) + ' autre(s)</div>';
        }
    } else {
        bootLongHTML += '<div class="detail-empty">Aucun demarrage detecte</div>';
    }

    // v1.6 : bloc CPU pour le panel Materiel
    var cpuHTML = '';
    var cpuCat = p.pc.CPUAgeCategory || 'Inconnu';
    var cpuBadgeCls = 'ok';
    if (cpuCat === 'Ancien')       cpuBadgeCls = 'danger';
    else if (cpuCat === 'Vieillissant') cpuBadgeCls = 'warning';
    else if (cpuCat === 'Recent')  cpuBadgeCls = 'ok';
    var cpuMetaParts = [];
    if (p.pc.CPUVendor) cpuMetaParts.push(p.pc.CPUVendor);
    if (p.pc.CPUGen)    cpuMetaParts.push('Gen ' + p.pc.CPUGen);
    if (p.pc.CPUYear)   cpuMetaParts.push(p.pc.CPUYear);
    if (p.pc.CPUAge !== null && p.pc.CPUAge !== undefined) cpuMetaParts.push(p.pc.CPUAge + ' an(s)');
    var cpuMetaStr = cpuMetaParts.join(' &middot; ');
    if (p.pc.CPUName) {
        cpuHTML = '<div class="cpu-info-block">' +
                    '<div class="cpu-name">' + p.pc.CPUName + '</div>' +
                    '<div class="cpu-meta">' +
                      (cpuMetaStr ? '<span>' + cpuMetaStr + '</span>' : '') +
                      '<span class="cpu-badge-inline ' + cpuBadgeCls + '">' + cpuCat + '</span>' +
                    '</div>' +
                  '</div>';
    } else {
        cpuHTML = '<div class="detail-empty">Nom CPU non disponible</div>';
    }

    var diskHTML = '';
    if (p.diskInfo.length > 0) {
        p.diskInfo.forEach(function(d) {
            var fillClass = 'ok';
            if (d.IsAlert) fillClass = 'danger';
            else if (d.PctFree < seuilDiskWarning) fillClass = 'warning';
            diskHTML += '<div class="disk-bar-wrap"><span class="disk-drive">' + d.Drive + '</span><div class="disk-bar"><div class="disk-bar-fill ' + fillClass + '" style="width:' + d.PctUsed + '%"></div></div><span class="disk-info-text">' + d.FreeGB + ' GB (' + d.PctFree + '%)</span></div>';
        });
    } else {
        diskHTML = '<div class="detail-empty">Pas d info disque</div>';
    }

    var crasherHTML = '';
    // v1.7 : on filtre les blacklist HARD (invisibles) et marque les SOFT (grises).
    var pcCrashers = (p.topCrashers || []).filter(function(tc) {
        return !isBlacklistedHard(tc.AppName);
    });
    if (pcCrashers.length > 0) {
        pcCrashers.slice(0, 5).forEach(function(tc) {
            var isSoft = isBlacklistedSoft(tc.AppName);
            var rowCls = isSoft ? 'crasher-item crasher-soft-noise' : 'crasher-item';
            var noiseTag = isSoft ? ' <span class="crasher-noise-tag" title="Crasher connu - bruit ambient">bruit</span>' : '';
            crasherHTML += '<div class="' + rowCls + '">' +
                             '<span class="crasher-name">' + tc.AppName + '</span> ' +
                             '<span class="crasher-count">(' + tc.CrashCount + ')</span>' +
                             noiseTag +
                           '</div>';
        });
    } else {
        crasherHTML = '<div class="detail-empty">Aucun crash app</div>';
    }

    // ========================================================
    // DRILL-DOWN HARDWARE (v5.2)
    //   Section 1 : Erreurs FATALES (WHEA Fatal + GPU TDR + Thermal)
    //     → une ligne par occurrence, rouge/violet selon gravite
    //   Section 2 : Erreurs CORRIGEES (WHEA Corrected)
    //     → une ligne par signature, avec compteur d'occurrences
    //     (la plupart des parcs en generent massivement sans que
    //      ca reflete un vrai probleme)
    // ========================================================
    var hwHTML = '';
    var hasFatal = (p.wheaFatal.length + p.gpuTDR.length + p.thermal.length) > 0;
    var hasCorrected = p.wheaCorrected.length > 0;

    if (hasFatal) {
        hwHTML += '<div style="font-size:10px;color:var(--red);text-transform:uppercase;letter-spacing:0.5px;margin-bottom:6px;font-weight:600">Fatales (' + (p.wheaFatal.length + p.gpuTDR.length + p.thermal.length) + ')</div>';

        // WHEA fatales : classification par composant
        p.wheaFatal.forEach(function(h) {
            var comp = (h.Component || 'Autre').toLowerCase();
            var badgeClass = 'whea-' + comp;
            if (comp !== 'cpu' && comp !== 'ram' && comp !== 'pcie') badgeClass = 'whea-cpu'; // fallback style
            var detail = h.ErrorSource || '';
            if (h.BDF) detail += (detail ? ' · ' : '') + h.BDF;
            hwHTML += '<div class="detail-item">' +
                '<span class="hw-badge ' + badgeClass + '">' + (h.Component || '?') + '</span>' +
                '<span class="detail-info">' + detail + '</span>' +
                '<span class="detail-date">' + h.Timestamp + '</span></div>';
        });

        // GPU TDR
        p.gpuTDR.forEach(function(h) {
            hwHTML += '<div class="detail-item">' +
                '<span class="hw-badge gpu-tdr">GPU</span>' +
                '<span class="detail-info">' + (h.Driver || '') + '</span>' +
                '<span class="detail-date">' + h.Timestamp + '</span></div>';
        });

        // Thermal
        p.thermal.forEach(function(h) {
            hwHTML += '<div class="detail-item">' +
                '<span class="hw-badge thermal">' + h.AlertType + '</span>' +
                '<span class="detail-info">' + (h.Temperature || '') + '</span>' +
                '<span class="detail-date">' + h.Timestamp + '</span></div>';
        });
    }

    if (hasCorrected) {
        if (hasFatal) hwHTML += '<div style="margin-top:12px"></div>';
        hwHTML += '<div style="font-size:10px;color:var(--text-muted);text-transform:uppercase;letter-spacing:0.5px;margin-bottom:6px;font-weight:600">Corrigees (' + p.wheaCorrectedTotal + ' occurrences)</div>';

        // Top 5 des signatures les plus frequentes
        p.wheaCorrected.slice(0, 5).forEach(function(h) {
            var comp = (h.Component || 'Autre').toLowerCase();
            var badgeClass = 'whea-' + comp;
            if (comp !== 'cpu' && comp !== 'ram' && comp !== 'pcie') badgeClass = 'whea-cpu';
            var detail = h.ErrorSource || '';
            if (h.BDF) detail += (detail ? ' · ' : '') + h.BDF;
            hwHTML += '<div class="detail-item">' +
                '<span class="hw-badge ' + badgeClass + '" style="opacity:0.7">' + (h.Component || '?') + '</span>' +
                '<span class="detail-info" style="flex:1">' + detail + '</span>' +
                '<span class="crasher-count">&times;' + h.Count + '</span>' +
                '<span class="detail-date">' + h.LastSeen + '</span></div>';
        });
        if (p.wheaCorrected.length > 5) {
            hwHTML += '<div class="detail-empty">... et ' + (p.wheaCorrected.length - 5) + ' autre(s) signature(s)</div>';
        }
    }

    if (!hasFatal && !hasCorrected) {
        hwHTML = '<div class="detail-empty">Aucune erreur materielle</div>';
    }

    var perfHTML = '';
    if (p.warnings.length > 0) {
        p.warnings.forEach(function(w) {
            var badge = '', badgeClass = '';
            if (w.Type === 'RAM exhaustion')      { badgeClass = 'ram-exhaustion'; badge = 'RAM'; }
            else if (w.Type === 'CPU throttling') { badgeClass = 'cpu-throttling'; badge = 'CPU'; }
            else if (w.Type === 'Disk full')      { badgeClass = 'disk-full';      badge = 'DISK'; }
            else if (w.Type === 'Disk slow')      { badgeClass = 'disk-slow';      badge = 'I/O'; }
            // v1.5 : affichage enrichi avec Count + burst (backward-compat v1.4-)
            var detailStr = w.Detail || '';
            var countStr  = '';
            var burstStr  = '';
            if (typeof w.Count === 'number' && w.Count > 1) {
                var durSec = 0;
                if (w.FirstSeen && w.LastSeen) {
                    try { durSec = Math.round((parseDate(w.LastSeen) - parseDate(w.FirstSeen)) / 1000); } catch(e) {}
                }
                countStr = ' &times;' + w.Count + ' events';
                if (durSec <= 1) countStr += ' en 1s';
                else if (durSec < 60) countStr += ' en ' + durSec + 's';
                else countStr += ' en ' + Math.round(durSec / 60) + 'min';
            }
            if (w.IsBurst === true) {
                burstStr = ' <span class="burst-badge" title="Burst massif : probable incident mat\u00e9riel">&#128293; BURST</span>';
            }
            perfHTML += '<div class="detail-item"><span class="warn-badge ' + badgeClass + '">' + badge + '</span><span class="warn-detail">' + detailStr + countStr + burstStr + '</span><span class="detail-date">' + w.Timestamp + '</span></div>';
        });
    }
    if (p.topRAM.length > 0) {
        if (perfHTML) perfHTML += '<div style="margin-top:10px;margin-bottom:6px;font-size:10px;color:var(--text-faint)">Top RAM</div>';
        p.topRAM.forEach(function(proc) {
            perfHTML += '<div class="detail-item"><span class="detail-info">' + proc.Name + '</span><span class="detail-dur">' + proc.WorkingSetMB + ' MB</span></div>';
        });
    }
    if (!perfHTML) perfHTML = '<div class="detail-empty">Aucun warning</div>';

    // ========================================================
    // DRILL-DOWN BATTERIE + EDR (v5.3)
    // On combine les deux dans une section pour economiser la place
    // ========================================================
    var batteryHTML = '';
    if (!p.battery) {
        batteryHTML = '<div class="detail-empty">Donn&eacute;es batterie absentes (Collector &lt; v5.3)</div>';
    } else if (!p.battery.HasBattery) {
        batteryHTML = '<div class="detail-empty">Pas de batterie d&eacute;tect&eacute;e (desktop ou VM)</div>';
    } else {
        var health = p.battery.HealthPercent || 0;
        var barClass = 'good';
        if (health < 60) barClass = 'danger';
        else if (health < 80) barClass = 'warning';

        batteryHTML += '<div class="batt-bar-wrap">' +
            '<div class="batt-bar"><div class="batt-bar-fill ' + barClass + '" style="width:' + Math.min(health, 100) + '%"></div></div>' +
            '<strong style="color:var(--text-dim);min-width:55px;text-align:right">' + health + '%</strong>' +
            '</div>';

        batteryHTML += '<div class="batt-meta">' +
            (p.battery.HealthCategory ? 'Etat : <strong>' + p.battery.HealthCategory + '</strong> &middot; ' : '') +
            'Charge : <strong>' + p.battery.CurrentChargePct + '%</strong> (' + p.battery.Status + ')' +
            '</div>';

        var metaParts = [];
        if (p.battery.Manufacturer)       metaParts.push('Fab: <strong>' + p.battery.Manufacturer + '</strong>');
        if (p.battery.Chemistry)          metaParts.push('Chimie: <strong>' + p.battery.Chemistry + '</strong>');
        if (p.battery.CycleCount != null) metaParts.push('Cycles: <strong>' + p.battery.CycleCount + '</strong>');
        if (p.battery.DesignCapacity)     metaParts.push('Design: <strong>' + Math.round(p.battery.DesignCapacity / 1000) + ' Wh</strong>');
        if (p.battery.FullChargeCapacity) metaParts.push('Actuel: <strong>' + Math.round(p.battery.FullChargeCapacity / 1000) + ' Wh</strong>');
        if (metaParts.length > 0) {
            batteryHTML += '<div class="batt-meta" style="margin-top:6px">' + metaParts.join(' &middot; ') + '</div>';
        }
    }

    // EDR SentinelOne dans la meme section
    var edrHTML = '';
    if (!p.sentinel) {
        edrHTML = '<div class="detail-empty">Donn&eacute;es EDR absentes (Collector &lt; v5.3)</div>';
    } else {
        var edrBadge, edrBadgeClass;
        if (!p.sentinel.Installed) {
            edrBadge = 'NON INSTALL&Eacute;'; edrBadgeClass = 'bsod';
        } else if (p.sentinelAlert) {
            edrBadge = p.sentinel.Status;     edrBadgeClass = 'hard-reset';
        } else {
            edrBadge = 'Running';              edrBadgeClass = 'freeze-app';
        }
        edrHTML = '<div class="detail-item">' +
            '<span class="crash-badge ' + edrBadgeClass + '">SentinelOne</span>' +
            '<span class="detail-info" style="flex:1">' + p.sentinel.DisplayName + ' (' + p.sentinel.ServiceName + ')</span>' +
            '<span class="detail-date">' + edrBadge + '</span>' +
            '</div>';
        if (p.sentinel.Installed) {
            edrHTML += '<div class="batt-meta" style="margin-top:6px">Type d&eacute;marrage : <strong>' + p.sentinel.StartType + '</strong></div>';
        }
    }

    // ========================================================
    // DRILL-DOWN BOOT PERFORMANCE (v5.4)
    // - Phases du dernier boot en barres proportionnelles
    // - Historique des 5 derniers boots
    // ========================================================
    var bootPerfHTML = '';
    if (!p.bootPerf || !p.bootPerf.LastBoot) {
        bootPerfHTML = '<div class="detail-empty">Aucune donn&eacute;e boot perf (aucun cold boot r&eacute;cent ou Collector &lt; v5.4)</div>';
    } else {
        var lb = p.bootPerf.LastBoot;
        var totalMs = Math.max(lb.BootTimeMs, 1);

        // Phases breakdown
        bootPerfHTML += '<div style="font-size:10px;color:var(--text-muted);text-transform:uppercase;letter-spacing:0.5px;margin-bottom:6px;font-weight:600">Dernier boot &mdash; ' + (lb.BootTimeMs/1000).toFixed(1) + 's total</div>';

        bootPerfHTML += '<div class="bootperf-phases">';
        function phase(label, ms, slowThresh) {
            var pct = Math.min((ms / totalMs) * 100, 100);
            var cls = 'bootperf-bar-fill';
            if (slowThresh && ms > slowThresh) cls += ' slow';
            else if (ms < 10000) cls += ' fast';
            return '<div class="bootperf-phase">' +
                   '<span class="bootperf-label">' + label + '</span>' +
                   '<div class="bootperf-bar"><div class="' + cls + '" style="width:' + pct + '%"></div></div>' +
                   '<span class="bootperf-value">' + (ms/1000).toFixed(1) + 's</span>' +
                   '</div>';
        }
        bootPerfHTML += phase('MainPath (OS&rarr;Logon)', lb.MainPathBootTimeMs, 60000);
        bootPerfHTML += phase('PostBoot (Logon&rarr;Idle)', lb.BootPostBootTimeMs, 90000);
        bootPerfHTML += phase('Profil utilisateur', lb.UserProfileProcessingTimeMs, 10000);
        bootPerfHTML += phase('Init Explorer', lb.ExplorerInitTimeMs, 15000);
        bootPerfHTML += '</div>';

        var bootMeta = [];
        bootMeta.push('Apps d&eacute;marrage: <strong>' + lb.NumStartupApps + '</strong>');
        bootMeta.push('Niveau Win: <strong>' + (lb.Level || '?') + '</strong>');
        if (lb.IsRebootAfterInstall) bootMeta.push('<strong style="color:var(--orange)">Post-MAJ Windows</strong>');
        bootPerfHTML += '<div class="bootperf-meta">' + bootMeta.join(' &middot; ') + '</div>';

        // Historique (skip le premier qui est LastBoot)
        if (p.bootPerf.History && p.bootPerf.History.length > 1) {
            bootPerfHTML += '<div style="margin-top:12px;font-size:10px;color:var(--text-muted);text-transform:uppercase;letter-spacing:0.5px;margin-bottom:4px;font-weight:600">Historique (' + (p.bootPerf.History.length) + ' boots)</div>';
            p.bootPerf.History.forEach(function(h, i) {
                if (i === 0) return; // deja affiche en "dernier boot"
                var rowCls = h.IsSlow ? 'bootperf-row slow' : 'bootperf-row';
                bootPerfHTML += '<div class="' + rowCls + '">' +
                    '<span class="date">' + h.Timestamp + '</span>' +
                    '<span>Main ' + (h.MainPathBootTimeMs/1000).toFixed(1) + 's</span>' +
                    '<span>Post ' + (h.BootPostBootTimeMs/1000).toFixed(1) + 's</span>' +
                    '<span class="total">' + (h.BootTimeMs/1000).toFixed(1) + 's</span>' +
                    '</div>';
            });
        }

        // Stats agregees
        if (p.bootPerf.Stats) {
            var st = p.bootPerf.Stats;
            bootPerfHTML += '<div class="bootperf-meta" style="margin-top:10px">' +
                'Moyennes : Main <strong>' + (st.AvgMainPathMs/1000).toFixed(1) + 's</strong> &middot; ' +
                'Post <strong>' + (st.AvgPostBootMs/1000).toFixed(1) + 's</strong> &middot; ' +
                'Total <strong>' + (st.AvgBootTimeMs/1000).toFixed(1) + 's</strong>' +
                ' &middot; Slow boots : <strong>' + st.SlowBootsCount + '/' + st.BootsAnalyzed + '</strong>' +
                '</div>';
        }
    }

    // ========================================================
    // DRILL-DOWN SMART (v5.4)
    // Une carte par disque physique avec les compteurs SMART
    // ========================================================
    var smartHTML = '';
    if (!p.diskHealth || p.diskHealth.length === 0) {
        smartHTML = '<div class="detail-empty">Aucune donn&eacute;e SMART (Collector &lt; v5.4)</div>';
    } else {
        p.diskHealth.forEach(function(d) {
            var healthBadgeCls = 'healthy';
            var healthLabel = d.HealthStatus || 'Unknown';
            if (healthLabel === 'Warning')          healthBadgeCls = 'warning';
            else if (healthLabel !== 'Healthy')     healthBadgeCls = 'unhealthy';

            function kv(label, val, cls) {
                var cc = cls || '';
                if (val === null || val === undefined || val === '')
                    return '<div class="smart-kv"><span class="k">' + label + '</span><span class="v na">N/A</span></div>';
                return '<div class="smart-kv"><span class="k">' + label + '</span><span class="v ' + cc + '">' + val + '</span></div>';
            }

            // Couleur temperature
            var tempCls = '';
            if (d.TemperatureC != null) {
                if (d.TemperatureC >= 70)      tempCls = 'ko';
                else if (d.TemperatureC >= 55) tempCls = 'warn';
                else                           tempCls = 'ok';
            }
            // Couleur wear
            var wearCls = '';
            if (d.WearPct != null) {
                if (d.WearPct >= 70)      wearCls = 'ko';
                else if (d.WearPct >= 40) wearCls = 'warn';
                else                      wearCls = 'ok';
            }

            smartHTML += '<div class="smart-disk">' +
                '<div class="smart-disk-head">' +
                  '<div>' +
                    '<div class="smart-disk-name">' + (d.FriendlyName || '?') + '</div>' +
                    '<div class="smart-disk-sub">' + (d.MediaType || '?') + ' &middot; ' + (d.BusType || '?') + ' &middot; ' + (d.SizeGB || 0) + ' GB</div>' +
                  '</div>' +
                  '<span class="smart-health-badge ' + healthBadgeCls + '">' + healthLabel + '</span>' +
                '</div>' +
                '<div class="smart-grid">' +
                  kv('Temp&eacute;rature', d.TemperatureC != null ? d.TemperatureC + '&deg;C' : null, tempCls) +
                  kv('Temp max historique', d.TemperatureMaxC != null ? d.TemperatureMaxC + '&deg;C' : null) +
                  kv('Wear', d.WearPct != null ? d.WearPct + '%' : null, wearCls) +
                  kv('Power-On Hours', d.PowerOnHours != null ? d.PowerOnHours + 'h' : null) +
                  kv('Read errors', d.ReadErrorsTotal) +
                  kv('Read uncorr.', d.ReadErrorsUncorrected, (d.ReadErrorsUncorrected > 0 ? 'ko' : '')) +
                  kv('Write errors', d.WriteErrorsTotal) +
                  kv('Write uncorr.', d.WriteErrorsUncorrected, (d.WriteErrorsUncorrected > 0 ? 'ko' : '')) +
                '</div>';
            if (d.AlertReasons && d.AlertReasons.length > 0) {
                smartHTML += '<div class="smart-alerts">';
                d.AlertReasons.forEach(function(r) {
                    smartHTML += '<span class="smart-alert-chip">' + r + '</span>';
                });
                smartHTML += '</div>';
            }
            smartHTML += '</div>';
        });
    }

    // ========================================================
    // MONITORS (v5.7) : inventaire ecrans externes du PC
    // Chaque ecran = une petite carte avec nom, fab, serie, age.
    // Les internes (laptop LCD) sont exclus par le Collector.
    // ========================================================
    var monitorsHTML = '';
    var monitorsList = p.monitors || [];
    if (monitorsList.length === 0) {
        monitorsHTML = '<div class="detail-empty">Aucun &eacute;cran secondaire branch&eacute; (ou Collector &lt; v5.5)</div>';
    } else {
        monitorsList.forEach(function(mon) {
            var isOld = (mon.AgeYears !== null && mon.AgeYears !== undefined && mon.AgeYears >= 7);
            var cardCls = 'monitor-card' + (isOld ? ' old' : '');
            var displayName = mon.Model || mon.ProductCode || '(mod&egrave;le inconnu)';
            var displayManuf = mon.Manufacturer || mon.ManufacturerCode || '?';

            var badges = '';
            if (mon.Active)  badges += '<span class="mon-badge active">Actif</span>';
            else             badges += '<span class="mon-badge inactive">D&eacute;branch&eacute;</span>';
            if (isOld)       badges += '<span class="mon-badge old">' + mon.AgeYears + ' ans</span>';

            var metaParts = [];
            if (mon.SerialNumber) metaParts.push('S/N : <strong>' + mon.SerialNumber + '</strong>');
            if (mon.YearOfManufacture) {
                var dateStr = mon.YearOfManufacture;
                if (mon.WeekOfManufacture) dateStr += ' S' + mon.WeekOfManufacture;
                metaParts.push('Fabriqu&eacute; : <strong>' + dateStr + '</strong>');
            }
            if (mon.ProductCode && mon.ProductCode !== mon.Model) {
                metaParts.push('Code : <strong>' + mon.ProductCode + '</strong>');
            }

            monitorsHTML += '<div class="' + cardCls + '">' +
                '<div class="monitor-head">' +
                  '<div class="monitor-name">' + displayName +
                    '<div class="monitor-manuf">' + displayManuf + '</div>' +
                  '</div>' +
                  '<div class="monitor-badges">' + badges + '</div>' +
                '</div>' +
                (metaParts.length > 0 ? '<div class="monitor-meta">' + metaParts.join(' &middot; ') + '</div>' : '') +
                '</div>';
        });
    }

    // ========================================================
    // v1.8 : RAM INVENTORY
    // ========================================================
    var ramHTML = '';
    if (!p.memory) {
        ramHTML = '<div class="detail-empty">Donn&eacute;es non disponibles (Collector &lt; v1.8)</div>';
    } else {
        var mem = p.memory;
        // Ligne de synthese (totale + slots + max + upgrade possible)
        var upgradeTag = mem.CanUpgrade
            ? '<span class="ram-badge upgrade-ok">Upgrade possible</span>'
            : '<span class="ram-badge upgrade-no">Max atteint</span>';
        ramHTML = '<div class="ram-summary">' +
                    '<div class="ram-summary-total">' +
                      '<span class="ram-big">' + mem.TotalInstalledGB + '</span>' +
                      '<span class="ram-unit">Go</span>' +
                      '<span class="ram-summary-slots">installes sur ' + mem.MaxCapacityGB + ' Go max</span>' +
                    '</div>' +
                    '<div class="ram-summary-meta">' +
                      '<span>' + mem.OccupiedSlots + '/' + mem.TotalSlots + ' slots</span>' +
                      (mem.FreeSlots > 0 ? '<span class="ram-meta-ok">' + mem.FreeSlots + ' libre(s)</span>' : '') +
                      upgradeTag +
                    '</div>' +
                  '</div>';
        // Detail des barrettes
        if (mem.Modules && mem.Modules.length > 0) {
            ramHTML += '<div class="ram-modules">';
            mem.Modules.forEach(function(mod) {
                var metaParts = [];
                if (mod.Type)         metaParts.push(mod.Type);
                if (mod.SpeedMHz > 0) metaParts.push(mod.SpeedMHz + ' MHz');
                if (mod.Manufacturer) metaParts.push(mod.Manufacturer);
                ramHTML += '<div class="ram-module-row">' +
                             '<span class="ram-slot-name">' + (mod.Slot || '?') + '</span>' +
                             '<span class="ram-capacity">' + mod.CapacityGB + ' Go</span>' +
                             '<span class="ram-module-meta">' + metaParts.join(' &middot; ') + '</span>' +
                           '</div>';
            });
            ramHTML += '</div>';
        }
    }

    // ========================================================
    // v1.8 : GPU INVENTORY
    // ========================================================
    var gpuHTML = '';
    var gpuList = p.gpuInventory || [];
    if (gpuList.length === 0) {
        gpuHTML = '<div class="detail-empty">Donn&eacute;es non disponibles (Collector &lt; v1.8)</div>';
    } else {
        gpuList.forEach(function(g) {
            // Marquer le driver comme "vieux" si > 2 ans (simple heuristique)
            var oldDriver = false;
            if (g.DriverDate) {
                try {
                    var dd = parseDate(g.DriverDate + ' 00:00:00');
                    var yearsDiff = (generatedAt - dd) / (365.25 * 24 * 3600 * 1000);
                    if (yearsDiff > 2) oldDriver = true;
                } catch (e) {}
            }
            var driverCls = oldDriver ? 'gpu-driver-old' : 'gpu-driver-ok';
            var metaParts = [];
            if (g.DriverVersion) metaParts.push('v' + g.DriverVersion);
            if (g.DriverDate)    metaParts.push(g.DriverDate);
            gpuHTML += '<div class="gpu-row">' +
                         '<div class="gpu-name">' + (g.Name || '(GPU inconnu)') + '</div>' +
                         '<div class="gpu-meta ' + driverCls + '">' + metaParts.join(' &middot; ') +
                           (oldDriver ? ' <span class="gpu-old-tag">driver &gt;2 ans</span>' : '') +
                         '</div>' +
                       '</div>';
        });
    }

    // ========================================================
    // v1.8 : CPU THROTTLING (Kernel-Processor-Power 35/55)
    // ========================================================
    var throttleHTML = '';
    var throttleList = (p.hardwareHealth && p.hardwareHealth.CPUThrottling) ? p.hardwareHealth.CPUThrottling : [];
    var throttleDays = throttleList.length;
    // Calcul du Count max sur une seule journee (detecte un gros pic meme si 1 seul jour)
    var throttleMaxCount = 0;
    throttleList.forEach(function(t) { if (t.Count > throttleMaxCount) throttleMaxCount = t.Count; });

    if (throttleDays === 0) {
        throttleHTML = '<div class="detail-empty">Aucun throttling CPU d&eacute;tect&eacute;</div>';
    } else {
        // v1.8 : 4 niveaux de severite
        //   1-7 jours    => Normal usage intensif (info bleu)
        //   8-20 jours   => A surveiller (jaune)
        //   >20 jours OU >500 events/jour => Critique (rouge)
        var sevCls, sevLabel, sevIcon;
        if (throttleDays > 20 || throttleMaxCount > 500) {
            sevCls   = 'throttle-sev-critical';
            sevIcon  = '&#128680;';  // rotating red light
            sevLabel = sevIcon + ' ' + throttleDays + ' jours de throttling' +
                       (throttleMaxCount > 500 ? ' (pic &agrave; ' + throttleMaxCount + ' events en 1j)' : '') +
                       ' - refroidissement &agrave; v&eacute;rifier en priorit&eacute;';
        } else if (throttleDays >= 8) {
            sevCls   = 'throttle-sev-watch';
            sevIcon  = '&#9888;&#65039;';  // warning
            sevLabel = sevIcon + ' ' + throttleDays + ' jours de throttling - &agrave; surveiller';
        } else {
            sevCls   = 'throttle-sev-info';
            sevIcon  = '&#8505;&#65039;';  // info
            sevLabel = sevIcon + ' ' + throttleDays + ' jour(s) avec throttling (usage intensif normal)';
        }
        throttleHTML = '<div class="throttle-summary ' + sevCls + '">' + sevLabel + '</div>';
        // Liste des journees (Collector agrege deja par jour)
        throttleList.slice(0, 5).forEach(function(t) {
            var typeShort = t.EventId === 35 ? 'Firmware limit' :
                            t.EventId === 55 ? 'Thermal reduction' : ('Event ' + t.EventId);
            throttleHTML += '<div class="throttle-row">' +
                              '<span class="throttle-day">' + t.Day + '</span>' +
                              '<span class="throttle-type">' + typeShort + '</span>' +
                              '<span class="throttle-count">&#215;' + t.Count + ' events</span>' +
                            '</div>';
        });
        if (throttleList.length > 5) {
            throttleHTML += '<div class="throttle-more">+ ' + (throttleList.length - 5) + ' journees plus anciennes</div>';
        }
    }

    // ========================================================
    // v5.5 : VUE D'ENSEMBLE (onglet par defaut)
    // Liste les alertes actives en clair sans diluer l'info.
    // Si le PC est tout vert, message encourageant.
    // ========================================================
    function ovAlert(lvl, icon, title, meta) {
        return '<div class="overview-alert ' + lvl + '">' +
               '<span class="ov-icon">' + icon + '</span>' +
               '<span class="ov-text">' + title + '</span>' +
               '<span class="ov-meta">' + (meta || '') + '</span>' +
               '</div>';
    }
    var overview = [];
    if (p.pc.IsOffline) {
        overview.push(ovAlert('critical', '&#128268;', '<strong>Machine hors-ligne</strong> depuis ' + timeAgo(p.pc.CollectedAt), ''));
    }
    if (p.sentinelAlert) {
        var edrStatus = p.sentinel && p.sentinel.Installed ? p.sentinel.Status : 'NON INSTALL&Eacute;';
        overview.push(ovAlert('critical', '&#128737;', '<strong>EDR SentinelOne</strong> en probl&egrave;me', edrStatus));
    }
    if (p.bsodCount > 0) {
        overview.push(ovAlert('critical', '&#128165;', '<strong>' + p.bsodCount + ' BSOD</strong> sur la p&eacute;riode', ''));
    }
    if (p.crashCount > 0) {
        overview.push(ovAlert('warning', '&#9888;', '<strong>' + p.crashCount + ' crash/freeze</strong> sur la p&eacute;riode', ''));
    }
    if (p.hwCount > 0) {
        overview.push(ovAlert('critical', '&#128268;', '<strong>' + p.hwCount + ' erreur(s) mat&eacute;rielle(s) fatale(s)</strong>', ''));
    }
    if (p.diskSmartAlert) {
        var reasons = [];
        p.diskSmartAlerts.forEach(function(d) { reasons.push(d.FriendlyName + ' (' + d.AlertReasons.join(', ') + ')'); });
        overview.push(ovAlert('critical', '&#128190;', '<strong>SMART alerte</strong> sur disque', reasons.join(' | ')));
    }
    if (p.diskAlertCount > 0) {
        overview.push(ovAlert('warning', '&#128190;', '<strong>Disque(s) satur&eacute;(s)</strong>', p.diskAlertCount + ' volume(s)'));
    }
    if (p.bootPerfAlert) {
        var lb2 = (p.bootPerf && p.bootPerf.LastBoot) ? p.bootPerf.LastBoot : null;
        var bootMeta = lb2 ? ((lb2.BootTimeMs / 1000).toFixed(1) + 's au dernier cold boot') : '';
        overview.push(ovAlert('warning', '&#9201;', '<strong>Boots lents</strong> (&gt; 90s post-boot)', bootMeta));
    }
    if (p.bootLongCount > 0) {
        overview.push(ovAlert('warning', '&#9201;', '<strong>' + p.bootLongCount + ' boot(s) longs</strong> (&gt; seuil)', ''));
    }
    if (p.batteryAlert && p.battery) {
        overview.push(ovAlert('warning', '&#128267;', '<strong>Batterie us&eacute;e</strong>', p.battery.HealthPercent + '% de sa capacit&eacute; d\'origine'));
    }
    if (p.oldMonitorAlert && p.oldMonitors && p.oldMonitors.length > 0) {
        var oldest = p.oldMonitors.reduce(function(a, b) { return (a.AgeYears || 0) > (b.AgeYears || 0) ? a : b; });
        overview.push(ovAlert('warning', '&#128250;', '<strong>' + p.oldMonitors.length + ' &eacute;cran(s) &acirc;g&eacute;(s)</strong>', 'le plus vieux : ' + oldest.AgeYears + ' ans'));
    }
    if (p.wheaCorrected && p.wheaCorrected.length > 0) {
        overview.push(ovAlert('info', '&#9432;', p.wheaCorrectedTotal + ' erreur(s) mat&eacute;rielle(s) corrig&eacute;es (t&eacute;l&eacute;m&eacute;trie)', p.wheaCorrected.length + ' signature(s)'));
    }

    var overviewHTML;
    // v1.6 : bandeau verdict global calcule en tete, avant la liste d'alertes
    var verdict = computeVerdict(p);
    var verdictHTML = '<div class="verdict-banner ' + verdict.cls + '">' +
                        '<span class="v-icon">' + verdict.icon + '</span>' +
                        '<span class="v-label">' + verdict.label + '</span>' +
                        (verdict.reasons.length > 0
                          ? '<span class="v-reasons">&middot; ' + verdict.reasons.join(' &middot; ') + '</span>'
                          : '') +
                      '</div>';
    if (overview.length === 0) {
        overviewHTML = verdictHTML + '<div class="overview-empty"><strong>&#10003; Tout va bien</strong>Aucune alerte active sur cette p&eacute;riode</div>';
    } else {
        overviewHTML = verdictHTML + '<div class="overview-list">' + overview.join('') + '</div>';
    }

    // ========================================================
    // v5.5 : STRUCTURE EN ONGLETS
    // On garde les 10 sections dans 5 onglets thematiques + vue d'ensemble.
    // Les compteurs sur les onglets aident l'utilisateur a savoir
    // instantanement ou il y a un probleme.
    // ========================================================
    // Compteurs par onglet (pour les badges)
    var cntStability = p.crashCount + p.bsodCount + p.hwCount + (p.wheaCorrected ? p.wheaCorrected.length : 0);
    var cntBoot      = p.bootLongCount + (p.bootPerfAlert ? 1 : 0);
    var cntMaterial  = p.diskAlertCount + (p.diskSmartAlert ? p.diskSmartAlerts.length : 0) + (p.batteryAlert ? 1 : 0) + (p.oldMonitorAlert ? p.oldMonitors.length : 0);
    var cntSecurity  = (p.sentinelAlert ? 1 : 0);
    var cntOverview  = overview.length;

    function tabBtn(key, label, count, icon) {
        var active = (key === 'overview') ? ' active' : '';
        var badge = '';
        if (count > 0) {
            badge = '<span class="tab-badge">' + count + '</span>';
        } else if (count === 0) {
            // badge neutre "0" seulement pour onglets standards (pas vue d'ensemble)
            if (key !== 'overview') badge = '<span class="tab-badge quiet">0</span>';
        }
        return '<button class="detail-tab' + active + '" data-tab="' + key + '" onclick="selectDetailTab(' + idx + ', \'' + key + '\')">' +
               '<span>' + icon + '</span>' + label + badge + '</button>';
    }

    var tabsHTML =
        tabBtn('overview',  'Vue d\'ensemble', cntOverview,  '&#128270;') +
        tabBtn('stability', 'Stabilit&eacute;',cntStability,'&#128165;') +
        tabBtn('boot',      'D&eacute;marrage',cntBoot,     '&#9889;')   +
        tabBtn('material',  'Mat&eacute;riel', cntMaterial, '&#128295;') +
        tabBtn('security',  'S&eacute;curit&eacute;',cntSecurity,'&#128274;');

    function panel(key, inner, active) {
        var cls = active ? 'detail-tab-panel active' : 'detail-tab-panel';
        return '<div class="' + cls + '" data-panel="' + key + '">' + inner + '</div>';
    }

    var panelOverview = '<div class="detail-box">' +
        '<div class="detail-section" style="flex:1 1 100%;border-left:none;padding-left:0">' +
        overviewHTML + '</div></div>';

    // v1.6 : section Signaux croises (correlations temporelles 10 min)
    var correlations = detectCorrelations(p);
    var correlationsHTML = '';
    if (correlations.length > 0) {
        correlationsHTML = '<div class="correlations-block">' +
                             '<h5>&#128269; Signaux crois&eacute;s</h5>';
        correlations.forEach(function(f) {
            correlationsHTML += '<div class="correlation-item ' + f.severity + '">' +
                                  '<span class="c-icon">' + f.icon + '</span>' +
                                  '<div>' +
                                    '<span class="c-severity">' + f.title + '</span>' +
                                    '<div class="c-detail">' + f.detail + '</div>' +
                                  '</div>' +
                                '</div>';
        });
        correlationsHTML += '</div>';
    }

    var panelStability = '<div class="detail-box">' +
        (correlationsHTML ? '<div style="flex:1 1 100%">' + correlationsHTML + '</div>' : '') +
        '<div class="detail-section sec-crash"><h4>Crash / Freeze</h4>' + crashHTML + '</div>' +
        '<div class="detail-section sec-hw"><h4>Hardware</h4>' + hwHTML + '</div>' +
        '<div class="detail-section sec-crasher"><h4>Top Crashers</h4>' + crasherHTML + '</div>' +
        '<div class="detail-section sec-perf"><h4>Performance</h4>' + perfHTML + '</div>' +
        '</div>';

    var panelBoot = '<div class="detail-box">' +
        '<div class="detail-section sec-boot"><h4>D&eacute;marrages r&eacute;cents</h4>' + bootLongHTML + '</div>' +
        '<div class="detail-section sec-bootperf"><h4>Boot Performance</h4>' + bootPerfHTML + '</div>' +
        '</div>';

    var panelMaterial = '<div class="detail-box">' +
        '<div class="detail-section sec-cpu"><h4>CPU</h4>' + cpuHTML + '</div>' +
        '<div class="detail-section sec-ram"><h4>M&eacute;moire RAM</h4>' + ramHTML + '</div>' +
        '<div class="detail-section sec-gpu"><h4>GPU &amp; Drivers</h4>' + gpuHTML + '</div>' +
        '<div class="detail-section sec-throttle"><h4>CPU Throttling</h4>' + throttleHTML + '</div>' +
        '<div class="detail-section sec-disk"><h4>Disque (remplissage)</h4>' + diskHTML + '</div>' +
        '<div class="detail-section sec-smart"><h4>SMART</h4>' + smartHTML + '</div>' +
        '<div class="detail-section sec-battery"><h4>Batterie</h4>' + batteryHTML + '</div>' +
        '<div class="detail-section sec-monitors"><h4>&Eacute;crans secondaires branch&eacute;s</h4>' + monitorsHTML + '</div>' +
        '</div>';

    var panelSecurity = '<div class="detail-box">' +
        '<div class="detail-section sec-edr" style="flex:1 1 100%;max-width:720px"><h4>EDR (SentinelOne)</h4>' + edrHTML + '</div>' +
        '</div>';

    return '<tr class="row-detail" id="detail-' + idx + '">' +
        '<td colspan="' + colspan + '" style="padding:0">' +
        '<div class="detail-tabs">' + tabsHTML + '</div>' +
        panel('overview',  panelOverview,  true) +
        panel('stability', panelStability, false) +
        panel('boot',      panelBoot,      false) +
        panel('material',  panelMaterial,  false) +
        panel('security',  panelSecurity,  false) +
        '</td></tr>';
}

function toggleDetail(id) {
    var detailRow = document.getElementById(id);
    var mainRow = detailRow.previousElementSibling;
    detailRow.classList.toggle('visible');
    mainRow.classList.toggle('open');
}

// ===== EXPORT CSV =====
function exportCSV() {
    var cutoff = new Date(generatedAt);
    cutoff.setDate(cutoff.getDate() - state.days);
    var searchTerm = document.getElementById('searchInput').value.trim().toLowerCase();
    var allEnriched = pcData.map(function(pc) { return enrichPC(pc, cutoff); });
    var visible = allEnriched.filter(function(p) {
        if (state.siteFilter && p.pc.Site !== state.siteFilter) return false;
        if (state.cpuFilter  && p.pc.CPUAgeCategory !== state.cpuFilter) return false;
        if (searchTerm) {
            if (p.pc.PC.toLowerCase().indexOf(searchTerm) === -1 &&
                (p.pc.CurrentUser || '').toLowerCase().indexOf(searchTerm) === -1) return false;
        }
        if (!matchKpiFilter(p, state.kpiFilter)) return false;
        if (state.maskHealthy && p.score === 0) return false;
        return true;
    });
    sortPCs(visible);

    var headers = ['PC', 'Site', 'IP', 'Utilisateur', 'Statut', 'Connexion', 'CPU', 'CategorieCPU',
                   'AnneeCPU', 'UptimeJours', 'DerniereActivite', 'Score',
                   'Crash', 'BSOD',
                   'WHEA_Fatal', 'WHEA_Corrected_Occurrences', 'WHEA_Corrected_Signatures',
                   'GPU_TDR', 'Thermal',
                   'BootsLongs',
                   'Boots_ColdBoot', 'Boots_FastStartup', 'Boots_Resume',
                   'DisquesCritiques', 'PctLibreMin', 'Warnings',
                   // v5.3 / v5.4
                   'BatterieSantePct', 'BatterieCategorie', 'BatterieCycles', 'BatterieAlerte',
                   'SentinelStatus', 'SentinelInstalle', 'SentinelAlerte',
                   'BootTimeMs', 'PostBootMs', 'BootPerfAlerte',
                   'DiskWorstWearPct', 'DiskSmartAlerte', 'DiskSmartRaisons',
                   // v5.7
                   'MonitorsCount', 'MonitorsOldCount', 'MonitorsList'];
    var rows = [headers.join(';')];

    visible.forEach(function(p) {
        var pc = p.pc;
        var minPctFree = p.diskInfo.length > 0
            ? p.diskInfo.reduce(function(m, d) { return d.PctFree < m ? d.PctFree : m; }, 100)
            : '';
        var bt = p.bootsByType || {};

        // v5.3 / v5.4 helpers
        var battPct   = (p.battery && p.battery.HasBattery) ? p.battery.HealthPercent : '';
        var battCat   = (p.battery && p.battery.HasBattery) ? p.battery.HealthCategory : '';
        var battCycle = (p.battery && p.battery.HasBattery && p.battery.CycleCount != null) ? p.battery.CycleCount : '';
        var battAlert = p.batteryAlert ? 'OUI' : 'NON';
        var senStat   = p.sentinel ? p.sentinel.Status : '';
        var senInst   = p.sentinel ? (p.sentinel.Installed ? 'OUI' : 'NON') : '';
        var senAlert  = p.sentinelAlert ? 'OUI' : 'NON';
        var btMs      = (p.bootPerf && p.bootPerf.LastBoot) ? p.bootPerf.LastBoot.BootTimeMs : '';
        var pbMs      = (p.bootPerf && p.bootPerf.LastBoot) ? p.bootPerf.LastBoot.BootPostBootTimeMs : '';
        var bpAlert   = p.bootPerfAlert ? 'OUI' : 'NON';
        var worstWear = (p.diskWorstWear != null) ? p.diskWorstWear : '';
        var smartAlert= p.diskSmartAlert ? 'OUI' : 'NON';
        var smartRaisons = '';
        if (p.diskSmartAlerts && p.diskSmartAlerts.length > 0) {
            smartRaisons = p.diskSmartAlerts.map(function(d) {
                return d.FriendlyName + ':' + d.AlertReasons.join('+');
            }).join(' | ');
        }

        // v5.7 helpers : moniteurs
        var monCount = (p.monitors || []).length;
        var monOldCount = (p.oldMonitors || []).length;
        var monList = '';
        if (monCount > 0) {
            monList = p.monitors.map(function(m) {
                var s = (m.Manufacturer || '?') + ' ' + (m.Model || m.ProductCode || '?');
                if (m.SerialNumber) s += ' [SN:' + m.SerialNumber + ']';
                if (m.YearOfManufacture) s += ' (' + m.YearOfManufacture + ')';
                return s;
            }).join(' | ');
        }

        var row = [
            pc.PC, pc.Site, pc.IP, pc.CurrentUser || '',
            pc.IsOffline ? 'OFFLINE' : 'OK',
            pc.ConnectionType || '',
            pc.CPUName || '', pc.CPUAgeCategory || '', pc.CPUYear || '',
            pc.UptimeDays != null ? pc.UptimeDays : '',
            pc.CollectedAt, p.score,
            p.crashCount, p.bsodCount,
            p.wheaFatal.length, p.wheaCorrectedTotal, p.wheaCorrected.length,
            p.gpuTDR.length, p.thermal.length,
            p.bootLongCount,
            bt.ColdBoot || 0, bt.FastStartup || 0, bt.Resume || 0,
            p.diskAlertCount, minPctFree, p.warningCount,
            // v5.3 / v5.4
            battPct, battCat, battCycle, battAlert,
            senStat, senInst, senAlert,
            btMs, pbMs, bpAlert,
            worstWear, smartAlert, smartRaisons,
            // v5.7
            monCount, monOldCount, monList
        ];
        rows.push(row.map(function(v) {
            var s = String(v).replace(/"/g, '""');
            return (s.indexOf(';') !== -1 || s.indexOf('"') !== -1 || s.indexOf('\n') !== -1) ? '"' + s + '"' : s;
        }).join(';'));
    });

    var csv = '\uFEFF' + rows.join('\n');   // BOM UTF-8 pour Excel FR
    var blob = new Blob([csv], { type: 'text/csv;charset=utf-8' });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    var dateStr = new Date().toISOString().slice(0, 10).replace(/-/g, '');
    a.download = 'PCPulse-Export-' + dateStr + '.csv';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

// ===== INIT =====
initSiteDropdown();
var initBtn = document.querySelectorAll('.range-btn')[3];
state.daysBtn = initBtn;
initBtn.classList.add('active');

if (state.maskHealthy) {
    document.getElementById('maskHealthyBtn').classList.add('active');
}

// v5.5 : restaurer la preference "Vue detaillee" depuis localStorage
try {
    if (localStorage.getItem('pcmon-advcols') === '1') {
        document.body.classList.add('advanced-cols');
        document.getElementById('advancedColsBtn').classList.add('active');
    }
} catch(e) {}

try {
    render();
} catch(e) {
    // v5.5 : on affiche l'erreur dans la zone des groupes KPI (plus visible)
    var errTarget = document.getElementById('kpiGroups') || document.getElementById('kpiGrid');
    errTarget.innerHTML =
        '<div style="color:var(--red);padding:24px;font-size:14px;background:var(--bg-danger);border-radius:10px">' +
        '<strong>Erreur JS :</strong> ' + e.message + '<br>' + e.stack + '</div>';
}
</script>

</body>
</html>
"@

# ============================================================
# EXPORT ET OUVERTURE
# ============================================================
Set-Content -Path $OutputHTML -Value $html -Encoding UTF8
Write-Host "[+] Dashboard genere : $OutputHTML" -ForegroundColor Green
Write-Host "[*] Ouverture dans le navigateur..." -ForegroundColor Cyan

Start-Process -FilePath $OutputHTML
