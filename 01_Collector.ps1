#Requires -Version 5.1
<#
.SYNOPSIS
    PCPulse Collector v1.8
.DESCRIPTION
    Collecte les evenements systeme (boot, crash, freeze, BSOD, hardware)
    et les exporte en JSON vers un dossier partage.
    A deployer via Intune, SmartDeploy ou GPO, execute en contexte SYSTEM.
.NOTES
    Auteur       : Damien Gouhier
    Repository   : https://github.com/Damien-Gouhier/pcpulse
    Licence      : MIT
    Version      : 1.8
    Runtime      : PowerShell 5.1+ (compatible parc Windows 10/11 natif)
.CHANGELOG
    v1.8 : Enrichissement panel Materiel (Wave 2 - hardware)
           - CPU Throttling detaille : nouveaux Events 35/55 Kernel-
             Processor-Power = vrai throttling firmware/thermique.
             Les Events 37/88/125 Kernel-Power (thermal critique)
             restent captures comme avant. Distinction CPUThrottling
             vs Thermal dans HardwareHealth pour granularite.
             Throttling recurrent = pate thermique seche ou ventilo
             encrasse, fix typique 15EUR par PC.
           - Inventaire RAM : Win32_PhysicalMemory + Win32_PhysicalMemoryArray
             pour remonter totale installee, nb de slots total/occupes,
             capacite max supportee par la carte mere, et detail des
             barrettes (taille/type/vitesse/fabricant).
             Utile pour decider upgrade vs remplacement.
           - Inventaire GPU : Win32_VideoController pour remonter nom,
             version driver, date driver. Sans fioritures (pas de
             VRAM car peu fiable via WMI), suffisant pour correler
             avec les Event 4101 TDR GPU.
    v1.7 : Extension du parser CPU (fix terrain)
           - Fix mapping Intel Gen 10 : 2019 -> 2020. Alignement sur
             l'annee de la grosse majorite des desktops Comet Lake
             (i5-10500, i7-10700, etc.). Accepte +1 an de decalage
             sur les mobiles Ice Lake 2019 (volume marginal).
           - Nouveau support : Intel Celeron N-series (crucial pour
             tablettes d'accueil, kiosques, mini-PC budget) :
             * N4000/N4100            -> 2017 (Gemini Lake)
             * N4500/N4505/N5100/N6000 -> 2021 (Jasper Lake)
             * N95/N97/N100/N200/N305 -> 2023 (Alder Lake-N)
             * N350/N355              -> 2025 (Twin Lake)
           - Nouveau support : Intel Pentium Gold (G6xxx/G7xxx desktop
             budget).
           - Nouveau support : AMD Athlon et A-Series (parc AMD legacy).
           - Avant v1.7, ces CPU ressortaient Category='Inconnu' dans
             le Dashboard, genant le verdict global v1.6.
    v1.6 : Enrichissement Event 41 Kernel-Power (classification fine)
           - Parsing XML nomme des Event 41 : lecture de BugcheckCode,
             SleepInProgress, PowerButtonTimestamp (via EventData par
             nom, pas par index Properties[N] qui est fragile entre
             versions Windows).
           - Nouveau champ CrashCause sur les Events Id=41, avec 5
             valeurs precises :
               * 'BSODSilent'        -> BSOD non affiche (BugCheck != 0)
               * 'SleepResumeFailed' -> reprise de veille ratee
               * 'UserForcedReset'   -> bouton power maintenu par user
               * 'PowerLoss'         -> coupure alim/thermal shutdown
               * 'FreezeApp' / 'FreezeUnknown' -> freeze classique
           - Nouvelle stat Stats.TotalHardCrash : count des vrais
             plantages durs (BSODSilent + SleepResumeFailed + PowerLoss),
             hors user action et hors freeze.
           - Type/Detail des Events existants INCHANGES (backward-compat
             Dashboard v1.5 total, aucun risque de regression).
    v1.5 : Clustering des Event 51 (Disk slow / I/O timeout)
           - Avant : chaque Event 51 ecrit en ligne distincte. Un seul
             incident matos pouvait generer des centaines de lignes
             identiques (ex parc pilote : 957 events pour ~8 incidents
             reels, dont un burst de 894 en 1 seconde).
           - Fix : regroupement par fenetre de 60 secondes. Une entree
             par cluster, avec Count, FirstSeen, LastSeen, IsBurst.
           - IsBurst=true si Count >= 50 (signal materiel fort :
             probable deconnexion disque momentanee ou slot PCIe fatigue).
           - JSON beaucoup plus petit sur les PC problematiques, lisible
             pour l'utilisateur final dans le Dashboard.
    v1.4 : Fix parsing Intel Core Gen 10+ avec suffixe lettre
    v1.3 : Fix CurrentUser + BootDurations (Fast Startup)
    v1.2 : Fix algorithme Uptime utilisateur
    v1.1 : Uptime utilisateur (Fast Startup aware)
    v1.0 : Release initiale.
    Voir CHANGELOG.md du repo pour l'historique complet.
.EXAMPLE
    .\01_Collector.ps1
    # Execution normale avec delai anti-collision aleatoire (0-60 min)
.EXAMPLE
    .\01_Collector.ps1 -NoDelay
    # Execution immediate, pour tests manuels
#>

param(
    # Chemin du partage ou dossier local ou le Collector depose son JSON
    # et lit sa config. Par defaut : C:\PCPulse (dossier local).
    # En production, typiquement un UNC type \\SERVEUR\PCPulse$.
    [string]$SharePath = 'C:\PCPulse',

    # Bypass le delai anti-collision (pour tests manuels).
    [switch]$NoDelay
)

# ============================================================
# CONSTANTES (non modifiables - structurelles)
# ============================================================
$SchemaVersion = '1.8'
$ConfigFile    = Join-Path $SharePath 'config.psd1'

# Valeurs par defaut utilisees si config.psd1 est absent/invalide
$DefaultConfig = @{
    HistoriqueJours = 30
    SeuilBootLong   = 2
    SeuilBootMax    = 30
    SeuilDiskAlert  = 10
}
# ============================================================

# ============================================================
# FONCTIONS UTILITAIRES
# ============================================================

# Ecrit une ligne dans le log machine (UTF-8 sans BOM)
# Fonction tolerante aux pannes : ne lance jamais d'exception
function Write-Log {
    param([string]$Message)
    try {
        $logFile = Join-Path $SharePath "logs\$($env:COMPUTERNAME).log"
        $logDir  = Split-Path $logFile
        if (-not (Test-Path $logDir)) {
            $null = New-Item $logDir -ItemType Directory -Force -ErrorAction SilentlyContinue
        }
        $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | $Message`r`n"
        [System.IO.File]::AppendAllText($logFile, $line, [System.Text.UTF8Encoding]::new($false))
    } catch {
        # On ne bloque jamais la collecte pour un log qui echoue
    }
}

# v6.1 : helper pour Get-WinEvent tolerant + log standardise.
# Avant, on avait 7 blocs try/catch/Write-Log identiques qui repetaient
# 15 lignes chacun. Cette fonction factorise tout ca. Retourne toujours
# un array (jamais $null) pour simplifier le code appelant.
#   - $Label   : libelle court affiche dans le log (ex: "Application 1000/1002")
#   - $Filter  : hashtable passe a Get-WinEvent -FilterHashtable
#   - $MaxEvt  : optionnel, limite (pour Event 100 Boot Perf)
#   - $Sort    : si $true, trie par TimeCreated croissant (pour les boots)
function Invoke-SafeGetWinEvent {
    param(
        [string]$Label,
        [hashtable]$Filter,
        [int]$MaxEvt = 0,
        [switch]$Sort
    )
    $events = @()
    try {
        if ($MaxEvt -gt 0) {
            $events = Get-WinEvent -FilterHashtable $Filter -ErrorAction SilentlyContinue -MaxEvents $MaxEvt
        } else {
            $events = Get-WinEvent -FilterHashtable $Filter -ErrorAction SilentlyContinue
        }
        # Get-WinEvent retourne $null si aucun event -> force un array vide
        if ($null -eq $events) { $events = @() }
        if ($Sort) { $events = @($events | Sort-Object TimeCreated) }
        # Forcer la lazy evaluation en array concret pour stabiliser .Count
        $events = @($events)
        Write-Log "$Label : $($events.Count) events"
    } catch {
        Write-Log "Erreur $Label : $_"
        $events = @()
    }
    return ,$events   # virgule = force retour array meme si 1 seul element
}

# Tronque un message d'event de maniere securisee (guard contre $null)
function Get-TruncatedMessage {
    param([string]$Text, [int]$MaxLength = 150)
    if (-not $Text) { return '' }
    $clean = $Text -replace '\s+', ' '
    return $clean.Substring(0, [math]::Min($MaxLength, $clean.Length))
}

# Charge config.psd1 et fusionne avec les defaults.
# Tolerant aux pannes : si le fichier est absent, corrompu ou
# inaccessible, on retombe sur les defaults sans lever d'erreur.
function Import-MonitorConfig {
    param(
        [string]$Path,
        [hashtable]$Defaults
    )
    if (-not (Test-Path $Path)) {
        Write-Log "Config absente ($Path), utilisation des defaults"
        return $Defaults.Clone()
    }
    try {
        $loaded = Import-PowerShellDataFile -Path $Path -ErrorAction Stop
        $merged = @{}
        foreach ($key in $Defaults.Keys) {
            if ($loaded.ContainsKey($key)) {
                $merged[$key] = $loaded[$key]
            } else {
                $merged[$key] = $Defaults[$key]
            }
        }
        Write-Log "Config chargee depuis $Path"
        return $merged
    } catch {
        Write-Log "Erreur lecture config ($_), utilisation des defaults"
        return $Defaults.Clone()
    }
}

# Rotation du log : garde uniquement les lignes des derniers
# $HistoriqueJours jours. Tourne une fois au debut de chaque
# execution du Collector.
function Invoke-LogCleanup {
    param([int]$HistoriqueJours)
    $logFile = Join-Path $SharePath "logs\$($env:COMPUTERNAME).log"
    if (-not (Test-Path $logFile)) { return }
    try {
        $cutoff = (Get-Date).AddDays(-$HistoriqueJours).ToString('yyyy-MM-dd')
        $content = Get-Content -Path $logFile -Raw -ErrorAction Stop
        $filtrees = $content -split "`n" | Where-Object {
            if ($_.Length -ge 10) {
                $dateLigne = $_.Substring(0, 10)
                $dateLigne -ge $cutoff
            }
        }
        # Ecriture UTF-8 sans BOM pour rester coherent avec Write-Log
        $finalText = if ($filtrees) { ($filtrees -join "`n") } else { '' }
        [System.IO.File]::WriteAllText($logFile, $finalText, [System.Text.UTF8Encoding]::new($false))
    } catch {
        # On ne bloque pas le Collector si le cleanup echoue
    }
}

# Retourne l'IPv4 "primaire" du poste, par ordre de priorite :
#   1. 10.x.x.x        (reseau corporate principal)
#   2. 172.16-31.x.x   (reseau corporate RFC1918)
#   3. toute autre IPv4 hors loopback / vEthernet / link-local
# Renvoie 'N/A' si aucune IP trouvee.
function Get-PrimaryIPv4 {
    $candidates = Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue | Where-Object {
        $_.InterfaceAlias -notlike 'Loopback*' -and
        $_.InterfaceAlias -notlike 'vEthernet*' -and
        $_.IPAddress      -notlike '169.254.*'  -and
        $_.IPAddress      -notlike '127.*'
    }
    if (-not $candidates) { return 'N/A' }

    $patterns = @(
        '^10\.',
        '^172\.(1[6-9]|2[0-9]|3[0-1])\.',
        '.*'
    )
    foreach ($pattern in $patterns) {
        $found = $candidates | Where-Object { $_.IPAddress -match $pattern } | Select-Object -First 1
        if ($found) { return $found.IPAddress }
    }
    return 'N/A'
}

# v1.3 : Recupere le nom de l'utilisateur interactif actuellement connecte.
# Necessaire car le Collector tourne en SYSTEM (via tache planifiee), donc
# $env:USERNAME retourne le nom machine ($env:COMPUTERNAME suivi de '$',
# ex: "PCNAME$"), qui n'est pas ce qu'on veut afficher dans le Dashboard.
#
# Strategie en cascade :
#   1. Win32_ComputerSystem.UserName (rapide, retourne "DOMAINE\user")
#   2. Si vide : chercher le proprietaire du process explorer.exe
#   3. Si vide : retourner "(aucune session)"
#
# Le nom de domaine est toujours strippe pour n'afficher que le username.
function Get-CurrentInteractiveUser {
    # Methode 1 : Win32_ComputerSystem
    try {
        $cs = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop
        if ($cs.UserName) {
            # Format "DOMAINE\user" -> on garde juste "user"
            return ($cs.UserName -replace '^[^\\]+\\', '')
        }
    } catch {
        # On continue avec la methode 2
    }

    # Methode 2 : owner du process explorer.exe
    try {
        $explorerProcess = Get-CimInstance -ClassName Win32_Process `
                                           -Filter "Name='explorer.exe'" `
                                           -ErrorAction SilentlyContinue |
                           Select-Object -First 1
        if ($explorerProcess) {
            $ownerResult = Invoke-CimMethod -InputObject $explorerProcess `
                                            -MethodName GetOwner `
                                            -ErrorAction Stop
            if ($ownerResult.User) {
                return $ownerResult.User
            }
        }
    } catch {
        # Methode 3 : pas de user
    }

    return '(aucune session)'
}

# Analyse le nom d'un CPU et en deduit :
#   - Vendor (Intel / AMD)
#   - Generation (int pour Intel Core classique, "UltraN" pour Ultra,
#                 "RyzenN" pour Ryzen, "RyzenAI" pour Ryzen AI PRO)
#   - Annee de sortie de la generation (approximative)
#   - Age en annees (par rapport a l'annee courante)
#   - Categorie : Recent (<=3 ans) / Vieillissant (4-6 ans) / Ancien (>6 ans)
# Les CPU non reconnus retournent Vendor=null et Category="Inconnu".
function Get-CpuProfile {
    param([string]$CpuName)

    $result = [PSCustomObject]@{
        Vendor   = $null
        Gen      = $null
        Year     = $null
        Age      = $null
        Category = 'Inconnu'
    }
    if (-not $CpuName) { return $result }

    # Normalisation : on retire les marqueurs (R) (TM), les espaces multiples
    $clean = $CpuName -replace '\(R\)', '' -replace '\(TM\)', '' -replace '\s+', ' '
    $clean = $clean.Trim()

    # --- Intel Core Ultra (Series N, ex: "Core Ultra 7 155H", "Core Ultra 9 285H") ---
    # On deduit la serie du premier chiffre apres le grade : 1xx=serie 1 (2023), 2xx=serie 2 (2024)
    if ($clean -match 'Intel.*Core\s+Ultra\s+[579]\s+(\d)\d{2}') {
        $result.Vendor = 'Intel'
        $series = [int]$matches[1]
        switch ($series) {
            1 { $result.Year = 2023; $result.Gen = 'Ultra1' }
            2 { $result.Year = 2024; $result.Gen = 'Ultra2' }
            default {
                # Extension automatique : serie 3 -> 2025, serie 4 -> 2026, etc.
                $result.Year = 2023 + ($series - 1)
                $result.Gen  = "Ultra$series"
            }
        }
    }
    # --- Intel Core classique i3/i5/i7/i9 (gen 1-14) ---
    # Format attendu : "Intel Core i7-12700" (5 chiffres, gen 12),
    #                  "Core i5-1345U"     (4 chiffres + suffixe, gen 13),
    #                  "Core i5-8350U"     (4 chiffres + suffixe, gen 8).
    #
    # v1.4 : fix du parsing gen 10+ avec suffixe lettre
    # Avant, la regex `(\d{1,2})\d{3}` etait greedy et capturait toujours
    # 1 seul chiffre sur les CPU a 4 chiffres (ex: 1345U -> Gen=1 au lieu
    # de 13). Seuls les CPU a 5 chiffres (i7-12700K) etaient bien detectes.
    # Fix : on capture les 4 ou 5 chiffres et on applique une heuristique :
    #   - 5 chiffres : gen = les 2 premiers
    #   - 4 chiffres commencant par "1[0-4]" : gen = les 2 premiers (10-14)
    #   - 4 chiffres autrement : gen = le 1er chiffre (1-9)
    elseif ($clean -match 'Intel.*Core\s+i[3579]-(\d{4,5})') {
        $result.Vendor = 'Intel'
        $numStr = $matches[1]
        $gen = 0
        if ($numStr.Length -eq 5) {
            # CPU Gen 10+ desktop/mobile sans suffixe (ex: i7-12700, i7-14700K)
            $gen = [int]$numStr.Substring(0, 2)
        } elseif ($numStr[0] -eq '1' -and $numStr[1] -match '[0-4]') {
            # CPU Gen 10-14 avec suffixe lettre (ex: i5-1345U, i5-1065G7)
            $gen = [int]$numStr.Substring(0, 2)
        } else {
            # CPU Gen 1-9 (ex: i5-4300U, i7-8350U, i3-6100U)
            $gen = [int]$numStr[0].ToString()
        }
        $intelMap = @{
            1=2010; 2=2011; 3=2012; 4=2013; 5=2015; 6=2015; 7=2016
            8=2017; 9=2018; 10=2020; 11=2020; 12=2021; 13=2023; 14=2024
        }
        # Note v1.7 : Gen 10 = 2020 pour coller au desktop Comet Lake
        # (mai 2020, majorite du parc Gen 10). Les rares mobiles Ice Lake
        # de fin 2019 apparaitront avec 1 an de decalage, acceptable.
        if ($intelMap.ContainsKey($gen)) {
            $result.Gen  = $gen
            $result.Year = $intelMap[$gen]
        }
    }
    # --- AMD Ryzen AI PRO (gamme 2024, ex: "Ryzen AI 9 PRO 365") ---
    # Doit etre teste AVANT Ryzen standard car contient "AI ... PRO"
    elseif ($clean -match 'AMD.*Ryzen\s+AI\s+\d\s+PRO') {
        $result.Vendor = 'AMD'
        $result.Gen    = 'RyzenAI'
        $result.Year   = 2024
    }
    # --- AMD Ryzen standard ou PRO non-AI ---
    # Format : "Ryzen [3|5|7|9] [PRO ] N..." ou N est le 1er chiffre de la serie.
    # Le "PRO " est optionnel (fix v5.0 : v4.1 ratait les Ryzen X PRO NNNN).
    elseif ($clean -match 'AMD.*Ryzen\s+[3579]\s+(?:PRO\s+)?(\d)') {
        $result.Vendor = 'AMD'
        $ryzenGen = [int]$matches[1]
        $ryzenMap = @{
            1=2017; 2=2018; 3=2019; 4=2020; 5=2020
            6=2022; 7=2022; 8=2024; 9=2024
        }
        if ($ryzenMap.ContainsKey($ryzenGen)) {
            $result.Gen  = "Ryzen$ryzenGen"
            $result.Year = $ryzenMap[$ryzenGen]
        }
    }
    # --- v1.7 : Intel Celeron N-series (tablettes accueil, mini-PC budget) ---
    # Format : "Intel(R) Celeron(R) N4500", "Celeron J6412", etc.
    # Table par plage de numero (plus robuste que nom par nom).
    elseif ($clean -match 'Intel.*Celeron.*\b[NJG](\d{2,4})\b') {
        $result.Vendor = 'Intel'
        $num = [int]$matches[1]
        # Mapping par plage (majoritaires dans le parc tablettes/mini-PC)
        if     ($num -ge 4000 -and $num -le 4199) { $result.Year = 2017; $result.Gen = 'GeminiLake' }     # N4000, N4100, J4xxx
        elseif ($num -ge 4500 -and $num -le 4599) { $result.Year = 2021; $result.Gen = 'JasperLake' }     # N4500, N4505
        elseif ($num -ge 5100 -and $num -le 6099) { $result.Year = 2021; $result.Gen = 'JasperLake' }     # N5100, N6000
        elseif ($num -ge 6400 -and $num -le 6499) { $result.Year = 2021; $result.Gen = 'JasperLake' }     # J6412
        elseif ($num -ge 90   -and $num -le 349)  { $result.Year = 2023; $result.Gen = 'AlderLakeN' }     # N95, N97, N100, N200, N305
        elseif ($num -ge 350  -and $num -le 399)  { $result.Year = 2025; $result.Gen = 'TwinLake' }       # N350, N355
        elseif ($num -ge 3000 -and $num -le 3999) { $result.Year = 2014; $result.Gen = 'BayTrail' }       # N3050, N3150, J3xxx (tres vieux)
    }
    # --- v1.7 : Intel Pentium Gold (desktops budget) ---
    # Format : "Intel(R) Pentium(R) Gold G6400", "Pentium Gold G7400"
    elseif ($clean -match 'Intel.*Pentium.*Gold\s+G(\d{4})') {
        $result.Vendor = 'Intel'
        $num = [int]$matches[1]
        if     ($num -ge 5000 -and $num -le 5999) { $result.Year = 2018; $result.Gen = 'CoffeeLake' }     # G5400, G5500 etc
        elseif ($num -ge 6000 -and $num -le 6999) { $result.Year = 2020; $result.Gen = 'CometLake' }      # G6400, G6500, G6600
        elseif ($num -ge 7000 -and $num -le 7999) { $result.Year = 2022; $result.Gen = 'AlderLake' }      # G7400
    }
    # --- v1.7 : AMD Athlon et A-Series (parc legacy) ---
    elseif ($clean -match 'AMD.*Athlon') {
        $result.Vendor = 'AMD'
        $result.Gen    = 'Athlon'
        # Pas d'annee precise : Athlon couvre 2017 (Zen) a 2021 (Gold 3150G).
        # On prend milieu de plage, donne "Vieillissant" (~5 ans).
        $result.Year   = 2019
    }
    elseif ($clean -match 'AMD.*A(4|6|8|10|12)-\d') {
        $result.Vendor = 'AMD'
        $result.Gen    = 'A-Series'
        # A-Series APU : 2011-2016. Majoritairement anciens.
        $result.Year   = 2014
    }
    # --- Fallback : au moins identifier le vendor si on a pu ---
    elseif ($clean -match 'Intel') { $result.Vendor = 'Intel' }
    elseif ($clean -match 'AMD')   { $result.Vendor = 'AMD' }

    # Calcul age + categorie si on a une annee
    if ($result.Year) {
        $result.Age = (Get-Date).Year - $result.Year
        if     ($result.Age -le 3) { $result.Category = 'Recent' }
        elseif ($result.Age -le 6) { $result.Category = 'Vieillissant' }
        else                       { $result.Category = 'Ancien' }
    }

    return $result
}

# Analyse un event WHEA-Logger via son XML structure.
# Retourne :
#   - Severity (int)   : 1=Critical, 2=Error, 3=Warning, 4=Info
#   - IsFatal (bool)   : vrai si Severity <= 2 (Critical ou Error)
#   - ErrorSource      : source reelle (Machine Check Exception,
#                        PCI Express Root Port, NMI Error, etc.)
#   - Component        : CPU | RAM | PCIe | GPU | Storage | Autre
#   - Signature        : cle canonique pour dedupliquer/agreger les
#                        erreurs corrigees repetitives (ex: BDF du PCIe)
#
# Source : l'ID de l'event (17/18/19/46/47/...) ne suffit PAS a
# deduire le composant. On parse ErrorSource dans le message/XML.
# - Event 17 = erreur corrigee (typiquement Warning / Level 3)
# - Event 18 = erreur fatale   (typiquement Error    / Level 2)
# - Event 19 = erreur corrigee memoire/cache
# - Event 46/47 = erreurs fatales specifiques
function Get-WHEAAnalysis {
    param($Event)

    $result = [PSCustomObject]@{
        Severity    = [int]$Event.Level      # 1=Critical, 2=Error, 3=Warning
        IsFatal     = ($Event.Level -le 2)
        EventId     = $Event.Id
        ErrorSource = ''
        Component   = 'Autre'
        Detail      = ''
        Signature   = ''
        BDF         = ''   # Bus:Device:Function pour PCIe
    }

    # Parsing XML pour extraire les champs structures
    # (plus fiable que le message localise en FR/EN)
    try {
        $xml = [xml]$Event.ToXml()
        # Les champs UserData\WHEAErrorRecord varient selon les cas.
        # On ratisse large : on cherche ErrorSource dans tout l'XML.
        $allText = $xml.OuterXml
        $msg     = [string]$Event.Message

        # --- ErrorSource : indique la source reelle du rapport ---
        # Valeurs typiques : "Machine Check Exception", "Advanced Error
        # Reporting (PCI Express)", "NMI Error", "Corrected Machine Check",
        # "PMEM Error Source", "Platform Memory"
        if ($msg -match 'Error Source\s*:\s*([^\r\n]+)') {
            $result.ErrorSource = $matches[1].Trim()
        } elseif ($msg -match 'Source de l''erreur\s*:\s*([^\r\n]+)') {
            # Localisation FR
            $result.ErrorSource = $matches[1].Trim()
        } elseif ($allText -match '<Data Name="ErrorSource">([^<]+)</Data>') {
            $result.ErrorSource = $matches[1].Trim()
        }

        # --- Component : deduit de ErrorSource, puis de mots-cles dans le msg ---
        $src = $result.ErrorSource
        if ($src -match 'PCI|Express') {
            $result.Component = 'PCIe'
        } elseif ($src -match 'Machine Check|MCE|Corrected Machine') {
            $result.Component = 'CPU'
        } elseif ($src -match 'NMI') {
            $result.Component = 'CPU'
        } elseif ($src -match 'Memory|Platform Memory|PMEM') {
            $result.Component = 'RAM'
        } elseif ($msg -match 'PCI\s*Express|\bPCIe\b') {
            $result.Component = 'PCIe'
        } elseif ($msg -match 'memory|m[eé]moire|cache') {
            $result.Component = 'RAM'
        } elseif ($msg -match 'processor|processeur|CPU') {
            $result.Component = 'CPU'
        } elseif ($msg -match 'GPU|graphic|graphique') {
            $result.Component = 'GPU'
        }

        # --- BDF pour PCIe : format "Bus:Device:Function : 0xX:0xY:0xZ" ---
        # Permet de dedupliquer les erreurs corrigees venant du meme port
        # (un SSD NVMe pourri peut generer des milliers d'events par jour)
        if ($msg -match 'Bus:\s*Device:\s*Function\s*:\s*(0x[0-9A-Fa-f]+):(0x[0-9A-Fa-f]+):(0x[0-9A-Fa-f]+)') {
            $result.BDF = "{0}:{1}:{2}" -f $matches[1], $matches[2], $matches[3]
        } elseif ($msg -match 'Bus[^:]*:P[eé]riph[eé]rique[^:]*:Fonction\s*:\s*(0x[0-9A-Fa-f]+):(0x[0-9A-Fa-f]+):(0x[0-9A-Fa-f]+)') {
            # Localisation FR
            $result.BDF = "{0}:{1}:{2}" -f $matches[1], $matches[2], $matches[3]
        }

        $result.Detail = Get-TruncatedMessage -Text $msg -MaxLength 180

        # --- Signature pour agregation ---
        # Une meme source + meme composant + meme BDF = meme probleme
        $result.Signature = "{0}|{1}|{2}|{3}" -f $result.Component, $result.EventId, $result.ErrorSource, $result.BDF

    } catch {
        # En cas d'echec de parsing, on garde ce qu'on peut
        $result.Detail = Get-TruncatedMessage -Text $Event.Message -MaxLength 180
        $result.Signature = "Unknown|$($Event.Id)"
    }

    return $result
}

# Extrait le BootType depuis le message d'un Event 27 Kernel-Boot.
#   0x0 = Cold boot (vrai arret + demarrage, ou Restart)
#   0x1 = Fast Startup / Hybrid Boot (Windows 8+)
#   0x2 = Resume from Hibernation
# Retourne : 'ColdBoot', 'FastStartup', 'Resume', ou 'Unknown'
#
# Important : en Fast Startup, les events 6005/6006/1074 peuvent
# manquer. L'event 27 est le seul indicateur fiable du type de boot.
function Get-BootTypeFromMessage {
    param([string]$Message)
    if (-not $Message) { return 'Unknown' }
    if     ($Message -match '0x0\b') { return 'ColdBoot' }
    elseif ($Message -match '0x1\b') { return 'FastStartup' }
    elseif ($Message -match '0x2\b') { return 'Resume' }
    return 'Unknown'
}

# Parse un Event 100 de Microsoft-Windows-Diagnostics-Performance.
# Retourne un PSCustomObject avec les champs utiles pour le monitoring
# boot/logon/post-boot, ou $null si le parsing echoue.
#
# L'Event 100 contient ~40 champs dans son EventData, on n'expose que
# les 8 qui nous interessent (evite un JSON obese sans valeur ajoutee).
function ConvertFrom-BootPerfEvent {
    param($Event)
    try {
        $xml = [xml]$Event.ToXml()
        $props = @{}
        foreach ($d in $xml.Event.EventData.Data) {
            $props[$d.Name] = $d.'#text'
        }

        # Helper pour convertir proprement en int64 (certaines valeurs
        # peuvent etre vides ou absentes selon les versions de Windows)
        $toInt64 = {
            param($v)
            $n = 0L
            if ([int64]::TryParse([string]$v, [ref]$n)) { return $n } else { return 0L }
        }

        [PSCustomObject]@{
            Timestamp                   = $Event.TimeCreated.ToString('yyyy-MM-dd HH:mm:ss')
            Level                       = [string]$Event.LevelDisplayName
            BootStartTime               = [string]$props['BootStartTime']
            BootTimeMs                  = & $toInt64 $props['BootTime']
            MainPathBootTimeMs          = & $toInt64 $props['MainPathBootTime']
            BootPostBootTimeMs          = & $toInt64 $props['BootPostBootTime']
            UserProfileProcessingTimeMs = & $toInt64 $props['BootUserProfileProcessingTime']
            ExplorerInitTimeMs          = & $toInt64 $props['BootExplorerInitTime']
            NumStartupApps              = [int](& $toInt64 $props['BootNumStartupApps'])
            IsRebootAfterInstall        = ($props['BootIsRebootAfterInstall'] -eq 'true')
        }
    } catch {
        $null
    }
}

# ============================================================
# CHARGEMENT DE LA CONFIG
# ============================================================
$cfg = Import-MonitorConfig -Path $ConfigFile -Defaults $DefaultConfig

$HistoriqueJours = [int]$cfg.HistoriqueJours
$SeuilBootLong   = [double]$cfg.SeuilBootLong
$SeuilBootMax    = [double]$cfg.SeuilBootMax
$SeuilDiskAlert  = [int]$cfg.SeuilDiskAlert

Write-Log "=== Debut collecte (schema $SchemaVersion) ==="

# ============================================================
# ANTI-COLLISION : delai aleatoire 0-15 min
# Evite que 800 postes ecrivent simultanement sur le share.
# Bypass via -NoDelay pour les tests locaux.
# ============================================================
if (-not $NoDelay) {
    $delaySeconds = Get-Random -Minimum 0 -Maximum 900
    if ($delaySeconds -gt 0) {
        Write-Log "Attente anti-collision : $delaySeconds secondes"
        Start-Sleep -Seconds $delaySeconds
    }
} else {
    Write-Log 'Mode NoDelay : execution immediate'
}

Invoke-LogCleanup -HistoriqueJours $HistoriqueJours

# ============================================================
# 1. INFOS DE BASE DE LA MACHINE
# ============================================================
$os              = Get-CimInstance -ClassName Win32_OperatingSystem
$kernelLastBoot  = $os.LastBootUpTime   # vrai cold boot kernel (pour info technique)
# NOTE v1.1 : on ne calcule PAS $uptimeDays ici. Le vrai uptime utilisateur
# (= temps depuis la derniere reprise d'activite : cold boot, Fast Startup,
# ou wake Modern Standby) est calcule plus bas a partir de $bootDurations.
# Win32_OperatingSystem.LastBootUpTime ne reflete QUE le vrai cold boot et
# donne donc un uptime trompeur (15 jours alors que l'user eteint chaque soir).
$currentIP       = Get-PrimaryIPv4

# ----- CPU -----
$cpuInfo = Get-CimInstance -ClassName Win32_Processor | Select-Object -First 1
$cpuName = $cpuInfo.Name
$cpu     = Get-CpuProfile -CpuName $cpuName

if (-not $cpu.Vendor -and $cpuName) {
    Write-Log "CPU non reconnu par le parser : $cpuName"
}

# ----- Type de connexion (Ethernet vs WiFi) -----
$connectionType = 'Inconnu'
$activeAdapters = Get-NetAdapter -ErrorAction SilentlyContinue | Where-Object {
    $_.Status -eq 'Up' -and $_.Virtual -eq $false
}
$ethernetUp = $activeAdapters | Where-Object {
    $_.PhysicalMediaType -eq '802.3' -or $_.Name -match 'Ethernet|LAN'
}
$wifiUp = $activeAdapters | Where-Object {
    $_.PhysicalMediaType -match '802.11' -or $_.Name -match 'Wi-Fi|Wireless|WLAN'
}
if     ($ethernetUp)     { $connectionType = 'Ethernet' }
elseif ($wifiUp)         { $connectionType = 'WiFi' }
elseif ($activeAdapters) { $connectionType = 'Autre' }
else                     { $connectionType = 'Deconnecte' }

# ============================================================
# CHASSIS TYPE (v5.6) - Detection du form factor
#    via Win32_SystemEnclosure.ChassisTypes
#    Reference DMTF : codes 1-36 avec sous-groupes clairs
#
#    Utilise pour :
#      1) Inventaire SI (savoir si un PC est laptop / desktop / AIO)
#      2) Filtre moniteurs : sur AIO, on doit exclure l'ecran integre
#         (qui remonte comme externe via HDMI car techniquement il
#         n'est pas en LVDS ni DP internal)
# ============================================================
$LaptopChassis  = @(8, 9, 10, 14, 30, 31, 32)   # Portable/Laptop/Notebook/SubNB/Tablet/Convertible/Detachable
$DesktopChassis = @(3, 4, 5, 6, 7, 15, 35, 36)   # Desktop/LowProf/PizzaBox/MiniTower/Tower/SpaceSaving/etc
$AIOChassis     = @(13)                           # All-in-One

$chassisType    = $null
$chassisLabel   = 'Inconnu'
$isLaptop       = $false
$isDesktop      = $false
$isAIO          = $false
try {
    $enclosure = Get-CimInstance -ClassName Win32_SystemEnclosure -ErrorAction Stop
    if ($enclosure -and $enclosure.ChassisTypes) {
        # ChassisTypes est un array, on prend le premier code
        $chassisType = [int]($enclosure.ChassisTypes | Select-Object -First 1)

        if     ($chassisType -in $LaptopChassis)  { $isLaptop  = $true; $chassisLabel = 'Laptop' }
        elseif ($chassisType -in $AIOChassis)     { $isAIO     = $true; $chassisLabel = 'All-In-One' }
        elseif ($chassisType -in $DesktopChassis) { $isDesktop = $true; $chassisLabel = 'Desktop' }
        else                                       { $chassisLabel = "Autre ($chassisType)" }
    }
} catch {
    Write-Log "ChassisType indisponible : $_"
}
$chassisInfo = [PSCustomObject]@{
    ChassisType  = $chassisType
    ChassisLabel = $chassisLabel
    IsLaptop     = $isLaptop
    IsDesktop    = $isDesktop
    IsAIO        = $isAIO
}
Write-Log "ChassisType : $chassisLabel (code=$chassisType)"

# v1.1 : LastBoot et UptimeDays sont initialises avec les valeurs KERNEL
# (LastBootUpTime de Win32_OperatingSystem). Ces valeurs sont REMPLACEES
# plus tard par les vraies valeurs utilisateur calculees a partir de
# $bootDurations. Les valeurs kernel restent disponibles via LastRealColdBoot.
$machineInfo = [PSCustomObject]@{
    PC                  = $env:COMPUTERNAME
    CollectedAt         = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    OS                  = $os.Caption
    LastBoot            = $kernelLastBoot.ToString('yyyy-MM-dd HH:mm:ss')  # ecrase plus bas
    UptimeDays          = [math]::Round(((Get-Date) - $kernelLastBoot).TotalDays, 1)  # ecrase plus bas
    LastRealColdBoot    = $kernelLastBoot.ToString('yyyy-MM-dd HH:mm:ss')  # NEW v1.1
    FastStartupEnabled  = $null  # rempli plus bas (registry)
    CurrentUser         = Get-CurrentInteractiveUser
    IP                  = $currentIP
    CPUName             = $cpuName
    CPUVendor           = $cpu.Vendor
    CPUGen              = $cpu.Gen
    CPUYear             = $cpu.Year
    CPUAge              = $cpu.Age
    CPUAgeCategory      = $cpu.Category
    ConnectionType      = $connectionType
    # v5.6 : chassis info pour l'inventaire
    ChassisInfo         = $chassisInfo
}

# ============================================================
# v1.1 : Detection du Fast Startup (Hibernate Boot)
# ============================================================
# HiberbootEnabled dans le registre = le Fast Startup est globalement active
# sur ce PC. Combine avec HibernateEnabled pour etre sur qu'il est utilisable.
try {
    $hiberBoot = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power' `
                                  -Name HiberbootEnabled -ErrorAction Stop
    $machineInfo.FastStartupEnabled = ($hiberBoot.HiberbootEnabled -eq 1)
    Write-Log "Fast Startup enabled : $($machineInfo.FastStartupEnabled)"
} catch {
    Write-Log "Fast Startup : impossible de lire HiberbootEnabled ($_)"
    $machineInfo.FastStartupEnabled = $false
}

# ============================================================
# 2. RECUPERATION CONSOLIDEE DES EVENTS
#    v5 : 6 appels Get-WinEvent au lieu de 11 en v4.1
# ============================================================
$dateDebut = (Get-Date).AddDays(-$HistoriqueJours)

# --- Appel 1 : gros lot System (boot/crash/shutdown/resource/thermal) ---
# On ramasse tous les IDs d'un coup puis on dispatche en memoire.
# Le filtrage par ID est fait par l'ETL Windows (rapide).
$systemIdsBulk = @(
    6005, 6006, 41, 1074, 6008,   # boot / crash / shutdown
    2004, 2019, 51,               # resource warnings (RAM / disk full / disk slow)
    37,   88,   125               # thermal (throttling / arret / surchauffe)
)
# --- Appel 1 : gros lot System (boot/crash/shutdown/resource/thermal) ---
# On ramasse tous les IDs d'un coup puis on dispatche en memoire.
# Le filtrage par ID est fait par l'ETL Windows (rapide).
$systemIdsBulk = @(
    6005, 6006, 41, 1074, 6008,   # boot / crash / shutdown
    2004, 2019, 51,               # resource warnings (RAM / disk full / disk slow)
    37,   88,   125               # thermal (throttling / arret / surchauffe)
)
$systemBulk = Invoke-SafeGetWinEvent -Label 'System bulk' -Filter @{
    LogName   = 'System'
    StartTime = $dateDebut
    Id        = $systemIdsBulk
}

# Dispatch en memoire (1 seul parcours pour tout trier, au lieu de 6 Where-Object)
# v6.1 : plus efficace sur gros volumes d'events
$rawEventsList           = [System.Collections.Generic.List[object]]::new()
$dirtyShutdownEventsList = [System.Collections.Generic.List[object]]::new()
$ramEventsList           = [System.Collections.Generic.List[object]]::new()
$diskFullEventsList      = [System.Collections.Generic.List[object]]::new()
$diskSlowEventsList      = [System.Collections.Generic.List[object]]::new()
$thermalEventsList       = [System.Collections.Generic.List[object]]::new()
foreach ($ev in $systemBulk) {
    switch ($ev.Id) {
        6005 { $rawEventsList.Add($ev) }
        6006 { $rawEventsList.Add($ev) }
        41   { $rawEventsList.Add($ev) }
        1074 { $rawEventsList.Add($ev) }
        6008 { $dirtyShutdownEventsList.Add($ev) }
        2004 { $ramEventsList.Add($ev) }
        2019 { $diskFullEventsList.Add($ev) }
        51   { $diskSlowEventsList.Add($ev) }
        37   { if ($ev.Level -le 3) { $thermalEventsList.Add($ev) } }
        88   { if ($ev.Level -le 3) { $thermalEventsList.Add($ev) } }
        125  { if ($ev.Level -le 3) { $thermalEventsList.Add($ev) } }
    }
}
$rawEvents           = @($rawEventsList           | Sort-Object TimeCreated)
$dirtyShutdownEvents = @($dirtyShutdownEventsList | Sort-Object TimeCreated)
$ramEvents           = @($ramEventsList)
$diskFullEvents      = @($diskFullEventsList)
$diskSlowEvents      = @($diskSlowEventsList)
$thermalEvents       = @($thermalEventsList)

# --- Appel 2 : Application (crashs / hangs) ---
$appEvents = Invoke-SafeGetWinEvent -Label 'Application 1000/1002' -Sort -Filter @{
    LogName   = 'Application'
    StartTime = $dateDebut
    Id        = @(1000, 1002)
}

# --- Appel 3 : Event 12 Kernel-General (debut vrai du boot) ---
$bootStartEvents = Invoke-SafeGetWinEvent -Label 'Kernel-General Event 12' -Sort -Filter @{
    LogName      = 'System'
    StartTime    = $dateDebut
    Id           = 12
    ProviderName = 'Microsoft-Windows-Kernel-General'
}

# --- Appel 3b : Event 27 Kernel-Boot (BootType : cold/fast/resume) ---
# NOUVEAU v5.2 : essentiel pour detecter les Fast Startup, invisibles
# via les events classiques 6005/6006/1074
$kernelBootEvents = Invoke-SafeGetWinEvent -Label 'Kernel-Boot Event 27' -Sort -Filter @{
    LogName      = 'System'
    StartTime    = $dateDebut
    Id           = 27
    ProviderName = 'Microsoft-Windows-Kernel-Boot'
}

# --- Appel 4 : Kernel-Processor-Power ---
# v1.8 : on recupere aussi les Events 35 et 55 (throttling firmware
# / thermique) en plus de l'Event 26 (info capabilities). Ces events
# sont RARES par nature, donc tres significatifs quand ils apparaissent.
#   26 = Capabilities (log normal au boot, informatif)
#   35 = Throttle limited par firmware (ventilo HS / pate thermique)
#   55 = Puissance reduite pour cause thermique (surchauffe active)
$cpuEvents = Invoke-SafeGetWinEvent -Label 'Kernel-Processor-Power Event 26/35/55' -Filter @{
    LogName      = 'System'
    StartTime    = $dateDebut
    Id           = @(26, 35, 55)
    ProviderName = 'Microsoft-Windows-Kernel-Processor-Power'
}

# --- Appel 5 : WHEA (CPU / RAM / PCIe) ---
$wheaEvents = Invoke-SafeGetWinEvent -Label 'WHEA-Logger' -Filter @{
    LogName      = 'System'
    ProviderName = 'Microsoft-Windows-WHEA-Logger'
    StartTime    = $dateDebut
}

# --- Appel 6 : Display 4101 (GPU TDR) ---
$gpuEvents = Invoke-SafeGetWinEvent -Label 'Display GPU TDR' -Filter @{
    LogName      = 'System'
    ProviderName = 'Display'
    Id           = 4101
    StartTime    = $dateDebut
}

# --- Appel 7 : Boot Performance (Event 100 Diagnostics-Performance) ---
# v5.4 : capte les metriques de chaque cold boot (MainPath, PostBoot,
# UserProfile...). Ecrit uniquement pour les cold boots, pas pour
# Fast Startup ou Resume (par design de Windows).
# On prend les 5 plus recents pour limiter la taille du JSON.
$bootPerfEvents = Invoke-SafeGetWinEvent -Label 'Boot Performance Event 100' -MaxEvt 5 -Filter @{
    LogName      = 'Microsoft-Windows-Diagnostics-Performance/Operational'
    ProviderName = 'Microsoft-Windows-Diagnostics-Performance'
    Id           = 100
    StartTime    = $dateDebut
}

# ============================================================
# 3. CLASSIFICATION DES EVENTS (BSOD / Hard reset / Freeze)
#    v1.6 : Ajout du champ CrashCause sur les Events Id=41
#           pour distinguer finement la cause du plantage.
# ============================================================

# Helper v1.6 : extrait une valeur nommee depuis l'EventData XML d'un event.
# Retourne $null si le champ est absent.
# (Plus robuste que Properties[N] qui peut varier entre versions Windows.)
function Get-EventDataByName {
    param($Event, [string]$Name)
    try {
        $xml = [xml]$Event.ToXml()
        $node = $xml.Event.EventData.Data | Where-Object { $_.Name -eq $Name } | Select-Object -First 1
        if ($node) { return $node.'#text' }
    } catch {}
    return $null
}

$events = $rawEvents | ForEach-Object {
    $type        = ''
    $detail      = ''
    $crashCause  = $null   # v1.6 : renseigne uniquement pour Id=41

    if ($_.Id -eq 41) {
        # v1.6 : lecture des champs XML EventData par nom
        $bugCheck    = $null
        $sleepInProg = $null
        $pwrBtnTs    = $null
        try {
            $v = Get-EventDataByName -Event $_ -Name 'BugcheckCode'
            if ($v) { $bugCheck = [uint64]$v }
        } catch {}
        try {
            $v = Get-EventDataByName -Event $_ -Name 'SleepInProgress'
            if ($v) { $sleepInProg = [int]$v }
        } catch {}
        try {
            $v = Get-EventDataByName -Event $_ -Name 'PowerButtonTimestamp'
            if ($v) { $pwrBtnTs = [uint64]$v }
        } catch {}

        # Fallback si le parsing XML echoue : on retombe sur Properties[0]
        # pour BugCheckCode (compat logique v1.5).
        if ($null -eq $bugCheck) {
            try { $bugCheck = [uint64]$_.Properties[0].Value } catch {}
        }

        if ($bugCheck -and $bugCheck -ne 0) {
            # BSOD classique (peut avoir ete affiche ou silencieux).
            # On conserve Type='BSOD' pour backward-compat Dashboard v1.5.
            $type       = 'BSOD'
            $detail     = '0x' + $bugCheck.ToString('X')
            $crashCause = 'BSODSilent'
        } else {
            $event41Time = $_.TimeCreated

            # Etape 1 : chercher un dirty shutdown (6008) DANS LES 30 MIN SUIVANTES
            $dirtyShutdown = $dirtyShutdownEvents | Where-Object {
                $_.TimeCreated -gt $event41Time -and
                ($_.TimeCreated - $event41Time).TotalMinutes -le 30
            } | Select-Object -First 1

            if ($dirtyShutdown) {
                # v1.6 : raffinement de la cause selon SleepInProgress et PowerButtonTimestamp
                $type = 'Hard reset'
                if ($sleepInProg -eq 1) {
                    $detail     = 'Sleep/Resume failed'
                    $crashCause = 'SleepResumeFailed'
                } elseif ($pwrBtnTs -and $pwrBtnTs -ne 0) {
                    $detail     = 'User forced power off'
                    $crashCause = 'UserForcedReset'
                } else {
                    $detail     = 'Dirty shutdown detected'
                    $crashCause = 'PowerLoss'
                }
            } else {
                # Etape 2 : chercher une app crash/hang (1000/1002) dans les 10 min AVANT
                $fenetre = $event41Time.AddMinutes(-10)
                $appSuspect = $appEvents | Where-Object {
                    $_.TimeCreated -ge $fenetre -and $_.TimeCreated -le $event41Time
                } | Sort-Object TimeCreated -Descending | Select-Object -First 1

                $type = 'Freeze'
                if ($appSuspect) {
                    try {
                        $appNom = $appSuspect.Properties[0].Value
                        if ($appNom) { $detail = $appNom }
                    } catch {}
                    $crashCause = 'FreezeApp'
                } else {
                    $crashCause = 'FreezeUnknown'
                }
            }
        }
    } else {
        $type = switch ($_.Id) {
            6005 { 'Boot' }
            6006 { 'Shutdown propre' }
            1074 { 'Redemarrage force' }
        }
    }

    [PSCustomObject]@{
        Timestamp  = $_.TimeCreated.ToString('yyyy-MM-dd HH:mm:ss')
        EventId    = $_.Id
        Type       = $type
        Detail     = $detail
        CrashCause = $crashCause   # v1.6 : null hors Id=41
        Message    = (Get-TruncatedMessage -Text $_.Message -MaxLength 120)
    }
}

# ============================================================
# 4. DUREES DE BOOT (v5.2 : integre Event 27 Kernel-Boot)
#    Deux objectifs :
#      a) calculer la duree de boot (temps mesurable)
#      b) identifier le type : ColdBoot / FastStartup / Resume
#
#    Methode v5.2 :
#      - Pour chaque Event 27 (signal "le boot est termine, type = X"),
#        on cherche l'Event 12 (Kernel-General) le plus proche avant.
#      - Si Event 12 trouve : duree = 27 - 12, Method = "Event12+27"
#      - Sinon fallback : on retombe sur la logique classique 6005 - (6006|41)
#
#    On capture AUSSI les BootType via les Event 27 isoles, meme si
#    on ne peut pas calculer de duree : ca permet de compter les
#    redemarrages reels (fast startup inclus).
# ============================================================
$bootDurations = [System.Collections.Generic.List[PSCustomObject]]::new()
$bootsByType   = @{ ColdBoot = 0; FastStartup = 0; Resume = 0; Unknown = 0 }

# Methode principale : itere sur les Event 27 (chaque 27 = un vrai demarrage)
foreach ($ev27 in $kernelBootEvents) {
    $bootEndTime   = $ev27.TimeCreated
    $bootType      = Get-BootTypeFromMessage -Message $ev27.Message
    $bootsByType[$bootType]++

    $durationMin   = $null
    $methodUsed    = ''
    $precedentType = ''

    # Chercher l'Event 12 le plus proche AVANT le 27, dans la fenetre
    $bestEvent12 = $null
    foreach ($ev12 in $bootStartEvents) {
        $diffMin = ($bootEndTime - $ev12.TimeCreated).TotalMinutes
        if ($diffMin -gt 0 -and $diffMin -le $SeuilBootMax) {
            if (-not $bestEvent12 -or $ev12.TimeCreated -gt $bestEvent12.TimeCreated) {
                $bestEvent12 = $ev12
            }
        }
    }

    if ($bestEvent12) {
        $durationMin   = ($bootEndTime - $bestEvent12.TimeCreated).TotalMinutes
        $methodUsed    = 'Event12+27'
        $precedentType = 'Kernel boot start (Event 12)'
    }

    # v1.3 : Ajout a BootDurations AVEC OU SANS Event 12
    # Avant v1.3 : seuls les boots avec Event 12 (= cold boots) etaient ajoutes
    #              -> les Fast Startup et Resume (pas d'Event 12 = pas de kernel
    #              boot complet) etaient detectes mais pas inscrits
    # Resultat : le Dashboard n'affichait que les cold boots dans le graphe
    # "Repartition des demarrages" alors que les Fast Startup etaient bien
    # comptes dans Stats.BootsByType.
    #
    # Fix v1.3 : on inscrit TOUS les Event 27 dans BootDurations.
    #   - Avec Event 12 (cold boot)  : Method='Event12+27',   DurationMin calculee
    #   - Sans Event 12 (FS/Resume)  : Method='Event27-only', DurationMin=0
    if ($null -ne $durationMin -and $durationMin -le $SeuilBootMax) {
        # Cas classique : Event 12 trouve, duree calculee et dans les bornes
        $bootDurations.Add([PSCustomObject]@{
            DateBoot      = $bootEndTime.ToString('yyyy-MM-dd HH:mm:ss')
            DurationMin   = [math]::Round($durationMin, 1)
            PrecedentType = $precedentType
            EstBootLong   = ($durationMin -gt $SeuilBootLong)
            Method        = $methodUsed
            BootType      = $bootType
        })
    } elseif ($bootType -in @('FastStartup', 'Resume')) {
        # v1.3 : cas Fast Startup / Resume (pas d'Event 12 attendu)
        # Duree non calculable mais on veut quand meme tracer l'evenement
        # (geste utilisateur "arreter+rallumer" capture via Event 27 seul).
        $bootDurations.Add([PSCustomObject]@{
            DateBoot      = $bootEndTime.ToString('yyyy-MM-dd HH:mm:ss')
            DurationMin   = 0
            PrecedentType = 'Aucun (Event 27 seul, pas d''Event 12 pour ce type)'
            EstBootLong   = $false
            Method        = 'Event27-only'
            BootType      = $bootType
        })
    }
}

# Fallback : si on a des Event 6005 sans Event 27 correspondant
# (cas des systemes sans Kernel-Boot logging, rare mais possible)
if ($bootDurations.Count -eq 0 -and $kernelBootEvents.Count -eq 0) {
    $sortedEvents = $rawEvents | Sort-Object TimeCreated
    for ($i = 0; $i -lt $sortedEvents.Count; $i++) {
        if ($sortedEvents[$i].Id -ne 6005) { continue }
        $bootEndTime = $sortedEvents[$i].TimeCreated
        $precedent = $null
        for ($j = $i - 1; $j -ge 0; $j--) {
            if ($sortedEvents[$j].Id -in @(6006, 41)) {
                $precedent = $sortedEvents[$j]
                break
            }
        }
        if ($precedent) {
            $durationMin = ($bootEndTime - $precedent.TimeCreated).TotalMinutes
            if ($durationMin -le $SeuilBootMax) {
                $bootDurations.Add([PSCustomObject]@{
                    DateBoot      = $bootEndTime.ToString('yyyy-MM-dd HH:mm:ss')
                    DurationMin   = [math]::Round($durationMin, 1)
                    PrecedentType = switch ($precedent.Id) { 6006 {'Shutdown propre (legacy)'}; 41 {'Crash/Freeze (legacy)'} }
                    EstBootLong   = ($durationMin -gt $SeuilBootLong)
                    Method        = 'Legacy'
                    BootType      = 'Unknown'
                })
                $bootsByType['Unknown']++
            }
        }
    }
}

Write-Log ("Boots par type : ColdBoot={0} FastStartup={1} Resume={2} Unknown={3}" -f `
    $bootsByType.ColdBoot, $bootsByType.FastStartup, $bootsByType.Resume, $bootsByType.Unknown)

# ============================================================
# v1.2 : RECALCUL DE LastBoot ET UptimeDays
#
# Algorithme corrige apres analyse terrain (v1.1 etait incomplete).
#
# Le probleme de v1.1 : $bootDurations capture les couples Event 12 + 27
# (Kernel-General + Kernel-Boot). Ces events sont emis pour les vrais
# boots et les Fast Startup, MAIS PAS pour les wakes Modern Standby.
#
# Sur un laptop moderne avec Fast Startup + Modern Standby, l'utilisateur
# clique "Arreter" le soir -> Event 42 (veille), et allume le matin ->
# Event 507 (wake Modern Standby), sans aucun Event 12.
#
# Solution : combiner les 2 sources
#   LastUserOn = MAX(
#     derniere entree de $bootDurations,  # couvre cold boot et Fast Startup
#     dernier Event 507                   # couvre wake Modern Standby
#   )
# ============================================================

# Recuperer le dernier Event 507 (wake Modern Standby)
$lastWakeModernStandby = $null
try {
    $wakeEvent = Get-WinEvent -FilterHashtable @{
        LogName      = 'System'
        ProviderName = 'Microsoft-Windows-Kernel-Power'
        Id           = 507
        StartTime    = $dateDebut
    } -MaxEvents 1 -ErrorAction SilentlyContinue

    if ($wakeEvent) {
        $lastWakeModernStandby = $wakeEvent.TimeCreated
        Write-Log "Wake Modern Standby detecte : $($lastWakeModernStandby.ToString('yyyy-MM-dd HH:mm:ss'))"
    } else {
        Write-Log "Wake Modern Standby : aucun Event 507 trouve (normal sur desktop)"
    }
} catch {
    Write-Log "Wake Modern Standby : erreur de lecture ($_)"
}

# Calculer la derniere reprise d'activite
$userLastBoot = $null
$bootTypeLabel = 'none'

if ($bootDurations.Count -gt 0) {
    $mostRecentBoot = $bootDurations |
        Sort-Object -Property @{ Expression = { [datetime]$_.DateBoot } } -Descending |
        Select-Object -First 1
    $userLastBoot = [datetime]$mostRecentBoot.DateBoot
    $bootTypeLabel = "BootDurations($($mostRecentBoot.BootType))"
}

# Si Event 507 plus recent que la derniere entree BootDurations, il gagne
if ($lastWakeModernStandby -and ($null -eq $userLastBoot -or $lastWakeModernStandby -gt $userLastBoot)) {
    $userLastBoot = $lastWakeModernStandby
    $bootTypeLabel = 'ModernStandbyWake(Event 507)'
}

# Appliquer le resultat
if ($userLastBoot) {
    $userUptimeDays = [math]::Round(((Get-Date) - $userLastBoot).TotalDays, 2)
    $machineInfo.LastBoot   = $userLastBoot.ToString('yyyy-MM-dd HH:mm:ss')
    $machineInfo.UptimeDays = $userUptimeDays

    Write-Log ("v1.2 Uptime corrige : LastBoot user = {0} (source: {1}) | UptimeDays = {2} | Kernel cold boot = {3}" -f `
        $machineInfo.LastBoot, $bootTypeLabel, $userUptimeDays, $machineInfo.LastRealColdBoot)
} else {
    Write-Log "v1.2 Uptime : aucune source trouvee, conservation des valeurs kernel"
}

# ============================================================
# 5. BSOD via Minidump (top 10 recents)
# ============================================================
$bsods = @()
$minidumpPath = 'C:\Windows\Minidump'
if (Test-Path $minidumpPath) {
    $bsods = Get-ChildItem -Path $minidumpPath -Filter '*.dmp' -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -ge $dateDebut } |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 10 |
        ForEach-Object {
            [PSCustomObject]@{
                Date   = $_.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss')
                Nom    = $_.Name
                Taille = "$([math]::Round($_.Length / 1MB, 1)) MB"
            }
        }
}

# ============================================================
# 6. RESOURCE WARNINGS
#    Consolides depuis les events deja collectes au point 2
# ============================================================
$resourceWarnings = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($ev in $ramEvents) {
    $detail = ''
    try {
        if ($ev.Message -match '(\w+\.exe)') { $detail = $matches[1] }
    } catch {}
    $resourceWarnings.Add([PSCustomObject]@{
        Timestamp = $ev.TimeCreated.ToString('yyyy-MM-dd HH:mm:ss')
        Type      = 'RAM exhaustion'
        Detail    = $detail
    })
}
foreach ($ev in $cpuEvents) {
    $resourceWarnings.Add([PSCustomObject]@{
        Timestamp = $ev.TimeCreated.ToString('yyyy-MM-dd HH:mm:ss')
        Type      = 'CPU throttling'
        Detail    = 'Performance reduite'
    })
}
foreach ($ev in $diskFullEvents) {
    $detail = ''
    try {
        if ($ev.Message -match '([A-Z]:)') { $detail = $matches[1] }
    } catch {}
    $resourceWarnings.Add([PSCustomObject]@{
        Timestamp = $ev.TimeCreated.ToString('yyyy-MM-dd HH:mm:ss')
        Type      = 'Disk full'
        Detail    = $detail
    })
}
# v1.5 : Clustering des Event 51 (Disk slow / I/O timeout)
# Windows emet souvent des dizaines voire centaines d'Event 51 en
# meme temps (un par secteur/IRP en timeout) pour UN seul incident.
# Ex parc pilote : 894 events sur 1 seule seconde = 1 incident matos,
# pas 894 evenements distincts.
# On groupe donc par fenetre de 60s. Chaque cluster ajoute 1 entree avec :
#   - Count      : nombre d'events dans le cluster
#   - FirstSeen  : premier timestamp du cluster
#   - LastSeen   : dernier timestamp du cluster
#   - IsBurst    : true si Count > 50 (signal matos fort)
# Le Dashboard affichera "x894 events en 1s" avec un badge si burst.
if ($diskSlowEvents.Count -gt 0) {
    $SeuilBurst = 50  # Au-dela on considere que c'est un incident materiel
    $FenetreSec = 60  # Fenetre de regroupement en secondes

    # Tri chronologique croissant pour le clustering
    $sortedSlow = @($diskSlowEvents | Sort-Object TimeCreated)

    $clusters = [System.Collections.Generic.List[object]]::new()
    $currentCluster = $null

    foreach ($ev in $sortedSlow) {
        if ($null -eq $currentCluster) {
            # Premier event : initialise le cluster
            $currentCluster = [PSCustomObject]@{
                FirstTime = $ev.TimeCreated
                LastTime  = $ev.TimeCreated
                Count     = 1
            }
        } else {
            $diffSec = ($ev.TimeCreated - $currentCluster.LastTime).TotalSeconds
            if ($diffSec -le $FenetreSec) {
                # Meme cluster : on etend
                $currentCluster.LastTime = $ev.TimeCreated
                $currentCluster.Count++
            } else {
                # Nouveau cluster : on ferme l'ancien et on en demarre un
                $clusters.Add($currentCluster)
                $currentCluster = [PSCustomObject]@{
                    FirstTime = $ev.TimeCreated
                    LastTime  = $ev.TimeCreated
                    Count     = 1
                }
            }
        }
    }
    # Fermer le dernier cluster ouvert
    if ($null -ne $currentCluster) { $clusters.Add($currentCluster) }

    # Ecrire 1 entree par cluster dans ResourceWarnings
    foreach ($c in $clusters) {
        $resourceWarnings.Add([PSCustomObject]@{
            Timestamp = $c.FirstTime.ToString('yyyy-MM-dd HH:mm:ss')
            Type      = 'Disk slow'
            Detail    = 'I/O timeout'
            Count     = $c.Count
            FirstSeen = $c.FirstTime.ToString('yyyy-MM-dd HH:mm:ss')
            LastSeen  = $c.LastTime.ToString('yyyy-MM-dd HH:mm:ss')
            IsBurst   = ($c.Count -ge $SeuilBurst)
        })
    }
}
$resourceWarnings = @($resourceWarnings | Sort-Object Timestamp -Descending)

# ============================================================
# 7. TOP 5 RAM (snapshot instant)
# ============================================================
$topRAM = @()
try {
    $topRAM = Get-Process | Sort-Object WorkingSet -Descending | Select-Object -First 5 | ForEach-Object {
        [PSCustomObject]@{
            Name         = $_.ProcessName
            WorkingSetMB = [math]::Round($_.WorkingSet / 1MB, 1)
            CPUSeconds   = [math]::Round($_.CPU, 1)
        }
    }
} catch {
    Write-Log "Erreur Top 5 RAM : $_"
}

# ============================================================
# 8. ESPACE DISQUE (volumes fixes)
# ============================================================
$diskInfo = [System.Collections.Generic.List[PSCustomObject]]::new()
try {
    $volumes = Get-CimInstance -ClassName Win32_LogicalDisk -Filter 'DriveType=3' -ErrorAction SilentlyContinue
    foreach ($vol in $volumes) {
        $totalGB = [math]::Round($vol.Size / 1GB, 1)
        $freeGB  = [math]::Round($vol.FreeSpace / 1GB, 1)
        $usedGB  = [math]::Round(($vol.Size - $vol.FreeSpace) / 1GB, 1)
        $pctUsed = if ($vol.Size -gt 0) { [math]::Round((($vol.Size - $vol.FreeSpace) / $vol.Size) * 100, 0) } else { 0 }
        $pctFree = 100 - $pctUsed
        $isAlert = ($pctFree -lt $SeuilDiskAlert)

        $diskInfo.Add([PSCustomObject]@{
            Drive   = $vol.DeviceID
            Label   = if ($vol.VolumeName) { $vol.VolumeName } else { 'Sans nom' }
            TotalGB = $totalGB
            UsedGB  = $usedGB
            FreeGB  = $freeGB
            PctUsed = $pctUsed
            PctFree = $pctFree
            IsAlert = $isAlert
        })
    }
} catch {
    Write-Log "Erreur espace disque : $_"
}

# ============================================================
# 9. BATTERIE (laptop uniquement, desktop/VM ignores)
#    - DesignCapacity (WMI root\wmi\BatteryStaticData) vs
#      FullChargeCapacity (WMI root\wmi\BatteryFullChargedCapacity)
#      => pourcentage d'usure reel de la batterie
#    - BatteryStatus decode via enum Win32_Battery (1=Discharging,
#      2=AC Power, 3=Fully Charged, 4=Low, 5=Critical, 6=Charging...)
#    - Alerte si health < 60%
#    - HasBattery=false si pas de batterie (normal pour desktop/VM)
# ============================================================
$batteryInfo = [PSCustomObject]@{
    HasBattery         = $false
    Manufacturer       = ''
    Chemistry          = ''
    DesignCapacity     = 0
    FullChargeCapacity = 0
    HealthPercent      = 0
    HealthCategory     = 'Inconnu'
    CycleCount         = $null
    CurrentChargePct   = 0
    Status             = 'N/A'
    IsAlert            = $false
}

try {
    # Win32_Battery : presence + etat instantane
    $battery = Get-CimInstance -ClassName Win32_Battery -ErrorAction SilentlyContinue | Select-Object -First 1

    if ($battery) {
        $batteryInfo.HasBattery       = $true
        $batteryInfo.CurrentChargePct = [int]$battery.EstimatedChargeRemaining

        # BatteryStatus : enumeration 1-11 -> libelle lisible
        $batteryInfo.Status = switch ([int]$battery.BatteryStatus) {
            1       { 'Discharging' }
            2       { 'AC Power' }
            3       { 'Fully Charged' }
            4       { 'Low' }
            5       { 'Critical' }
            6       { 'Charging' }
            7       { 'Charging High' }
            8       { 'Charging Low' }
            9       { 'Charging Critical' }
            10      { 'Undefined' }
            11      { 'Partially Charged' }
            default { "Unknown ($($battery.BatteryStatus))" }
        }

        # Chimie : enumeration Win32_Battery.Chemistry
        if ($battery.Chemistry) {
            $batteryInfo.Chemistry = switch ([int]$battery.Chemistry) {
                1       { 'Other' }
                2       { 'Unknown' }
                3       { 'Lead Acid' }
                4       { 'Nickel Cadmium' }
                5       { 'Nickel Metal Hydride' }
                6       { 'Lithium Ion' }
                7       { 'Zinc Air' }
                8       { 'Lithium Polymer' }
                default { "Type $($battery.Chemistry)" }
            }
        }

        # Capacites reelles : 2 methodes en cascade pour maximiser la
        # couverture du parc.
        #
        # Essai 1 : BatteryStaticData + BatteryFullChargedCapacity (WMI)
        # Rapide mais n'est PAS expose par tous les firmwares ACPI.
        # Cas connus de defaillance : AMD Ryzen AI (2024+), certains
        # Intel Gen 13+, OEMs avec firmwares specifiques.
        #
        # Essai 2 (fallback) : powercfg /batteryreport /xml
        # Lit directement le firmware ACPI sans passer par le provider WMI.
        # Fonctionne sur tous les laptops Windows 8+. Un peu plus lent
        # (~1-2s pour generer le XML) mais tres fiable.
        $capacitiesResolved = $false

        try {
            $designData    = Get-CimInstance -Namespace 'root\wmi' -ClassName 'BatteryStaticData'          -ErrorAction Stop | Select-Object -First 1
            $fullChargeData= Get-CimInstance -Namespace 'root\wmi' -ClassName 'BatteryFullChargedCapacity' -ErrorAction Stop | Select-Object -First 1

            if ($designData -and $fullChargeData -and $designData.DesignedCapacity -gt 0) {
                $batteryInfo.DesignCapacity     = [int]$designData.DesignedCapacity
                $batteryInfo.FullChargeCapacity = [int]$fullChargeData.FullChargedCapacity

                # ManufactureName est un tableau de caracteres -> on reassemble
                if ($designData.ManufactureName) {
                    $mfgRaw = if ($designData.ManufactureName -is [array]) {
                        -join ($designData.ManufactureName | ForEach-Object { [char]$_ })
                    } else {
                        [string]$designData.ManufactureName
                    }
                    $batteryInfo.Manufacturer = $mfgRaw.Trim([char]0, ' ')
                }

                $capacitiesResolved = $true
            }
        } catch {
            Write-Log "BatteryStaticData indisponible ($_) - fallback powercfg"
        }

        # Fallback : powercfg /batteryreport (AMD Ryzen AI, etc.)
        if (-not $capacitiesResolved) {
            try {
                $reportFile = Join-Path $env:TEMP "battreport_$($env:COMPUTERNAME).xml"
                if (Test-Path $reportFile) {
                    Remove-Item $reportFile -Force -ErrorAction SilentlyContinue
                }

                # Execution silencieuse, duration 1 jour pour minimiser la generation
                # Remarque : on redirige stdout+stderr pour ne pas polluer le pipeline
                $null = & powercfg.exe /batteryreport /xml /duration 1 /output $reportFile 2>&1

                if (Test-Path $reportFile) {
                    [xml]$report = Get-Content $reportFile -Raw -ErrorAction Stop
                    # En cas de multi-batterie (rare), on prend la premiere
                    $batt = @($report.BatteryReport.Batteries.Battery)[0]

                    if ($batt -and $batt.DesignCapacity -and [int64]$batt.DesignCapacity -gt 0) {
                        $batteryInfo.DesignCapacity     = [int64]$batt.DesignCapacity
                        $batteryInfo.FullChargeCapacity = [int64]$batt.FullChargeCapacity

                        if ($batt.Manufacturer) {
                            $batteryInfo.Manufacturer = ([string]$batt.Manufacturer).Trim()
                        }
                        # powercfg peut fournir Chemistry en string (LION, LIPO...)
                        # Ne pas ecraser si Win32_Battery a deja donne une valeur propre
                        if (($batteryInfo.Chemistry -eq '' -or $batteryInfo.Chemistry -eq 'Unknown') -and $batt.Chemistry) {
                            $batteryInfo.Chemistry = ([string]$batt.Chemistry).Trim()
                        }
                        # CycleCount de secours (le bloc BatteryCycleCount WMI en aval peut l'ecraser)
                        if (-not $batteryInfo.CycleCount -and $batt.CycleCount) {
                            $cc = 0
                            if ([int]::TryParse([string]$batt.CycleCount, [ref]$cc) -and $cc -gt 0) {
                                $batteryInfo.CycleCount = $cc
                            }
                        }

                        $capacitiesResolved = $true
                        Write-Log "Batterie : capacites resolues via powercfg (fallback)"
                    } else {
                        Write-Log 'powercfg batteryreport : XML parse OK mais DesignCapacity vide'
                    }

                    # Nettoyage du fichier temporaire (best-effort)
                    Remove-Item $reportFile -Force -ErrorAction SilentlyContinue
                } else {
                    Write-Log 'powercfg batteryreport : fichier XML non genere'
                }
            } catch {
                Write-Log "Fallback powercfg echec : $_"
            }
        }

        # Calcul final du pourcentage de sante (commun aux 2 methodes)
        if ($capacitiesResolved -and $batteryInfo.DesignCapacity -gt 0) {
            $healthPct = [math]::Round(($batteryInfo.FullChargeCapacity * 100.0) / $batteryInfo.DesignCapacity, 1)
            # Cap a 100 au cas ou (peut arriver sur batterie neuve surchargee par firmware)
            if ($healthPct -gt 100) { $healthPct = 100 }

            $batteryInfo.HealthPercent   = $healthPct
            $batteryInfo.HealthCategory  = if     ($healthPct -ge 80) { 'Bonne' }
                                           elseif ($healthPct -ge 50) { 'Degradee' }
                                           else                      { 'Critique' }
            $batteryInfo.IsAlert         = ($healthPct -lt 60)
        }

        # CycleCount : optionnel, absent chez beaucoup d'OEMs grand public
        try {
            $cycleData = Get-CimInstance -Namespace 'root\wmi' -ClassName 'BatteryCycleCount' -ErrorAction Stop | Select-Object -First 1
            if ($cycleData -and $cycleData.CycleCount -gt 0) {
                $batteryInfo.CycleCount = [int]$cycleData.CycleCount
            }
        } catch {
            # Normal si non supporte par l'OEM, on ignore silencieusement
        }

        Write-Log ("Batterie : Health={0}% ({1}) Charge={2}% Status={3}" -f `
            $batteryInfo.HealthPercent, $batteryInfo.HealthCategory,
            $batteryInfo.CurrentChargePct, $batteryInfo.Status)
    } else {
        Write-Log 'Pas de batterie detectee (desktop ou VM)'
    }
} catch {
    Write-Log "Erreur batterie : $_"
}

# ============================================================
# 10. SERVICES CRITIQUES (EDR)
#    - v5.3 : surveillance du service Sentinel Agent uniquement
#    - Structure extensible : on pourra ajouter FortiClient, Intune,
#      Dell Command Update dans les versions futures sans casser
#      le schema cote Dashboard
#    - Match sur nom interne (SentinelAgent) ET display name
#      (Sentinel Agent) pour robustesse multi-versions
# ============================================================
$servicesHealth = [PSCustomObject]@{
    SentinelAgent = [PSCustomObject]@{
        DisplayName = 'Sentinel Agent'
        ServiceName = 'SentinelAgent'
        Installed   = $false
        Status      = 'NotInstalled'
        StartType   = 'N/A'
        IsAlert     = $false
    }
}

try {
    # Tentative 1 : nom interne (le plus robuste)
    $sentinelSvc = Get-Service -Name 'SentinelAgent' -ErrorAction SilentlyContinue
    # Tentative 2 : display name (au cas ou une version future renommerait)
    if (-not $sentinelSvc) {
        $sentinelSvc = Get-Service -DisplayName 'Sentinel Agent' -ErrorAction SilentlyContinue
    }

    if ($sentinelSvc) {
        $servicesHealth.SentinelAgent.Installed = $true
        $servicesHealth.SentinelAgent.Status    = $sentinelSvc.Status.ToString()
        $servicesHealth.SentinelAgent.StartType = $sentinelSvc.StartType.ToString()
        # Alerte si service present mais pas en etat Running
        $servicesHealth.SentinelAgent.IsAlert   = ($sentinelSvc.Status -ne 'Running')

        Write-Log ("SentinelAgent : Status={0} StartType={1} Alert={2}" -f `
            $sentinelSvc.Status, $sentinelSvc.StartType,
            $servicesHealth.SentinelAgent.IsAlert)
    } else {
        # Service totalement absent : c'est une alerte serieuse
        # (SentinelOne doit etre deploye via Intune sur tout le parc)
        $servicesHealth.SentinelAgent.IsAlert = $true
        Write-Log 'SentinelAgent : service NON INSTALLE (alerte)'
    }
} catch {
    Write-Log "Erreur verification SentinelAgent : $_"
}

# ============================================================
# 11. BOOT PERFORMANCE (v5.4)
#    - Analyse fine des cold boots via Event 100
#    - 5 derniers boots (~1 semaine d'historique) + stats agregees
#    - MainPathBootTime : logo Windows -> logon screen
#    - BootPostBootTime : logon screen -> systeme utilisable (80% idle)
#    - UserProfileProcessingTime : chargement profil user
#    - Seuil d'alerte : BootPostBootTime > 90s OU >= 2 slow boots sur 5
#    - Note : les Fast Startup et Resume ne genèrent pas d'Event 100
#      (c'est normal, ce ne sont pas des vrais boots)
# ============================================================
$SeuilPostBootMs = 90000   # 90 secondes = alerte

$bootPerfHistory = [System.Collections.Generic.List[PSCustomObject]]::new()
try {
    foreach ($ev in $bootPerfEvents) {
        $parsed = ConvertFrom-BootPerfEvent -Event $ev
        if ($parsed) {
            # Flag individuel : slow boot si post-boot > seuil
            $parsed | Add-Member -NotePropertyName 'IsSlow' -NotePropertyValue ($parsed.BootPostBootTimeMs -gt $SeuilPostBootMs)
            $bootPerfHistory.Add($parsed)
        }
    }
} catch {
    Write-Log "Erreur parsing Boot Perf : $_"
}

# Tri par date decroissante (le plus recent en premier)
$bootPerfHistorySorted = @($bootPerfHistory | Sort-Object Timestamp -Descending)

# Stats agregees sur les boots remontes
$slowBootsCount   = @($bootPerfHistorySorted | Where-Object { $_.IsSlow }).Count
$avgBootTimeMs    = 0
$avgMainPathMs    = 0
$avgPostBootMs    = 0
$maxBootTimeMs    = 0

if ($bootPerfHistorySorted.Count -gt 0) {
    $avgBootTimeMs = [int]([math]::Round((($bootPerfHistorySorted | Measure-Object -Property BootTimeMs         -Average).Average), 0))
    $avgMainPathMs = [int]([math]::Round((($bootPerfHistorySorted | Measure-Object -Property MainPathBootTimeMs -Average).Average), 0))
    $avgPostBootMs = [int]([math]::Round((($bootPerfHistorySorted | Measure-Object -Property BootPostBootTimeMs -Average).Average), 0))
    $maxBootTimeMs = [int](($bootPerfHistorySorted  | Measure-Object -Property BootTimeMs         -Maximum).Maximum)
}

# LastBoot = le plus recent, ou $null si aucun event remonte
$lastBoot = if ($bootPerfHistorySorted.Count -gt 0) { $bootPerfHistorySorted[0] } else { $null }
$lastBootAlert = if ($lastBoot) { $lastBoot.IsSlow } else { $false }

$bootPerformance = [PSCustomObject]@{
    LastBoot = $lastBoot
    History  = @($bootPerfHistorySorted)
    Stats    = [PSCustomObject]@{
        BootsAnalyzed    = $bootPerfHistorySorted.Count
        AvgBootTimeMs    = $avgBootTimeMs
        AvgMainPathMs    = $avgMainPathMs
        AvgPostBootMs    = $avgPostBootMs
        MaxBootTimeMs    = $maxBootTimeMs
        SlowBootsCount   = $slowBootsCount
    }
    # Alerte globale : soit le dernier boot est lent, soit >= 2 sur les 5
    IsAlert  = ($lastBootAlert -or ($slowBootsCount -ge 2))
}

if ($lastBoot) {
    Write-Log ("Boot Perf : Last={0}ms (Main={1}ms Post={2}ms) Avg={3}ms SlowBoots={4}/{5}" -f `
        $lastBoot.BootTimeMs, $lastBoot.MainPathBootTimeMs, $lastBoot.BootPostBootTimeMs,
        $avgBootTimeMs, $slowBootsCount, $bootPerfHistorySorted.Count)
} else {
    Write-Log 'Boot Perf : aucun Event 100 dans la periode (machine sans cold boot recent ?)'
}

# ============================================================
# 12. DISK HEALTH SMART (v5.4)
#    - Get-PhysicalDisk         : infos basiques + HealthStatus global
#    - Get-StorageReliabilityCounter : SMART detaille (Temp, Wear, Errors)
#    - Seuils d'alerte :
#        * HealthStatus != Healthy
#        * Wear >= 70%       (SSD : 100% = fin de vie estimee)
#        * Temperature >= 70C
#        * ReadErrorsUncorrected  > 0
#        * WriteErrorsUncorrected > 0
#    - Un element du tableau par disque physique (multi-disque supporte)
#    - Gestion defensive des champs vides (certains OEMs ne remontent
#      pas toutes les metriques)
# ============================================================
$SeuilDiskWear     = 70    # %
$SeuilDiskTempC    = 70    # °C

# v5.5 : whitelist des BusType consideres comme disques "internes" fiables.
# Une cle USB branchee ou un disque externe n'a pas de reelle metrique SMART
# exploitable et polluerait le dashboard. Pattern inspire de la doc NinjaOne.
$InternalDiskBusTypes = @('NVMe', 'SATA', 'SAS', 'SCSI', 'RAID', 'ATA', 'ATAPI')

$diskHealth = [System.Collections.Generic.List[PSCustomObject]]::new()
$diskSkipped = 0
try {
    $allDisks = @(Get-PhysicalDisk -ErrorAction SilentlyContinue)
    # Filtrage : on garde uniquement les BusType whitelisted, et on exclut
    # tout ce qui ressemble a un support amovible (PhysicalLocation *USB*,
    # MediaType 'Removable Media', etc.)
    $physicalDisks = @($allDisks | Where-Object {
        $busType     = [string]$_.BusType
        $location    = [string]$_.PhysicalLocation
        $mediaType   = [string]$_.MediaType

        ($busType -in $InternalDiskBusTypes) -and
        ($location -notlike '*USB*') -and
        ($busType  -notlike '*USB*') -and
        ($busType  -ne 'File Backed Virtual') -and
        ($mediaType -ne 'Removable Media')
    })
    $diskSkipped = $allDisks.Count - $physicalDisks.Count
    if ($diskSkipped -gt 0) {
        Write-Log ("DiskHealth : {0} disque(s) externe(s)/USB ignore(s) (filtre BusType)" -f $diskSkipped)
    }

    foreach ($pd in $physicalDisks) {
        $reliability = $null
        try {
            $reliability = Get-StorageReliabilityCounter -PhysicalDisk $pd -ErrorAction Stop
        } catch {
            # Certains disques (virtuels, USB, RAID) refusent le counter
            # On continue avec juste les infos basiques Get-PhysicalDisk
        }

        # v6.1 : acces direct aux proprietes au lieu d'un helper scriptblock.
        # Avant : $getRel = { param($name) ... }   +   & $getRel 'Temperature' x8
        # Ce pattern etait elegant mais : (a) recreait le scriptblock a chaque
        # disque, (b) payait le cout de param()+return a chaque appel.
        # Maintenant : un ternaire inline, 2x plus rapide sur PS 5.1.
        if ($reliability) {
            $temperature     = $reliability.Temperature
            $temperatureMax  = $reliability.TemperatureMax
            $wear            = $reliability.Wear
            $powerOnHours    = $reliability.PowerOnHours
            $readErrsTotal   = $reliability.ReadErrorsTotal
            $readErrsUncorr  = $reliability.ReadErrorsUncorrected
            $writeErrsTotal  = $reliability.WriteErrorsTotal
            $writeErrsUncorr = $reliability.WriteErrorsUncorrected
        } else {
            $temperature = $temperatureMax = $wear = $powerOnHours = $null
            $readErrsTotal = $readErrsUncorr = $writeErrsTotal = $writeErrsUncorr = $null
        }

        # Determination des alertes
        $alertReasons = [System.Collections.Generic.List[string]]::new()
        if ($pd.HealthStatus -and [string]$pd.HealthStatus -ne 'Healthy') {
            $alertReasons.Add("Health=$($pd.HealthStatus)")
        }
        if ($null -ne $wear -and [int]$wear -ge $SeuilDiskWear) {
            $alertReasons.Add("Wear>=$SeuilDiskWear%")
        }
        if ($null -ne $temperature -and [int]$temperature -ge $SeuilDiskTempC) {
            $alertReasons.Add("Temp>=$SeuilDiskTempC" + "C")
        }
        if ($null -ne $readErrsUncorr -and [int64]$readErrsUncorr -gt 0) {
            $alertReasons.Add("ReadErrUncorr>0")
        }
        if ($null -ne $writeErrsUncorr -and [int64]$writeErrsUncorr -gt 0) {
            $alertReasons.Add("WriteErrUncorr>0")
        }

        $diskHealth.Add([PSCustomObject]@{
            FriendlyName           = [string]$pd.FriendlyName
            MediaType              = [string]$pd.MediaType
            BusType                = [string]$pd.BusType
            SizeGB                 = [math]::Round(($pd.Size / 1GB), 1)
            OperationalStatus      = [string]$pd.OperationalStatus
            HealthStatus           = [string]$pd.HealthStatus
            TemperatureC           = if ($null -ne $temperature)    { [int]$temperature }    else { $null }
            TemperatureMaxC        = if ($null -ne $temperatureMax) { [int]$temperatureMax } else { $null }
            WearPct                = if ($null -ne $wear)           { [int]$wear }           else { $null }
            PowerOnHours           = if ($null -ne $powerOnHours)   { [int64]$powerOnHours } else { $null }
            ReadErrorsTotal        = if ($null -ne $readErrsTotal)  { [int64]$readErrsTotal } else { $null }
            ReadErrorsUncorrected  = if ($null -ne $readErrsUncorr) { [int64]$readErrsUncorr } else { $null }
            WriteErrorsTotal       = if ($null -ne $writeErrsTotal) { [int64]$writeErrsTotal } else { $null }
            WriteErrorsUncorrected = if ($null -ne $writeErrsUncorr){ [int64]$writeErrsUncorr } else { $null }
            IsAlert                = ($alertReasons.Count -gt 0)
            AlertReasons           = @($alertReasons)
        })
    }
} catch {
    Write-Log "Erreur DiskHealth : $_"
}

# Stats globales disk
$diskWorstWear       = 0
$diskHealthAlert     = $false
foreach ($d in $diskHealth) {
    if ($null -ne $d.WearPct -and [int]$d.WearPct -gt $diskWorstWear) {
        $diskWorstWear = [int]$d.WearPct
    }
    if ($d.IsAlert) { $diskHealthAlert = $true }
}

Write-Log ("DiskHealth : {0} disques analyses, WorstWear={1}% Alert={2}" -f `
    $diskHealth.Count, $diskWorstWear, $diskHealthAlert)

# ============================================================
# 12b. MONITORS (v5.5) - Inventaire des ecrans EXTERNES branches
#    - WmiMonitorID           : EDID de chaque ecran (nom, serie, fab)
#    - WmiMonitorConnectionParams : VideoOutputTechnology pour distinguer
#      les ecrans internes (laptop LCD) des ecrans externes
#    - On exclut les ecrans internes (pas interessant pour l'inventaire
#      secondaire, on sait deja que le laptop a un ecran)
#    - Decodage EDID : les champs sont des byte[] ASCII avec padding 0
#    - Table de 15 fabricants majeurs, code brut conserve pour les autres
# ============================================================

# Codes internes = ecran LCD laptop / panel embarque (on les exclut)
#   6  = LVDS (classique laptop)
#   11 = DisplayPort Internal
#   13 = Unified Display Interface Embedded
#   16 = Internal
#   2147483648 = Internal (valeur etendue Windows)
$InternalVideoOutputTypes = @(6, 11, 13, 16, 2147483648)

# Table des fabricants EDID majeurs. Pour les codes non listes, on garde
# le code brut (3 lettres) dans le JSON.
$EdidManufacturers = @{
    'DEL' = 'Dell Inc.'
    'HPN' = 'HP Inc.'
    'HWP' = 'HP Inc.'
    'SAM' = 'Samsung Electronics'
    'SNG' = 'Samsung Electronics'
    'GSM' = 'LG Electronics'
    'LGD' = 'LG Display'
    'ACR' = 'Acer'
    'AUS' = 'ASUS'
    'AUO' = 'AU Optronics'
    'VSC' = 'ViewSonic'
    'LEN' = 'Lenovo'
    'PHL' = 'Philips'
    'AOC' = 'AOC'
    'BNQ' = 'BenQ'
    'ENC' = 'Eizo'
    'NEC' = 'NEC'
    'IVM' = 'Iiyama'
    'MSI' = 'MSI'
    'MEI' = 'Panasonic'
    'CMN' = 'Chimei Innolux'
    'BOE' = 'BOE Technology'
    'SHP' = 'Sharp'
    'CMO' = 'Chi Mei Optoelectronics'
    'IBM' = 'IBM'
    'APP' = 'Apple'
}

# Helper : decode un byte[] EDID (ASCII avec padding 0) en string propre
function ConvertFrom-EdidBytes {
    param([uint16[]]$Bytes)
    if ($null -eq $Bytes -or $Bytes.Count -eq 0) { return '' }
    try {
        # Filtrer les zeros de padding, convertir chaque code en char
        $chars = $Bytes | Where-Object { $_ -ne 0 } | ForEach-Object { [char]$_ }
        return (-join $chars).Trim()
    } catch {
        return ''
    }
}

$monitors = [System.Collections.Generic.List[PSCustomObject]]::new()
try {
    $monIds = @(Get-CimInstance -Namespace 'root\wmi' -ClassName 'WmiMonitorID' -ErrorAction SilentlyContinue)
    $monConns = @{}
    try {
        $connRaw = @(Get-CimInstance -Namespace 'root\wmi' -ClassName 'WmiMonitorConnectionParams' -ErrorAction SilentlyContinue)
        foreach ($c in $connRaw) {
            # On map par InstanceName pour retrouver la techno de sortie
            $monConns[[string]$c.InstanceName] = [uint32]$c.VideoOutputTechnology
        }
    } catch {
        Write-Log "WmiMonitorConnectionParams indisponible : $_"
    }

    $totalFound = $monIds.Count
    $internalSkipped = 0
    $aioSkipped      = 0

    # v5.6 : patterns d'ecran integre d'AIO a blacklister quand IsAIO
    # Les AIO ont un "ecran interne" qui remonte comme HDMI (tech 5) au
    # lieu de LVDS/Internal -> impossible de le filtrer par VideoOutputTech.
    # On se base sur le nom du modele qui contient typiquement le nom du
    # chassis du PC.
    $AIOMonitorPatterns = @(
        'AIO',
        'All-in-One',
        'All In One',
        'OptiPlex',     # Dell OptiPlex AIO
        'ThinkCentre',  # Lenovo ThinkCentre AIO
        'ProOne',       # HP ProOne
        'EliteOne',     # HP EliteOne
        'iMac'          # Apple iMac (tres improbable en entreprise mais safe)
    )

    foreach ($m in $monIds) {
        $instName = [string]$m.InstanceName
        $videoTech = $null
        if ($monConns.ContainsKey($instName)) { $videoTech = $monConns[$instName] }

        # Decodage EDID (utile avant le filtre AIO qui regarde le Model)
        $manuCode = ConvertFrom-EdidBytes -Bytes $m.ManufacturerName
        $model    = ConvertFrom-EdidBytes -Bytes $m.UserFriendlyName
        $serial   = ConvertFrom-EdidBytes -Bytes $m.SerialNumberID
        $prodCode = ConvertFrom-EdidBytes -Bytes $m.ProductCodeID

        # -------------------------------------------------------
        # Filtre 1 : ecrans internes (laptop LCD) via VideoOutputTech
        # -------------------------------------------------------
        if ($null -ne $videoTech -and $videoTech -in $InternalVideoOutputTypes) {
            $internalSkipped++
            continue
        }

        # -------------------------------------------------------
        # Filtre 2 (v5.6) : si machine AIO, on skip l'ecran integre
        # qui est identifie par un pattern dans son UserFriendlyName.
        # L'objectif metier : compter uniquement les ecrans "secondaires"
        # branches par l'utilisateur, pas l'ecran du chassis AIO lui-meme.
        # -------------------------------------------------------
        if ($isAIO -and $model) {
            $isAIOMonitor = $false
            foreach ($pattern in $AIOMonitorPatterns) {
                if ($model -like "*$pattern*") { $isAIOMonitor = $true; break }
            }
            if ($isAIOMonitor) {
                $aioSkipped++
                Write-Log "  Monitor AIO integre ignore : $model"
                continue
            }
        }

        # Traduction code fabricant -> nom lisible
        $manuName = $manuCode
        if ($manuCode -and $EdidManufacturers.ContainsKey($manuCode)) {
            $manuName = $EdidManufacturers[$manuCode]
        }

        # Annee/semaine de fabrication (si non expose, reste a 0)
        $year  = if ($m.YearOfManufacture) { [int]$m.YearOfManufacture } else { $null }
        $week  = if ($m.WeekOfManufacture) { [int]$m.WeekOfManufacture } else { $null }

        # Calcul de l'age (en annees) pour detecter les ecrans anciens
        $ageYears = $null
        if ($year -and $year -gt 1990 -and $year -lt 2100) {
            $ageYears = (Get-Date).Year - $year
        }

        $monitors.Add([PSCustomObject]@{
            ManufacturerCode  = $manuCode
            Manufacturer      = $manuName
            Model             = $model
            SerialNumber      = $serial
            ProductCode       = $prodCode
            YearOfManufacture = $year
            WeekOfManufacture = $week
            AgeYears          = $ageYears
            Active            = [bool]$m.Active
            # VideoOutputTechnology : null si WmiMonitorConnectionParams KO
            VideoOutputTech   = $videoTech
        })
    }

    $msg = "Monitors : {0} detecte(s), {1} interne(s) exclu(s)" -f $totalFound, $internalSkipped
    if ($aioSkipped -gt 0) { $msg += ", {0} ecran(s) AIO integre(s) exclu(s)" -f $aioSkipped }
    $msg += ", {0} ecran(s) secondaire(s) gardes" -f $monitors.Count
    Write-Log $msg
} catch {
    Write-Log "Erreur Monitors : $_"
}

# ============================================================
# 12c. MEMORY INVENTORY (v1.8)
#      - Win32_PhysicalMemoryArray : capacite max supportee par la
#        carte mere + nombre de slots total.
#      - Win32_PhysicalMemory       : detail des barrettes presentes
#        (taille, type DDRx, vitesse, fabricant, banque/slot).
#      - Permet de decider : upgrade possible (slots libres ou
#        barrette plus grosse) vs remplacement du PC.
# ============================================================
$memoryInventory = [PSCustomObject]@{
    TotalInstalledGB = 0
    MaxCapacityGB    = 0
    TotalSlots       = 0
    OccupiedSlots    = 0
    FreeSlots        = 0
    CanUpgrade       = $false
    Modules          = @()
}
try {
    # MaxCapacity en Ko dans Win32_PhysicalMemoryArray (oui, Ko, c'est Microsoft)
    $memArray = Get-CimInstance Win32_PhysicalMemoryArray -ErrorAction Stop |
                Select-Object -First 1
    if ($memArray) {
        $memoryInventory.MaxCapacityGB = [math]::Round($memArray.MaxCapacity / 1MB, 0)   # Ko -> Go
        $memoryInventory.TotalSlots    = [int]$memArray.MemoryDevices
    }

    $modules = @(Get-CimInstance Win32_PhysicalMemory -ErrorAction Stop)
    $memoryInventory.OccupiedSlots = $modules.Count
    $memoryInventory.FreeSlots     = [math]::Max(0, $memoryInventory.TotalSlots - $modules.Count)

    # Type DDR : SMBIOSMemoryType est plus fiable que MemoryType (souvent 0)
    # Codes standards : 20=DDR, 21=DDR2, 24=DDR3, 26=DDR4, 34=DDR5
    $typeMap = @{ 20='DDR'; 21='DDR2'; 22='DDR2 FB-DIMM'; 24='DDR3'; 26='DDR4'; 27='LPDDR3'; 28='LPDDR4'; 29='Logical'; 30='HBM'; 31='HBM2'; 32='DDR5'; 34='DDR5' }

    $totalBytes = 0
    foreach ($m in $modules) {
        $totalBytes += [int64]$m.Capacity
        $typeCode = [int]($m.SMBIOSMemoryType)
        $typeStr  = if ($typeMap.ContainsKey($typeCode)) { $typeMap[$typeCode] } else { "Type=$typeCode" }

        $memoryInventory.Modules += [PSCustomObject]@{
            Slot         = $m.DeviceLocator
            Bank         = $m.BankLabel
            CapacityGB   = [math]::Round($m.Capacity / 1GB, 0)
            Type         = $typeStr
            SpeedMHz     = [int]$m.Speed
            Manufacturer = ($m.Manufacturer -replace '\s+$', '')
            PartNumber   = ($m.PartNumber   -replace '\s+$', '')
        }
    }
    $memoryInventory.TotalInstalledGB = [math]::Round($totalBytes / 1GB, 0)

    # CanUpgrade : vrai si slots libres OU si la carte mere supporte
    # beaucoup plus que la quantite actuelle (barrettes remplacables par
    # des plus grosses).
    if ($memoryInventory.FreeSlots -gt 0) {
        $memoryInventory.CanUpgrade = $true
    } elseif ($memoryInventory.MaxCapacityGB -gt 0 -and
              $memoryInventory.TotalInstalledGB -lt ($memoryInventory.MaxCapacityGB * 0.75)) {
        # Moins de 75% de la capacite max installee -> upgrade possible
        # par remplacement de barrettes
        $memoryInventory.CanUpgrade = $true
    }

    Write-Log ("RAM : {0} Go installes, {1}/{2} slots, max={3} Go, upgrade={4}" -f `
        $memoryInventory.TotalInstalledGB, $memoryInventory.OccupiedSlots,
        $memoryInventory.TotalSlots, $memoryInventory.MaxCapacityGB, $memoryInventory.CanUpgrade)
} catch {
    Write-Log "Erreur MemoryInventory : $_"
}

# ============================================================
# 12d. GPU INVENTORY (v1.8)
#      - Win32_VideoController : nom, driver version, date du driver.
#      - On garde tous les GPU presents (laptop gaming = integre + dedie,
#        workstation = multiples cartes pro). On filtre les "pseudo" GPU
#        qui apparaissent parfois (Remote RDP adapter, Meta Hyper-V, etc.).
#      - On se limite a ces 3 champs pour rester sobre ; la VRAM via WMI
#        est peu fiable (AdapterRAM est un int32 qui overflow a 4 Go).
# ============================================================
$gpuInventory = @()
try {
    $gpus = @(Get-CimInstance Win32_VideoController -ErrorAction Stop | Where-Object {
        # Filtre des "GPU" non physiques
        $_.Name -and
        $_.Name -notmatch 'Remote|Mirror|RDP|Hyper-V|Basic Display|Meta |VNC'
    })
    foreach ($g in $gpus) {
        $driverDate = ''
        if ($g.DriverDate) {
            try { $driverDate = ([datetime]$g.DriverDate).ToString('yyyy-MM-dd') } catch {}
        }
        $gpuInventory += [PSCustomObject]@{
            Name          = $g.Name
            DriverVersion = $g.DriverVersion
            DriverDate    = $driverDate
        }
    }
    Write-Log ("GPU : {0} adapteur(s) detecte(s)" -f $gpuInventory.Count)
} catch {
    Write-Log "Erreur GPUInventory : $_"
}

# ============================================================
# 13. TOP CRASHERS (apps qui crashent le plus, agregees)
# ============================================================
$topCrashers = @()
try {
    if ($appEvents.Count -gt 0) {
        $crasherCounts = @{}
        foreach ($ev in $appEvents) {
            $appName = $null
            try { $appName = $ev.Properties[0].Value } catch {}
            if ($appName) {
                $key = $appName.ToLower()
                if ($crasherCounts.ContainsKey($key)) { $crasherCounts[$key]++ }
                else { $crasherCounts[$key] = 1 }
            }
        }
        $topCrashers = $crasherCounts.GetEnumerator() |
            Sort-Object Value -Descending |
            Select-Object -First 10 |
            ForEach-Object {
                [PSCustomObject]@{
                    AppName    = $_.Key
                    CrashCount = $_.Value
                }
            }
    }
} catch {
    Write-Log "Erreur Top Crashers : $_"
}

# ============================================================
# 14. HARDWARE HEALTH (v5.2 refondu)
#     - WHEA parses via Get-WHEAAnalysis (XML + multi-langue)
#     - Separation Fatal (Critical/Error) vs Corrected (Warning)
#     - Agregation des Corrected par signature (evite le spam de
#       1000 erreurs PCIe identiques du meme port)
#     - Le score de sante ne prend en compte que les Fatal
#     - Classification par composant reelle (CPU/RAM/PCIe/Autre),
#       plus par ID d'event (qui etait faux)
# ============================================================
$hardwareHealth = [PSCustomObject]@{
    # Erreurs fatales non corrigees : UNE par occurrence, a traiter
    # chacune comme un incident serieux (BSOD imminent ou passe)
    WHEA_Fatal     = [System.Collections.Generic.List[PSCustomObject]]::new()

    # Erreurs corrigees : AGREGEES par signature. Un seul item par
    # source distincte, avec un compteur. Volumetrie controlee.
    WHEA_Corrected = [System.Collections.Generic.List[PSCustomObject]]::new()

    GPU_TDR        = [System.Collections.Generic.List[PSCustomObject]]::new()
    Thermal        = [System.Collections.Generic.List[PSCustomObject]]::new()

    # v1.8 : CPU throttling detaille (Events 35/55 Kernel-Processor-Power).
    # Separe des Events 37/88/125 thermal (Kernel-Power) qui sont des
    # signaux materiels critiques plus larges. Ici on veut capter le
    # moment ou le CPU reduit sa frequence pour cause thermique, sans
    # forcement atteindre un arret securite. Signal avant-coureur utile.
    CPUThrottling  = [System.Collections.Generic.List[PSCustomObject]]::new()
}

# --- WHEA : parse + dispatch Fatal / Corrected ---
$correctedAgg = @{}   # Signature -> { Count, FirstSeen, LastSeen, Analysis }

foreach ($ev in $wheaEvents) {
    $a = Get-WHEAAnalysis -Event $ev

    if ($a.IsFatal) {
        # Les fatales sont rares et serieuses : on garde toutes les instances
        $hardwareHealth.WHEA_Fatal.Add([PSCustomObject]@{
            Timestamp   = $ev.TimeCreated.ToString('yyyy-MM-dd HH:mm:ss')
            EventId     = $a.EventId
            Severity    = $a.Severity
            Component   = $a.Component
            ErrorSource = $a.ErrorSource
            BDF         = $a.BDF
            Detail      = $a.Detail
        })
    } else {
        # Les corrigees peuvent etre massives : agregation par signature
        $sig = $a.Signature
        if (-not $correctedAgg.ContainsKey($sig)) {
            $correctedAgg[$sig] = [PSCustomObject]@{
                Signature   = $sig
                Count       = 0
                FirstSeen   = $ev.TimeCreated
                LastSeen    = $ev.TimeCreated
                EventId     = $a.EventId
                Component   = $a.Component
                ErrorSource = $a.ErrorSource
                BDF         = $a.BDF
                Detail      = $a.Detail
            }
        }
        $agg = $correctedAgg[$sig]
        $agg.Count++
        if ($ev.TimeCreated -lt $agg.FirstSeen) { $agg.FirstSeen = $ev.TimeCreated }
        if ($ev.TimeCreated -gt $agg.LastSeen)  { $agg.LastSeen  = $ev.TimeCreated }
    }
}

# Finalise la liste Corrected : trie par nombre d'occurrences decroissant
$corrSorted = $correctedAgg.Values | Sort-Object -Property Count -Descending
foreach ($c in $corrSorted) {
    $hardwareHealth.WHEA_Corrected.Add([PSCustomObject]@{
        Component   = $c.Component
        EventId     = $c.EventId
        ErrorSource = $c.ErrorSource
        BDF         = $c.BDF
        Count       = $c.Count
        FirstSeen   = $c.FirstSeen.ToString('yyyy-MM-dd HH:mm:ss')
        LastSeen    = $c.LastSeen.ToString('yyyy-MM-dd HH:mm:ss')
        Detail      = $c.Detail
    })
}

# --- GPU TDR ---
foreach ($ev in $gpuEvents) {
    $driverName = ''
    try {
        if     ($ev.Message -match '(\w+\.sys)')                 { $driverName = $matches[1] }
        elseif ($ev.Message -match '(NVIDIA|AMD|Intel|Radeon)')  { $driverName = $matches[1] }
    } catch {}
    $hardwareHealth.GPU_TDR.Add([PSCustomObject]@{
        Timestamp = $ev.TimeCreated.ToString('yyyy-MM-dd HH:mm:ss')
        Driver    = $driverName
        Detail    = Get-TruncatedMessage -Text $ev.Message
    })
}

# --- Thermal (37 = throttling, 88 = arret securite, 125 = surchauffe) ---
foreach ($ev in $thermalEvents) {
    $temperature = 'N/A'
    try {
        if     ($ev.Message -match '(\d+)\s*[^0-9]?C')  { $temperature = "$($matches[1])C" }
        elseif ($ev.Message -match '(\d+)\s*degr')      { $temperature = "$($matches[1])C" }
    } catch {}

    $alertType = switch ($ev.Id) {
        37      { 'Throttling' }
        88      { 'Arret Securite' }
        125     { 'Surchauffe Critique' }
        default { 'Alerte Thermique' }
    }
    $hardwareHealth.Thermal.Add([PSCustomObject]@{
        Timestamp   = $ev.TimeCreated.ToString('yyyy-MM-dd HH:mm:ss')
        AlertType   = $alertType
        Temperature = $temperature
        Detail      = Get-TruncatedMessage -Text $ev.Message
    })
}

# --- v1.8 : CPU Throttling (Events 35/55 Kernel-Processor-Power) ---
# L'Event 26 est juste informatif (capabilities au boot), on l'ignore.
# Events 35 et 55 sont les VRAIS signaux de throttling actif.
# On agrege par jour pour limiter la volumetrie (un PC qui throttle
# peut generer des centaines d'events par heure).
$throttleAgg = @{}
foreach ($ev in $cpuEvents) {
    if ($ev.Id -eq 26) { continue }  # Event 26 = info capabilities, skip

    $dayKey = $ev.TimeCreated.ToString('yyyy-MM-dd')
    $sig    = "$($ev.Id)|$dayKey"

    if ($throttleAgg.ContainsKey($sig)) {
        $throttleAgg[$sig].Count++
        if ($ev.TimeCreated -lt $throttleAgg[$sig].FirstSeenRaw) {
            $throttleAgg[$sig].FirstSeen    = $ev.TimeCreated.ToString('yyyy-MM-dd HH:mm:ss')
            $throttleAgg[$sig].FirstSeenRaw = $ev.TimeCreated
        }
        if ($ev.TimeCreated -gt $throttleAgg[$sig].LastSeenRaw) {
            $throttleAgg[$sig].LastSeen    = $ev.TimeCreated.ToString('yyyy-MM-dd HH:mm:ss')
            $throttleAgg[$sig].LastSeenRaw = $ev.TimeCreated
        }
    } else {
        $throttleType = switch ($ev.Id) {
            35      { 'Firmware throttle limited' }
            55      { 'Thermal power reduction' }
            default { "Event $($ev.Id)" }
        }
        $throttleAgg[$sig] = @{
            Day          = $dayKey
            EventId      = $ev.Id
            Type         = $throttleType
            Count        = 1
            FirstSeen    = $ev.TimeCreated.ToString('yyyy-MM-dd HH:mm:ss')
            FirstSeenRaw = $ev.TimeCreated
            LastSeen     = $ev.TimeCreated.ToString('yyyy-MM-dd HH:mm:ss')
            LastSeenRaw  = $ev.TimeCreated
            Detail       = Get-TruncatedMessage -Text $ev.Message -MaxLength 160
        }
    }
}
# Output trie par date decroissante
$throttleAgg.Values |
    Sort-Object LastSeenRaw -Descending |
    ForEach-Object {
        $hardwareHealth.CPUThrottling.Add([PSCustomObject]@{
            Day       = $_.Day
            EventId   = $_.EventId
            Type      = $_.Type
            Count     = $_.Count
            FirstSeen = $_.FirstSeen
            LastSeen  = $_.LastSeen
            Detail    = $_.Detail
        })
    }

# Nombre total d'erreurs corrigees (somme des compteurs agreges)
$totalCorrected = 0
foreach ($c in $hardwareHealth.WHEA_Corrected) { $totalCorrected += $c.Count }

Write-Log ("Hardware : WHEA_Fatal={0} WHEA_Corrected_sigs={1} (total={2}) GPU={3} Thermal={4}" -f `
    $hardwareHealth.WHEA_Fatal.Count, $hardwareHealth.WHEA_Corrected.Count, $totalCorrected,
    $hardwareHealth.GPU_TDR.Count, $hardwareHealth.Thermal.Count)

# ============================================================
# 15. ASSEMBLAGE ET EXPORT JSON
# ============================================================
# IMPORTANT : on force tous les arrays avec @() pour contourner le
# bug de serialisation PS 5.1 (un array a 1 element sort en {} au
# lieu de [{}]). A retirer quand le Collector passera en PS 7.
$payload = [PSCustomObject]@{
    SchemaVersion    = $SchemaVersion
    Machine          = $machineInfo
    Events           = @($events)
    BootDurations    = @($bootDurations)
    BSODs            = @($bsods)
    ResourceWarnings = @($resourceWarnings)
    TopRAM           = @($topRAM)
    DiskInfo         = @($diskInfo)
    BatteryInfo      = $batteryInfo
    ServicesHealth   = $servicesHealth
    BootPerformance  = $bootPerformance
    DiskHealth       = @($diskHealth)
    Monitors         = @($monitors)
    # v1.8 : nouveaux blocs hardware
    MemoryInventory  = $memoryInventory
    GPUInventory     = @($gpuInventory)
    TopCrashers      = @($topCrashers)
    HardwareHealth   = [PSCustomObject]@{
        WHEA_Fatal     = @($hardwareHealth.WHEA_Fatal)
        WHEA_Corrected = @($hardwareHealth.WHEA_Corrected)
        GPU_TDR        = @($hardwareHealth.GPU_TDR)
        Thermal        = @($hardwareHealth.Thermal)
        # v1.8 : CPU throttling distinct du thermal
        CPUThrottling  = @($hardwareHealth.CPUThrottling)
    }
    Stats            = [PSCustomObject]@{
        # Compteurs events (events 6005 historiques, conserves pour compat)
        TotalBoots       = @($events | Where-Object { $_.EventId -eq 6005 }).Count
        TotalCrashFreeze = @($events | Where-Object { $_.EventId -eq 41 }).Count
        TotalBSOD        = @($bsods).Count

        # v1.6 : total des "vrais" plantages durs.
        # Exclut UserForcedReset (action volontaire de l'user) et
        # FreezeApp/FreezeUnknown (pas forcement un plantage materiel).
        TotalHardCrash   = @($events | Where-Object {
            $_.EventId -eq 41 -and $_.CrashCause -in @('BSODSilent','SleepResumeFailed','PowerLoss')
        }).Count

        # NOUVEAU v5.2 : comptage reel des demarrages par type
        BootsByType      = [PSCustomObject]@{
            ColdBoot    = $bootsByType.ColdBoot
            FastStartup = $bootsByType.FastStartup
            Resume      = $bootsByType.Resume
            Unknown     = $bootsByType.Unknown
        }
        # Total des "vrais" redemarrages = cold boots + fast startups
        # (les Resume d'hibernation ne sont pas des vrais redemarrages)
        TotalRealBoots   = $bootsByType.ColdBoot + $bootsByType.FastStartup

        BootsLongs       = @($bootDurations | Where-Object { $_.EstBootLong }).Count
        ResourceWarnings = @($resourceWarnings).Count
        DiskAlerts       = @($diskInfo | Where-Object { $_.IsAlert }).Count
        TopCrasherApp    = if (@($topCrashers).Count -gt 0) { @($topCrashers)[0].AppName } else { '' }
        TopCrasherCount  = if (@($topCrashers).Count -gt 0) { @($topCrashers)[0].CrashCount } else { 0 }

        # NOUVEAU v5.2 : distinction Fatal vs Corrected
        TotalWHEAFatal      = @($hardwareHealth.WHEA_Fatal).Count
        TotalWHEACorrected  = $totalCorrected           # somme des compteurs agreges
        WHEACorrectedUnique = @($hardwareHealth.WHEA_Corrected).Count   # nb de signatures distinctes

        TotalGPU         = @($hardwareHealth.GPU_TDR).Count
        TotalThermal     = @($hardwareHealth.Thermal).Count
        # v1.8 : total des jours avec throttling CPU (somme de Count
        # serait trompeuse : un meme probleme en 1 jour fait plein d'events)
        TotalCPUThrottling = @($hardwareHealth.CPUThrottling).Count
        # TotalHardware : seules les vraies alertes (Fatal + GPU + Thermal)
        # Les corrected sont de la telemetrie, pas une alerte
        TotalHardware    = @($hardwareHealth.WHEA_Fatal).Count + @($hardwareHealth.GPU_TDR).Count + @($hardwareHealth.Thermal).Count

        DerniereActivite = ($events | Sort-Object Timestamp | Select-Object -Last 1).Timestamp
        Event12Trouve    = $bootStartEvents.Count
        Event27Trouve    = $kernelBootEvents.Count

        # NOUVEAU v5.3 : raccourcis batterie et EDR pour le dashboard
        BatteryHealthPct = $batteryInfo.HealthPercent
        BatteryAlert     = $batteryInfo.IsAlert
        SentinelRunning  = ($servicesHealth.SentinelAgent.Status -eq 'Running')
        SentinelAlert    = $servicesHealth.SentinelAgent.IsAlert

        # NOUVEAU v5.4 : raccourcis boot perf et disk health
        LastBootTimeMs     = if ($lastBoot) { $lastBoot.BootTimeMs }         else { 0 }
        LastPostBootTimeMs = if ($lastBoot) { $lastBoot.BootPostBootTimeMs } else { 0 }
        BootPerfAlert      = $bootPerformance.IsAlert
        DiskWorstWear      = $diskWorstWear
        DiskHealthAlert    = $diskHealthAlert

        # NOUVEAU v5.5 : inventaire moniteurs externes
        MonitorsCount      = $monitors.Count
    }
}

$outputFile  = Join-Path $SharePath "$($env:COMPUTERNAME).json"
$localBuffer = Join-Path $env:ProgramData "PCPulse\$($env:COMPUTERNAME).json"
$null = New-Item $SharePath -ItemType Directory -Force -ErrorAction SilentlyContinue
$null = New-Item (Split-Path $localBuffer) -ItemType Directory -Force -ErrorAction SilentlyContinue

# Serialisation en memoire (une seule fois)
$jsonContent  = $payload | ConvertTo-Json -Depth 5
$utf8NoBom    = [System.Text.UTF8Encoding]::new($false)
$bufferWritten = $false
$shareWritten  = $false

# --- Etape 1 : ecrire le buffer local (priorite pour la resilience) ---
# Si le fichier existe avec des permissions incoherentes (buffer d'une
# ancienne version avec un autre compte), on tente un nettoyage forcé.
try {
    if (Test-Path $localBuffer) {
        try {
            # Desactive l'attribut ReadOnly au cas ou, puis supprime
            Set-ItemProperty -Path $localBuffer -Name IsReadOnly -Value $false -ErrorAction SilentlyContinue
            Remove-Item -Path $localBuffer -Force -ErrorAction SilentlyContinue
        } catch {}
    }
    [System.IO.File]::WriteAllText($localBuffer, $jsonContent, $utf8NoBom)
    $bufferWritten = $true
    Write-Log "Buffer local OK : $localBuffer"
} catch {
    Write-Log "Erreur ecriture buffer local : $_"
}

# --- Etape 2 : copier le buffer sur le share ---
# On NE copie que si le buffer vient d'etre ecrit avec succes.
# Sinon on tentera une ecriture directe sur le share en etape 3.
if ($bufferWritten) {
    try {
        Copy-Item -Path $localBuffer -Destination $outputFile -Force -ErrorAction Stop
        $shareWritten = $true
    } catch {
        Write-Log "Erreur copie buffer vers share : $_"
    }
}

# --- Etape 3 : fallback - ecriture directe sur le share ---
# Utilisee si l'etape 1 OU l'etape 2 a echoue.
# On sacrifie la resilience buffer mais on evite la corruption silencieuse.
if (-not $shareWritten) {
    try {
        [System.IO.File]::WriteAllText($outputFile, $jsonContent, $utf8NoBom)
        $shareWritten = $true
        Write-Log "Ecriture directe sur share OK (sans buffer local)"
    } catch {
        Write-Log "Erreur ecriture directe share : $_"
    }
}

# --- Bilan final ---
if ($shareWritten) {
    Write-Log ("Export OK - RealBoots:{0} (CB:{1}/FS:{2}/R:{3}) Crash:{4} BSOD:{5} Uptime:{6}j WHEA_Fatal:{7} WHEA_Corr:{8} Batt:{9}% Sentinel:{10} BootPerf:{11}ms Disk:{12} Monitors:{13}" -f `
        $payload.Stats.TotalRealBoots,
        $payload.Stats.BootsByType.ColdBoot,
        $payload.Stats.BootsByType.FastStartup,
        $payload.Stats.BootsByType.Resume,
        $payload.Stats.TotalCrashFreeze,
        $payload.Stats.TotalBSOD,
        $uptimeDays,
        $payload.Stats.TotalWHEAFatal,
        $payload.Stats.TotalWHEACorrected,
        $(if ($batteryInfo.HasBattery) { $batteryInfo.HealthPercent } else { 'N/A' }),
        $(if ($servicesHealth.SentinelAgent.Installed) { $servicesHealth.SentinelAgent.Status } else { 'NotInstalled' }),
        $(if ($lastBoot) { $lastBoot.BootTimeMs } else { 'N/A' }),
        $(if ($diskHealthAlert) { "ALERT-WorstWear$($diskWorstWear)%" } else { "OK-WorstWear$($diskWorstWear)%" }),
        $monitors.Count)
} else {
    Write-Log "=== EXPORT ECHEC : ni buffer local ni share n'ont pu etre ecrits ==="
}
