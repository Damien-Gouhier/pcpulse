# Changelog

Toutes les modifications notables de PCPulse sont documentees ici.

Format base sur [Keep a Changelog](https://keepachangelog.com/fr/1.1.0/),
versionnement respectant [Semantic Versioning](https://semver.org/lang/fr/).

---

## [1.4.0] — 2026-04-23

### 🐛 Fix parsing Intel Core Gen 10+ avec suffixe lettre

Les CPU Intel Core de 10e à 14e génération avec suffixe lettre (ex: `Core i5-1345U`, `Core i7-1165G7`, `Core i5-1335U`) étaient incorrectement classés en **Gen 1 / 2010 / Ancien**.

**Cause** : la regex utilisée pour extraire la génération (`(\d{1,2})\d{3}`) était greedy et capturait toujours 1 seul chiffre sur les CPU à 4 chiffres. Seuls les CPU à 5 chiffres (i7-12700, i7-14700K) étaient correctement parsés.

**Fix** : nouvelle heuristique basée sur la longueur du numéro de modèle :
- 5 chiffres : gen = 2 premiers chiffres (i7-12700 → Gen 12)
- 4 chiffres commençant par `1[0-4]` : gen = 2 premiers chiffres (i5-1345U → Gen 13)
- 4 chiffres autrement : gen = 1er chiffre (i5-4300U → Gen 4)

**Impact** : tous les laptops du parc équipés Intel Gen 10-14 avec suffixe U/H/P/G (ex: i5-1135G7, i7-1165G7, i5-1235U, i5-1335U, i5-1345U) vont être correctement reclassés de "Ancien" en "Récent" au prochain cycle de collecte.

---

## [1.3.0] — 2026-04-23

### 🐛 Fix affichage CurrentUser

Le Collector tourne en SYSTEM via une tâche planifiée. `$env:USERNAME` retournait donc le compte machine AD (ex: `LAPTOP001$`), non lisible dans le Dashboard.

**Fix** : nouvelle fonction `Get-CurrentInteractiveUser` qui utilise en cascade :
1. `Win32_ComputerSystem.UserName` (méthode principale, fiable et rapide)
2. Owner du process `explorer.exe` (fallback si aucune session renvoyée)
3. `(aucune session)` si personne n'est connecté

Le nom de domaine est automatiquement strippé (`MONDOMAINE\jean.dupont` → `jean.dupont`).

### 🐛 Fix BootDurations incomplet

Les Fast Startup et Resume étaient détectés (compteur `Stats.BootsByType` correct) mais **pas inscrits dans `BootDurations`**, ce qui faisait que le Dashboard affichait "Aucun démarrage détecté sur la période" même sur des parcs actifs.

**Cause** : le code exigeait une durée de boot calculable pour ajouter une entrée à `BootDurations`, or les Fast Startup et Resume n'ont pas d'Event 12 associé (pas de vrai kernel boot).

**Fix** : tous les Event 27 sont maintenant inscrits dans `BootDurations` :
- Avec Event 12 (cold boot) : `Method='Event12+27'`, durée calculée
- Sans Event 12 (Fast Startup / Resume) : `Method='Event27-only'`, `DurationMin=0`

**Impact** : le panneau "Répartition des démarrages" du Dashboard reflète désormais la réalité du parc (ex: un PC avec 3 cold boots + 14 Fast Startup sur 30j affiche bien 17 entrées au lieu de 3).

---

## [1.2.0] — 2026-04-23

### ✨ Nouveautes majeures

**Auto-updater** — le parc se met a jour tout seul :

- Nouveau script `PCPulse-Updater.ps1`, deploye sur chaque PC a cote du
  Collector. Il devient la cible de la tache planifiee (au lieu du Collector
  direct) et verifie a chaque cycle horaire si une nouvelle version est
  disponible dans `\\SERVEUR\share\release\`.
- Workflow de release : l'admin copie le nouveau `01_Collector.ps1` et
  met a jour `version.txt` sur le serveur. Les PC se mettent a jour
  automatiquement dans l'heure qui suit, sans intervention locale.
- Verification SHA256 post-copie (integrite)
- Backup automatique de la version precedente (rolling 5 dernieres)
- Lock file anti-collision (gestion des instances concurrentes)
- Log dedie `updater.log` pour le diagnostic
- Aucun rollback automatique : en cas d'echec, le Collector local continue
  de tourner (pas de boucle infernale possible)

**Scripts de deploiement inclus** :

- `Setup-Server.ps1` : initialisation complete d'un serveur Windows
  (role File Server, partage SMB3 encrypte, NTFS, resolution des groupes
  AD via SIDs well-known pour compatibilite FR/EN)
- `Install-Client.ps1` v2.0 : installe le Collector + l'Updater +
  version.txt + tache planifiee pointant sur l'Updater

### 🐛 Correctif important : Uptime utilisateur realiste

Le champ `UptimeDays` du JSON `Machine` reflete maintenant le **temps depuis
la derniere reprise d'activite utilisateur**, et non plus le temps depuis
le dernier cold boot du kernel.

**Pourquoi** : sur les laptops modernes avec Fast Startup active (cas par
defaut des Dell Latitude recents, par exemple), le bouton "Arreter" de
Windows fait en realite un hibernate partiel + Modern Standby. L'ancien
calcul base sur `Win32_OperatingSystem.LastBootUpTime` affichait des
uptimes de 15+ jours meme pour des utilisateurs qui eteignent chaque soir,
faussant toute interpretation.

**Algorithme** : `MAX(BootDurations.last, LastEvent507)` — combine la
detection des cold boots / Fast Startup (Event 12+27) avec les wakes
Modern Standby (Event 507, provider Microsoft-Windows-Kernel-Power).

### ✨ Nouveaux champs JSON (schema 1.2)

- `Machine.LastRealColdBoot` : date du dernier vrai cold boot kernel
  (utile pour savoir quand les patches Windows Update noyau ont ete
  appliques)
- `Machine.FastStartupEnabled` : boolean, etat du registry
  `HiberbootEnabled` (info d'inventaire)

### 🎨 Dashboard

- Formatage humain de la colonne Uptime : `45min`, `3h`, `1j 8h`, `12j`
  au lieu de `0.13j`, `1.85j`, etc.
- Message d'info "Fast Startup domine sur le parc" mis a jour pour
  refleter la detection correcte

### 📋 Compatibilite

- Schema JSON : 1.0 → 1.2 (les anciens JSON restent lisibles par le
  nouveau Dashboard — les champs absents s'affichent en "—")

---

## [1.0.0] — 2026-04-21

Premiere version publique de PCPulse.

### ✨ Fonctionnalites initiales

**Collector (`01_Collector.ps1`)** — a deployer sur chaque endpoint :

- Collecte des evenements systeme sur les N derniers jours (30 par defaut)
- Detection des crashs, freezes, BSOD, shutdowns dirty
- Mesure des durees de boot + type de boot (Cold / Fast Startup / Resume)
- Analyse **Boot Performance** detaillee (MainPath, PostBoot, UserProfile,
  Explorer init) via Event 100 Diagnostics-Performance
- Sante disque **SMART** via Get-StorageReliabilityCounter
  (temperature, wear, power-on hours, read/write errors)
- Sante **batterie** avec pourcentage d'usure (2 methodes en cascade
  WMI + powercfg pour couvrir les CPU AMD Ryzen AI et Intel Gen 13+)
- Monitoring **EDR SentinelOne** (installe / running / startup type)
- Inventaire **moniteurs externes** via EDID (fabricant, modele, serial,
  annee de fabrication, age)
- Detection **type de chassis** (Laptop / Desktop / All-In-One) via
  `Win32_SystemEnclosure.ChassisTypes`
- Classification **WHEA** Fatal vs Corrected avec parsing XML
- Detection **GPU TDR** (Timeout Detection Recovery)
- Detection **throttling CPU** via Event 26 Kernel-Processor-Power
- Export JSON atomique avec buffer local en cas d'echec SMB
- Delai anti-collision aleatoire (0-60 min) par defaut
- Mode `-NoDelay` pour tests manuels
- Compatibilite **PowerShell 5.1** (parc natif Windows 10/11)

**Dashboard (`02_Dashboard.ps1`)** — a executer sur le poste admin :

- Lecture de tous les JSON d'un dossier partage, agregation en un
  seul rapport HTML autonome (aucune dependance externe)
- Regroupement des KPIs en **4 familles** : Securite, Stabilite,
  Performance, Usure materielle
- Tableau detaille par appareil avec tri, filtres (periode, CPU,
  site, recherche), vue compacte ou detaillee
- Mapping IP/hostname → Site via CSV externe configurable
- **Drill-down** par PC en 5 onglets : Vue d'ensemble, Stabilite,
  Demarrage, Materiel, Securite
- Panneaux agreges parc : repartition des demarrages, top crashers
  transverse, inventaire ecrans secondaires
- Export CSV du rapport
- **Dark mode + light mode** avec persistence localStorage
- Scoring automatique de la sante de chaque PC (ponderation
  configurable via `config.psd1`)
- Retrocompatibilite totale avec les futurs schemas JSON

### 🔧 Architecture

- **Zero dependance externe** : que du PowerShell natif et HTML inline
- Pas de base de donnees, pas de service, pas d'agent
- Deployement adapte aux contextes **Intune**, **SmartDeploy**, GPO
- Donnees stockees dans un simple dossier partage SMB

### 📋 Notes de compatibilite

- **OS** : Windows 10 / 11 (collector) ; Windows 10/11 + PS 7 (dashboard)
- **PowerShell** : 5.1+ (collector), 7.0+ (dashboard)
- **Schema JSON** : 1.0 (stable, retrocompatible avec futurs schemas)
