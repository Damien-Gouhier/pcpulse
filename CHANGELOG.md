# Changelog

Toutes les modifications notables de PCPulse sont documentees ici.

Format base sur [Keep a Changelog](https://keepachangelog.com/fr/1.1.0/),
versionnement respectant [Semantic Versioning](https://semver.org/lang/fr/).

---

## [2.0.0] — 2026-05-04

### 🛡️ Hardening sécurité majeur (release pre-pentest)

Cette release transforme PCPulse d'un outil "fonctionnel" en un outil **pentest-ready**, avec un trust model documenté, des ACLs hardenées par défaut, un mécanisme killswitch sécurisé, et des sanity-checks stricts côté Dashboard. C'est une **release majeure** parce qu'elle introduit un changement de comportement structurant côté serveur (`Setup-Server.ps1` réécrit) et un bump du schema JSON. Le Collector et le Dashboard restent compatibles entre eux mais ne sont plus rétrocompatibles avec les schemas antérieurs.

#### Côté serveur : `Setup-Server.ps1` v2.0

Réécriture complète du script d'initialisation du share serveur, focalisée sur la fermeture d'une **faille majeure** identifiée lors de l'audit pré-pentest :

> *"Si un attaquant compromet n'importe quel PC du parc, il peut écrire dans `\release\` (Domain Computers en Modify héritage) et remplacer le Collector par sa version malveillante. Au prochain cycle, les 800 PC du parc téléchargent et exécutent son code en SYSTEM."*

→ **Escalation latérale massive** : 1 endpoint compromis = RCE SYSTEM sur tout le parc.

Le nouveau `Setup-Server.ps1` :

* **Casse l'héritage NTFS** sur la racine du share
* **Crée les sous-dossiers** `\release\`, `\killed\`, `\logs\` avec **ACLs explicites** par dossier (au lieu d'hériter)
* **Rétrograde Domain Computers en `ReadAndExecute`** sur `\release\` (lecture seule pour les PC, écriture admin uniquement)
* **Retire `BUILTIN\Users`** des permissions trop larges (un user lambda du domaine ne peut plus créer de fichiers sur le share)
* **Réassigne les owners** aux Domain Admins (au lieu de comptes personnels) pour audit propre
* **Support gMSA** : nouveau paramètre `-gMSAName` pour ajouter un compte de service dédié aux ACLs (recommandé en remplacement de SYSTEM)
* **Préservation de groupes existants** : nouveau paramètre `-ExtraReadGroups` pour conserver les conventions infra de votre organisation (groupes admins, supervision, etc.)
* **Idempotent** : peut être relancé sans risque pour re-aligner les ACLs après une modification manuelle
* **Mode dry-run** via `-WhatIf` pour valider la cible avant exécution
* **Backup automatique** des ACLs originales dans `\\SERVER\PCPulse$\.acl-backup\` avant toute modification (rollback possible)
* **3 tests post-application** automatiques (Domain Computers peut lire / ne peut pas modifier / gMSA peut lire `\release\`)

#### Côté Updater : `PCPulse-Updater.ps1` v1.1 — Killswitch

Ajout d'un mécanisme **killswitch** permettant de désinstaller à distance les Collectors de tous les PC du parc, sans toucher physiquement aux machines.

**Cas d'usage** :
* Décommissionner PCPulse proprement (remplacement par un autre outil)
* Stop d'urgence si un bug critique est détecté
* Migration majeure (kill v1 → install v2 propre)

**Workflow** :

1. L'admin dépose un fichier sentinelle dans `\\SERVER\PCPulse$\release\` (par défaut `KILLSWITCH.txt` contenant `CONFIRM-UNINSTALL-PCPULSE`)
2. Au prochain cycle horaire, chaque PC :
   * Détecte le fichier
   * Vérifie le contenu exact (sinon ignore avec un log WARN)
   * Écrit un rapport dans `\killed\<HOSTNAME>.txt` (timestamp, version, hostname)
   * Supprime sa tâche planifiée `PCPulse-Collector`
   * Lance un cleanup différé qui supprime `C:\ProgramData\PCPulse\`
   * Exit
3. L'admin retire le fichier sentinelle quand tous les PC sont apparus dans `\killed\` (les PC qui se rallument après coup se kill aussi)

**Personnalisable via `config.psd1`** :

```powershell
KillSwitch = @{
    Enabled  = $true
    Filename = 'NOM_PERSONNALISE.txt'
    Phrase   = 'PHRASE-COMPLEXE-IMPOSSIBLE-A-DEVINER'
}
```

Les valeurs par défaut sont publiques (puisque dans le repo open source). En production, **changer la phrase** pour éviter tout déclenchement par accident. La vraie protection reste les ACLs sur `\release\` (cf. `SECURITY.md`).

#### Côté Collector : `01_Collector.ps1` v2.0 — Bump SchemaVersion

* **SchemaVersion bumpé `1.8` → `2.0`** : tous les JSON produits par le Collector v2.0 portent désormais ce numéro de version, qui est obligatoire pour être accepté par le Dashboard v2.0.
* Aucun changement fonctionnel côté collecte : les blocs `MemoryInventory`, `GPUInventory`, `HardwareHealth.CPUThrottling`, `Stats.TotalHardCrash`, `Stats.TotalCPUThrottling` (introduits en v1.6/v1.8) restent identiques.

#### Côté Installeur client : `Install-Client.ps1` v2.0 — Support gMSA optionnel

* **Nouveau paramètre `-gMSAName`** : permet de créer la tâche planifiée avec un compte gMSA dédié au lieu de SYSTEM (recommandé pour audit propre, voir `SECURITY.md`).
* **Validation préalable** via `Test-ADServiceAccount` avant la création de la tâche : si le PC n'est pas autorisé à utiliser le gMSA, le script échoue immédiatement avec un message clair (pas de fallback silencieux sur SYSTEM).
* **`-InitialVersion`** par défaut bumpé `1.0` → `2.0` pour aligner sur la release.
* **Mention du killswitch** dans le résumé final : pour la désinstallation propre, l'admin est désormais redirigé vers la procédure killswitch (voir `SECURITY.md`).

#### Côté Dashboard : `02_Dashboard.ps1` v2.0 — Sanity-checks stricts + sanitization

Réécriture du bloc de lecture des JSON pour fermer une **deuxième faille** identifiée lors de l'audit pré-pentest :

> *"Un PC compromis dépose un JSON contenant `<script>alert('xss')</script>` dans le champ `Machine.PC` ou `CurrentUser`. Quand l'admin ouvre le rapport HTML, le navigateur exécute le script."*

**Validation à 4 niveaux** (rejet si KO) via la nouvelle fonction `Test-PCPulseJson` :

1. **Parsing JSON** — fichier malformé / illisible = rejet
2. **SchemaVersion** — whitelist stricte `@('2.0')` (plus de tolérance retro-compat sur les anciens schemas). Tout JSON sans `SchemaVersion = '2.0'` est rejeté.
3. **Machine.PC** — présent, non vide, alphanumérique + tirets/underscores/points (regex `'^[A-Za-z0-9_.-]{1,63}$'`)
4. **Match nom de fichier ↔ Machine.PC** : un PC compromis ne peut pas se faire passer pour un autre PC du parc (anti-spoofing). Si `LAPTOP-042.json` contient `Machine.PC = "CEO-WORKSTATION"`, c'est rejeté.

**Sanitization en place** des champs string sensibles avant injection dans le HTML généré (via la nouvelle fonction `Get-SafeString`) :

* `Machine.CurrentUser`, `Machine.IPAddress`, `Machine.CPUName`, `Machine.OSCaption`, `Machine.Manufacturer`, `Machine.Model`, `Machine.Site`
* HTML-encoding des 5 caractères dangereux : `& < > " '` (dans le bon ordre, `&` en premier)
* Si un PC compromis met `<script>` dans son `CurrentUser`, ça devient `&lt;script&gt;` et s'affiche comme du texte inoffensif

**Nouvelle section HTML "JSON suspects"** affichée en tête du rapport si des rejets sont détectés :

* Panneau visuel rouge/orange pliable
* Liste chaque rejet avec : nom de fichier, raison, détail technique
* Indicateur fort pour le pentester : *"vos tentatives sont détectées et journalisées"*
* Renvoi vers `SECURITY.md` pour le rationale complet

**Content-Security-Policy stricte** (`<meta http-equiv="Content-Security-Policy">`) injectée dans le `<head>` du HTML généré, en dernière ligne de défense au cas où la sanitization laisserait passer quelque chose :

```
default-src 'none';
script-src 'unsafe-inline';
style-src 'unsafe-inline';
img-src data: 'self';
font-src 'self';
connect-src 'none';
frame-src 'none';
object-src 'none';
base-uri 'none';
form-action 'none';
```

* `default-src 'none'` → tout est interdit par défaut, on autorise au cas par cas
* `connect-src 'none'` → impossible de faire du fetch/XHR/WebSocket = **pas d'exfiltration de données** possible même si du code malveillant s'exécutait
* `script-src 'unsafe-inline'` + `style-src 'unsafe-inline'` → JS/CSS inline OK (par design : zéro CDN), mais **pas de script externe**
* `img-src data: 'self'` → images en data URI (favicon SVG base64) + relatives OK, pas de tracker tiers
* `frame-src`, `object-src`, `base-uri`, `form-action` à `'none'` → ferme toutes les voies non utilisées (pas d'iframe, pas de Flash/embed, pas de redéfinition de base URL, pas de soumission de formulaire)

→ Combinaison **input filtering** (sanity-checks) + **output policy** (CSP) = défense en profondeur.

#### Documentation : `SECURITY.md` (nouveau)

Document de sécurité complet, à lire avant tout déploiement :

* **Trust model** : qui peut faire quoi sur le share, scénario d'attaque principal, défense
* **ACLs recommandées** : tableau des permissions par dossier, justifications, pourquoi `ReadAndExecute` et pas `Modify`, pourquoi pas `Authenticated Users`, etc.
* **Modèle de menace du killswitch** : ce que ça protège, ce que ça ne protège PAS, pourquoi la phrase secrète est de la défense en profondeur (pas la vraie protection)
* **Surface d'attaque résiduelle** : risques connus avec criticité (🟡/🟢) et mitigations actuelles
* **Roadmap de hardening** : Fait / En cours / Planifié / À l'étude
* **Procédure de signalement** des vulnérabilités via GitHub Security Advisories
* **Références** Microsoft (tier model, gMSA, well-known SIDs)

#### Documentation : `README.md` enrichi

* Nouvelle section **"Pourquoi je développe ça"** expliquant les 3 motivations du projet (curiosité, open source, besoin métier)
* Nouvelle section **"Killswitch"** synthétique (avec renvoi vers SECURITY.md pour les détails)
* Nouvelle section **"Sécurité"** qui pointe vers SECURITY.md
* Architecture (ASCII art) mise à jour pour montrer `\release\`, `\killed\`, `KILLSWITCH.txt`

#### `config.psd1.example` enrichi

Nouvelle section `KillSwitch` documentée, avec note de sécurité expliquant le modèle de menace (le pentest se fait sur l'écriture du fichier sentinelle, pas sur la lecture de la phrase).

#### `version.txt` à la racine du repo (nouveau)

Fichier de référence indiquant la version actuelle du Collector dans ce snapshot du repo. Sert :
- de modèle pour l'utilisateur qui copie sur son serveur (`\\SERVER\PCPulse$\release\version.txt`)
- de source par défaut pour `Install-Client.ps1` lors de l'installation initiale
- de marqueur de version pour les futurs outils de check (CI/CD, dependency tracking, etc.)

Contenu : `2.0`.

#### `README.en.md` mis à jour

La version anglaise du README a été synchronisée avec les changements du `README.md` français (sections "Why I'm building this", "Killswitch", "Security").

#### JSON de démo (`examples/demo/`) mis à jour

Les 5 JSON de démo (`LAPTOP-001` à `OFFLINE-005`) sont enrichis pour refléter la structure complète v2.0 :

* `SchemaVersion: "2.0"` (au lieu de `"1.0"`)
* Ajout des blocs introduits depuis : `MemoryInventory`, `GPUInventory`, `HardwareHealth.CPUThrottling`, `Stats.TotalHardCrash`, `Stats.TotalCPUThrottling`
* Champs `Machine.LastRealColdBoot` et `Machine.FastStartupEnabled` ajoutés
* `CrashCause` ajouté sur les Event 41 du LAPTOP-002 (BSOD scenario)

Le scénario `LAPTOP-002` (cas problématique) gagne en cohérence : 10 jours de CPU throttling viennent expliquer les boots lents, la batterie HS, et corréler avec les BSOD MEMORY_MANAGEMENT (probable problème thermique chronique). Démo plus pédagogique pour le visiteur du repo qui découvre PCPulse.

### 📋 Migration depuis v1.x

Si tu as déjà déployé PCPulse v1.x en production :

1. **Côté serveur** : exécuter le nouveau `Setup-Server.ps1` (en mode `-WhatIf` d'abord pour valider). Backup automatique des ACLs avant modification.
2. **Côté clients** : remplacer `PCPulse-Updater.ps1` par la v1.1. Comme l'auto-update ne met à jour que le Collector et pas l'Updater lui-même, il faut un push manuel (Intune, SmartDeploy, GPO, ou push SMB direct).
3. **Côté Collector** : déposer le nouveau `01_Collector.ps1` v2.0 dans `\release\` + bumper `version.txt` à `2.0`. Les PC se mettront à jour automatiquement au prochain cycle horaire.
4. **Côté Dashboard** : remplacer `02_Dashboard.ps1` par la v2.0 sur ton poste admin.
5. **Optionnel** : créer un gMSA dédié dans AD et l'ajouter aux ACLs via `Setup-Server.ps1 -gMSAName 'svc-pcpulse$'`. Migrer la tâche planifiée des PC vers ce gMSA (recommandé pour audit propre vs SYSTEM).

⚠️ **Pendant la migration** : le Dashboard v2.0 va rejeter tous les JSON déposés par les Collector v1.x (SchemaVersion non whitelisté). C'est attendu et visible dans la nouvelle section "JSON suspects". Une fois tous les PC à jour, ces rejets disparaissent. Si tu veux éviter cette transition visible, déploie le Collector d'abord, puis le Dashboard.

### 🛣️ Roadmap (cf. SECURITY.md)

Hardening prévu pour les versions suivantes :

* **CSP inline HTML** : Content-Security-Policy meta tag dans le HTML généré (anti-XSS, défense en profondeur des sanity-checks)
* **Code signing** : signature numérique du Collector et Updater via certificat AD CS
* **Vérification de signature dans l'Updater** : refus du téléchargement si signature invalide
* **Détection d'anomalies** : JSON futur, PC qui spam, hostnames inhabituels
* **Audit logs** : log immuable des actions admin

---

## [1.8.0] — 2026-04-24

### 🆕 Panel Matériel enrichi : RAM, GPU, CPU Throttling

Trois nouveaux blocs d'informations matérielles viennent compléter le drill-down par PC pour aider à la prise de décision (upgrade vs remplacement, diagnostic de refroidissement défaillant, corrélation driver/crash).

#### Côté Collector

**Inventaire RAM** (nouveau bloc `MemoryInventory`) — via `Win32_PhysicalMemory` et `Win32_PhysicalMemoryArray` :
- Capacité totale installée, nombre de slots total/occupés, capacité max supportée par la carte mère
- Détail barrette par barrette : emplacement, taille, type (DDR3/DDR4/DDR5/LPDDRx), vitesse, fabricant, PartNumber
- Flag `CanUpgrade` calculé (slots libres OU moins de 75% de la capa max installée)

**Inventaire GPU** (nouveau bloc `GPUInventory`) — via `Win32_VideoController` :
- Nom, version driver, date driver pour chaque adaptateur physique
- Filtrage des "pseudo-GPU" (Remote RDP, Hyper-V, Mirror, Meta, Basic Display, VNC)
- VRAM volontairement exclue : `AdapterRAM` est un `int32` qui overflow au-dessus de 4 Go, donnée peu fiable via WMI

**CPU Throttling détaillé** (nouveau bloc `HardwareHealth.CPUThrottling`) :
- Capture des Events 35 (firmware throttle limited) et 55 (thermal power reduction) du provider `Kernel-Processor-Power`
- Agrégation **par jour** pour limiter la volumétrie (un PC qui throttle peut générer des centaines d'events/h)
- Format par jour : `{Day, EventId, Type, Count, FirstSeen, LastSeen, Detail}`
- Distinct du bloc `Thermal` existant (Events 37/88/125 Kernel-Power) qui capture les alertes critiques de surchauffe

Nouvelle stat `Stats.TotalCPUThrottling` = nombre de jours distincts avec throttling.

#### Côté Dashboard

Panel Matériel restructuré en 8 sections :
1. CPU
2. 🆕 **Mémoire RAM** — gros chiffre de synthèse + détail des barrettes, badge "Upgrade possible" vert ou "Max atteint" rouge
3. 🆕 **GPU & Drivers** — un bloc par adaptateur, tag jaune "driver >2 ans" pour inciter à la MàJ
4. 🆕 **CPU Throttling** — bandeau de sévérité à 4 niveaux :
   - `0 jour` → message neutre
   - `1-7 jours` → info cyan (usage intensif normal sur laptop moderne)
   - `8-20 jours` → warning jaune "à surveiller"
   - `>20 jours` **ou** `Count > 500 events en un seul jour` → critique rouge "refroidissement à vérifier en priorité"
5. Disque (remplissage)
6. SMART
7. Batterie
8. Écrans secondaires branchés

Le seuil à 4 niveaux a été calibré après tests terrain sur laptops modernes (Ryzen AI Pro, Core Ultra) qui throttlent naturellement lors de charges soutenues sans que ce soit pathologique. Le seuil initial "≥3 jours = alerte" générait trop de faux positifs.

**Backward-compat totale** : sur un JSON de Collector antérieur (< v1.8), les trois nouvelles sections affichent discrètement "Données non disponibles (Collector < v1.8)" sans rien casser.

---

## [1.7.0] — 2026-04-24

### 🆕 Débruitage des Top Crashers : Signal vs Bruit ambient

Le panel Top Crashers était pollué par des processus qui crashent naturellement en arrière-plan sur toute une flotte (ex: `microsoftsearchinbing.exe`, `shellexperiencehost.exe`, utilitaires constructeurs) et masquaient les vrais signaux locaux exploitables. v1.7 introduit un système de scoring qui promeut les crashs concentrés (1-2 PC) et déprime les crashs dispersés (5+ PC).

#### Scoring automatique

Chaque processus reçoit un score calculé par :

```
score = crashs_par_PC / sqrt(PC_impactés)
```

La racine carrée pénalise la dispersion entre PC : un processus à **8 crashs sur 1 PC** (score 8.0) ressort plus haut qu'un processus à **81 crashs sur 10 PC** (score 2.6), ce dernier étant typiquement du bruit de fond.

#### Trois sections dans le panel Top Crashers global

- 🎯 **Signaux locaux** (score ≥ 3) — concentrés sur peu de PC, à investiguer en priorité
- ⚠️ **Problèmes répartis** (2 ≤ score < 3) — possibles bugs applicatifs touchant plusieurs machines
- 🔇 **Bruit ambient** (score < 2) — replié par défaut, à consulter seulement si curiosité

#### Deux blacklists complémentaires

- **Blacklist HARD** : processus totalement écartés de l'agrégation, invisibles dans les 3 sections. Réservée aux crashers "inactionnables par construction".
  - `microsoftsearchinbing.exe` — bug Microsoft récurrent avec typo interne ("search**IN**bing" et non "searchBing"), pollue déjà depuis l'époque Nexthink
- **Blacklist SOFT** : processus toujours affichés mais forcés en section "Bruit ambient" même si leur score les aurait remontés plus haut.
  - `dellosd.exe`, `shellexperiencehost.exe`, `gamebar.exe`, `asussystemanalysis.exe`, `asusverifyjwt.exe`, `dell.techhub.diagnostics.subagent`

Dans le drill-down par PC, les bruiteurs SOFT apparaissent **grisés** avec un badge `BRUIT`, les HARD sont totalement absents.

### 🐛 Fix parser CPU étendu (terrain)

- **Intel Gen 10 = 2020** (était 2019) — alignement sur la date de sortie effective du desktop Comet Lake (mai 2020), qui représente la grosse majorité du parc Gen 10. Les rares mobiles Ice Lake de fin 2019 auront 1 an de décalage, acceptable.
- **Nouveau support Intel Celeron N-series** (crucial pour les tablettes d'accueil) :
  - N4000/N4100 → Gemini Lake (2017)
  - N4500/N4505/N5100/N6000 → Jasper Lake (2021)
  - N95/N97/N100/N200/N305 → Alder Lake-N (2023)
  - N350/N355 → Twin Lake (2025)
  - N3000/N3050/N3150 → Bay Trail (2014)
- **Nouveau support Intel Pentium Gold** (desktops budget G6xxx/G7xxx)
- **Nouveau support AMD Athlon et A-Series** (parc AMD legacy)

Avant v1.7, ces CPU ressortaient `Category='Inconnu'` dans le Dashboard, ce qui gênait le calcul du verdict global introduit en v1.6.

---

## [1.6.0] — 2026-04-24

### 🆕 Intelligence de diagnostic : Wave 1

Transforme le Dashboard de "liste de faits" en véritable outil de diagnostic en corrélant automatiquement les données brutes pour en sortir du sens.

#### Bandeau verdict global par PC

Chaque drill-down affiche maintenant un bandeau de synthèse en tête, à 4 niveaux :

| Niveau | Couleur | Déclencheurs (règles OR) |
|---|---|---|
| 🟢 **Sain** | vert | aucun signal |
| 🟡 **À surveiller** | jaune | CPU vieillissant / batterie 50-70% / SSD wear 50-80% / burst I/O isolé |
| 🟠 **Incident probable** | orange | burst I/O massif (≥50 events) / WHEA corrigées fréquentes / 1+ BSOD / 2-4 hard crashs/30j |
| 🔴 **Critique** | rouge | WHEA Fatal / ≥5 hard crashs/30j / batterie <50% / SSD wear >80% / CPU ancien+crashs |

Format : `🟠 Incident probable · burst I/O massif (>=50 events) · 3 hard crashs/30j`. Les raisons sont listées en metadata pour que le diagnostic soit explicite.

#### Section "Signaux croisés" (corrélations temporelles)

Détection automatique de 5 patterns suspects à fenêtre **10 minutes** :

1. **Burst I/O ≥50 events → Hard crash** dans les 10 min = probable panne disque
2. **WHEA Corrected PCIe → crash système** dans les 10 min = slot PCIe suspect
3. **Thermal event + pic de boot > 2× la moyenne** = refroidissement dégradé
4. **≥2 BSOD avec même BugCheck sur 7 jours** = crash récurrent à cause identifiée
5. **≥2 hard crashs en 24h** = instabilité marquée (alim / thermique / drivers)

Affiché uniquement dans l'onglet Stabilité, seulement si au moins 1 pattern est détecté (pas de pollution visuelle si tout est propre).

#### Enrichissement Event 41 Kernel-Power (classification fine)

Parsing XML par nom des champs `BugcheckCode`, `SleepInProgress`, `PowerButtonTimestamp` (plus robuste que `Properties[N]` qui varie entre versions Windows). Nouveau champ `CrashCause` sur chaque Event 41 avec 5 valeurs précises :

| CrashCause | Condition | Libellé Dashboard |
|---|---|---|
| `BSODSilent` | BugCheckCode != 0 | BSOD silencieux |
| `SleepResumeFailed` | SleepInProgress = 1 | Reprise veille ratée |
| `UserForcedReset` | PowerButtonTimestamp != 0 | User bouton power |
| `PowerLoss` | dirty shutdown sans les 3 ci-dessus | Coupure alim / thermal |
| `FreezeApp` / `FreezeUnknown` | sinon | Freeze applicatif / inconnu |

Nouvelle stat `Stats.TotalHardCrash` = count des `BSODSilent` + `SleepResumeFailed` + `PowerLoss` (exclut user action et freezes).

#### Fixes divers

- Nom du CPU affiché dans le panel Matériel (via `CPUName`, déjà collecté mais pas montré)
- Fix batterie : vérifie `HasBattery === true` avant d'évaluer `HealthPercent`, sinon desktops tombaient en "Critique batterie 0%"

---

## [1.5.0] — 2026-04-23

### 🆕 Clustering des Event 51 (Disk slow / I/O timeout)

Les Event 51 du provider `Disk` (I/O timeout) arrivent souvent en **rafales massives** lors d'un décrochage SSD ponctuel (glitch PCIe, firmware qui hoquette, câble SATA qui faiblit). Un seul incident physique peut générer des centaines d'events identiques en quelques secondes, noyant l'info utile.

#### Côté Collector

Agrégation automatique des Event 51 consécutifs (fenêtre glissante 60 secondes) en **clusters** :

```json
{
  "Timestamp": "2026-04-23 14:01:45",
  "Type": "Disk slow",
  "Detail": "I/O timeout",
  "Count": 81,
  "IsBurst": true,
  "FirstSeen": "2026-04-23 14:01:45",
  "LastSeen":  "2026-04-23 14:01:46"
}
```

Flag `IsBurst: true` si `Count >= 50` events dans la fenêtre = signal matériel fort.

#### Côté Dashboard

Affichage enrichi dans le panel Performance :
- `I/O timeout` → event isolé (ancien comportement)
- `I/O timeout ×26 events en 1s` → cluster multi-events
- `I/O timeout ×81 events en 1s 🔥 BURST` → badge rouge en plus, signal matériel critique

Sur les JSON v1.4 ou antérieurs, l'ancien format est préservé (backward-compat totale).

**Impact concret** mesuré sur le parc :
- PC #A : 957 lignes I/O illisibles → 8 lignes dont 1 burst de 894 events/1s
- PC #B : 385 lignes → 1 burst unique immédiatement identifiable
- PC #C : 117 lignes → 5 clusters

La corrélation burst-I/O → hard crash (intégrée en v1.6) devient alors possible, et exploite directement ce clustering.

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
  disponible dans `\SERVEUR\share
elease\`.
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
