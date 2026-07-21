# Changelog

Format inspiré de [Keep a Changelog](https://keepachangelog.com).

Chaque script embarque son propre bloc `.CHANGELOG` détaillé dans son en-tête ;
ce fichier consolide les évolutions notables au niveau du projet.

> À partir de la **2.2.0**, le Collecteur et le Dashboard partagent le même
> numéro de version. Auparavant ils étaient versionnés indépendamment (le
> Collecteur jusqu'à 2.1.5, le Dashboard jusqu'à 2.1.14). Pour l'historique fin
> de chaque composant, voir le bloc `.CHANGELOG` en tête de `01_Collector.ps1`
> et `02_Dashboard.ps1`.

---

## [2.4.6]

**Durcissement de la chaîne de mise à jour (suite audit) + robustesse Dashboard.**
Collecteur → 2.4.6, Dashboard → 2.4.6, Updater → 1.9. `version.txt` → 2.4.6
(redéploiement du Collecteur). Schéma JSON inchangé.

### Sécurité — chaîne d'update

- **TOCTOU fermé.** La signature Authenticode est désormais vérifiée sur le **fichier
  local (`.tmp`)** réellement installé puis exécuté, et non plus sur le fichier du
  partage. Un attaquant avec write sur `release\` ne peut plus faire vérifier un
  fichier légitime puis en installer un autre (permutation entre la vérif et la copie).
- **Anti-downgrade de l'Updater.** Le self-update lit la version dans le fichier signé
  à installer et **refuse toute version antérieure** → un ancien Updater, même
  légitimement signé, ne peut plus être ré-injecté (le maillon critique est protégé).
- **Anti-downgrade Collecteur (deux verrous).** (1) Une `version.txt` serveur **non
  parseable** en `[version]` est refusée (elle contournait la garde). (2) La version
  du Collecteur est désormais lue dans le **fichier signé** (marqueur `$CollectorVersion`)
  et non plus dans `version.txt` du share (non signé) → un attaquant ne peut plus
  ré-injecter un vieux Collecteur signé en gonflant `version.txt`. Fail-closed si le
  marqueur est illisible.

### Robustesse — Dashboard

- **Tri `DateBoot` tolérant.** Un seul JSON avec une date `DateBoot` illisible faisait
  planter **toute** la génération (même classe que le `CollectedAt` corrigé en 2.4.5) ;
  le tri utilise maintenant `TryParse` et la génération continue.

### Divers

- `config.psd1.example` : bloc killswitch en `Enabled = $false` (cohérent avec le
  défaut opt-in).

---

## [2.4.5]

**Correctif critique de collecte + correctifs Dashboard.** Le Collecteur ayant
évolué, cette version **nécessite un redéploiement** du parc (`version.txt` → 2.4.5).
Schéma JSON inchangé.

### Corrigé — critique (production)

- **Le nettoyage de log ne bloque plus la collecte.** `Invoke-LogCleanup` s'exécute
  **avant** la collecte ; sur un journal corrompu et gonflé (37 Mo, une seule ligne
  de ~9 millions de caractères — mojibake composé hérité de l'incident 2.2.x), il
  se bloquait ou mourait (OOM) **en gardant le fichier ouvert**, si bien que la
  collecte ne démarrait jamais : **~247 postes figés hors ligne**, journaux
  impossibles à supprimer (« ouvert dans System »). Correctif : lecture **par
  octets bornée** (dernier Mo) au lieu de `Get-Content -Tail` (qui charge des
  lignes entières et s'étouffe sur une ligne géante). **Auto-guérison** en local,
  sans aucune intervention manuelle.

### Corrigé — Dashboard

- **Double-encodage HTML.** Certains champs (`CurrentUser`, `LastLoggedUser`, `IP`,
  `CollectorRunAs`, `CPUName`, `Site`) étaient échappés une première fois à la
  validation (`Test-PCPulseJson`) puis une seconde fois à l'affichage → un nom
  comme `O'Brien` s'affichait `O&#39;Brien`. L'échappement se fait désormais **une
  seule fois**, à l'embed (`ConvertTo-HtmlSafe`). Vérifié : chaque champ concerné
  reste échappé au point de sortie, **aucune perte de protection XSS**.
- **`CollectedAt` illisible ne casse plus la génération.** Le cast en `[datetime]`
  est désormais sous `try/catch` : un seul JSON avec une date corrompue faisait
  planter **toute** la génération du tableau de bord ; le poste concerné s'affiche
  maintenant hors ligne et la génération continue.

### Updater (1.8)

- **Vérification de signature ancrée sur le thumbprint épinglé.** Le contrôle
  exigeait auparavant un statut Authenticode `Valid` (donc la confiance de chaîne
  côté OS) ; tant que le certificat de signature n'est pas déployé dans les magasins
  de confiance des postes, un script pourtant correctement signé est refusé
  (`UnknownError`) — ce qui pouvait figer les mises à jour du parc. Désormais : on
  exige la signature par le **thumbprint épinglé** + l'intégrité (rejet si
  `NotSigned`/`HashMismatch`) et on **tolère** une chaîne non encore de confiance.
  Le pin est l'ancre de confiance ; la confiance de chaîne OS devient un bonus
  (statut `Valid` automatique une fois le certificat déployé par GPO).
- **Nettoyage local à chaque cycle.** La purge du dossier `backup\` (backups legacy
  antérieurs au zéro-backup) et le balayage des `*.new.tmp` d'installation
  interrompue s'exécutent désormais à **chaque** cycle, et plus seulement lors d'une
  mise à jour — un poste déjà à jour se nettoie aussi (auto-guérison).

### Divers

- `.gitignore` : ajout de `*.log` / `log/` (les journaux peuvent contenir des noms
  de postes et chemins réels — protection contre un `git add` accidentel).

---

## [2.4.4]

**Fiabilité des logs + hygiène des postes.** Suite à l'observation d'un log de
poste gonflé à 37 Mo (résidu de l'incident 2.2.x). **Schéma JSON inchangé (`2.2`)**.
Nécessite un redéploiement (Collecteur modifié).

### Corrigé

- **Log Collecteur : mojibake qui se composait.** La rotation lisait le journal en
  ANSI (défaut PS 5.1) alors qu'il est écrit en UTF-8 → à chaque nettoyage, les
  caractères accentués d'un message d'erreur .NET étaient ré-encodés de travers et
  **grossissaient** ; une seule ligne a atteint ~9 millions de caractères (37 Mo).
  Lecture repassée en UTF-8 (cohérente avec l'écriture) + **plafond par ligne**
  (2000 caractères) qui tronque toute ligne monstrueuse et auto-guérit les journaux
  déjà corrompus au prochain passage.

### Durcissement

- **Zéro backup local + installation atomique vérifiée (Updater → 1.6).** L'ancien
  schéma « backup → copie → restaure si échec » est remplacé par une **installation
  sûre sans backup** : copie vers un fichier temporaire → contrôle SHA (copie
  fidèle) → **renommage atomique**. Le fichier vivant n'est jamais remplacé par une
  copie partielle/corrompue, donc **aucun backup n'est nécessaire** — et c'est même
  plus robuste. Plus **aucun code ancien exploitable** stocké sur les postes ;
  `Remove-OldBackups` supprime en prime le dossier `backup\` laissé par les versions
  antérieures.

---

## [2.4.3]

**Durcissement sécurité** (suite à un audit externe). **Schéma JSON inchangé
(`2.2`)**, additif. Le Collecteur ayant évolué (chemin de config), cette version
**nécessite un redéploiement** du parc (`version.txt` 2.4.2 → 2.4.3).

### Sécurité

- **Authenticité de la chaîne de mise à jour (Updater → 1.5).** Le SHA256 ne
  vérifiait que l'intégrité (calculé sur le même partage) : un accès en écriture à
  `release\` permettait d'exécuter du code SYSTEM sur tout le parc. Ajout de la
  **vérification de signature Authenticode + empreinte (thumbprint) épinglée** dans
  l'Updater — appliquée **d'abord au self-update de l'Updater** (maillon critique),
  puis au Collecteur, **avant toute adoption**. Empreinte(s) épinglée(s) **en dur**
  dans l'Updater (jamais dans `config.psd1`, écrivable par les postes). Liste vide
  = vérification désactivée (déploiement progressif). Statut de refus journalisé.
  **Anti-downgrade** : refus d'une version serveur inférieure à la locale (empêche
  le rejeu d'un ancien binaire légitimement signé).
- **XSS stocké (passe 3).** `DiskInfo` (dont `Drive`) et des champs numériques
  (`DurationMin`, `OSBuild`, tailles disque) partaient bruts vers `innerHTML` : un
  poste compromis (ou un EDID forgé) pouvait injecter du JS dans le navigateur de
  l'administrateur. Échappement / cast appliqués ; `DiskInfo` et `BootDurations`
  ajoutés au garde-fou de complétude.
- **`config.psd1` déplacé dans `release\`** (lecture seule pour les postes) au lieu
  de la racine du partage (où « Ordinateurs du domaine » a Modify) : un poste
  compromis ne peut plus altérer la config (phrase killswitch, seuils). Lecture
  avec repli sur la racine pour une migration douce.
- **Killswitch en opt-in.** Défaut passé à `Enabled = $false` : un killswitch actif
  par défaut avec une phrase/nom publics = risque de déni de service du parc si les
  ACL tombent. Il ne s'active que si `config.psd1` déclare explicitement
  `KillSwitch.Enabled = $true` avec des valeurs personnalisées.

### Correctif (Updater 1.4, inclus)

- Rotation de `updater.log` : plafond de taille (lecture par `-Tail` au-delà) —
  `Get-Content -Raw` levait une `OutOfMemoryException` sur un journal de plusieurs
  Go (observé), la rotation échouait et le journal ne redescendait jamais.

---

## [2.4.2]

**N° de série machine + audit du cycle de vie.** Collecteur et Dashboard
**réunifiés à 2.4.2** (le Collecteur repasse sur le numéro commun ; il était resté
en 2.3.2 pour la release Dashboard 2.4.0). **Schéma JSON inchangé (`2.2`)**,
additif et rétrocompatible. Le Collecteur ayant évolué (n° de série), cette
version **nécessite un redéploiement** du parc pour peupler le champ.
(Le Collecteur a été test-déployé en interne sous l'étiquette 2.4.1 — n° de série
seul ; publié en **2.4.2** après ajout du correctif de robustesse WMI ci-dessous,
pour déclencher proprement l'auto-update du parc.)

### Ajouté

- **N° de série machine (service tag).** Le Collecteur relève
  `Machine.SerialNumber` depuis `Win32_BIOS.SerialNumber` (service tag Dell,
  serial HP…) ; les valeurs BIOS factices (`Default string`, `To be filled by
  O.E.M.`…) sont normalisées à `null`. Côté Dashboard : affiché dans le drill-down
  Matériel (carte OS), **colonne `NumeroSerie` dans l'export CSV**, et recherche
  étendue au n° de série. Utile pour l'inventaire, la garantie et le report de
  décommissionnement vers Excel. **Nécessite un redéploiement** du Collecteur
  pour peupler le champ sur le parc (additif : les postes en collecteur antérieur
  restent lisibles, champ vide).
- **Panneau « Cycle de vie du parc » (stats / audit décommissionnement).** Lit le
  registre **complet** (`decommissioning.json`, y compris les postes dont le JSON
  a été purgé), indépendamment des filtres de la vue : compteurs décommissionnés
  / à faire / en retard, **délai moyen marqué→fait**, répartition **par
  technicien** et **par mois**. Valorise le registre comme journal d'audit
  (reporting hiérarchie).

### Corrigé

- **Collecteur : robustesse WMI.** Sur un poste au WMI matériel incomplet (ni
  modèle, ni n° de série), `Win32_OperatingSystem.LastBootUpTime` revenait `null`
  et l'appel de méthode qui suivait faisait **planter toute la collecte** (aucun
  JSON produit → poste invisible au tableau de bord). Les champs de boot noyau
  sont rendus null-safe : la collecte se poursuit et le poste redevient visible.
- **Updater (→ 1.4) : rotation du journal.** La rotation de `updater.log`
  chargeait tout le fichier en mémoire (`Get-Content -Raw`) et levait une
  `OutOfMemoryException` sur un journal devenu énorme (observé : 2 Go) — la
  rotation échouait alors à chaque cycle et le journal ne redescendait jamais.
  Ajout d'un plafond de taille : au-delà, lecture par la fin (`-Tail`, mémoire
  bornée). Même classe de correctif que le Collecteur 2.3.2, côté Updater.

---

## [2.4.0]

Release **Dashboard + outillage**. Le **Collecteur reste en 2.3.2** (aucun
changement de code) : pas de redéploiement du parc, `version.txt` non touché.
**Schéma JSON inchangé (`2.2`)**, additif et rétrocompatible.

### Ajouté

- **Décommissionnement — cycle de vie des postes (🚧 en développement).** Nouvel
  axe, **additif et voué à évoluer**. Un outil interactif `Decommission-PC.ps1`
  (v2.0) marque un poste « À faire » puis « Fait » (workflow), assigne un
  technicien et journalise l'historique, dans un **registre séparé**
  (`decommissioning.json`) — jamais dans le JSON du poste (que le Collecteur
  réécrit à chaque cycle). Le registre vit dans un dossier où les techniciens
  écrivent avec leur **compte de session normal** (hors partage durci) ; le
  Dashboard le lit via `DecommissionRegistryPath` et affiche une famille KPI
  **Cycle de vie** (badges par poste : à décommissionner / en retard / fait mais
  en ligne / faite à purger, et filtres associés). La purge des postes
  décommissionnés reste **manuelle** pour l'instant (suppression du `<PC>.json`
  par un compte disposant de l'écriture sur le partage) ; une purge automatique
  est envisagée. Réglages non sensibles (techniciens, admins, délai) dans
  `decom-config.psd1` à côté du registre.
- **Tuile « En ligne 24h » cliquable** : filtre les postes hors ligne d'un clic.
- **Colonne CPU triable** (du plus ancien au plus récent).

### Modifié

- **KPI Sécurité recomposé** : « EDR arrêté » / « EDR absent » / « OS en fin de
  support » remplacent l'ancien « PC offline » (non pertinent comme signal de
  sécurité). Le score et les autres familles sont inchangés.

---

## [2.3.2]

**Correctif critique** + affinage backlog. **Schéma JSON inchangé (`2.2`)**,
additif et rétrocompatible. (Regroupe le lot testé sous l'étiquette interne 2.3.1
+ les finitions suivantes ; publié en 2.3.2 pour déclencher proprement
l'auto-update du parc.)

### Corrigé — critique (production)

- **Publication du JSON sur le partage réparée.** L'écriture « atomique » via
  `[System.IO.File]::Replace` échoue sur un chemin réseau UNC/SMB (« le chemin
  d'accès n'a pas une forme conforme ») : l'export échouait à chaque exécution,
  les fichiers du partage n'étaient plus mis à jour (postes « figés » côté
  tableau de bord) et les journaux gonflaient sans fin. Repli sur un renommage
  compatible SMB (`Move-Item -Force`) quand le renommage atomique n'est pas
  supporté.
- **Plafond de taille des journaux.** La rotation ne purgeait que par ancienneté
  (30 jours), sans limite de taille : un poste en échec répété voyait son journal
  atteindre plusieurs Go et saturer le partage. Ajout d'un plafond (troncature de
  la fin au-delà d'un seuil) qui s'auto-répare à l'exécution suivante.

### Ajouté

- **Lisibilité des crashers** (drill-down) : « X plantés / Y figés », origine
  déduite du module fautif (application / interne mémoire-système / .NET /
  conflit de composant) et code d'exception traduit en clair
  (`0xc0000005` → accès mémoire invalide, etc.). Plus aucun code brut ni « hang »
  affiché ; le détail technique reste en infobulle pour l'administrateur.
- **`CollectorRunAs`** : compte réel d'exécution de l'agent (écrit à chaque
  relevé), exporté dans le CSV — permet d'auditer le compte de service utilisé
  sur tout le parc sans se connecter aux postes.
- **Jusqu'à 25 applications en échec** listées par poste (au lieu de 10) — la
  limite tronquait les postes cumulant beaucoup d'applications distinctes.

### Sécurité / robustesse

- **Cast strict des champs numériques** du Dashboard (défense en profondeur) :
  un champ censé être un nombre est forcé au type numérique (ou `null`), jamais
  interprété comme du texte.
- **Garde-fou de complétude** étendu aux sous-objets (crashers, erreurs
  matérielles, moniteurs, disques, barrettes mémoire) : un champ ajouté au
  collecteur mais oublié à l'affichage est signalé à la génération.
- **Fusion de configuration récursive** : un `ScoreWeights` partiel dans
  `config.psd1` ne remet plus à zéro (en silence) les poids non redéfinis ; seuls
  les poids explicitement fournis sont surchargés.

### Corrigé

- Résumé de fin de collecte qui n'affichait pas l'uptime.
- Profondeur de sérialisation JSON augmentée (marge pour les objets imbriqués).
- Documentation du délai anti-collision (0-15 min) alignée sur le code.
- Nettoyage de branches de sanitisation visant des champs inexistants (no-op
  silencieux) ; assainissement aligné sur les vrais champs.
- Utilitaire de démo (`Refresh-DemoDates`) : gestion de l'heure UTC et de
  l'alignement des dates sans heure.

---

## [2.3.0]

Release unifiée regroupant les évolutions publiées entre-temps (le parc a été
déployé sous les étiquettes internes 2.2.1 → 2.2.4 ; la version publique repart
d'un numéro commun Collecteur + Dashboard, comme acté en 2.2.0). **Schéma JSON
inchangé (`2.2`)** : tous les ajouts ci-dessous sont additifs et rétrocompatibles
(le Dashboard décode aussi côté client, donc un parc pas encore régénéré reste
lisible).

### Ajouté

- **Inventaire OS Windows 10 / 11** (initialement 2.2.1). Quatre champs additifs
  dérivés du build (`OSProduct` — jamais `ProductName`/`Caption` —, `OSBuild`,
  `OSDisplayVersion`, `OSEdition`), colonne + filtre « OS » dans le tableau,
  bloc OS dans le drill-down Matériel, colonnes OS dans l'export CSV. Utile pour
  l'inventaire de parc et le suivi de fin de support Windows 10.
- **Dernier utilisateur connu** (`Machine.LastLoggedUser`). Quand aucune session
  n'est active au moment de la collecte (« (aucune session) »), le Dashboard
  affiche le dernier utilisateur ayant ouvert une session (une seule entrée, pas
  d'historique). Recherche et export CSV étendus à ce champ.
- **Inventaire mémoire enrichi.** Le fabricant des barrettes est décodé depuis
  l'identifiant JEDEC brut (ex. `80AD000080AD` → `SK Hynix`) ; l'identifiant brut
  reste disponible (`ManufacturerRaw`).
- **Écrans à EDID non transmis.** Les écrans branchés via dock / adaptateur / KVM
  qui ne relaient pas l'EDID (fabricant `@@@`, code produit `0000`) sont
  détectés, tentés en repli via l'EDID en cache du registre, et sinon libellés
  « Écran non identifié » (champ `Identified`) — exclus du top fabricants et du
  calcul d'âge au lieu d'afficher des identifiants parasites.

### Modifié

- **UI drill-down Matériel** : cartes RAM et SMART aérées (largeurs, espacements
  et tailles de police revus ; elles étaient tassées en colonne étroite).
- **Écriture atomique du JSON** côté Collecteur (buffer temporaire puis bascule
  par renommage) — fiabilise la lecture concurrente par le Dashboard.

### Corrigé

- **Rotation du log machine** (`Invoke-LogCleanup`) : gestion CRLF corrigée (le
  log ne s'élaguait pas de façon fiable).

### Sécurité / déploiement (scripts d'installation)

- **Updater (→ 1.3)** : self-update de l'agent par empreinte SHA256 ; nettoyage
  automatique des traces post-déploiement sur les postes (`Invoke-LocalCleanup`) ;
  journal déplacé sur le partage (`logs\<HOST>.updater.log`, plus aucun log local
  sur le poste) ; durcissement des ACL du dossier runtime (SYSTEM + Administrateurs).
- **Install-Client (→ 2.3)** : auto-nettoyage du dossier source en fin
  d'installation et pose des ACL durcies dès l'installation.

---

## [2.2.0]

Généralisation du projet pour publication open source. Le code ne nomme plus
aucun produit ni aucune spécificité d'environnement : tout le paramétrable vit
dans `config.psd1` (non publié ; un modèle `config.psd1.example` est fourni).

### Modifié — BREAKING (schéma JSON `2.1` → `2.2`)

- **Services critiques pilotés par configuration.** Le Collecteur ne surveille
  plus un service codé en dur : il lit la clé `MonitoredServices` de
  `config.psd1`, une liste dont chaque entrée porte `Id`, `ServiceName`,
  `DisplayName`, `Role`, `AlertIfNotInstalled`. Le match se fait sur
  `ServiceName` puis, en repli, sur `DisplayName` (robustesse multi-versions).
  `AlertIfNotInstalled` (défaut : vrai) : un service déclaré mais absent lève
  une alerte.
- **Schéma `ServicesHealth`.** L'ancien objet unique (un seul service, codé en
  dur) devient une **liste** `ServicesHealth.Monitored`. Le champ `Role`
  permet au Dashboard de repérer le service structurant (par exemple l'EDR) pour
  le score et le badge — l'interprétation reste côté Dashboard.
- **`ScoreWeights`.** Le poids de l'EDR est renommé `EDRDown` (le code ne nomme
  plus le produit). Un repli côté JavaScript protège une configuration qui ne le
  déclarerait pas encore.
- **Colonnes CSV** de l'EDR renommées (`EDRStatus`, `EDRInstalle`, `EDRAlerte`).

### Ajouté

- **`PriorityApps` externalisée** dans `config.psd1` : la section « À investiguer
  en priorité » du panneau Top Crashers est désormais alimentée par une liste
  définie par chaque site (applis métier, socle bureautique, sécurité/réseau…).
  Le défaut générique se limite au socle bureautique/collaboratif.

### Sécurité / compatibilité

- **Whitelist SchemaVersion élargie à `2.1` et `2.2`.** Le Dashboard lit les deux
  pour absorber un déploiement poste par poste : un poste encore en Collecteur
  2.1 (JSON `2.1`) affiche l'EDR « en attente » (neutre) plutôt qu'une fausse
  alerte. La liste sera resserrée une fois le parc entièrement migré.
- **Défaut générique sûr :** sans `MonitoredServices` ni `PriorityApps` dans
  `config.psd1`, aucun service n'est suivi et la liste de priorité retombe sur le
  défaut — aucune donnée d'environnement n'est jamais codée dans le dépôt.

> **Ordre de déploiement recommandé :** publier le Dashboard 2.2 **d'abord** (il
> accepte 2.1 et 2.2), puis pousser le Collecteur 2.2 poste par poste. Lors d'un
> changement de clés de configuration, déposer le `config.psd1` mis à jour sur le
> share **avant** de publier la nouvelle version du Collecteur, pour qu'aucun
> poste ne lance le nouveau code contre une configuration incomplète.

---

## Versions antérieures

### Collecteur

- **2.1.5** — Top Crashers : détail figé/planté, module fautif et code exception
  (Phase 1). Quatre champs additifs, schéma `2.1` inchangé (rétrocompatible).
- **2.1.4** — Correction du profilage CPU AMD récent (Ryzen série 200, Ryzen AI)
  qui étaient classés à tort « ancien ».
- **2.1.3** — Distinction crash réel vs échec applicatif récurrent (champ `Type`
  sur chaque entrée Top Crashers).
- **2.1.2** — Fiabilisation de Top Crashers : agrégation par clé normalisée (fin
  des doublons de type events tiers écrits sous l'ID 1000).
- **2.1.1** — Correction de la détection throttling / thermal (suppression de
  faux positifs sur des machines saines).
- **2.1 et antérieur** — Passage du schéma en `2.1` puis `2.0` (durcissement
  sécurité : whitelist de schéma, limites de taille JSON, échappement HTML),
  inventaire matériel, collecte boot / crash / BSOD. Détail dans l'en-tête du
  script.

### Dashboard

- **2.1.14** — Bouton « Tout effacer » réinitialisant tous les filtres en une
  fois (chips, menus, recherche).
- **2.1.13** — Piste matérielle « mémoire » (co-occurrence d'un crasher mémoire
  `0xc0000005` / `0xc0000374` et d'une erreur WHEA sur composant RAM) ; recopie
  des champs additifs du Collecteur 2.1.5 vers l'affichage.
- **2.1.12** — Panneau Top Crashers : section « À investiguer en priorité » et
  filtrage du tableau par application au clic.
- **2.1.11** — Refonte UX des anomalies : signalées **par PC** (badge discret +
  filtre dédié) au lieu d'un bandeau global, pour ne pas primer sur l'axe santé.
- **2.1.10** — Message de troncature rendu factuel (information de volumétrie,
  plus un faux diagnostic de « reboot loop »).
- **2.1.9** — Retrait de la détection « SimilarHostname » (faux positifs
  structurels sur une nomenclature séquentielle).
- **Antérieur** — Détection d'anomalies et validation des JSON (`Test-PCPulseJson`,
  `Test-PCPulseAnomalies`), durcissement sécurité, KPIs et drill-down par PC.
  Détail dans l'en-tête du script.
