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
