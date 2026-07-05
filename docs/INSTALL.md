# Installation — PCPulse

Ce guide couvre un **déploiement réel sur un parc** (serveur + postes + Dashboard).
Pour juste voir le Dashboard sans rien installer, reste sur le **Quick Start** du
[README](../README.md) avec les JSON de démo.

Le principe : chaque PC dépose un JSON sur un **partage SMB** ; un poste admin lit
ces JSON et génère un **rapport HTML autonome**. Il y a donc trois briques à mettre
en place, dans cet ordre : **le serveur (partage)**, **la configuration**, puis
**les postes**. On génère le Dashboard à la fin.

---

## Prérequis

- Un **domaine Active Directory** (l'authentification des postes au partage se fait
  via le compte machine, en Kerberos).
- Un **serveur de fichiers Windows** joint au domaine pour héberger le partage.
- Des **droits admin** : admin local (ou Domain Admin) sur le serveur ; admin local
  sur chaque poste pour l'installation.
- Sur le **poste admin** qui génère le Dashboard : **PowerShell 7** (`pwsh`).
  `winget install Microsoft.PowerShell`. ⚠️ Le Dashboard **ne tourne pas** en
  PowerShell 5.1.
- *(Optionnel, recommandé)* un **gMSA** créé au préalable dans l'AD si tu veux que la
  tâche planifiée tourne sous un compte de service dédié au lieu de `SYSTEM`.

> Le Collector, lui, tourne en **PowerShell 5.1** natif sur les postes — rien à
> installer côté client.

---

## Étape 1 — Le serveur (partage SMB)

### 1.1 Créer le partage `PCPulse$`

> ⚠️ **`Setup-Server.ps1` ne crée pas le partage**, il ne fait que **durcir** ses
> droits. Il faut donc créer le partage d'abord.

Sur le serveur, en admin, crée un dossier et partage-le en donnant l'accès en
écriture aux comptes machine du domaine (les vrais droits seront resserrés à
l'étape suivante par les ACLs NTFS) :

```powershell
New-Item -Path 'D:\PCPulse' -ItemType Directory -Force
New-SmbShare -Name 'PCPulse$' -Path 'D:\PCPulse' `
    -FullAccess 'Domain Admins' `
    -ChangeAccess 'Domain Computers'
```

Adapte le chemin (`D:\PCPulse`) à ton serveur. Le `$` final rend le partage caché,
c'est volontaire.

### 1.2 Durcir les ACLs avec `Setup-Server.ps1`

Ce script pose les ACLs NTFS **explicites** : il crée les sous-dossiers `release\`,
`killed\`, `logs\`, casse l'héritage, met `Domain Computers` en **lecture seule sur
`release\`** (le point critique : empêche un poste compromis de pousser un Collector
malveillant → RCE parc), retire `BUILTIN\Users`, et met l'owner sur `Domain Admins`.
Il **sauvegarde les ACLs existantes** dans `.acl-backup\` avant toute modif, et il
est **idempotent** (relançable sans risque).

D'abord un **essai à blanc** pour voir ce qu'il ferait, sans rien changer :

```powershell
.\Setup-Server.ps1 -WhatIf
```

Puis, pour appliquer :

```powershell
# Classique (tâche en SYSTEM)
.\Setup-Server.ps1

# Avec un gMSA dédié
.\Setup-Server.ps1 -gMSAName 'svc-pcpulse$'

# En préservant un groupe admin infra existant (lecture) :
.\Setup-Server.ps1 -gMSAName 'svc-pcpulse$' -ExtraReadGroups @('GR-ADMINS-IT')
```

Le script termine par une **vérification** : il doit afficher que `Domain Computers`
peut **lire** `release\` mais **pas le modifier**. Le détail du modèle de sécurité est
dans [SECURITY.md](../SECURITY.md).

---

## Étape 2 — La configuration (sur le partage)

`config.psd1` et `ip-ranges.csv` se placent **à la racine du partage** (`\\SERVEUR\PCPulse$\`).
Ils sont volontairement **exclus du dépôt** (`.gitignore`) : tu les crées à partir des
modèles `.example`.

### 2.1 `config.psd1` — important pour l'EDR

Copie `config.psd1.example` en `config.psd1` sur le partage et édite-le. Le point à
ne pas rater : **`MonitoredServices`**. Tant que tu n'y déclares pas ton EDR, la
colonne / le badge EDR du Dashboard reste **vide** — et on croit à un bug alors que
tout va bien.

Déclare le service Windows de ton EDR avec `Role = 'EDR'` (c'est ce `Role` que le
Dashboard repère pour le score et le badge) :

```powershell
MonitoredServices = @(
    @{ Id = 'edr'; DisplayName = 'Mon EDR'; ServiceName = 'NomDuServiceWindows'; Role = 'EDR' }
    # tu peux en ajouter d'autres (VPN, sauvegarde...) avec Role = $null
)
```

Le même fichier permet d'ajuster les seuils, la pondération du score, le titre du
Dashboard, la liste `PriorityApps` (« à investiguer en priorité »), et la
**phrase du killswitch** (à personnaliser en prod). Tout est commenté dans
`config.psd1.example`.

### 2.2 `ip-ranges.csv` — optionnel (colonne Site)

Si tu veux la colonne **Site**, copie `ip-ranges.example.csv` en `ip-ranges.csv` sur
le partage. Le Dashboard 2.2 mappe les sites **par plage IP (CIDR) uniquement**, sur
l'IP de la machine :

```csv
Pattern1,Entity
10.10.0.0/24,HQ-Paris
10.20.0.0/24,Branch-Lyon
192.168.50.0/24,RemoteSite-A
```

- `Pattern1` = une plage **CIDR** (doit contenir un `/`) ; `Entity` = le libellé du site.
  Les noms de colonnes sont imposés (c'est ce que le Dashboard cherche).
- **Première plage qui contient l'IP gagne** : mets les plages les plus spécifiques en haut.
- Une IP dans aucune plage → site `Inconnu`. La colonne n'apparaît que si au moins une
  plage CIDR valide est chargée.
- Le matching se fait **par IP/CIDR uniquement** (pas par nom d'hôte).
- Fichier lu **côté Dashboard** uniquement : pas besoin de redémarrer les Collectors.
- Absent → la colonne Site est simplement masquée, l'outil marche quand même.
- Détails complets dans [`ip-ranges.README.md`](../ip-ranges.README.md).

---

## Étape 3 — Déployer le Collector sur les postes

### 3.1 Sur un poste pilote

Sur un poste (en admin), lance `Install-Client.ps1` en pointant le partage :

```powershell
# Classique (tâche en SYSTEM)
.\Install-Client.ps1 -ServerPath "\\SERVEUR\PCPulse$"

# Avec gMSA
.\Install-Client.ps1 -ServerPath "\\SERVEUR\PCPulse$" -gMSAName 'svc-pcpulse$'
```

Le script vérifie l'accès au partage (port 445), copie le Collector + l'Updater dans
`C:\ProgramData\PCPulse\`, crée la **tâche planifiée `PCPulse-Collector`** (qui pointe
sur l'**Updater**, pas sur le Collector directement — c'est ce qui permet la mise à
jour automatique), puis lance une première collecte.

### 3.2 Vérifier

Le Collector a un **délai anti-collision aléatoire** (~0-10 min) : le premier JSON
peut ne pas apparaître tout de suite. Attends quelques minutes puis vérifie qu'un
fichier `NOM-DU-PC.json` est bien arrivé à la racine du partage.

### 3.3 Passage à l'échelle

Une fois le pilote validé, déploie sur le reste du parc via ton canal habituel
(Intune, GPO, SmartDeploy…) : tu pousses les fichiers + crées la même tâche
planifiée. Les postes se mettront à jour tout seuls quand tu publieras une nouvelle
version dans `release\` (l'Updater compare les SHA256).

---

## Étape 4 — Générer le Dashboard

Depuis le poste admin, **en PowerShell 7** :

```powershell
pwsh .\02_Dashboard.ps1 -SharePath "\\SERVEUR\PCPulse$"
```

Le script lit tous les JSON du partage, applique `config.psd1` + `ip-ranges.csv`, et
génère un **HTML autonome** qui s'ouvre dans le navigateur.

> Si Windows bloque l'exécution (`... not digitally signed`), c'est la protection par
> défaut. Ponctuel : ajoute `-ExecutionPolicy Bypass`. Permanent (conseillé, une fois
> en admin) : `Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned`.

---

## Pièges fréquents

- **« Le Dashboard plante / ne se lance pas »** → mauvais PowerShell. Le Dashboard
  exige **PS7 (`pwsh`)**, pas 5.1.
- **« Rien ne s'affiche, exécution bloquée »** → ExecutionPolicy (voir ci-dessus).
- **« La colonne / le badge EDR est vide »** → `MonitoredServices` pas renseigné dans
  `config.psd1` (ou pas de `Role = 'EDR'`). Ce n'est pas un bug.
- **« Aucun JSON n'arrive sur le partage »** → les comptes machine n'ont pas le droit
  d'écrire. Vérifie les droits **de partage** (`Domain Computers` en Change) — les
  droits NTFS sont gérés par `Setup-Server.ps1`.
- **« Setup-Server dit que le partage n'existe pas »** → tu as sauté l'étape 1.1 (il
  ne crée pas le partage, il le durcit).
- **« La colonne Site n'apparaît pas »** → pas de `ip-ranges.csv` sur le partage, ou
  `Pattern1` qui n'est pas un CIDR (le matching par nom d'hôte n'existe pas en 2.2).

---

## Désinstaller / décommissionner

- **À distance (propre)** : le **killswitch** — dépose le fichier sentinelle dans
  `release\` avec la phrase de confirmation. Chaque poste se désinstalle seul (tâche
  planifiée + dossier local supprimés) et laisse un rapport dans `killed\`. Procédure
  complète dans [SECURITY.md](../SECURITY.md).
- **Manuel sur un poste** :

  ```powershell
  Unregister-ScheduledTask -TaskName 'PCPulse-Collector' -Confirm:$false
  Remove-Item 'C:\ProgramData\PCPulse' -Recurse -Force
  ```

---

## Récapitulatif de l'ordre

1. Créer le partage `PCPulse$` (**1.1**)
2. Durcir les ACLs : `Setup-Server.ps1` (**1.2**)
3. Déposer `config.psd1` (avec `MonitoredServices`) + `ip-ranges.csv` sur le partage (**2**)
4. `Install-Client.ps1` sur un poste pilote (**3.1**) → attendre le JSON (**3.2**)
5. Générer le Dashboard en PS7 (**4**)
6. Passer à l'échelle sur le reste du parc (**3.3**)
