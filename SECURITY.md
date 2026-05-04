# Security policy

PCPulse est conçu pour être **déployé sur du parc en production** (jusqu'à plusieurs milliers de PC). À ce titre, la sécurité du projet est prise au sérieux et fait partie intégrante de l'architecture.

Ce document explique :
- Le **modèle de confiance** (trust model) utilisé
- Les **ACLs recommandées** sur le partage SMB
- Comment **fonctionne le killswitch** (auto-désinstallation à distance)
- Comment **signaler une vulnérabilité**
- La **roadmap de hardening** (ce qui est fait, ce qui est en cours, ce qui reste)

---

## 🛡️ Trust model

PCPulse repose sur **un partage SMB** comme point central d'échange entre les PC du parc et l'admin. Comprendre qui peut faire quoi sur ce partage est essentiel pour bien sécuriser le déploiement.

### Acteurs

| Acteur | Rôle | Niveau de confiance |
|---|---|---|
| **Admin** (toi, ou ton équipe) | Pousse les nouvelles versions du Collector. Lit le Dashboard. Déclenche le killswitch. | ✅ Confiance totale |
| **Compte machine du PC** (`PC-XXX$`) | Lit le Collector + version.txt. Écrit son JSON à la racine. Écrit son rapport killed le cas échéant. | 🟡 Confiance limitée (partiellement compromis si le PC est compromis) |
| **gMSA dédié** (optionnel, ex: `svc-pcpulse$`) | Identique au compte machine. Recommandé en remplacement de SYSTEM pour audit propre. | 🟡 Confiance limitée |
| **User lambda du domaine** (`DOMAIN\toto`) | Aucun rôle dans PCPulse. | 🔴 Aucune confiance (= attaquant potentiel) |

### Le scénario d'attaque principal

Le risque numéro 1 dans ce design, c'est :

> **"Un attaquant compromet 1 PC du parc → il peut écrire dans `\release\` → il remplace `01_Collector.ps1` par sa version malveillante → au prochain cycle, les 800 PC téléchargent et exécutent son code en SYSTEM."**

C'est une **escalation latérale massive** : 1 endpoint compromis = RCE SYSTEM sur tout le parc.

### La défense

Pour fermer ce scénario, le partage doit être configuré pour que :
- **Domain Computers** puisse **lire** `\release\` (pour télécharger le Collector)
- **Domain Computers** ne puisse **PAS modifier** `\release\` (pour ne pas pouvoir y pousser du code)
- Seuls les **admins** (ou le compte de service de déploiement) peuvent écrire dans `\release\`

C'est exactement ce que fait `Setup-Server.ps1` v2.0.

---

## 🔒 ACLs recommandées

### Vue d'ensemble

| Dossier | SYSTEM | Admins du domaine | Domain Computers | gMSA (optionnel) | Rationale |
|---|---|---|---|---|---|
| `\` (racine) | FullControl | FullControl | **Modify** | Modify | Les PC déposent leurs JSON ici |
| `\release\` ⭐ | FullControl | FullControl | **ReadAndExecute** | ReadAndExecute | **Lecture seule pour les PC** (anti-RCE) |
| `\killed\` | FullControl | FullControl | Modify | Modify | Les PC déposent leur rapport de mort ici |
| `\logs\` | FullControl | FullControl | Modify | Modify | Les PC peuvent écrire des logs |

⭐ **Le dossier `\release\` est le plus critique.** Tout l'enjeu sécurité tient ici.

### Pourquoi `ReadAndExecute` et pas juste `Read` ?

`ReadAndExecute` est le nom Windows pour "lecture autorisée d'un fichier exécutable". Il **n'autorise pas Windows à exécuter automatiquement** le fichier — il dit juste *"si quelqu'un veut le lancer, ça passe la vérification d'accès."*. C'est le bon niveau pour un fichier `.ps1` qui sera **téléchargé puis lancé localement par le PC**.

Concrètement, `ReadAndExecute` autorise :
- ✅ Ouvrir et lire le contenu
- ✅ Lancer le fichier en script (si on en a envie)

Et **interdit** :
- ❌ Modifier le fichier
- ❌ Le supprimer
- ❌ Le remplacer
- ❌ Créer de nouveaux fichiers à côté

### Pourquoi pas de `Authenticated Users` ou `Domain Users` ?

Beaucoup de tutoriels recommandent `Authenticated Users` ou `Domain Users` pour simplifier. **Ne le faites pas pour PCPulse.** Ces groupes incluent **tous les comptes utilisateur**, ce qui crée plusieurs problèmes :
- Un user lambda peut lire les JSON et savoir qui a quoi comme PC
- Si vous combinez avec `Modify`, un user lambda peut polluer le share

PCPulse a besoin d'**accès machine** (`Domain Computers`), pas user. Les comptes machine s'authentifient via Kerberos quand un PC fait `\\SERVEUR\share\`. C'est natif et plus sûr.

### Le rôle de `BUILTIN\Users`

Sur Windows, `BUILTIN\Users` **inclut implicitement Authenticated Users**. Donc si vous laissez `BUILTIN\Users` avec `CreateFiles + AppendData` (configuration par défaut sur certains serveurs), vous offrez à n'importe quel user du domaine la possibilité de créer des fichiers dans le share.

→ **`Setup-Server.ps1` retire ce groupe** des ACLs lors du hardening.

### Owner du share

Eviter les **comptes utilisateur personnels** comme owner. Préférer un **groupe AD** (genre `Domain Admins`) ou un **compte de service**. Bénéfices :
- Audit plus clair (un pentester voit *"l'admin du share, c'est l'équipe X"* et pas *"l'admin du share, c'est Robert qui est parti il y a 2 ans"*)
- Continuité si la personne quitte la structure
- Plus difficile de cibler une personne précise pour escalation sociale

→ `Setup-Server.ps1` réassigne automatiquement les owners aux `Domain Admins`.

### Casser l'héritage NTFS

L'héritage NTFS fait propager les ACLs de la racine du share vers tous les sous-dossiers. C'est **pratique mais dangereux** : ça force à mettre des ACLs uniformes sur toute l'arborescence, alors que `\release\` doit être plus restrictif que la racine.

→ `Setup-Server.ps1` **casse l'héritage** sur les sous-dossiers et applique des ACLs **explicites** par dossier.

### Backups d'ACLs

Avant toute modification, `Setup-Server.ps1` exporte les ACLs actuelles dans `\\SERVER\PCPulse$\.acl-backup\acl-<dossier>-<timestamp>.txt`. Tu peux toujours revenir en arrière en cas de problème.

---

## ☠️ Killswitch — auto-désinstallation à distance

PCPulse intègre un **mécanisme killswitch** qui permet de désinstaller à distance les Collectors de tous les PC du parc, sans avoir à toucher physiquement aux machines.

### Cas d'usage

- **Décommissionnement** : tu remplaces PCPulse par autre chose, tu veux nettoyer proprement
- **Stop d'urgence** : tu détectes un bug critique sur le Collector (genre fuite mémoire), tu veux tout arrêter en urgence
- **Migration majeure** : tu veux redéployer PCPulse v2 from scratch (kill v1 → install v2)

### Comment ça fonctionne

1. **L'admin** dépose un fichier sentinelle dans `\\SERVER\PCPulse$\release\`. Par défaut : `KILLSWITCH.txt` contenant exactement la chaîne `CONFIRM-UNINSTALL-PCPULSE`.

2. **Au prochain cycle horaire**, chaque PC qui exécute `PCPulse-Updater.ps1` :
   - Détecte le fichier sentinelle
   - Vérifie le contenu exact (sinon ignore)
   - Écrit un rapport dans `\killed\<HOSTNAME>.txt`
   - Supprime sa tâche planifiée `PCPulse-Collector`
   - Lance un cleanup différé qui supprime `C:\ProgramData\PCPulse\`
   - Exit

3. **L'admin** retire le fichier sentinelle quand tous les PC sont apparus dans `\killed\` (cleanup manuel). Sinon les PC qui se rallument après coup se kill aussi (utile : ne laisse pas d'agents zombies).

### Personnalisation (recommandée en prod)

Les valeurs par défaut sont **publiques** (puisque dans le repo open source). En prod, **change la phrase** via `config.psd1` pour que personne ne puisse trigger le kill par accident :

```powershell
# Dans config.psd1
KillSwitch = @{
    Enabled  = $true
    Filename = 'NOM_PERSONNALISE.txt'
    Phrase   = 'PHRASE-COMPLEXE-IMPOSSIBLE-A-DEVINER'
}
```

### Modèle de menace du killswitch

#### Ce que protège le killswitch
- ✅ **Erreur humaine** : un admin pose un fichier `test.txt` dans `\release\` par erreur → ne déclenche rien (mauvais nom de fichier)
- ✅ **Attaquant en lecture seule** sur le share → ne peut pas déclencher de kill (pas le droit d'écrire le fichier sentinelle)
- ✅ **Compte machine compromis** → s'il a uniquement `ReadAndExecute` sur `\release\` (configuration recommandée), ne peut pas écrire de fichier sentinelle

#### Ce que NE protège PAS le killswitch
- ❌ **Attaquant avec écriture sur `\release\`** : si vos ACLs autorisent l'écriture (mauvaise configuration), un attaquant peut déclencher un kill du parc. **C'est le scénario d'attaque majeur de PCPulse**, et c'est précisément pour ça que `Setup-Server.ps1` v2.0 verrouille `\release\` en lecture seule pour les comptes machine.
- ❌ **Admin malveillant** : un admin avec écriture sur `\release\` peut bien sûr déclencher le kill. C'est attendu (c'est un outil admin).

#### La phrase secrète, c'est de la défense en profondeur

La phrase secrète dans `config.psd1` n'est **pas une protection magique**. Si un attaquant a déjà l'écriture sur `\release\`, il peut aussi lire `config.psd1` (sauf si vous le placez ailleurs). La phrase sert principalement à :
- Empêcher les **erreurs humaines** (pas de kill par fausse manip)
- Empêcher les **scans automatiques** (un scanner qui poserait `KILLSWITCH.txt` sans contenu spécifique ne déclencherait rien)
- Forcer une **action volontaire et documentée** pour kill le parc

→ **La vraie protection, ce sont les ACLs** sur `\release\`.

---

## 🔬 Surface d'attaque résiduelle

Même avec le hardening complet, il reste des risques. Les voici, classés par criticité.

### 🟡 Moyen : compromission d'un compte admin
Si un compte avec écriture sur `\release\` est compromis (phishing d'un admin, vol de credentials), l'attaquant peut pousser un Collector malveillant. **Mitigation** :
- Limiter le nombre d'admins ayant écriture
- Imposer MFA sur les comptes admin
- Voir la roadmap : **code signing** rendra cette attaque encore plus difficile

### 🟡 Moyen : injection via JSON malformé
Le Dashboard parse les JSON déposés par les PC. Un PC compromis pourrait déposer un JSON manipulé pour faire du XSS dans le rapport HTML. **Mitigation** :
- Le Dashboard utilise `ConvertFrom-Json` natif (pas d'injection PowerShell)
- Voir la roadmap : **sanity-checks** stricts sur les JSON + **CSP inline** dans le HTML

### 🟢 Faible : fuite d'information
Les rapports HTML contiennent les noms des PC, IPs, modèles de hardware. **Mitigation** :
- Le Dashboard est généré localement par l'admin (pas exposé sur le réseau)
- Si tu veux le partager, supprime ou anonymise avant

### 🟢 Faible : déni de service
Un PC compromis pourrait remplir le share en spam de fausses entrées. **Mitigation** :
- Quotas sur le share au niveau Windows
- Voir la roadmap : **détection d'anomalies** (futur)

---

## 🛣️ Roadmap de hardening

### ✅ Fait (v2.0)

- **Setup serveur hardenisé** : `Setup-Server.ps1` v2.0 (ACLs explicites, héritage cassé, owner groupe AD)
- **Killswitch configurable** : phrase et fichier personnalisables via `config.psd1`
- **Documentation trust model** (ce document)
- **Support gMSA** : possibilité de remplacer SYSTEM par un compte de service dédié

### 🔄 En cours

- **Sanity-checks Dashboard** : rejet des JSON aux schémas inattendus, validation `Machine.PC` = nom de fichier, types contrôlés
- **CSP inline HTML** : Content-Security-Policy meta tag dans le HTML généré pour bloquer toute injection XSS résiduelle

### 📅 Planifié

- **Code signing** : signature numérique du Collector et de l'Updater via certificat AD CS interne. L'Updater refusera tout fichier dont la signature n'est pas valide ou dont le certificat n'est pas dans le store Trusted Publishers.
- **Vérification de signature dans l'Updater** : refus du téléchargement si signature invalide (anti-tampering supplémentaire)
- **Détection d'anomalies** : JSON dans le futur, PC qui spam, hostnames inhabituels
- **Audit logs** : log immuable des actions admin (push de version, kill, etc.)
- **Migration scheduled task → gMSA** : remplacement de l'exécution SYSTEM par un compte de service dédié sur tous les PC

### 🤔 À l'étude

- **Chiffrement SMB3 obligatoire** : déjà recommandé, à formaliser
- **Auto-rotation de la phrase killswitch** : pour limiter le risque de fuite
- **Audit AD du gMSA** : monitoring des droits du gMSA pour détecter une éventuelle élévation

---

## 📢 Signaler une vulnérabilité

Si vous découvrez une vulnérabilité dans PCPulse, **merci de NE PAS l'ouvrir en Issue publique**.

Préférez :
1. **GitHub Security Advisories** : <https://github.com/Damien-Gouhier/pcpulse/security/advisories/new> (recommandé, processus privé)
2. **Email direct** : ouvrir le profil GitHub <https://github.com/Damien-Gouhier> pour récupérer un contact
3. **Issue chiffrée** si rien d'autre n'est possible

Engagement :
- **Réponse initiale** : sous 7 jours
- **Triage et validation** : sous 30 jours
- **Fix** : selon la criticité (critique = 30 jours, élevé = 60 jours, moyen = release suivante)
- **Crédit** : tu seras crédité dans le CHANGELOG sauf si tu préfères rester anonyme

---

## 📚 Références

- [Microsoft Tier Model](https://learn.microsoft.com/en-us/security/privileged-access-workstations/privileged-access-access-model) — pour comprendre la séparation T0/T1/T2
- [Microsoft well-known SIDs](https://learn.microsoft.com/en-us/windows/security/identity-protection/access-control/security-identifiers) — référence pour les SIDs utilisés dans `Setup-Server.ps1`
- [gMSA — Group Managed Service Accounts](https://learn.microsoft.com/en-us/windows-server/security/group-managed-service-accounts/group-managed-service-accounts-overview) — pour comprendre la migration SYSTEM → gMSA
- [NIST SP 800-167 — Application Whitelisting](https://csrc.nist.gov/publications/detail/sp/800-167/final) — pour le code signing prévu

---

*Dernière mise à jour : 2026-05-04 — release v2.0* 


