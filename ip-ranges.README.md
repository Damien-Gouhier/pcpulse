# ip-ranges.csv — Mapping IP → Site

PCPulse peut enrichir chaque machine avec un libellé de **site** (bureau, agence,
site distant, etc.) en fonction de son adresse IP. Le Dashboard s'en sert pour :

- Afficher la colonne **Site** dans le tableau
- Filtrer les machines par site via le menu déroulant de la toolbar
- Regrouper les agrégats si tu gères un parc multi-sites

Si le fichier est absent, la colonne Site est simplement masquée — l'outil fonctionne
sans.

## Format

Un CSV à **2 colonnes**, avec cet en-tête exact en ligne 1 :

```csv
Pattern1,Entity
10.10.0.0/24,HQ-Paris
10.20.0.0/24,Branch-Lyon
192.168.50.0/24,RemoteSite-A
```

| Colonne    | Rôle                                                                    |
|------------|-------------------------------------------------------------------------|
| `Pattern1` | Une plage **CIDR** (doit contenir un `/`), comparée à l'IP de la machine |
| `Entity`   | Le libellé du site affiché dans le Dashboard                            |

> Les noms de colonnes `Pattern1` et `Entity` sont **imposés** : c'est ce que le
> Dashboard va chercher dans le CSV. Ne les renomme pas.

Le matching se fait **uniquement par plage IP (CIDR)**, sur l'IP principale de la
machine. Les lignes dont `Pattern1` n'est pas un CIDR (pas de `/`) sont ignorées.

## Ordre d'évaluation

Les règles sont évaluées **dans l'ordre du fichier** : la **première plage qui
contient l'IP gagne** et définit le site. Mets donc les plages les plus spécifiques
(les plus petites) **en haut**, les plus larges (fallback) en bas.

Une machine dont l'IP ne tombe dans **aucune** plage est affichée avec le site
`Inconnu`. La colonne Site n'apparaît que si **au moins une plage CIDR valide** est
chargée.

## Préparer ton propre fichier

1. Copie `ip-ranges.example.csv` en `ip-ranges.csv` à la racine de ton `$SharePath`
   (le dossier que tu passes au Dashboard, ex. `\\SERVEUR\PCPulse$\`).
2. Remplace les lignes d'exemple par tes propres plages.
3. Relance le Dashboard — les sites remontent automatiquement.

Le fichier est lu **côté Dashboard uniquement** : pas besoin de redémarrer les
Collectors.

## Édition depuis Excel

Excel peut ouvrir et éditer le fichier directement. Pense à :

- Sauvegarder en **CSV UTF-8** (pour éviter les soucis d'accents)
- Garder l'en-tête `Pattern1,Entity` en **ligne 1**
- Les guillemets autour des valeurs sont optionnels (Excel peut en ajouter, ce n'est
  pas gênant)
