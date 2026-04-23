# ip-ranges.csv — Mapping IP / Hostname → Site

PC-Monitor enrichit chaque machine avec un libellé de **site** (bureau, agence, datacenter, etc.)
en appliquant les règles définies dans ce fichier CSV. Le Dashboard l'utilise pour :

- Afficher la colonne **Site** dans le tableau
- Filtrer les machines par site via le dropdown de la toolbar
- Grouper les agrégats si tu gères un parc multi-sites

Si le fichier est absent, la colonne Site est simplement masquée — l'outil fonctionne sans.

## Format

Le fichier est un CSV standard à 5 colonnes :

| Colonne      | Rôle                                                                            |
|--------------|---------------------------------------------------------------------------------|
| `Field1`     | Nom du champ à matcher : `primary_local_ip` ou `name`                           |
| `Pattern1`   | Valeur ou pattern à matcher (CIDR pour IP, wildcard `*` pour nom)              |
| `Field2`     | (Optionnel) Second champ pour un match combiné                                  |
| `Pattern2`   | (Optionnel) Second pattern. Si rempli, les DEUX critères doivent matcher (AND) |
| `Entity`     | Libellé du site affiché dans le Dashboard                                       |

## Champs supportés

- **`primary_local_ip`** — IP principale du PC (celle de l'interface active, Ethernet
  en priorité, puis Wi-Fi). Le pattern doit être au format CIDR (ex. `10.10.10.0/24`).
- **`name`** — Nom du PC (hostname). Le pattern supporte le wildcard `*`
  (ex. `SRV-*` matche `SRV-DB01`, `SRV-WEB02`, etc.).

## Exemples inclus dans le fichier fourni

```csv
"primary_local_ip","10.10.10.0/24","","","HQ-Paris"
```
→ Tous les PC ayant une IP dans `10.10.10.0/24` sont étiquetés "HQ-Paris".

```csv
"name","LAB-*","","","R&D-Lab"
```
→ Tous les PC dont le nom commence par `LAB-` sont étiquetés "R&D-Lab".

```csv
"primary_local_ip","10.99.99.0/24","name","VIP-*","Executive-Floor"
```
→ Un PC doit à la fois être dans le subnet `10.99.99.0/24` **ET** avoir un nom commençant
par `VIP-` pour être étiqueté "Executive-Floor".

## Ordre d'évaluation

Les règles sont évaluées **dans l'ordre du fichier**. La **première règle qui matche** gagne
et définit le site du PC. Si aucune règle ne matche, le champ Site reste vide et le PC
n'apparaît pas dans les filtres par site (mais reste visible dans le tableau).

**Astuce** : mets les règles les plus spécifiques en haut, les plus larges (fallback) en bas.

## Préparer ton propre fichier

1. Copie `ip-ranges.example.csv` en `ip-ranges.csv` dans ton `$SharePath`
   (par défaut `C:\PCPulse\`)
2. Remplace les lignes d'exemple par tes propres règles
3. Relance le Dashboard — les sites remontent automatiquement

Pas besoin de redémarrer le Collector : ce fichier est lu **côté Dashboard** uniquement.

## Édition depuis Excel

Excel peut ouvrir et éditer le fichier directement. Pense à :
- Sauvegarder en **CSV UTF-8** (pour éviter les soucis d'accents)
- Garder les **guillemets** autour des valeurs (format standard CSV)
- Garder l'en-tête `Field1,Pattern1,Field2,Pattern2,Entity` en ligne 1
