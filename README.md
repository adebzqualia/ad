# POPS Check

POPS Check compare les classeurs Excel envoyés aux pays avec les classeurs retournés et génère des rapports de contrôle opérationnel. Il recherche les causes racines structurelles, puis compare les formules, les dépendances entre feuilles et les valeurs explicitement contrôlées après réalignement des coordonnées.

> [!IMPORTANT]
> POPS Check contrôle la structure du template. Il ne valide pas la justesse d'un budget, d'une prévision ou d'un autre contenu métier.

## Fonctionnalités

- appariement automatique des fichiers de référence et des fichiers reçus ;
- détection des feuilles ajoutées, supprimées ou réordonnées ;
- détection des lignes et colonnes ajoutées, supprimées ou déplacées ;
- comparaison sémantique des formules après réalignement, avec distinction entre formule ajoutée, supprimée, remplacée par une valeur, réellement modifiée ou contenant `#REF!` ;
- inventaire des dépendances inter-feuilles et des liens externes statiques, avec rattachement des impacts confirmés, probables ou possibles à la cause racine ;
- comparaison configurable des valeurs dans des plages contrôlées ou critiques, tandis que les valeurs des zones éditables restent autorisées ;
- prise en compte des formules, libellés stables, styles, fusions, tables, validations et dimensions comme indices structurels ;
- actions recommandées et message pays directement exploitables dans le rapport ;
- traitement de tout le lot même lorsqu'un classeur individuel est absent, illisible ou invalide ;
- rapport HTML global, rapport détaillé par pays et export JSON structuré ;
- exécution entièrement locale, sans Excel et sans service externe.

## Comment les faux positifs sont limités

Comparer seulement le nombre de lignes et de colonnes ne suffit pas : une ligne peut être supprimée puis remplacée ailleurs, tandis qu'une saisie normale peut remplir de nombreuses cellules auparavant vides. POPS Check construit donc une signature structurelle pour chaque ligne et chaque colonne, puis aligne les séquences attendue et reçue.

Les principaux garde-fous sont les suivants :

1. Les nombres, dates et autres valeurs métier ne sont pas utilisés pour identifier une ligne ou une colonne.
2. Un texte ne sert d'ancre que s'il existe dans les deux classeurs et qu'il ne se trouve pas dans une plage déclarée modifiable.
3. Les formules sont normalisées relativement à leur cellule. Par exemple, `=B5+$C$2` déplacée avec sa ligne et devenue `=B8+$C$2` conserve la même topologie.
4. Les styles ne sont qu'un signal faible. Les fusions, tables, validations de données et dimensions apportent des indices complémentaires.
5. Les lignes et colonnes sont alignées en deux passes et chaque passe compare uniquement l'intersection déjà appariée sur l'autre axe. Une colonne supprimée ne contamine donc pas la signature de toutes les lignes, et inversement.
6. Seules les signatures exactes, uniques et suffisamment informatives deviennent des ancres fortes. Les blocs restants sont alignés par un script d'édition monotone qui préfère un ajout ou une suppression à une série de mauvaises correspondances.
7. Les déplacements sont recherchés après l'alignement monotone afin qu'une paire croisée ne vole pas une ancre à une insertion ou une suppression ordinaire.
8. Une suppression suivie d'un ajout à la même position reste visible comme deux causes structurelles lorsque leurs signatures ne correspondent pas ; elle n'est plus assimilée automatiquement à un élément inchangé.
9. Les ajouts ou suppressions contigus sont regroupés en une seule anomalie de plage, avec un impact égal au nombre de lignes ou colonnes concernées.
10. Si plusieurs lignes ou colonnes sont structurellement indiscernables, POPS Check préfère émettre un avertissement d'ambiguïté plutôt que d'inventer une position.
11. Par défaut, une valeur isolée saisie au-delà du template ne suffit pas à créer une nouvelle ligne ou colonne structurelle.

Ces règles privilégient la réduction des faux positifs. Elles ne constituent pas une preuve mathématique qu'aucune modification n'a eu lieu ; les [limites fondamentales](#limites-fondamentales) restent applicables.

## Formats et prérequis

- Windows 10 ou Windows 11 ;
- Python 3.11 ou version ultérieure ;
- `openpyxl` 3.1 ou version ultérieure et antérieure à 4, installé automatiquement avec le projet ;
- classeurs `.xlsx`, `.xlsm`, `.xltx` ou `.xltm`.

Les anciens formats binaires `.xls` et `.xlt` ne sont pas pris en charge. Les classeurs chiffrés ou protégés par mot de passe doivent être déchiffrés avant l'analyse.

Excel n'a pas besoin d'être installé. Les macros éventuellement présentes dans un fichier `.xlsm` ou `.xltm` ne sont jamais exécutées.

## Installation sous Windows

Ouvrez PowerShell dans le dossier du projet :

```powershell
Set-Location 'C:\chemin\vers\POPS_AD'
```

Créez un environnement Python isolé :

```powershell
python -m venv .venv
```

Si votre installation utilise le lanceur Python Windows, la commande équivalente est :

```powershell
py -3.11 -m venv .venv
```

Installez POPS Check et ses dépendances :

```powershell
.\.venv\Scripts\python.exe -m pip install -e .
```

Cette installation rend disponibles les deux formes de commande suivantes :

```powershell
.\.venv\Scripts\python.exe -m popscheck --help
.\.venv\Scripts\popscheck.exe --help
```

Il n'est pas nécessaire d'activer l'environnement virtuel. Cela évite notamment les erreurs de stratégie d'exécution PowerShell liées à `Activate.ps1`.

## Préparer les fichiers

Par défaut, POPS Check utilise cette arborescence :

```text
data/
├── sent/
│   ├── France.xlsx
│   ├── Germany.xlsx
│   └── Spain.xlsm
└── received/
    ├── France.xlsx
    ├── Germany.xlsx
    └── Spain.xlsm
```

Créez les dossiers si nécessaire :

```powershell
New-Item -ItemType Directory -Force `
  -Path .\data\sent, .\data\received, .\rapports |
  Out-Null
```

Placez dans `data\sent` les fichiers effectivement envoyés aux pays, puis leurs retours dans `data\received`. POPS Check ouvre les fichiers en lecture et ne les modifie pas.

## Utilisation rapide

Depuis la racine du projet :

```powershell
.\.venv\Scripts\python.exe -m popscheck
```

Les valeurs par défaut sont :

- références : `data\sent` ;
- fichiers reçus : `data\received` ;
- rapports : `rapports` ;
- configuration : `popscheck.toml` s'il existe dans le dossier courant, sinon les valeurs internes par défaut.

Une invocation entièrement explicite est préférable pour une exécution automatisée :

```powershell
.\.venv\Scripts\python.exe -m popscheck `
  --sent .\data\sent `
  --received .\data\received `
  --reports .\rapports `
  --config .\popscheck.toml
```

Après installation, la même analyse peut être lancée avec l'exécutable :

```powershell
.\.venv\Scripts\popscheck.exe `
  --sent .\data\sent `
  --received .\data\received `
  --reports .\rapports `
  --config .\popscheck.toml
```

Les chemins relatifs sont résolus depuis le dossier courant, et non depuis le dossier contenant le fichier TOML.

## Options de la ligne de commande

| Option | Valeur par défaut | Description |
| --- | --- | --- |
| `--sent CHEMIN` | `data/sent` | Dossier des fichiers de référence envoyés aux pays. |
| `--received CHEMIN` | `data/received` | Dossier des fichiers retournés. |
| `--reports CHEMIN` | `rapports` | Dossier de sortie des rapports HTML et JSON. |
| `--config FICHIER` | `popscheck.toml` s'il existe | Fichier TOML à charger. Le chemin fourni explicitement doit exister. |
| `--fail-on-issues` | désactivé | Retourne le code 1 si une anomalie, un fichier manquant, un fichier sans référence ou une erreur métier est rencontré. |
| `--verbose` | désactivé | Affiche la trace Python complète en cas d'erreur fatale. |
| `--version` | — | Affiche la version et quitte. |
| `--help` | — | Affiche l'aide et quitte. |

## Règles d'appariement des fichiers

L'appariement utilise le chemin relatif sans extension, normalisé en Unicode NFC. La comparaison ne tient pas compte de la casse par défaut.

Ainsi, dans les dossiers non récursifs :

- `sent\France.xlsx` correspond à `received\france.xlsx` ;
- `sent\Spain.xlsx` peut correspondre à `received\Spain.xlsm` ;
- `France.xlsx` et `France.xlsm` présents du même côté créent une collision, car ils ont la même clé sans extension.

POPS Check ne choisit jamais arbitrairement entre plusieurs candidats. Une collision est affichée avec le statut **Erreur**. Les fichiers temporaires créés par Excel et dont le nom commence par `~$` sont ignorés.

Lorsque `recursive = true`, les sous-dossiers font partie de la clé. Par exemple, `Europe\France.xlsx` correspond à `Europe\France.xlsm`, mais pas à `France.xlsx` placé directement à la racine.

## Configuration TOML

Le fichier [`popscheck.toml`](popscheck.toml) fourni avec le projet contient des valeurs prudentes adaptées à une première analyse. Sans option `--config`, il est chargé automatiquement uniquement s'il existe dans le dossier courant.

Exemple complet des paramètres reconnus :

```toml
[analysis]
extensions = [".xlsx", ".xlsm", ".xltx", ".xltm"]
recursive = false
case_sensitive_names = false
strict_sheet_order = true
use_stable_text_anchors = true

# Seuils de comparaison, compris entre 0 et 1.
min_axis_similarity = 0.62
move_min_similarity = 0.84
ambiguity_margin = 0.08

# Limites de sûreté.
max_cells_per_sheet = 500000
max_rows = 50000
max_columns = 2000
max_style_gap = 25

# Comportement anti-faux-positifs.
detect_value_only_expansion = false
report_ambiguities = true

# Contrôles sémantiques et limites du graphe de dépendances.
compare_formulas = true
compare_dependencies = true
compare_controlled_values = true
max_dependency_depth = 8
max_dependency_samples = 20
numeric_absolute_tolerance = 0.0
numeric_relative_tolerance = 0.0

[[sheet_rules]]
pattern = "Budget"
editable_ranges = ["D8:N100"]
formula_allowed_ranges = []
controlled_ranges = ["B2:B6"]
critical_ranges = ["B2:B2"]
monitored_ranges = ["A1:N120"]

[[sheet_rules]]
pattern = "Notes*"
ignore = true
```

Les clés inconnues, plages invalides et seuils hors limites arrêtent la commande avec le code 2 afin d'éviter une analyse fondée sur une configuration mal comprise.

### Paramètres `[analysis]`

| Paramètre | Rôle |
| --- | --- |
| `extensions` | Liste des extensions recherchées. Le point initial est ajouté automatiquement s'il manque. |
| `recursive` | Recherche également dans les sous-dossiers et inclut leur chemin dans la clé d'appariement. |
| `case_sensitive_names` | Rend l'appariement sensible à la casse. La valeur par défaut est `false`. |
| `strict_sheet_order` | Signale une modification de l'ordre relatif des feuilles communes. |
| `use_stable_text_anchors` | Utilise comme ancres les libellés textuels communs aux deux classeurs. |
| `min_axis_similarity` | Similarité minimale utilisée pour aligner prudemment les lignes ou colonnes restantes. |
| `move_min_similarity` | Similarité minimale, supérieure ou égale à la précédente, requise pour annoncer un déplacement. |
| `ambiguity_margin` | Écart minimal entre le meilleur candidat et les autres avant de considérer un appariement comme discriminant. |
| `max_cells_per_sheet` | Nombre maximal de cellules matérialisées par feuille avant de classer le classeur en erreur. |
| `max_rows` | Dernière ligne analysable. Les cellules au-delà sont ignorées avec un avertissement. |
| `max_columns` | Dernière colonne analysable. Les cellules au-delà sont ignorées avec un avertissement. |
| `max_style_gap` | Distance maximale permettant à des styles ou dimensions isolés de prolonger la zone structurelle utile. |
| `detect_value_only_expansion` | Si `true`, une simple valeur au-delà du template peut être considérée comme une extension de ligne ou colonne. |
| `report_ambiguities` | Ajoute aux avertissements les changements d'ordre qui ne peuvent pas être localisés avec une confiance suffisante. |
| `compare_formulas` | Active la comparaison des formules après réalignement des lignes et colonnes. |
| `compare_dependencies` | Active les dépendances inter-feuilles, liens externes et conséquences rattachées aux causes structurelles. Fonctionne indépendamment de `compare_formulas`. |
| `compare_controlled_values` | Compare les valeurs littérales uniquement dans les plages contrôlées ou critiques. |
| `max_dependency_depth` | Profondeur maximale de propagation statique vers les feuilles aval. |
| `max_dependency_samples` | Nombre maximal de localisations d'exemple conservées par groupe d'impact. |
| `numeric_absolute_tolerance` | Tolérance absolue pour les nombres des zones contrôlées. |
| `numeric_relative_tolerance` | Tolérance relative pour les nombres des zones contrôlées. |

Les seuils par défaut sont volontairement conservateurs. Ne les modifiez qu'après validation sur un échantillon représentatif de fichiers POPS. En particulier, abaisser `move_min_similarity` augmente le risque d'annoncer de faux déplacements.

### Règles `[[sheet_rules]]`

Chaque règle s'applique aux feuilles dont le nom correspond à `pattern`. La comparaison est insensible à la casse et accepte des jokers tels que `*` et `?`.

- `editable_ranges` autorise les changements de **valeurs** saisis par le pays. Une formule ou un lien externe ajouté dans cette zone reste contrôlé.
- `formula_allowed_ranges` autorise explicitement les formules et leurs dépendances à varier sans anomalie.
- `controlled_ranges` compare les valeurs littérales après réalignement des coordonnées.
- `critical_ranges` applique les mêmes contrôles avec une gravité `error` et rend critique la suppression d'une ligne ou colonne qui traverse la plage.
- `monitored_ranges` limite l'analyse interne de la feuille aux plages indiquées. Sans cette clé, toute la zone utile est analysée.
- `ignore = true` désactive la comparaison des lignes et colonnes internes de la feuille.

Une feuille ignorée reste néanmoins contrôlée au niveau du classeur : son absence, son ajout ou son changement de position peut toujours être signalé.

Plusieurs règles peuvent correspondre à la même feuille. Leurs plages sont cumulées ; les zones contrôlées ou critiques ont priorité sur une zone éditable ou une zone de formules autorisées. Pour `ignore`, la dernière règle correspondante l'emporte. Placez donc les règles générales avant les règles particulières.

## Rapports générés

Une exécution réussie crée ou met à jour :

```text
rapports/
├── index.html
├── France.html
├── Germany.html
└── resultats.json
```

Les noms des pages pays sont transformés en noms ASCII sûrs. En cas de collision, un suffixe dérivé d'un hash les distingue.

Le rapport global affiche :

- le nombre total de dossiers issus de l'union de `sent` et `received` ;
- les dossiers conformes ;
- les dossiers contenant des anomalies ;
- les fichiers manquants ou sans référence ;
- les erreurs d'analyse ;
- le nombre total d'anomalies structurelles.

Chaque ligne mène à un rapport détaillé organisé selon la hiérarchie **fichier → feuille → cause racine → conséquences**. Une carte indique l'identifiant stable, le code, la localisation Excel canonique (`'Forecast'!F:F`, `'Budget'!25:28` ou `'Summary'!B12`), les états attendu et reçu, la confiance, les impacts confirmés/probables/possibles, l'action recommandée et un message prêt à transmettre au pays. Une synthèse des dimensions attendues et reçues apparaît pour les feuilles dont le nombre de lignes ou de colonnes diffère.

Le fichier `resultats.json` contient les mêmes résultats sous une forme exploitable par un autre outil. `root_cause_count` compte les cartes de causes racines, `total_anomalies` compte les éléments directement concernés, `root_counts_by_code` compte les causes, `counts_by_code` compte les éléments, et `impact_summary` sépare les objets impactés confirmés, probables et possibles. `validation_level` expose un niveau normalisé `ok`, `warning` ou `error`. Les conséquences ne sont pas recopiées comme anomalies indépendantes.

Ouvrez le rapport global sous Windows avec :

```powershell
Start-Process (Resolve-Path .\rapports\index.html)
```

Les fichiers de l'exécution courante sont remplacés de manière atomique, mais les anciennes pages pays qui ne sont plus produites ne sont pas automatiquement supprimées. Utilisez un dossier de rapports dédié par exécution si leur conservation peut créer une confusion.

## Comprendre les statuts

| Statut | Signification |
| --- | --- |
| **Conforme** | Aucune cause racine détectée. Des avertissements non bloquants peuvent néanmoins être présents. |
| **Anomalies** | Au moins une cause racine de structure, formule, dépendance ou valeur contrôlée a été détectée. |
| **Fichier manquant** | Une référence existe dans `sent`, mais aucun retour correspondant n'existe dans `received`. |
| **Sans référence** | Un fichier reçu existe, mais aucune référence correspondante n'existe dans `sent`. |
| **Erreur** | Le classeur est illisible, dépasse une limite de sûreté, provoque une erreur d'analyse ou entre en collision avec un autre nom. |

Le rapport distingue en plus le niveau de validation : **OK**, **Avertissement** ou **Erreur structurelle**. Le statut machine historique reste inchangé pour préserver la compatibilité, tandis que `validation_level` expose ce niveau dans le JSON.

Une modification de visibilité de feuille et une localisation structurelle ambiguë sont actuellement des avertissements. Elles ne sont pas ajoutées au nombre d'anomalies et, à elles seules, ne déclenchent pas le code 1 avec `--fail-on-issues`.

## Codes de sortie

| Code | Signification |
| ---: | --- |
| `0` | L'analyse et les rapports ont été générés. Sans `--fail-on-issues`, ce code est également utilisé lorsque des anomalies métier existent. |
| `1` | `--fail-on-issues` est actif et au moins un dossier n'est pas conforme : anomalie, fichier manquant, fichier sans référence ou erreur individuelle. |
| `2` | Erreur fatale de ligne de commande, configuration, dossier d'entrée ou écriture des rapports. |

Exemple PowerShell pour une tâche automatisée :

```powershell
& .\.venv\Scripts\python.exe -m popscheck `
  --sent .\data\sent `
  --received .\data\received `
  --reports .\rapports `
  --config .\popscheck.toml `
  --fail-on-issues

switch ($LASTEXITCODE) {
  0 { Write-Host 'Analyse terminée : tous les dossiers sont conformes.' }
  1 { Write-Warning 'Analyse terminée : des problèmes ont été détectés.' }
  2 { Write-Error 'POPS Check n’a pas pu terminer l’analyse.' }
}
```

Une erreur limitée à un seul classeur devient un résultat **Erreur** dans les rapports : le reste du lot continue d'être traité. Le code 2 est réservé à une impossibilité de terminer le traitement global.

## Périmètre actuel

POPS Check contrôle actuellement :

- présence, absence, type et ordre relatif des feuilles ;
- ajout, suppression et déplacement de lignes ou colonnes ;
- ajout, suppression, remplacement par une valeur et modification d'une formule dans une cellule survivante, après projection des coordonnées reçues vers le template ;
- références statiques entre feuilles, références `#REF!`, feuilles référencées manquantes et liens externes présents dans les formules ;
- noms définis et tables utilisés par une formule lorsqu'ils disparaissent ou que leur définition change ;
- valeurs littérales des seules plages déclarées dans `controlled_ranges` ou `critical_ranges` ;
- état visible/masqué d'une feuille sous forme d'avertissement ;
- libellés stables, styles, dimensions, plages fusionnées, validations et tables comme indices d'appariement structurel.

L'analyse des formules est volontairement conservatrice : elle normalise les références et ignore les différences d'espacement, mais ce n'est ni le moteur de calcul d'Excel ni une preuve d'équivalence algébrique. Elle ne recalcule pas les résultats. Des parenthèses redondantes ou un opérateur unaire peuvent donc rester visibles comme modification de logique.

Les points suivants nécessitent encore une vérification humaine ou un outil Excel spécialisé :

- résolution certaine des dépendances dynamiques construites avec `INDIRECT`, `OFFSET` ou du texte ; elles sont signalées comme impacts possibles lorsque la feuille cible peut être identifiée ;
- propagation cellule par cellule à l'intérieur d'une même feuille : les consommateurs directs sont localisés, puis la propagation statique aval est synthétisée au niveau des feuilles ;
- renommage d'une feuille : sans règle d'alias ou identifiant métier, il est rendu prudemment comme une suppression et un ajout ;
- équivalence algébrique complète, résultat recalculé, cache de calcul et comportement des fonctions volatiles ;
- audit indépendant des styles, formats, fusions, validations, tables et noms qui ne sont référencés par aucune formule ;
- connexions de données, Power Query, mises en forme conditionnelles, graphiques, tableaux croisés dynamiques et segments ;
- contenu et comportement des macros ;
- validité métier des valeurs saisies hors des plages contrôlées configurées.

Les références 3-D et les dépendances opaques sont conservées comme avertissements ou impacts possibles : le rapport ne les présente jamais comme une causalité certaine.

## Limites fondamentales

La position exacte d'une ligne ou d'une colonne ne peut pas toujours être prouvée. Si plusieurs éléments ont les mêmes formules, libellés, styles, validations et contexte, leur permutation produit une structure observable identique. Aucun programme ne peut alors déterminer de façon fiable quelle occurrence a été déplacée sans identifiant stable ou règle métier supplémentaire.

Dans ce cas, POPS Check adopte une stratégie conservatrice :

- il n'annonce pas un déplacement précis avec une confiance insuffisante ;
- il peut signaler une ambiguïté dans les avertissements ;
- un changement de cardinalité reste détectable, mais sa position exacte peut rester incertaine.

Pour améliorer la précision sur un template homogène :

1. déclarez les zones de saisie dans `editable_ranges` ;
2. limitez l'analyse aux zones pertinentes avec `monitored_ranges` ;
3. conservez des libellés, formules ou formats stables servant d'ancres ;
4. validez les seuils sur des retours réels avant de les modifier.

Avec `detect_value_only_expansion = false`, une ligne ou colonne située hors de la zone attendue et ne contenant que de nouvelles valeurs littérales peut volontairement être ignorée. Activez cette option si le template interdit aussi ce type d'extension, en acceptant un risque supérieur de faux positifs.

## Confidentialité et sécurité

- L'analyse est locale ; POPS Check n'envoie aucun classeur ni résultat sur Internet.
- Les fichiers d'entrée sont lus mais jamais enregistrés ou modifiés.
- Les macros VBA ne sont pas exécutées.
- Les formules et liens externes sont lus comme des éléments du classeur ; POPS Check ne contacte pas leurs sources.
- Les rapports contiennent les noms de pays, noms de feuilles, détails d'anomalies et chemins absolus des fichiers comparés.

Traitez donc le dossier `rapports` comme une donnée potentiellement confidentielle. Vérifiez son contenu avant de le partager et appliquez les règles de conservation habituelles de votre organisation.

## Résolution des problèmes

### `python` n'est pas reconnu

Installez Python 3.11 ou ultérieur depuis une source approuvée et activez l'option d'ajout à `PATH`, ou utilisez le lanceur Windows :

```powershell
py -3.11 --version
```

### `popscheck.toml` n'est pas chargé

La détection automatique recherche ce fichier dans le dossier courant. Indiquez son chemin explicitement si vous lancez la commande depuis un autre emplacement :

```powershell
.\.venv\Scripts\python.exe -m popscheck --config 'C:\chemin\vers\popscheck.toml'
```

### Un pays n'est pas apparié

Vérifiez le radical du nom, le sous-dossier lorsque le mode récursif est activé et les éventuelles collisions entre extensions. Consultez le rapport détaillé : POPS Check indique les références absentes et les doublons normalisés.

### Un classeur est indiqué comme invalide

Fermez-le dans Excel, vérifiez qu'il n'est pas chiffré, qu'il possède une extension prise en charge et qu'il s'ouvre normalement. Une corruption ZIP/XML, une protection par mot de passe ou un dépassement des limites configurées produit également ce statut.

### Une saisie normale crée trop d'anomalies

Commencez par déclarer précisément ses plages dans `editable_ranges` et, si possible, la zone utile dans `monitored_ranges`. Conservez les seuils par défaut pendant ce diagnostic. Les avertissements d'ambiguïté expliquent les cas où les signatures ne sont pas suffisamment discriminantes.

### Une valeur ajoutée hors du template n'est pas signalée

C'est le comportement anti-faux-positifs par défaut. Utilisez :

```toml
[analysis]
detect_value_only_expansion = true
```

après validation sur un échantillon représentatif.

### Le rapport affiche encore une ancienne page pays

L'index ne référence que l'exécution courante, mais une ancienne page HTML peut rester physiquement dans le même dossier. Utilisez un nouveau dossier `--reports` pour chaque campagne ou nettoyez consciemment le dossier avant une nouvelle exécution.

## Développement et tests

La suite automatisée repose sur `unittest` et génère ses classeurs dans des dossiers temporaires. Depuis la racine du projet :

```powershell
.\.venv\Scripts\python.exe -m unittest discover `
  -s .\tests `
  -p "test_*.py" `
  -v
```

Les tests couvrent notamment les saisies normales, les ajouts, suppressions et déplacements d'axes, les remplacements à cardinalité constante, les modifications simultanées de lignes et colonnes, les suppressions physiques laissant des dimensions Excel résiduelles, les limites de `monitored_ranges`, les formules décalées, références cassées, dépendances, liens externes, noms définis, tables, valeurs contrôlées, l'ordre des feuilles, les fichiers absents ou corrompus, l'immuabilité des fichiers d'entrée, les rapports HTML et les codes de sortie de la CLI.
