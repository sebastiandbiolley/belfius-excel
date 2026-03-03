# Bank Export

Nettoyage et catégorisation des exports CSV Belfius, puis export vers un tableau de bord Excel et un fichier CSV propre.

---

## Installation

```bash
pip install -r requirements.txt
```

---

## Utilisation rapide

1. Placez votre export CSV Belfius dans le dossier `input/`.
2. Vérifiez que `config/categories.json` existe (modifiez-le selon vos catégories).
3. Exécutez :

```bash
python export.py
```

**Résultats :**
- `output/clean_dashboard.xlsx` — Feuilles : Subcategories, Lists, Transactions (listes déroulantes), Summary (formules liées), Pivot (tableau croisé dynamique réel + graphique)
- `output/transactions_clean.csv`

**Tableau croisé dynamique :** Nécessite Excel et xlwings. Lignes = trimestre, mois. Colonnes = catégorie. Valeurs = somme des montants. Le graphique se met à jour automatiquement avec le pivot.

---

## Structure du projet

```
bank-export/
  input/           Exports CSV Belfius
  config/          categories.json (définition des catégories)
  output/          clean_dashboard.xlsx, transactions_clean.csv
  export.py        Script principal
```

---

## Documentation des fonctions

Le script `export.py` est organisé en quatre blocs : parsing CSV Belfius, moteur de règles, tableau de bord Excel, et flux principal.

### 1. Parsing CSV Belfius

#### `find_first_csv() -> Path | None`

Recherche le premier fichier CSV dans le dossier `input/`.

- Parcourt `input/` de manière triée
- Retourne le premier fichier avec extension `.csv`
- Retourne `None` si le dossier n’existe pas ou ne contient aucun CSV
- Utilisé pour localiser automatiquement le fichier à traiter

---

#### `parse_belfius_csv(path: Path) -> pd.DataFrame`

Parse un fichier CSV Belfius avec séparateur `;`, détection d’en-tête et encodage cp1252/latin1/utf-8.

**Détails :**
- Teste les encodages dans l’ordre : `cp1252`, `latin1`, `utf-8`
- Cherche la ligne d’en-tête commençant par `Compte;Date de comptabilisation`
- Lit le CSV avec `pandas.read_csv`, `sep=";"`, en sautant jusqu’à la ligne d’en-tête
- Toutes les colonnes sont chargées en texte (`dtype=str`)
- Lève une erreur si aucun encodage ne convient ou si l’en-tête Belfius est absent

---

#### `parse_european_amount(s: str) -> float`

Convertit un montant au format européen (ex. `1.234,56` ou `-750`) en nombre décimal.

**Détails :**
- Supprime les espaces
- Remplace le point millier (`.`) et la virgule décimale (`,`) pour obtenir un format exploitable par Python
- Retourne `0.0` si la chaîne est vide ou invalide

---

#### `parse_date(s: str) -> str | None`

Convertit une date au format `DD/MM/YYYY` en `YYYY-MM-DD`.

- Utilise `datetime.strptime` avec le format `%d/%m/%Y`
- Retourne `None` si la chaîne est vide ou invalide

---

#### `clean_and_normalize(df: pd.DataFrame) -> pd.DataFrame`

Nettoie et normalise les colonnes Belfius vers un schéma standard.

**Colonnes Belfius → colonnes internes :**
- `Compte` → `account_iban`
- `Date de comptabilisation` → `booking_date`
- `Numéro d'extrait` → `extract_nr`
- `Numéro de transaction` → `transaction_nr`
- `Compte contrepartie` → `counterparty_account`
- `Nom contrepartie contient` → `counterparty`
- `Rue et numéro` → `street`
- `Code postal et localité` → `city`
- `Transaction` → `description`
- `Date valeur` → `value_date`
- `Montant` → `amount`
- `Devise` → `currency`
- `BIC` → `bic`
- `Code pays` → `country_code`
- `Communications` → `communications`

**Colonnes calculées :**
- `raw_type` : début de la description (ex. `VIREMENT`, `PAIEMENT DEBITMASTERCARD`) via une regex
- `direction` : `"in"` si montant ≥ 0, sinon `"out"`
- `month` : format `YYYY-MM` à partir de `booking_date`
- `quarter` : format `YYYY-Q1/Q2/Q3/Q4` à partir du mois

Les dates sont passées dans `parse_date`, les montants dans `parse_european_amount`, et les chaînes sont nettoyées (`.fillna()`, `.strip()`).

---

### 2. Moteur de règles (catégorisation)

#### `_text(desc: str, counterparty: str) -> str`

Construit une chaîne unique pour la recherche (description + contrepartie en majuscules).

Permet de vérifier les mots-clés indépendamment de la casse et du champ (description ou contrepartie).

---

#### `apply_rules(row: pd.Series, categories: dict) -> tuple[str, str]`

Assigne une catégorie et une sous-catégorie à une transaction selon des règles basées sur des mots-clés.

**Paramètres :**
- `row` : une ligne du DataFrame (transaction)
- `categories` : dictionnaire `{catégorie: [sous-catégories]}` chargé depuis `categories.json`

**Logique :**
- Combine `description` et `counterparty` via `_text()` pour la recherche
- Parcourt une liste de règles ordonnées : `(mots_clés, catégorie, sous_catégorie)`
- Si un mot-clé est trouvé dans la chaîne ou dans `raw_type` :
  - Vérifie que la catégorie et la sous-catégorie existent dans `categories`
  - Retourne `(catégorie, sous_catégorie)` ou `(catégorie, "")` si la sous-catégorie n’existe pas
- Si aucune règle ne correspond : `("", "")` — à remplir manuellement dans Excel

**Ordre important :** Les règles Administration et Capital & Financing sont placées avant Marketing pour éviter des faux positifs (ex. `ROOFWANDER` dans des références notariales/capital).

**Exemples de règles :**
- Frais bancaires, fiduciaire, notaire → Administration
- Investissement, versamento, capitale, ROOFWANDER ACCOUNT → Capital & Financing / Founder Contributions
- Google Ads, LeBonCoin → Marketing / Paid Advertising
- Magnis Group, Cmonevent, etc. → Marketing / Events & Offline Marketing
- Render, ShareTribe, Cleverbridge → Technology
- VERSEMENT DE → Revenue / Other Operating Revenue

---

### 3. Tableau de bord Excel

#### `_cat_to_range_name(cat: str) -> str`

Transforme un nom de catégorie en nom de plage Excel valide (sans espaces, `&`, parenthèses, etc.).

Ex. `Capital & Financing` → `Capital__Financing` (pour les plages nommées Excel).

---

#### `_add_excel_pivot_table(filepath: str, n_tx: int) -> None`

Crée un tableau croisé dynamique réel et un graphique via xlwings (nécessite Excel installé).

**Détails :**
- Lance Excel en arrière-plan (`xlwings.App(visible=False)`)
- Ouvre le fichier sauvegardé et accède aux feuilles `Transactions` et `Pivot`
- Crée un pivot cache à partir de la plage de données Transactions
- Configure le pivot :
  - **Lignes :** `quarter`, puis `month`
  - **Colonnes :** `category`
  - **Valeurs :** somme de `amount`
- Crée un graphique en colonnes groupées (`xlColumnClustered`) lié au pivot
- Sauvegarde et ferme Excel
- En cas d’erreur (ex. Excel absent) : message informatif et fermeture propre

---

#### `create_excel_dashboard(df: pd.DataFrame, categories: dict) -> None`

Génère le fichier `clean_dashboard.xlsx` avec cinq feuilles.

**Feuille Subcategories :**
- Une colonne par catégorie, avec les sous-catégories en lignes
- Crée une plage nommée Excel par catégorie (ex. `Revenue`, `Marketing`) pour les listes déroulantes dépendantes

**Feuille Lists :**
- Vue de référence : catégorie → liste des sous-catégories (depuis `config`)

**Feuille Transactions :**
- Toutes les colonnes du DataFrame + `category`, `subcategory`
- **Liste déroulante catégorie :** liste des catégories du config
- **Liste déroulante sous-catégorie :** dépendante de la catégorie sélectionnée (`INDIRECT`, `SUBSTITUTE`, `ADDRESS`, `ROW`)
- Les lignes sans correspondance de règles restent vides pour sélection manuelle

**Feuille Summary :**
- Nombre total de transactions
- Total entrées : `=SUMIF(..., ">0")`
- Total sorties : `=ABS(SUMIF(..., "<0"))`
- Net : `=SUM(...)`
- Tableau par catégorie/sous-catégorie : une ligne par combinaison config + `(choose)` pour les non catégorisées, avec formules `SUMIFS` liées à Transactions

**Feuille Pivot :**
- Texte d’accroche puis création du véritable tableau croisé via `_add_excel_pivot_table`

---

### 4. Flux principal

#### `main() -> int`

Orchestre l’exécution complète.

1. **Vérifications :**
   - Recherche un CSV dans `input/` via `find_first_csv()`
   - Vérifie l’existence de `config/categories.json`
   - Quitte avec code 1 si l’un des deux manque

2. **Chargement :** Lit `categories.json` et le CSV via `parse_belfius_csv()`

3. **Traitement :**
   - `clean_and_normalize()` sur le brut
   - `apply_rules()` sur chaque ligne pour remplir `category` et `subcategory`

4. **Export :**
   - Sauvegarde du DataFrame en CSV propre (`output/transactions_clean.csv`)
   - Création du tableau de bord Excel via `create_excel_dashboard()`

5. Retourne `0` en cas de succès, `1` en cas d’erreur.

---

## Format de `config/categories.json`

Structure attendue :

```json
{
  "Catégorie1": ["SousCat1", "SousCat2", "..."],
  "Catégorie2": ["SousCatA", "SousCatB", "..."]
}
```

Les noms de catégories et sous-catégories doivent correspondre aux règles dans `apply_rules()` pour que l’auto-catégorisation fonctionne. Les transactions non matchées restent à catégoriser manuellement dans Excel.
