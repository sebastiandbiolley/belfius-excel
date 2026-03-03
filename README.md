# Bank Export

Clean and categorize Belfius bank CSV exports, then export to an Excel dashboard and a clean CSV file.

---

## Installation

```bash
pip install -r requirements.txt
```

---

## Quick Start

1. Place your Belfius CSV export in the `input/` folder.
2. Ensure `config/categories.json` exists (edit it to match your categories).
3. Run:

```bash
python export.py
```

**Output:**
- `output/clean_dashboard.xlsx` — Sheets: Subcategories, Lists, Transactions (dropdowns), Summary (linked formulas), Pivot (real pivot table + chart)
- `output/transactions_clean.csv`

**Pivot table:** Requires Excel and xlwings. Rows = quarter, month. Columns = category. Values = sum of amounts. The chart updates automatically with the pivot.

---

## Project Structure

```
bank-export/
  input/           Belfius CSV exports
  config/          categories.json (category definitions)
  output/           clean_dashboard.xlsx, transactions_clean.csv
  export.py        Main script
```

---

## Function Documentation

The `export.py` script is organized into four blocks: Belfius CSV parsing, rule engine, Excel dashboard, and main flow.

### 1. Belfius CSV Parsing

#### `find_first_csv() -> Path | None`

Finds the first CSV file in the `input/` folder.

- Iterates over `input/` in sorted order
- Returns the first file with `.csv` extension
- Returns `None` if the folder does not exist or contains no CSV
- Used to automatically locate the file to process

---

#### `parse_belfius_csv(path: Path) -> pd.DataFrame`

Parses a Belfius CSV file with semicolon separator, header detection, and cp1252/latin1/utf-8 encoding.

**Details:**
- Tries encodings in order: `cp1252`, `latin1`, `utf-8`
- Finds the header row starting with `Compte;Date de comptabilisation`
- Reads the CSV with `pandas.read_csv`, `sep=";"`, skipping rows until the header
- All columns are loaded as text (`dtype=str`)
- Raises an error if no encoding works or the Belfius header is missing

---

#### `parse_european_amount(s: str) -> float`

Converts a European-format amount (e.g. `1.234,56` or `-750`) to a decimal number.

**Details:**
- Strips spaces
- Replaces thousand separator (`.`) and decimal comma (`,`) to get a Python-ready format
- Returns `0.0` if the string is empty or invalid

---

#### `parse_date(s: str) -> str | None`

Converts a date from `DD/MM/YYYY` format to `YYYY-MM-DD`.

- Uses `datetime.strptime` with format `%d/%m/%Y`
- Returns `None` if the string is empty or invalid

---

#### `clean_and_normalize(df: pd.DataFrame) -> pd.DataFrame`

Cleans and normalizes Belfius columns into a standard schema.

**Belfius columns → internal columns:**
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

**Calculated columns:**
- `raw_type`: Start of description (e.g. `VIREMENT`, `PAIEMENT DEBITMASTERCARD`) via regex
- `direction`: `"in"` if amount ≥ 0, otherwise `"out"`
- `month`: `YYYY-MM` format from `booking_date`
- `quarter`: `YYYY-Q1/Q2/Q3/Q4` format from month

Dates go through `parse_date`, amounts through `parse_european_amount`, and strings are cleaned (`.fillna()`, `.strip()`).

---

### 2. Rule Engine (Categorization)

#### `_text(desc: str, counterparty: str) -> str`

Builds a single searchable string (description + counterparty in uppercase).

Used to check keywords case-insensitively across both description and counterparty fields.

---

#### `apply_rules(row: pd.Series, categories: dict) -> tuple[str, str]`

Assigns a category and subcategory to a transaction based on keyword rules.

**Parameters:**
- `row`: A DataFrame row (transaction)
- `categories`: Dictionary `{category: [subcategories]}` loaded from `categories.json`

**Logic:**
- Combines `description` and `counterparty` via `_text()` for search
- Iterates through an ordered list of rules: `(keywords, category, subcategory)`
- If a keyword is found in the string or in `raw_type`:
  - Verifies the category and subcategory exist in `categories`
  - Returns `(category, subcategory)` or `(category, "")` if the subcategory does not exist
- If no rule matches: `("", "")` — to be filled manually in Excel

**Order matters:** Administration and Capital & Financing rules come before Marketing to avoid false positives (e.g. `ROOFWANDER` in notary/capital references).

**Example rules:**
- Banking fees, fiduciary, notary → Administration
- Investissement, versamento, capitale, ROOFWANDER ACCOUNT → Capital & Financing / Founder Contributions
- Google Ads, LeBonCoin → Marketing / Paid Advertising
- Magnis Group, Cmonevent, etc. → Marketing / Events & Offline Marketing
- Render, ShareTribe, Cleverbridge → Technology
- VERSEMENT DE → Revenue / Other Operating Revenue

---

### 3. Excel Dashboard

#### `_cat_to_range_name(cat: str) -> str`

Converts a category name to a valid Excel range name (no spaces, `&`, parentheses, etc.).

E.g. `Capital & Financing` → `Capital__Financing` (for Excel named ranges).

---

#### `_add_excel_pivot_table(filepath: str, n_tx: int) -> None`

Creates a real pivot table and chart via xlwings (requires Excel installed).

**Details:**
- Launches Excel in the background (`xlwings.App(visible=False)`)
- Opens the saved file and accesses `Transactions` and `Pivot` sheets
- Creates a pivot cache from the Transactions data range
- Configures the pivot:
  - **Rows:** `quarter`, then `month`
  - **Columns:** `category`
  - **Values:** sum of `amount`
- Creates a clustered column chart (`xlColumnClustered`) linked to the pivot
- Saves and closes Excel
- On error (e.g. Excel not installed): informative message and clean shutdown

---

#### `create_excel_dashboard(df: pd.DataFrame, categories: dict) -> None`

Generates `clean_dashboard.xlsx` with five sheets.

**Subcategories sheet:**
- One column per category, with subcategories in rows
- Creates an Excel named range per category (e.g. `Revenue`, `Marketing`) for dependent dropdowns

**Lists sheet:**
- Reference view: category → list of subcategories (from config)

**Transactions sheet:**
- All DataFrame columns + `category`, `subcategory`
- **Category dropdown:** list of categories from config
- **Subcategory dropdown:** depends on selected category (`INDIRECT`, `SUBSTITUTE`, `ADDRESS`, `ROW`)
- Rows with no rule match stay empty for manual selection

**Summary sheet:**
- Total transaction count
- Total inflows: `=SUMIF(..., ">0")`
- Total outflows: `=ABS(SUMIF(..., "<0"))`
- Net: `=SUM(...)`
- Table by category/subcategory: one row per config combination + `(choose)` for uncategorized, with `SUMIFS` formulas linked to Transactions

**Pivot sheet:**
- Placeholder text, then real pivot table creation via `_add_excel_pivot_table`

---

### 4. Main Flow

#### `main() -> int`

Orchestrates the full execution.

1. **Checks:**
   - Finds a CSV in `input/` via `find_first_csv()`
   - Verifies `config/categories.json` exists
   - Exits with code 1 if either is missing

2. **Loading:** Reads `categories.json` and CSV via `parse_belfius_csv()`

3. **Processing:**
   - `clean_and_normalize()` on raw data
   - `apply_rules()` on each row to fill `category` and `subcategory`

4. **Export:**
   - Saves DataFrame to clean CSV (`output/transactions_clean.csv`)
   - Creates Excel dashboard via `create_excel_dashboard()`

5. Returns `0` on success, `1` on error.

---

## `config/categories.json` Format

Expected structure:

```json
{
  "Category1": ["SubCat1", "SubCat2", "..."],
  "Category2": ["SubCatA", "SubCatB", "..."]
}
```

Category and subcategory names must match the rules in `apply_rules()` for auto-categorization to work. Unmatched transactions remain for manual categorization in Excel.
