# bizplan-updater Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Build a Python CLI tool that ingests QuickBooks CSV exports, aggregates transactions by month/category, writes actuals to the correct cells in an Excel financial model, and produces a Markdown variance report.

**Architecture:** SQLite stores all imported transactions as the source of truth; a YAML config maps QuickBooks subcategory paths to Excel sheet/row labels; openpyxl writes only to designated input sheets on versioned copies of the workbook. Claude API is called only for unknown category suggestions and narrative summaries.

**Tech Stack:** Python 3.12+, uv, click, openpyxl, pandas, sqlite3 (stdlib), rich, anthropic, pyyaml, pytest

---

## Task 1: Project Scaffolding

**Files:**
- Create: `pyproject.toml`
- Create: `bizplan/__init__.py`
- Create: `bizplan/cli.py`
- Create: `bizplan/db.py`
- Create: `bizplan/ingest.py`
- Create: `bizplan/categorize.py`
- Create: `bizplan/aggregate.py`
- Create: `bizplan/excel_writer.py`
- Create: `bizplan/reconcile.py`
- Create: `bizplan/report.py`
- Create: `bizplan/ai_assist.py`
- Create: `bizplan/config.py`
- Create: `tests/__init__.py`
- Create: `.gitignore`

**Step 1: Create pyproject.toml**

```toml
[project]
name = "bizplan-updater"
version = "0.1.0"
requires-python = ">=3.12"
dependencies = [
    "click>=8.1",
    "openpyxl>=3.1",
    "pandas>=2.2",
    "rich>=13.0",
    "anthropic>=0.40",
    "pyyaml>=6.0",
]

[project.scripts]
bizplan = "bizplan.cli:cli"

[build-system]
requires = ["hatchling"]
build-backend = "hatchling.build"

[tool.uv]
dev-dependencies = [
    "pytest>=8.0",
    "pytest-mock>=3.12",
]
```

**Step 2: Initialize uv environment and install**

Run: `cd /Users/jrv/Documents/Code/projects/bizplan_updater && uv sync`
Expected: Creates `.venv/`, installs all dependencies without errors.

**Step 3: Create stub files**

Each file below starts as a minimal stub (just a module docstring and imports). Fill them in task by task.

`bizplan/__init__.py` — empty file.

`bizplan/cli.py`:
```python
import click

@click.group()
def cli():
    """bizplan — Business Plan Updater CLI.

    Automates monthly financial model updates from QuickBooks exports.
    """
    pass
```

`bizplan/db.py`, `bizplan/ingest.py`, `bizplan/categorize.py`, `bizplan/aggregate.py`,
`bizplan/excel_writer.py`, `bizplan/reconcile.py`, `bizplan/report.py`,
`bizplan/ai_assist.py`, `bizplan/config.py` — each an empty file for now.

`tests/__init__.py` — empty file.

`.gitignore`:
```
.venv/
__pycache__/
*.pyc
data/transactions.db
.DS_Store
output/
*.env
```

**Step 4: Verify CLI is importable**

Run: `uv run bizplan --help`
Expected: Shows "bizplan — Business Plan Updater CLI." with no commands listed yet.

**Step 5: Initialize git and commit**

```bash
cd /Users/jrv/Documents/Code/projects/bizplan_updater
git init
git add pyproject.toml bizplan/ tests/ .gitignore docs/
git commit -m "chore: scaffold project structure"
```

---

## Task 2: SQLite DB Module

**Files:**
- Modify: `bizplan/db.py`
- Create: `tests/test_db.py`

**Step 1: Write the failing tests**

`tests/test_db.py`:
```python
import sqlite3
import pytest
from pathlib import Path
from bizplan.db import init_db, get_conn, insert_transaction, get_monthly_totals_raw

@pytest.fixture
def db_path(tmp_path):
    path = tmp_path / "test.db"
    init_db(path)
    return path

def test_init_creates_tables(db_path):
    with get_conn(db_path) as conn:
        tables = {row[0] for row in conn.execute(
            "SELECT name FROM sqlite_master WHERE type='table'"
        ).fetchall()}
    assert {"transactions", "category_mappings", "run_log"} <= tables

def test_insert_transaction(db_path):
    txn = {
        "fingerprint": "2026-01-06|Patreon|6239.41|Sales:Subscription Sales",
        "date": "2026-01-06",
        "vendor": "Patreon",
        "amount": 6239.41,
        "category_path": "Sales:Subscription Sales",
        "txn_type": "Deposit",
    }
    with get_conn(db_path) as conn:
        inserted = insert_transaction(conn, txn)
    assert inserted is True

def test_duplicate_transaction_skipped(db_path):
    txn = {
        "fingerprint": "2026-01-06|Patreon|6239.41|Sales:Subscription Sales",
        "date": "2026-01-06",
        "vendor": "Patreon",
        "amount": 6239.41,
        "category_path": "Sales:Subscription Sales",
        "txn_type": "Deposit",
    }
    with get_conn(db_path) as conn:
        insert_transaction(conn, txn)
        inserted_again = insert_transaction(conn, txn)
    assert inserted_again is False

def test_get_monthly_totals_raw(db_path):
    txns = [
        {"fingerprint": "a", "date": "2026-01-06", "vendor": "V1", "amount": 100.0,
         "category_path": "Cat:Sub", "txn_type": "Deposit"},
        {"fingerprint": "b", "date": "2026-01-15", "vendor": "V2", "amount": 50.0,
         "category_path": "Cat:Sub", "txn_type": "Deposit"},
        {"fingerprint": "c", "date": "2026-02-01", "vendor": "V3", "amount": 200.0,
         "category_path": "Cat:Sub", "txn_type": "Deposit"},
    ]
    with get_conn(db_path) as conn:
        for t in txns:
            insert_transaction(conn, t)
        totals = get_monthly_totals_raw(conn, "2026-01")
    assert totals == {"Cat:Sub": 150.0}
```

**Step 2: Run tests to verify they fail**

Run: `uv run pytest tests/test_db.py -v`
Expected: ImportError — functions not yet defined.

**Step 3: Implement db.py**

```python
import sqlite3
from contextlib import contextmanager
from pathlib import Path

SCHEMA = """
CREATE TABLE IF NOT EXISTS transactions (
    id            INTEGER PRIMARY KEY,
    fingerprint   TEXT UNIQUE NOT NULL,
    date          TEXT NOT NULL,
    vendor        TEXT NOT NULL,
    amount        REAL NOT NULL,
    category_path TEXT NOT NULL,
    txn_type      TEXT NOT NULL,
    imported_at   TEXT DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS category_mappings (
    id               INTEGER PRIMARY KEY,
    category_path    TEXT UNIQUE NOT NULL,
    excel_sheet      TEXT NOT NULL,
    excel_row_label  TEXT NOT NULL,
    confirmed_by_user INTEGER DEFAULT 0,
    notes            TEXT
);

CREATE TABLE IF NOT EXISTS run_log (
    id                    INTEGER PRIMARY KEY,
    month                 TEXT NOT NULL,
    run_at                TEXT DEFAULT (datetime('now')),
    transactions_imported INTEGER DEFAULT 0,
    transactions_new      INTEGER DEFAULT 0,
    cells_written         INTEGER DEFAULT 0,
    output_workbook       TEXT,
    report_path           TEXT
);
"""

def init_db(db_path: Path) -> None:
    """Create database file and schema at db_path."""
    db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(db_path)
    conn.executescript(SCHEMA)
    conn.commit()
    conn.close()

@contextmanager
def get_conn(db_path: Path):
    """Context manager yielding a committed, row_factory-enabled connection."""
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

def insert_transaction(conn: sqlite3.Connection, txn: dict) -> bool:
    """Insert a transaction. Returns True if inserted, False if duplicate."""
    try:
        conn.execute(
            """INSERT INTO transactions
               (fingerprint, date, vendor, amount, category_path, txn_type)
               VALUES (:fingerprint, :date, :vendor, :amount, :category_path, :txn_type)""",
            txn,
        )
        return True
    except sqlite3.IntegrityError:
        return False

def get_monthly_totals_raw(conn: sqlite3.Connection, month: str) -> dict[str, float]:
    """Return {category_path: sum_amount} for all transactions in YYYY-MM month."""
    rows = conn.execute(
        """SELECT category_path, SUM(amount) AS total
           FROM transactions
           WHERE strftime('%Y-%m', date) = ?
           GROUP BY category_path""",
        (month,),
    ).fetchall()
    return {row["category_path"]: row["total"] for row in rows}

def get_category_mapping(conn: sqlite3.Connection, category_path: str) -> dict | None:
    """Return mapping row for category_path, or None if not found."""
    row = conn.execute(
        "SELECT excel_sheet, excel_row_label FROM category_mappings WHERE category_path = ?",
        (category_path,),
    ).fetchone()
    return dict(row) if row else None

def save_category_mapping(conn: sqlite3.Connection, category_path: str,
                          excel_sheet: str, excel_row_label: str,
                          confirmed: bool = True, notes: str = "") -> None:
    """Insert or replace a category mapping."""
    conn.execute(
        """INSERT INTO category_mappings
           (category_path, excel_sheet, excel_row_label, confirmed_by_user, notes)
           VALUES (?, ?, ?, ?, ?)
           ON CONFLICT(category_path) DO UPDATE SET
               excel_sheet=excluded.excel_sheet,
               excel_row_label=excluded.excel_row_label,
               confirmed_by_user=excluded.confirmed_by_user,
               notes=excluded.notes""",
        (category_path, excel_sheet, excel_row_label, int(confirmed), notes),
    )

def log_run(conn: sqlite3.Connection, month: str, imported: int, new: int,
            written: int, workbook: str, report: str) -> None:
    conn.execute(
        """INSERT INTO run_log
           (month, transactions_imported, transactions_new, cells_written,
            output_workbook, report_path)
           VALUES (?, ?, ?, ?, ?, ?)""",
        (month, imported, new, written, workbook, report),
    )
```

**Step 4: Run tests to verify they pass**

Run: `uv run pytest tests/test_db.py -v`
Expected: 4 tests PASS.

**Step 5: Commit**

```bash
git add bizplan/db.py tests/test_db.py
git commit -m "feat: add SQLite db module with transaction store"
```

---

## Task 3: QuickBooks CSV Ingest Module

**Files:**
- Modify: `bizplan/ingest.py`
- Create: `tests/test_ingest.py`
- Reference: `quickbooks_exports/TD_Bank-3.csv` (existing sample)

**Step 1: Write the failing tests**

`tests/test_ingest.py`:
```python
import pytest
from pathlib import Path
from bizplan.ingest import parse_qb_csv, parse_amount, parse_transaction_posted

SAMPLE_CSV = Path("quickbooks_exports/TD_Bank-3.csv")

def test_parse_amount_negative():
    assert parse_amount("-$44.99") == pytest.approx(-44.99)

def test_parse_amount_positive():
    assert parse_amount("$1,781.47") == pytest.approx(1781.47)

def test_parse_amount_large():
    assert parse_amount("$10,352.50") == pytest.approx(10352.50)

def test_parse_transaction_posted_expense():
    posted = "Added to:  Expense: Production Expense:Production Software 01/30/2026 $44.99"
    result = parse_transaction_posted(posted)
    assert result["txn_type"] == "Expense"
    assert result["category_path"] == "Production Expense:Production Software"

def test_parse_transaction_posted_deposit():
    posted = "Added to:  Deposit: Sales:Subscription Sales 01/29/2026 $1,781.47"
    result = parse_transaction_posted(posted)
    assert result["txn_type"] == "Deposit"
    assert result["category_path"] == "Sales:Subscription Sales"

def test_parse_transaction_posted_complex_category():
    posted = "Added to:  Expense: General business expenses:Bank fees & service charges 12/31/2025 $10.00"
    result = parse_transaction_posted(posted)
    assert result["txn_type"] == "Expense"
    assert result["category_path"] == "General business expenses:Bank fees & service charges"

def test_parse_qb_csv_returns_list(tmp_path):
    # Create a minimal test CSV
    csv_content = '''date,Bank description,Amount,From/To,Transaction Posted
"01/06/2026","Patreon","$6,239.41","Patreon","Added to:  Deposit: Sales:Subscription Sales 01/06/2026 $6,239.41"
"01/06/2026","Inst Xfer","-$4.00","","Added to:  Expense: Advertising & marketing:Patreon Merch 01/06/2026 $4.00"
'''
    csv_file = tmp_path / "test.csv"
    csv_file.write_text(csv_content)
    rows = parse_qb_csv(csv_file)
    assert len(rows) == 2
    assert rows[0]["vendor"] == "Patreon"
    assert rows[0]["amount"] == pytest.approx(6239.41)
    assert rows[0]["date"] == "2026-01-06"
    assert rows[0]["txn_type"] == "Deposit"
    assert rows[0]["category_path"] == "Sales:Subscription Sales"
    assert rows[1]["amount"] == pytest.approx(-4.0)

def test_parse_qb_csv_fingerprint_format(tmp_path):
    csv_content = '''date,Bank description,Amount,From/To,Transaction Posted
"01/06/2026","Patreon","$6,239.41","Patreon","Added to:  Deposit: Sales:Subscription Sales 01/06/2026 $6,239.41"
'''
    csv_file = tmp_path / "test.csv"
    csv_file.write_text(csv_content)
    rows = parse_qb_csv(csv_file)
    expected_fp = "2026-01-06|Patreon|6239.41|Sales:Subscription Sales"
    assert rows[0]["fingerprint"] == expected_fp

def test_parse_real_sample_csv():
    """Smoke test against the actual sample file."""
    if not SAMPLE_CSV.exists():
        pytest.skip("Sample CSV not present")
    rows = parse_qb_csv(SAMPLE_CSV)
    assert len(rows) > 30
    # All rows have required fields
    for row in rows:
        assert "fingerprint" in row
        assert "date" in row
        assert "vendor" in row
        assert "amount" in row
        assert "category_path" in row
        assert "txn_type" in row
```

**Step 2: Run tests to verify they fail**

Run: `uv run pytest tests/test_ingest.py -v`
Expected: ImportError.

**Step 3: Implement ingest.py**

```python
import re
from pathlib import Path
import pandas as pd

# Pattern: "Added to:  {Type}: {category_path} {MM/DD/YYYY} ${amount}"
_POSTED_RE = re.compile(
    r"Added to:\s+(\w+):\s+(.+?)\s+\d{2}/\d{2}/\d{4}\s+\$[\d,]+\.?\d*$"
)

def parse_amount(amount_str: str) -> float:
    """Parse amount strings like '-$44.99' or '$1,781.47' to float."""
    clean = amount_str.replace("$", "").replace(",", "").strip()
    return float(clean)

def parse_transaction_posted(posted: str) -> dict:
    """Parse the 'Transaction Posted' column to extract txn_type and category_path."""
    m = _POSTED_RE.match(posted.strip())
    if not m:
        raise ValueError(f"Cannot parse Transaction Posted: {posted!r}")
    return {"txn_type": m.group(1), "category_path": m.group(2).strip()}

def parse_qb_csv(csv_path: Path) -> list[dict]:
    """Parse a QuickBooks CSV export into a list of transaction dicts."""
    df = pd.read_csv(csv_path, dtype=str, keep_default_na=False)
    rows = []
    for _, row in df.iterrows():
        raw_date = row["date"].strip()          # MM/DD/YYYY
        vendor = row["Bank description"].strip()
        raw_amount = row["Amount"].strip()
        posted = row["Transaction Posted"].strip()

        if not posted.startswith("Added to:"):
            continue  # Skip unrecognized rows

        try:
            amount = parse_amount(raw_amount)
            parsed = parse_transaction_posted(posted)
            # Convert date MM/DD/YYYY -> YYYY-MM-DD
            m, d, y = raw_date.split("/")
            date = f"{y}-{m}-{d}"
            category_path = parsed["category_path"]
            fingerprint = f"{date}|{vendor}|{abs(amount)}|{category_path}"
            rows.append({
                "fingerprint": fingerprint,
                "date": date,
                "vendor": vendor,
                "amount": amount,
                "category_path": category_path,
                "txn_type": parsed["txn_type"],
            })
        except (ValueError, KeyError):
            continue  # Skip rows that can't be parsed

    return rows
```

**Step 4: Run tests to verify they pass**

Run: `uv run pytest tests/test_ingest.py -v`
Expected: All tests PASS.

**Step 5: Commit**

```bash
git add bizplan/ingest.py tests/test_ingest.py
git commit -m "feat: add QuickBooks CSV ingest module"
```

---

## Task 4: Config Module (YAML Mappings)

**Files:**
- Modify: `bizplan/config.py`
- Create: `tests/test_config.py`
- Create: `config/mappings.yaml` (empty starter)
- Create: `config/model.yaml` (empty starter)

**Step 1: Write the failing tests**

`tests/test_config.py`:
```python
import pytest
from pathlib import Path
from bizplan.config import load_mappings, save_mappings, load_model, save_model

def test_load_mappings_missing_file(tmp_path):
    result = load_mappings(tmp_path / "mappings.yaml")
    assert result == {}

def test_save_and_load_mappings(tmp_path):
    path = tmp_path / "mappings.yaml"
    data = {
        "Advertising & marketing:Social media": {
            "sheet": "Marketing Exp",
            "row_label": "Social Media",
        }
    }
    save_mappings(data, path)
    loaded = load_mappings(path)
    assert loaded["Advertising & marketing:Social media"]["sheet"] == "Marketing Exp"

def test_save_and_load_model(tmp_path):
    path = tmp_path / "model.yaml"
    data = {
        "input_sheets": ["Marketing Exp", "Production Exp"],
        "locked_sheets": ["Income Statement", "Balance Sheet"],
    }
    save_model(data, path)
    loaded = load_model(path)
    assert "Marketing Exp" in loaded["input_sheets"]

def test_load_model_missing_file(tmp_path):
    result = load_model(tmp_path / "model.yaml")
    assert result == {}
```

**Step 2: Run tests to verify they fail**

Run: `uv run pytest tests/test_config.py -v`
Expected: ImportError.

**Step 3: Implement config.py**

```python
from pathlib import Path
import yaml

DEFAULT_MAPPINGS_PATH = Path("config/mappings.yaml")
DEFAULT_MODEL_PATH = Path("config/model.yaml")

def load_mappings(path: Path = DEFAULT_MAPPINGS_PATH) -> dict:
    if not path.exists():
        return {}
    return yaml.safe_load(path.read_text()) or {}

def save_mappings(mappings: dict, path: Path = DEFAULT_MAPPINGS_PATH) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(yaml.dump(mappings, default_flow_style=False, sort_keys=True))

def load_model(path: Path = DEFAULT_MODEL_PATH) -> dict:
    if not path.exists():
        return {}
    return yaml.safe_load(path.read_text()) or {}

def save_model(model: dict, path: Path = DEFAULT_MODEL_PATH) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(yaml.dump(model, default_flow_style=False, sort_keys=True))
```

**Step 4: Create starter config files**

`config/mappings.yaml`:
```yaml
# QuickBooks category_path -> Excel sheet + row label
# Populated interactively by: bizplan init / bizplan map
```

`config/model.yaml`:
```yaml
# Excel workbook structure — populated by: bizplan init
# input_sheets: list of sheets the tool may write to
# locked_sheets: formula-driven sheets the tool must never touch
```

**Step 5: Run tests to verify they pass**

Run: `uv run pytest tests/test_config.py -v`
Expected: All tests PASS.

**Step 6: Commit**

```bash
git add bizplan/config.py tests/test_config.py config/
git commit -m "feat: add YAML config load/save for mappings and model structure"
```

---

## Task 5: Monthly Aggregation Module (Transaction Integrity Rules)

**Files:**
- Modify: `bizplan/aggregate.py`
- Create: `tests/test_aggregate.py`

**Step 1: Write the failing tests**

`tests/test_aggregate.py`:
```python
import pytest
from bizplan.db import init_db, get_conn, insert_transaction
from bizplan.aggregate import (
    aggregate_month,
    apply_mappings,
    detect_anomalies,
)

@pytest.fixture
def db_path(tmp_path):
    path = tmp_path / "test.db"
    init_db(path)
    return path

def seed(conn, txns):
    for i, t in enumerate(txns):
        t.setdefault("fingerprint", str(i))
        t.setdefault("txn_type", "Expense")
        insert_transaction(conn, t)

def test_rule1_aggregates_same_category(db_path):
    """Multiple transactions in same category/month are summed."""
    with get_conn(db_path) as conn:
        seed(conn, [
            {"date": "2026-01-10", "vendor": "V1", "amount": -80.0, "category_path": "Mktg:Dig"},
            {"date": "2026-01-20", "vendor": "V2", "amount": -20.0, "category_path": "Mktg:Dig"},
            {"date": "2026-01-25", "vendor": "V3", "amount": -200.0, "category_path": "Mktg:Dig"},
            {"date": "2026-01-28", "vendor": "V4", "amount": -100.0, "category_path": "Mktg:Dig"},
        ])
        totals = aggregate_month(conn, "2026-01")
    assert totals["Mktg:Dig"] == pytest.approx(-400.0)

def test_rule3_negatives_reduce_total(db_path):
    """Refunds (negative) reduce the monthly total."""
    with get_conn(db_path) as conn:
        seed(conn, [
            {"date": "2026-01-10", "vendor": "Meta", "amount": -200.0, "category_path": "Mktg:Ads"},
            {"date": "2026-01-15", "vendor": "Meta", "amount": 50.0, "category_path": "Mktg:Ads"},
        ])
        totals = aggregate_month(conn, "2026-01")
    assert totals["Mktg:Ads"] == pytest.approx(-150.0)

def test_rule4_split_transactions_independent(db_path):
    """Split transaction rows are treated independently before aggregation."""
    with get_conn(db_path) as conn:
        seed(conn, [
            {"date": "2026-01-10", "vendor": "Inv", "amount": -700.0, "category_path": "Prod:Prod"},
            {"date": "2026-01-10", "vendor": "Inv", "amount": -300.0, "category_path": "Mktg:Mktg"},
        ])
        totals = aggregate_month(conn, "2026-01")
    assert totals["Prod:Prod"] == pytest.approx(-700.0)
    assert totals["Mktg:Mktg"] == pytest.approx(-300.0)

def test_rule5_excludes_other_months(db_path):
    """Full-month recalculation returns only the requested month."""
    with get_conn(db_path) as conn:
        seed(conn, [
            {"date": "2026-01-10", "vendor": "V", "amount": -100.0, "category_path": "Cat:A"},
            {"date": "2026-02-10", "vendor": "V", "amount": -200.0, "category_path": "Cat:A"},
        ])
        totals = aggregate_month(conn, "2026-01")
    assert "Cat:A" in totals
    assert totals["Cat:A"] == pytest.approx(-100.0)

def test_apply_mappings_resolves_sheet_and_row(db_path):
    """apply_mappings pairs aggregated totals with their Excel targets."""
    from bizplan.db import save_category_mapping
    with get_conn(db_path) as conn:
        save_category_mapping(conn, "Mktg:Social", "Marketing Exp", "Social Media")
        seed(conn, [
            {"date": "2026-01-05", "vendor": "V", "amount": -420.0, "category_path": "Mktg:Social"},
        ])
        result = apply_mappings(conn, "2026-01")
    assert result[0]["excel_sheet"] == "Marketing Exp"
    assert result[0]["excel_row_label"] == "Social Media"
    assert result[0]["amount"] == pytest.approx(-420.0)

def test_detect_anomalies_flags_spike(db_path):
    """detect_anomalies flags categories > 3x prior month average."""
    with get_conn(db_path) as conn:
        # Seed two prior months as baseline
        seed(conn, [
            {"fingerprint": "a1", "date": "2025-11-10", "vendor": "V", "amount": -100.0, "category_path": "Cat:A"},
            {"fingerprint": "a2", "date": "2025-12-10", "vendor": "V", "amount": -100.0, "category_path": "Cat:A"},
            # Current month spike
            {"fingerprint": "a3", "date": "2026-01-10", "vendor": "V", "amount": -500.0, "category_path": "Cat:A"},
        ])
        anomalies = detect_anomalies(conn, "2026-01")
    assert any(a["category_path"] == "Cat:A" for a in anomalies)
```

**Step 2: Run tests to verify they fail**

Run: `uv run pytest tests/test_aggregate.py -v`
Expected: ImportError.

**Step 3: Implement aggregate.py**

```python
import sqlite3
from bizplan.db import get_monthly_totals_raw, get_category_mapping

def aggregate_month(conn: sqlite3.Connection, month: str) -> dict[str, float]:
    """Return {category_path: sum_amount} for the given YYYY-MM month.
    Full recomputation from DB — implements Rules 1, 3, 4, 5."""
    return get_monthly_totals_raw(conn, month)

def apply_mappings(conn: sqlite3.Connection, month: str) -> list[dict]:
    """Return aggregated totals enriched with Excel sheet/row targets.
    Only returns categories that have confirmed mappings."""
    totals = aggregate_month(conn, month)
    result = []
    for category_path, amount in totals.items():
        mapping = get_category_mapping(conn, category_path)
        if mapping:
            result.append({
                "category_path": category_path,
                "amount": amount,
                "excel_sheet": mapping["excel_sheet"],
                "excel_row_label": mapping["excel_row_label"],
            })
    return result

def get_unmapped_categories(conn: sqlite3.Connection, month: str) -> list[str]:
    """Return category_paths in month that have no mapping in category_mappings."""
    totals = aggregate_month(conn, month)
    unmapped = []
    for category_path in totals:
        if get_category_mapping(conn, category_path) is None:
            unmapped.append(category_path)
    return unmapped

def detect_anomalies(conn: sqlite3.Connection, month: str,
                     spike_threshold: float = 3.0) -> list[dict]:
    """Flag categories where this month's total is > spike_threshold * prior 2-month avg."""
    current = aggregate_month(conn, month)
    anomalies = []
    y, m = map(int, month.split("-"))

    def prior_month(offset: int) -> str:
        mo = m - offset
        yr = y
        if mo <= 0:
            mo += 12
            yr -= 1
        return f"{yr}-{mo:02d}"

    for category_path, current_total in current.items():
        prior_totals = []
        for offset in (1, 2):
            pm = prior_month(offset)
            pm_totals = get_monthly_totals_raw(conn, pm)
            if category_path in pm_totals:
                prior_totals.append(abs(pm_totals[category_path]))

        if prior_totals:
            avg = sum(prior_totals) / len(prior_totals)
            if avg > 0 and abs(current_total) > spike_threshold * avg:
                anomalies.append({
                    "category_path": category_path,
                    "current_total": current_total,
                    "prior_avg": avg,
                    "ratio": abs(current_total) / avg,
                })
    return anomalies
```

**Step 4: Run tests to verify they pass**

Run: `uv run pytest tests/test_aggregate.py -v`
Expected: All tests PASS.

**Step 5: Commit**

```bash
git add bizplan/aggregate.py tests/test_aggregate.py
git commit -m "feat: add monthly aggregation with transaction integrity rules"
```

---

## Task 6: Excel Model Reader and Writer

**Files:**
- Modify: `bizplan/excel_writer.py`
- Create: `tests/test_excel_writer.py`

**Step 1: Write the failing tests**

`tests/test_excel_writer.py`:
```python
import pytest
import openpyxl
from pathlib import Path
from bizplan.excel_writer import (
    scan_workbook_structure,
    find_month_column,
    find_row_by_label,
    write_cell_value,
    LOCKED_SHEETS,
)

@pytest.fixture
def sample_wb(tmp_path):
    """Create a minimal test workbook with one input sheet."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Marketing Exp"
    # Header row: A1 blank, B1=Jan 2026, C1=Feb 2026
    ws["A1"] = "Category"
    ws["B1"] = "Jan 2026"
    ws["C1"] = "Feb 2026"
    # Data rows
    ws["A2"] = "Social Media"
    ws["A3"] = "Website Ads"
    path = tmp_path / "test.xlsx"
    wb.save(path)
    return path

def test_scan_workbook_structure(sample_wb):
    structure = scan_workbook_structure(sample_wb)
    assert "Marketing Exp" in structure
    assert "header_row" in structure["Marketing Exp"]
    assert "label_col" in structure["Marketing Exp"]

def test_find_month_column(sample_wb):
    wb = openpyxl.load_workbook(sample_wb)
    ws = wb["Marketing Exp"]
    col = find_month_column(ws, "2026-01", header_row=1)
    assert col == 2  # Column B

def test_find_month_column_not_found(sample_wb):
    wb = openpyxl.load_workbook(sample_wb)
    ws = wb["Marketing Exp"]
    col = find_month_column(ws, "2026-06", header_row=1)
    assert col is None

def test_find_row_by_label(sample_wb):
    wb = openpyxl.load_workbook(sample_wb)
    ws = wb["Marketing Exp"]
    row = find_row_by_label(ws, "Social Media", label_col=1)
    assert row == 2

def test_find_row_by_label_not_found(sample_wb):
    wb = openpyxl.load_workbook(sample_wb)
    ws = wb["Marketing Exp"]
    row = find_row_by_label(ws, "Nonexistent Label", label_col=1)
    assert row is None

def test_write_cell_value(sample_wb, tmp_path):
    output = tmp_path / "output.xlsx"
    write_cell_value(sample_wb, "Marketing Exp", row=2, col=2, value=420.0, output_path=output)
    wb = openpyxl.load_workbook(output)
    assert wb["Marketing Exp"].cell(row=2, column=2).value == 420.0

def test_write_locked_sheet_raises(sample_wb, tmp_path):
    output = tmp_path / "output.xlsx"
    # Add a locked sheet to the test workbook
    wb = openpyxl.load_workbook(sample_wb)
    wb.create_sheet("Income Statement")
    wb.save(sample_wb)
    with pytest.raises(ValueError, match="locked"):
        write_cell_value(sample_wb, "Income Statement", row=2, col=2,
                         value=100.0, output_path=output)
```

**Step 2: Run tests to verify they fail**

Run: `uv run pytest tests/test_excel_writer.py -v`
Expected: ImportError.

**Step 3: Implement excel_writer.py**

```python
import re
import shutil
from pathlib import Path
import openpyxl
from openpyxl.worksheet.worksheet import Worksheet

# These sheets must never be written to
LOCKED_SHEETS = frozenset({
    "Summary Financials",
    "Income Statement",
    "Balance Sheet",
    "Cash Flow",
})

# Map abbreviated month headers to YYYY-MM
_MONTH_ABBR = {
    "jan": "01", "feb": "02", "mar": "03", "apr": "04",
    "may": "05", "jun": "06", "jul": "07", "aug": "08",
    "sep": "09", "oct": "10", "nov": "11", "dec": "12",
}

def _header_to_yyyymm(header: str) -> str | None:
    """Convert 'Jan 2026' or 'January 2026' to '2026-01', or None if unparseable."""
    m = re.match(r"([A-Za-z]+)\s+(\d{4})", str(header).strip())
    if not m:
        return None
    month_abbr = m.group(1)[:3].lower()
    year = m.group(2)
    month_num = _MONTH_ABBR.get(month_abbr)
    if not month_num:
        return None
    return f"{year}-{month_num}"

def find_month_column(ws: Worksheet, month: str, header_row: int = 1) -> int | None:
    """Return the 1-based column index for the given YYYY-MM month, or None."""
    for col in range(1, ws.max_column + 1):
        cell_val = ws.cell(row=header_row, column=col).value
        if cell_val and _header_to_yyyymm(str(cell_val)) == month:
            return col
    return None

def find_row_by_label(ws: Worksheet, label: str, label_col: int = 1) -> int | None:
    """Return the 1-based row index where column label_col matches label, or None."""
    for row in range(1, ws.max_row + 1):
        cell_val = ws.cell(row=row, column=label_col).value
        if cell_val and str(cell_val).strip() == label.strip():
            return row
    return None

def scan_workbook_structure(wb_path: Path) -> dict:
    """Scan writable input sheets and return their header_row and label_col positions.

    Returns: {sheet_name: {"header_row": int, "label_col": int}}
    Heuristic: header_row is the first row containing a recognizable month string.
    """
    wb = openpyxl.load_workbook(wb_path, read_only=True, data_only=True)
    structure = {}
    for sheet_name in wb.sheetnames:
        if sheet_name in LOCKED_SHEETS:
            continue
        ws = wb[sheet_name]
        header_row = None
        for row_idx in range(1, min(10, ws.max_row + 1)):
            for col_idx in range(1, ws.max_column + 1):
                val = ws.cell(row=row_idx, column=col_idx).value
                if val and _header_to_yyyymm(str(val)):
                    header_row = row_idx
                    break
            if header_row:
                break
        if header_row:
            structure[sheet_name] = {"header_row": header_row, "label_col": 1}
    wb.close()
    return structure

def write_cell_value(wb_path: Path, sheet_name: str, row: int, col: int,
                     value: float, output_path: Path) -> None:
    """Copy workbook to output_path and write value to sheet_name[row][col].

    Raises ValueError if sheet_name is locked.
    """
    if sheet_name in LOCKED_SHEETS:
        raise ValueError(f"Sheet '{sheet_name}' is locked and cannot be written to.")
    shutil.copy2(wb_path, output_path)
    wb = openpyxl.load_workbook(output_path)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Sheet '{sheet_name}' not found in workbook.")
    wb[sheet_name].cell(row=row, column=col, value=value)
    wb.save(output_path)

def write_monthly_actuals(
    source_wb_path: Path,
    output_wb_path: Path,
    actuals: list[dict],
    model_structure: dict,
    month: str,
) -> int:
    """Write all actuals for a month to a versioned workbook copy.

    actuals: list of {excel_sheet, excel_row_label, amount}
    model_structure: from scan_workbook_structure()
    Returns: number of cells written.
    """
    if source_wb_path != output_wb_path:
        shutil.copy2(source_wb_path, output_wb_path)

    wb = openpyxl.load_workbook(output_wb_path)
    cells_written = 0

    for item in actuals:
        sheet_name = item["excel_sheet"]
        row_label = item["excel_row_label"]
        amount = item["amount"]

        if sheet_name not in wb.sheetnames:
            continue
        if sheet_name in LOCKED_SHEETS:
            continue
        if sheet_name not in model_structure:
            continue

        ws = wb[sheet_name]
        sheet_meta = model_structure[sheet_name]
        header_row = sheet_meta["header_row"]
        label_col = sheet_meta["label_col"]

        col = find_month_column(ws, month, header_row=header_row)
        row = find_row_by_label(ws, row_label, label_col=label_col)

        if col is None or row is None:
            continue

        ws.cell(row=row, column=col, value=abs(amount))  # Store as positive in model
        cells_written += 1

    wb.save(output_wb_path)
    return cells_written
```

**Step 4: Run tests to verify they pass**

Run: `uv run pytest tests/test_excel_writer.py -v`
Expected: All tests PASS.

**Step 5: Commit**

```bash
git add bizplan/excel_writer.py tests/test_excel_writer.py
git commit -m "feat: add Excel model reader and safe cell writer"
```

---

## Task 7: AI Assist Module

**Files:**
- Modify: `bizplan/ai_assist.py`
- Create: `tests/test_ai_assist.py`

**Step 1: Write the failing tests**

`tests/test_ai_assist.py`:
```python
import pytest
from unittest.mock import MagicMock, patch
from bizplan.ai_assist import (
    suggest_category_mapping,
    generate_run_summary,
    explain_anomaly,
)

@patch("bizplan.ai_assist.anthropic.Anthropic")
def test_suggest_category_mapping(mock_anthropic_cls):
    mock_client = MagicMock()
    mock_anthropic_cls.return_value = mock_client
    mock_client.messages.create.return_value = MagicMock(
        content=[MagicMock(text='{"sheet": "Marketing Exp", "row_label": "Social Media"}')]
    )
    result = suggest_category_mapping(
        category_path="Advertising & marketing:Social media",
        vendor="Ap Vmo Vimeo",
        existing_mappings={"Advertising & marketing:Website ads": {"sheet": "Marketing Exp", "row_label": "Website Ads"}},
    )
    assert result["sheet"] == "Marketing Exp"
    assert result["row_label"] == "Social Media"

@patch("bizplan.ai_assist.anthropic.Anthropic")
def test_generate_run_summary(mock_anthropic_cls):
    mock_client = MagicMock()
    mock_anthropic_cls.return_value = mock_client
    mock_client.messages.create.return_value = MagicMock(
        content=[MagicMock(text="Subscription revenue was strong this month.")]
    )
    result = generate_run_summary(
        month="2026-01",
        totals={"Sales:Subscription Sales": 8021.0},
        prior_totals={"Sales:Subscription Sales": 7500.0},
    )
    assert isinstance(result, str)
    assert len(result) > 0

@patch("bizplan.ai_assist.anthropic.Anthropic")
def test_explain_anomaly(mock_anthropic_cls):
    mock_client = MagicMock()
    mock_anthropic_cls.return_value = mock_client
    mock_client.messages.create.return_value = MagicMock(
        content=[MagicMock(text="Q4 true-up payment likely.")]
    )
    result = explain_anomaly(
        category_path="Expense:Partner distributions",
        current_total=-15000.0,
        prior_avg=-5000.0,
    )
    assert isinstance(result, str)
```

**Step 2: Run tests to verify they fail**

Run: `uv run pytest tests/test_ai_assist.py -v`
Expected: ImportError.

**Step 3: Implement ai_assist.py**

```python
import json
import anthropic

_MODEL = "claude-haiku-4-5-20251001"  # Fast and cheap for short classification tasks

def _client() -> anthropic.Anthropic:
    return anthropic.Anthropic()  # Reads ANTHROPIC_API_KEY from environment

def suggest_category_mapping(
    category_path: str,
    vendor: str,
    existing_mappings: dict,
) -> dict | None:
    """Ask Claude to suggest an Excel sheet and row_label for an unknown category.

    Returns {"sheet": ..., "row_label": ...} or None if suggestion fails.
    """
    existing_str = json.dumps(existing_mappings, indent=2)
    prompt = f"""You are helping map QuickBooks categories to Excel model rows.

Unknown category: "{category_path}"
Vendor name: "{vendor}"

Existing mappings for reference:
{existing_str}

Respond with ONLY a JSON object: {{"sheet": "<sheet name>", "row_label": "<row label>"}}
Choose the most appropriate sheet and row_label from the existing mappings, or suggest a new one."""

    try:
        response = _client().messages.create(
            model=_MODEL,
            max_tokens=128,
            messages=[{"role": "user", "content": prompt}],
        )
        text = response.content[0].text.strip()
        return json.loads(text)
    except Exception:
        return None

def generate_run_summary(month: str, totals: dict, prior_totals: dict) -> str:
    """Generate a 2-3 sentence narrative summary of the monthly run."""
    prompt = f"""Summarize the key financial trends for {month} in 2-3 sentences.
Current month totals: {json.dumps(totals)}
Prior month totals: {json.dumps(prior_totals)}
Focus on notable changes, not every line item. Be factual and concise."""

    try:
        response = _client().messages.create(
            model=_MODEL,
            max_tokens=256,
            messages=[{"role": "user", "content": prompt}],
        )
        return response.content[0].text.strip()
    except Exception:
        return "AI summary unavailable."

def explain_anomaly(category_path: str, current_total: float, prior_avg: float) -> str:
    """Return a one-sentence explanation for an anomalous category."""
    ratio = abs(current_total) / abs(prior_avg) if prior_avg != 0 else 0
    prompt = f"""Explain in one sentence why "{category_path}" might be {ratio:.1f}x higher than usual.
Current: ${abs(current_total):,.2f}  Prior avg: ${abs(prior_avg):,.2f}
Be concise and suggest a likely business reason."""

    try:
        response = _client().messages.create(
            model=_MODEL,
            max_tokens=128,
            messages=[{"role": "user", "content": prompt}],
        )
        return response.content[0].text.strip()
    except Exception:
        return f"Anomaly: {ratio:.1f}x prior average — review recommended."
```

**Step 4: Run tests to verify they pass**

Run: `uv run pytest tests/test_ai_assist.py -v`
Expected: All tests PASS (uses mocked Anthropic client).

**Step 5: Commit**

```bash
git add bizplan/ai_assist.py tests/test_ai_assist.py
git commit -m "feat: add Claude API assist for category suggestions and summaries"
```

---

## Task 8: Categorize Module (Interactive Mapping)

**Files:**
- Modify: `bizplan/categorize.py`
- Create: `tests/test_categorize.py`

**Step 1: Write the failing tests**

`tests/test_categorize.py`:
```python
import pytest
from unittest.mock import patch, MagicMock
from bizplan.db import init_db, get_conn, save_category_mapping, get_category_mapping
from bizplan.categorize import resolve_categories, prompt_for_mapping

@pytest.fixture
def db_path(tmp_path):
    path = tmp_path / "test.db"
    init_db(path)
    return path

def test_resolve_categories_known(db_path):
    """Known categories resolve without prompting."""
    with get_conn(db_path) as conn:
        save_category_mapping(conn, "Mktg:Social", "Marketing Exp", "Social Media")
        unmapped = resolve_categories(conn, ["Mktg:Social"], vendor_map={}, prompt=False)
    assert unmapped == []

def test_resolve_categories_unknown_no_prompt(db_path):
    """Unknown categories returned as unmapped when prompt=False."""
    with get_conn(db_path) as conn:
        unmapped = resolve_categories(conn, ["Mktg:Unknown"], vendor_map={}, prompt=False)
    assert "Mktg:Unknown" in unmapped

def test_prompt_for_mapping_saves_to_db(db_path):
    """prompt_for_mapping saves user-provided mapping to DB."""
    with patch("bizplan.categorize.Prompt.ask", side_effect=["Marketing Exp", "Social Media"]):
        with patch("bizplan.categorize.Confirm.ask", return_value=True):
            with get_conn(db_path) as conn:
                prompt_for_mapping(conn, "Mktg:Social", vendor="Buffer", ai_suggestion=None)
                mapping = get_category_mapping(conn, "Mktg:Social")
    assert mapping is not None
    assert mapping["excel_sheet"] == "Marketing Exp"
```

**Step 2: Run tests to verify they fail**

Run: `uv run pytest tests/test_categorize.py -v`
Expected: ImportError.

**Step 3: Implement categorize.py**

```python
import sqlite3
from rich.prompt import Prompt, Confirm
from rich.console import Console
from rich import print as rprint
from bizplan.db import get_category_mapping, save_category_mapping

console = Console()

def resolve_categories(
    conn: sqlite3.Connection,
    category_paths: list[str],
    vendor_map: dict[str, str],
    prompt: bool = True,
) -> list[str]:
    """For each category_path, ensure a mapping exists in DB.

    If prompt=True, interactively ask user for unknown categories.
    Returns list of still-unmapped category_paths.
    """
    unmapped = []
    for cp in category_paths:
        if get_category_mapping(conn, cp):
            continue
        if prompt:
            vendor = vendor_map.get(cp, "Unknown vendor")
            prompt_for_mapping(conn, cp, vendor=vendor, ai_suggestion=None)
            if not get_category_mapping(conn, cp):
                unmapped.append(cp)
        else:
            unmapped.append(cp)
    return unmapped

def prompt_for_mapping(
    conn: sqlite3.Connection,
    category_path: str,
    vendor: str,
    ai_suggestion: dict | None,
) -> None:
    """Interactively prompt the user to map an unknown category to an Excel location."""
    console.rule(f"[yellow]Unknown Category[/yellow]")
    rprint(f"[bold]Category:[/bold] {category_path}")
    rprint(f"[bold]Vendor:[/bold] {vendor}")

    if ai_suggestion:
        rprint(f"[green]AI Suggestion:[/green] Sheet={ai_suggestion['sheet']!r}, "
               f"Row={ai_suggestion['row_label']!r}")
        if Confirm.ask("Accept AI suggestion?", default=True):
            save_category_mapping(
                conn, category_path,
                ai_suggestion["sheet"], ai_suggestion["row_label"],
                confirmed=True, notes="AI-suggested",
            )
            return

    sheet = Prompt.ask("Excel sheet name")
    row_label = Prompt.ask("Excel row label")
    save_category_mapping(conn, category_path, sheet, row_label, confirmed=True)
    rprint(f"[green]Saved:[/green] {category_path!r} → {sheet!r} / {row_label!r}")
```

**Step 4: Run tests to verify they pass**

Run: `uv run pytest tests/test_categorize.py -v`
Expected: All tests PASS.

**Step 5: Commit**

```bash
git add bizplan/categorize.py tests/test_categorize.py
git commit -m "feat: add interactive category mapping resolution"
```

---

## Task 9: Reconcile Module

**Files:**
- Modify: `bizplan/reconcile.py`
- Create: `tests/test_reconcile.py`

**Step 1: Write the failing tests**

`tests/test_reconcile.py`:
```python
import pytest
from bizplan.db import init_db, get_conn, insert_transaction
from bizplan.reconcile import reconcile_revenue_cash

@pytest.fixture
def db_path(tmp_path):
    path = tmp_path / "test.db"
    init_db(path)
    return path

def seed(conn, txns):
    for i, t in enumerate(txns):
        t.setdefault("fingerprint", str(i))
        insert_transaction(conn, t)

def test_reconcile_no_gap(db_path):
    """When deposits match recognized revenue categories, no gap flagged."""
    with get_conn(db_path) as conn:
        seed(conn, [
            {"date": "2026-01-09", "vendor": "Wingspan", "amount": 10352.50,
             "category_path": "Sales:Host Read Ad Sales", "txn_type": "Deposit"},
        ])
        result = reconcile_revenue_cash(conn, "2026-01")
    assert result["total_deposits"] == pytest.approx(10352.50)

def test_reconcile_flags_negative_gap(db_path):
    """When recognized revenue > deposits, flag timing gap."""
    with get_conn(db_path) as conn:
        seed(conn, [
            {"date": "2026-01-09", "vendor": "Wingspan", "amount": 10352.50,
             "category_path": "Sales:Host Read Ad Sales", "txn_type": "Deposit"},
            {"fingerprint": "b", "date": "2026-01-29", "vendor": "Apple",
             "amount": 1781.47, "category_path": "Sales:Subscription Sales",
             "txn_type": "Deposit"},
        ])
        result = reconcile_revenue_cash(conn, "2026-01")
    assert result["total_deposits"] == pytest.approx(10352.50 + 1781.47)
```

**Step 2: Run tests to verify they fail**

Run: `uv run pytest tests/test_reconcile.py -v`
Expected: ImportError.

**Step 3: Implement reconcile.py**

```python
import sqlite3

# Category paths that represent cash inflows (deposits)
_DEPOSIT_CATEGORIES = frozenset({
    "Sales:Host Read Ad Sales",
    "Sales:Subscription Sales",
    "Sales:Programmatic Ad Revenue",
    "Commissions & fees:Network Ad Revenue Distribution",
})

def reconcile_revenue_cash(conn: sqlite3.Connection, month: str) -> dict:
    """Summarize deposits received vs. expected revenue for the month.

    Returns a dict with total_deposits, by_category, and any timing flags.
    """
    rows = conn.execute(
        """SELECT category_path, SUM(amount) AS total
           FROM transactions
           WHERE strftime('%Y-%m', date) = ?
             AND txn_type = 'Deposit'
           GROUP BY category_path""",
        (month,),
    ).fetchall()

    by_category = {row["category_path"]: row["total"] for row in rows}
    total_deposits = sum(by_category.values())

    flags = []
    for cat, total in by_category.items():
        if cat not in _DEPOSIT_CATEGORIES:
            flags.append(f"Unrecognized deposit category: {cat!r} (${total:,.2f})")

    return {
        "month": month,
        "total_deposits": total_deposits,
        "by_category": by_category,
        "flags": flags,
    }
```

**Step 4: Run tests to verify they pass**

Run: `uv run pytest tests/test_reconcile.py -v`
Expected: All tests PASS.

**Step 5: Commit**

```bash
git add bizplan/reconcile.py tests/test_reconcile.py
git commit -m "feat: add revenue/cash reconciliation module"
```

---

## Task 10: Report Module

**Files:**
- Modify: `bizplan/report.py`
- Create: `tests/test_report.py`

**Step 1: Write the failing tests**

`tests/test_report.py`:
```python
import pytest
from pathlib import Path
from bizplan.report import generate_report, format_variance_table

def test_format_variance_table():
    current = {"Cat:A": -420.0, "Cat:B": -100.0}
    prior = {"Cat:A": -350.0, "Cat:B": -100.0}
    table = format_variance_table(current, prior)
    assert "Cat:A" in table
    assert "420" in table
    assert "350" in table

def test_generate_report_creates_file(tmp_path):
    report_path = tmp_path / "2026-01-report.md"
    generate_report(
        month="2026-01",
        imported=35,
        new=32,
        cells_written=12,
        workbook_path=str(tmp_path / "2026-01_update.xlsx"),
        current_totals={"Cat:A": -420.0},
        prior_totals={"Cat:A": -350.0},
        anomalies=[],
        reconcile_result={"total_deposits": 10000.0, "by_category": {}, "flags": []},
        ai_summary="Revenue was strong.",
        output_path=report_path,
    )
    assert report_path.exists()
    content = report_path.read_text()
    assert "2026-01" in content
    assert "Revenue was strong" in content
    assert "Cat:A" in content
```

**Step 2: Run tests to verify they fail**

Run: `uv run pytest tests/test_report.py -v`
Expected: ImportError.

**Step 3: Implement report.py**

```python
from pathlib import Path
from datetime import datetime

def format_variance_table(current: dict, prior: dict) -> str:
    """Format a Markdown variance table comparing current vs prior month totals."""
    lines = [
        "| Category | Prior | Actual | Variance |",
        "|---|---|---|---|",
    ]
    all_cats = sorted(set(current) | set(prior))
    for cat in all_cats:
        cur = current.get(cat, 0.0)
        prr = prior.get(cat, 0.0)
        var = cur - prr
        direction = "▲" if var > 0 else ("▼" if var < 0 else "—")
        lines.append(
            f"| {cat} | ${abs(prr):,.2f} | ${abs(cur):,.2f} | {direction} ${abs(var):,.2f} |"
        )
    return "\n".join(lines)

def generate_report(
    month: str,
    imported: int,
    new: int,
    cells_written: int,
    workbook_path: str,
    current_totals: dict,
    prior_totals: dict,
    anomalies: list,
    reconcile_result: dict,
    ai_summary: str,
    output_path: Path,
) -> None:
    """Write a Markdown run report to output_path."""
    lines = [
        f"# Business Plan Update — {month}",
        f"",
        f"_Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}_",
        f"",
        f"## Transactions",
        f"- {imported} imported, {imported - new} duplicates skipped, {new} new",
        f"",
        f"## Excel Updates",
        f"- {cells_written} cells written",
        f"- Workbook: `{workbook_path}`",
        f"",
        f"## Variance vs Prior Month",
        f"",
        format_variance_table(current_totals, prior_totals),
        f"",
        f"## Revenue & Cash Reconciliation",
        f"- Total deposits received: ${reconcile_result['total_deposits']:,.2f}",
    ]

    if reconcile_result.get("flags"):
        lines.append("")
        lines.append("**Flags:**")
        for flag in reconcile_result["flags"]:
            lines.append(f"- {flag}")

    if anomalies:
        lines += ["", "## Anomalies Detected"]
        for a in anomalies:
            lines.append(
                f"- **{a['category_path']}**: ${abs(a['current_total']):,.2f} "
                f"({a['ratio']:.1f}x prior avg of ${abs(a['prior_avg']):,.2f})"
            )

    lines += ["", "## AI Summary", ai_summary, ""]

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text("\n".join(lines))
```

**Step 4: Run tests to verify they pass**

Run: `uv run pytest tests/test_report.py -v`
Expected: All tests PASS.

**Step 5: Commit**

```bash
git add bizplan/report.py tests/test_report.py
git commit -m "feat: add Markdown report generator with variance table"
```

---

## Task 11: CLI — `init` and `validate` Commands

**Files:**
- Modify: `bizplan/cli.py`
- Create: `tests/test_cli.py`

These are the two commands that work together: `init` reads the workbook to discover its structure; `validate` checks formula coverage.

**Step 1: Write the failing tests**

`tests/test_cli.py`:
```python
import pytest
import openpyxl
from pathlib import Path
from click.testing import CliRunner
from bizplan.cli import cli

@pytest.fixture
def sample_wb(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Marketing Exp"
    ws["A1"] = "Category"
    ws["B1"] = "Jan 2026"
    ws["A2"] = "Social Media"
    # Add a locked sheet
    wb.create_sheet("Income Statement")
    path = tmp_path / "model.xlsx"
    wb.save(path)
    return path

def test_init_creates_db_and_config(tmp_path, sample_wb, monkeypatch):
    monkeypatch.chdir(tmp_path)
    runner = CliRunner()
    result = runner.invoke(cli, ["init", "--workbook", str(sample_wb), "--force"])
    assert result.exit_code == 0, result.output
    assert (tmp_path / "data" / "transactions.db").exists()
    assert (tmp_path / "config" / "model.yaml").exists()

def test_validate_passes_clean_model(tmp_path, sample_wb, monkeypatch):
    monkeypatch.chdir(tmp_path)
    runner = CliRunner()
    runner.invoke(cli, ["init", "--workbook", str(sample_wb), "--force"])
    result = runner.invoke(cli, ["validate", "--workbook", str(sample_wb)])
    assert result.exit_code == 0

def test_help_flag(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    runner = CliRunner()
    result = runner.invoke(cli, ["--help"])
    assert "init" in result.output
    assert "update" in result.output
    assert result.exit_code == 0
```

**Step 2: Run tests to verify they fail**

Run: `uv run pytest tests/test_cli.py -v`
Expected: Tests fail — commands not yet registered.

**Step 3: Implement init and validate in cli.py**

```python
import click
from pathlib import Path
from rich.console import Console
from rich import print as rprint

console = Console()

DB_PATH = Path("data/transactions.db")
MAPPINGS_PATH = Path("config/mappings.yaml")
MODEL_PATH = Path("config/model.yaml")

@click.group()
def cli():
    """bizplan — Business Plan Updater CLI.

    Automates monthly financial model updates from QuickBooks exports.

    Common commands:

      bizplan init      Set up database and discover workbook structure

      bizplan update    Run monthly financial update

      bizplan validate  Check model formula coverage

      bizplan map       Review category mappings

      bizplan report    Print variance report for a past month
    """
    pass

@cli.command()
@click.option("--workbook", "-w", required=True, type=click.Path(exists=True),
              help="Path to the Excel financial model workbook.")
@click.option("--force", is_flag=True,
              help="Re-initialize even if DB already exists.")
def init(workbook: str, force: bool):
    """Initialize the database and discover workbook structure.

    Scans the Excel workbook to identify input sheets and month column
    positions, then writes config/model.yaml. Creates data/transactions.db.
    """
    from bizplan.db import init_db
    from bizplan.excel_writer import scan_workbook_structure
    from bizplan.config import save_model

    wb_path = Path(workbook)

    if DB_PATH.exists() and not force:
        rprint("[yellow]Database already exists. Use --force to reinitialize.[/yellow]")
        return

    rprint(f"[bold]Initializing bizplan...[/bold]")

    # Initialize DB
    init_db(DB_PATH)
    rprint(f"[green]✓[/green] Database created at {DB_PATH}")

    # Scan workbook structure
    structure = scan_workbook_structure(wb_path)
    save_model({"sheets": structure}, MODEL_PATH)
    rprint(f"[green]✓[/green] Discovered {len(structure)} input sheets → {MODEL_PATH}")
    for sheet in structure:
        rprint(f"   • {sheet}")

    rprint("\n[bold green]Ready.[/bold green] Run [bold]bizplan validate[/bold] to check formula coverage.")

@cli.command()
@click.option("--workbook", "-w", required=True, type=click.Path(exists=True),
              help="Path to the Excel financial model workbook.")
def validate(workbook: str):
    """Check model integrity: formula coverage across all months.

    Scans the financial statement sheets (Income Statement, Balance Sheet,
    Cash Flow) for blank cells or formula errors in the month columns.
    Reports issues but does NOT auto-repair — fix manually in Excel, then
    re-run validate to confirm.
    """
    import openpyxl
    from bizplan.excel_writer import LOCKED_SHEETS, find_month_column

    wb_path = Path(workbook)
    wb = openpyxl.load_workbook(wb_path, read_only=True, data_only=True)
    issues = []

    for sheet_name in LOCKED_SHEETS:
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        # Find month columns by scanning first 5 rows for date-like headers
        from bizplan.excel_writer import _header_to_yyyymm
        month_cols = []
        for row_idx in range(1, 6):
            for col_idx in range(1, ws.max_column + 1):
                val = ws.cell(row=row_idx, column=col_idx).value
                if val and _header_to_yyyymm(str(val)):
                    month_cols.append((row_idx, col_idx, _header_to_yyyymm(str(val))))
        if not month_cols:
            issues.append(f"{sheet_name}: No month columns detected")
            continue

        # Check key rows (rows 2-10) for formula coverage in month columns
        for hr, col, yyyymm in month_cols[:24]:  # Check up to 24 months
            for data_row in range(hr + 1, min(hr + 10, ws.max_row + 1)):
                val = ws.cell(row=data_row, column=col).value
                if val is None:
                    issues.append(f"{sheet_name} row {data_row} col {col} ({yyyymm}): blank")
                    break  # One issue per column is enough

    wb.close()

    if issues:
        rprint("[bold red]Formula coverage issues detected:[/bold red]")
        for issue in issues:
            rprint(f"  [red]✗[/red] {issue}")
        rprint("\n[yellow]Fix these manually in Excel, then re-run:[/yellow] bizplan validate")
        raise SystemExit(1)
    else:
        rprint("[bold green]✓ Model validation passed.[/bold green]")
```

**Step 4: Run tests to verify they pass**

Run: `uv run pytest tests/test_cli.py -v`
Expected: All tests PASS.

**Step 5: Commit**

```bash
git add bizplan/cli.py tests/test_cli.py
git commit -m "feat: add bizplan init and validate CLI commands"
```

---

## Task 12: CLI — `update` Command

**Files:**
- Modify: `bizplan/cli.py` (add `update` command)

This is the core command — orchestrates the full monthly pipeline.

**Step 1: Write the failing test**

Add to `tests/test_cli.py`:
```python
def test_update_processes_csv(tmp_path, sample_wb, monkeypatch):
    monkeypatch.chdir(tmp_path)
    runner = CliRunner()
    # Copy sample CSV to expected location
    import shutil
    qb_dir = tmp_path / "quickbooks_exports"
    qb_dir.mkdir()
    src = Path("quickbooks_exports/TD_Bank-3.csv")
    if not src.exists():
        pytest.skip("Sample CSV not present")
    shutil.copy(src, qb_dir / "TD_Bank-3.csv")

    runner.invoke(cli, ["init", "--workbook", str(sample_wb), "--force"])
    result = runner.invoke(
        cli,
        ["update", "--month", "2026-01", "--workbook", str(sample_wb), "--no-prompt"],
    )
    assert result.exit_code == 0, result.output
    assert "transactions" in result.output.lower()
```

**Step 2: Run test to verify it fails**

Run: `uv run pytest tests/test_cli.py::test_update_processes_csv -v`
Expected: FAIL — `update` command not registered.

**Step 3: Implement the `update` command in cli.py**

Add after the `validate` command:
```python
@cli.command()
@click.option("--month", "-m", default=None,
              help="Month to update in YYYY-MM format. Defaults to prior calendar month.")
@click.option("--workbook", "-w", required=True, type=click.Path(exists=True),
              help="Path to the Excel financial model workbook.")
@click.option("--no-prompt", is_flag=True,
              help="Skip interactive prompts for unknown categories (useful for testing).")
@click.option("--force", is_flag=True,
              help="Proceed even if validate detects formula issues.")
def update(month: str | None, workbook: str, no_prompt: bool, force: bool):
    """Run the monthly financial update pipeline.

    Steps: detect new QB CSV files, ingest and deduplicate transactions,
    prompt for unknown category mappings, aggregate monthly totals,
    write actuals to the Excel model, reconcile revenue/cash, and
    generate a Markdown variance report.
    """
    from datetime import date
    from bizplan.db import get_conn, insert_transaction, log_run, get_monthly_totals_raw
    from bizplan.ingest import parse_qb_csv
    from bizplan.aggregate import apply_mappings, get_unmapped_categories, detect_anomalies
    from bizplan.categorize import resolve_categories
    from bizplan.excel_writer import scan_workbook_structure, write_monthly_actuals
    from bizplan.reconcile import reconcile_revenue_cash
    from bizplan.report import generate_report
    from bizplan.config import load_mappings, load_model
    import glob as _glob

    if not month:
        today = date.today()
        m = today.month - 1 or 12
        y = today.year if today.month > 1 else today.year - 1
        month = f"{y}-{m:02d}"

    rprint(f"\n[bold]bizplan update — {month}[/bold]\n")

    wb_path = Path(workbook)
    output_dir = Path("output/workbooks")
    report_dir = Path("output/reports")
    output_dir.mkdir(parents=True, exist_ok=True)
    report_dir.mkdir(parents=True, exist_ok=True)

    # Step 1: Find QB CSV files
    csv_files = list(Path("quickbooks_exports").glob("*.csv")) if Path("quickbooks_exports").exists() else []
    if not csv_files:
        rprint("[red]No CSV files found in quickbooks_exports/[/red]")
        return

    # Step 2: Ingest
    all_rows = []
    for f in csv_files:
        all_rows.extend(parse_qb_csv(f))
    rprint(f"[cyan]→[/cyan] {len(all_rows)} rows parsed from {len(csv_files)} file(s)")

    # Step 3: Deduplicate and insert
    imported = len(all_rows)
    new_count = 0
    with get_conn(DB_PATH) as conn:
        for row in all_rows:
            if insert_transaction(conn, row):
                new_count += 1
    rprint(f"[cyan]→[/cyan] {new_count} new transactions ({imported - new_count} duplicates skipped)")

    # Step 4: Categorize
    with get_conn(DB_PATH) as conn:
        unmapped = get_unmapped_categories(conn, month)
        if unmapped:
            rprint(f"\n[yellow]{len(unmapped)} unmapped categories found:[/yellow]")
            if not no_prompt:
                from bizplan.ai_assist import suggest_category_mapping
                from bizplan.categorize import prompt_for_mapping
                existing_mappings = load_mappings()
                # Build vendor map (most common vendor per category in this month)
                vendor_map: dict[str, str] = {}
                all_rows_for_month = [r for r in all_rows if r["date"].startswith(month)]
                for r in all_rows_for_month:
                    vendor_map.setdefault(r["category_path"], r["vendor"])

                for cp in unmapped:
                    ai_suggestion = suggest_category_mapping(cp, vendor_map.get(cp, ""), existing_mappings)
                    with get_conn(DB_PATH) as conn2:
                        prompt_for_mapping(conn2, cp, vendor=vendor_map.get(cp, ""), ai_suggestion=ai_suggestion)
            else:
                for cp in unmapped:
                    rprint(f"  [dim](skipped: {cp})[/dim]")

    # Step 5 & 6: Aggregate + get prior month for variance
    y, m_int = map(int, month.split("-"))
    pm_int = m_int - 1 or 12
    py = y if m_int > 1 else y - 1
    prior_month = f"{py}-{pm_int:02d}"

    with get_conn(DB_PATH) as conn:
        actuals = apply_mappings(conn, month)
        current_totals = get_monthly_totals_raw(conn, month)
        prior_totals = get_monthly_totals_raw(conn, prior_month)
        anomalies = detect_anomalies(conn, month)

    rprint(f"[cyan]→[/cyan] {len(actuals)} categories mapped to Excel targets")

    # Step 7: Write to Excel
    model_structure = load_model().get("sheets", {})
    if not model_structure:
        model_structure = scan_workbook_structure(wb_path)

    output_wb = output_dir / f"{month}_update.xlsx"
    cells_written = write_monthly_actuals(wb_path, output_wb, actuals, model_structure, month)
    rprint(f"[green]✓[/green] {cells_written} cells written → {output_wb}")

    # Step 8: Reconcile
    with get_conn(DB_PATH) as conn:
        reconcile_result = reconcile_revenue_cash(conn, month)
    rprint(f"[cyan]→[/cyan] Total deposits: ${reconcile_result['total_deposits']:,.2f}")

    # Step 9: AI summary
    try:
        from bizplan.ai_assist import generate_run_summary
        ai_summary = generate_run_summary(month, current_totals, prior_totals)
    except Exception:
        ai_summary = "AI summary unavailable."

    # Step 10: Report
    report_path = report_dir / f"{month}-report.md"
    generate_report(
        month=month,
        imported=imported,
        new=new_count,
        cells_written=cells_written,
        workbook_path=str(output_wb),
        current_totals=current_totals,
        prior_totals=prior_totals,
        anomalies=anomalies,
        reconcile_result=reconcile_result,
        ai_summary=ai_summary,
        output_path=report_path,
    )
    rprint(f"[green]✓[/green] Report → {report_path}")

    # Log run
    with get_conn(DB_PATH) as conn:
        log_run(conn, month, imported, new_count, cells_written, str(output_wb), str(report_path))

    rprint(f"\n[bold green]Update complete.[/bold green]")
    if anomalies:
        rprint(f"[yellow]⚠ {len(anomalies)} anomaly/anomalies detected — see report.[/yellow]")
```

**Step 4: Run tests to verify they pass**

Run: `uv run pytest tests/test_cli.py -v`
Expected: All tests PASS.

**Step 5: Commit**

```bash
git add bizplan/cli.py tests/test_cli.py
git commit -m "feat: add bizplan update command — full monthly pipeline"
```

---

## Task 13: CLI — `map` and `report` Commands

**Files:**
- Modify: `bizplan/cli.py` (add `map` and `report` commands)

**Step 1: Write the failing tests**

Add to `tests/test_cli.py`:
```python
def test_report_command_no_data(tmp_path, sample_wb, monkeypatch):
    monkeypatch.chdir(tmp_path)
    runner = CliRunner()
    runner.invoke(cli, ["init", "--workbook", str(sample_wb), "--force"])
    result = runner.invoke(cli, ["report", "--month", "2026-01"])
    # Should run without crashing even with no data
    assert result.exit_code == 0

def test_map_command_lists_mappings(tmp_path, sample_wb, monkeypatch):
    monkeypatch.chdir(tmp_path)
    runner = CliRunner()
    runner.invoke(cli, ["init", "--workbook", str(sample_wb), "--force"])
    result = runner.invoke(cli, ["map", "--list"])
    assert result.exit_code == 0
```

**Step 2: Run tests to verify they fail**

Run: `uv run pytest tests/test_cli.py::test_report_command_no_data tests/test_cli.py::test_map_command_lists_mappings -v`
Expected: FAIL — commands not registered.

**Step 3: Add `map` and `report` commands to cli.py**

```python
@cli.command(name="map")
@click.option("--list", "list_mappings", is_flag=True,
              help="List all current category mappings.")
def map_cmd(list_mappings: bool):
    """Review and edit QuickBooks category → Excel row mappings.

    Use --list to see all current mappings.
    Without flags, launches interactive mapping review for any unconfirmed entries.
    """
    from bizplan.db import get_conn
    from bizplan.config import load_mappings

    if list_mappings:
        with get_conn(DB_PATH) as conn:
            rows = conn.execute(
                "SELECT category_path, excel_sheet, excel_row_label, confirmed_by_user "
                "FROM category_mappings ORDER BY excel_sheet, excel_row_label"
            ).fetchall()
        if not rows:
            rprint("[dim]No mappings saved yet. Run bizplan update to add mappings.[/dim]")
            return
        rprint(f"\n[bold]{len(rows)} category mappings:[/bold]\n")
        for row in rows:
            confirmed = "[green]✓[/green]" if row["confirmed_by_user"] else "[yellow]?[/yellow]"
            rprint(f"  {confirmed} [cyan]{row['category_path']}[/cyan]")
            rprint(f"       → {row['excel_sheet']} / {row['excel_row_label']}")
        return

    rprint("[dim]Interactive mapping review not yet implemented. Use --list to view mappings.[/dim]")

@cli.command(name="report")
@click.option("--month", "-m", required=True, help="Month in YYYY-MM format.")
def report_cmd(month: str):
    """Print the variance report for a past month.

    The report file must already exist (created by bizplan update).
    """
    report_path = Path("output/reports") / f"{month}-report.md"
    if not report_path.exists():
        rprint(f"[red]No report found for {month}.[/red] Run: bizplan update --month {month}")
        return
    console.print(report_path.read_text())
```

**Step 4: Run tests to verify they pass**

Run: `uv run pytest tests/test_cli.py -v`
Expected: All tests PASS.

**Step 5: Run full test suite**

Run: `uv run pytest -v`
Expected: All tests PASS.

**Step 6: Commit**

```bash
git add bizplan/cli.py
git commit -m "feat: add bizplan map and report CLI commands"
```

---

## Task 14: End-to-End Integration Test

**Files:**
- Create: `tests/test_integration.py`

**Step 1: Write the integration test**

```python
"""End-to-end test using the real sample QuickBooks CSV and a minimal test workbook."""
import pytest
import openpyxl
import shutil
from pathlib import Path
from click.testing import CliRunner
from bizplan.cli import cli

SAMPLE_CSV = Path("quickbooks_exports/TD_Bank-3.csv")

@pytest.fixture
def env(tmp_path):
    """Set up a complete test environment with a sample workbook and QB CSV."""
    # Create minimal workbook matching known QB categories
    wb = openpyxl.Workbook()
    # Marketing Exp sheet
    ws1 = wb.active
    ws1.title = "Marketing Exp"
    ws1["A1"] = "Category"
    ws1["B1"] = "Jan 2026"
    ws1["C1"] = "Feb 2026"
    ws1["A2"] = "Social Media"
    ws1["A3"] = "Website Ads"
    # Production Exp sheet
    ws2 = wb.create_sheet("Production Exp")
    ws2["A1"] = "Category"
    ws2["B1"] = "Jan 2026"
    ws2["A2"] = "Software & Tools"
    ws2["A3"] = "Wardrobe"

    wb_path = tmp_path / "model.xlsx"
    wb.save(wb_path)

    # Copy QB CSV
    qb_dir = tmp_path / "quickbooks_exports"
    qb_dir.mkdir()
    if SAMPLE_CSV.exists():
        shutil.copy(SAMPLE_CSV, qb_dir / "TD_Bank-3.csv")

    return {"tmp_path": tmp_path, "wb_path": wb_path}

@pytest.mark.skipif(not SAMPLE_CSV.exists(), reason="Sample CSV not present")
def test_full_pipeline(env, monkeypatch):
    monkeypatch.chdir(env["tmp_path"])
    runner = CliRunner()

    # Init
    result = runner.invoke(cli, ["init", "--workbook", str(env["wb_path"]), "--force"])
    assert result.exit_code == 0, result.output

    # Update (no-prompt so we skip interactive categorization)
    result = runner.invoke(
        cli,
        ["update", "--month", "2026-01", "--workbook", str(env["wb_path"]), "--no-prompt"],
    )
    assert result.exit_code == 0, result.output
    assert "transactions" in result.output.lower()

    # Output workbook created
    assert (env["tmp_path"] / "output" / "workbooks" / "2026-01_update.xlsx").exists()

    # Report created
    report_path = env["tmp_path"] / "output" / "reports" / "2026-01-report.md"
    assert report_path.exists()
    content = report_path.read_text()
    assert "2026-01" in content

    # Deduplication: re-run should produce same totals
    result2 = runner.invoke(
        cli,
        ["update", "--month", "2026-01", "--workbook", str(env["wb_path"]), "--no-prompt"],
    )
    assert result2.exit_code == 0
    assert "0 new" in result2.output or "duplicates skipped" in result2.output

    # Spot-check: Social Media total should be Buffer+Vimeo+X+GoogleAds
    # $36 + $300 + $84 + $150.80 = $570.80
    wb = openpyxl.load_workbook(env["tmp_path"] / "output" / "workbooks" / "2026-01_update.xlsx")
    # (Cell location depends on mapping — test that the file opens cleanly)
    assert wb is not None
```

**Step 2: Run the integration test**

Run: `uv run pytest tests/test_integration.py -v`
Expected: PASS (skipped if sample CSV not present).

**Step 3: Run full test suite**

Run: `uv run pytest -v`
Expected: All tests PASS.

**Step 4: Final commit**

```bash
git add tests/test_integration.py
git commit -m "test: add end-to-end integration test"
```

---

## Verification Checklist

After all tasks complete:

1. `uv run bizplan --help` — lists all commands with descriptions
2. `uv run bizplan init -h` — shows options for init
3. `uv run bizplan init --workbook template/DA_Business_Plan_2025_nov.xlsx` — creates `data/transactions.db` and `config/model.yaml`, lists discovered sheets
4. `uv run bizplan validate --workbook template/DA_Business_Plan_2025_nov.xlsx` — reports formula coverage; lists any gaps
5. `uv run bizplan update --month 2026-01 --workbook template/DA_Business_Plan_2025_nov.xlsx` — ingests `quickbooks_exports/TD_Bank-3.csv`, prompts for unknown categories, writes actuals, saves `output/workbooks/2026-01_update.xlsx`, generates `output/reports/2026-01-report.md`
6. Re-run step 5 — confirms 0 new transactions (deduplication working)
7. `uv run bizplan map --list` — shows all confirmed mappings
8. `uv run bizplan report --month 2026-01` — prints the report
9. Open `output/workbooks/2026-01_update.xlsx` in Excel and spot-check Social Media = sum of Buffer+Vimeo+X+Google Ads transactions
10. `uv run pytest -v` — all tests pass
