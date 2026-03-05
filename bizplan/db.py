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
