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
