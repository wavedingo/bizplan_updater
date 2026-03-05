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
