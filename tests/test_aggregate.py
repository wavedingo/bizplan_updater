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
