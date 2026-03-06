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

def test_parse_qb_csv_duplicate_within_file_gets_sequence(tmp_path):
    """Two identical transactions in one file get distinct fingerprints."""
    csv_content = '''date,Bank description,Amount,From/To,Transaction Posted
"01/09/2026","Online Xfer Transfer","-$5,000.00","","Added to:  Expense: Expense:Partner distributions 01/09/2026 $5,000.00"
"01/09/2026","Online Xfer Transfer","-$5,000.00","","Added to:  Expense: Expense:Partner distributions 01/09/2026 $5,000.00"
'''
    csv_file = tmp_path / "test.csv"
    csv_file.write_text(csv_content)
    rows = parse_qb_csv(csv_file)
    assert len(rows) == 2
    assert rows[0]["fingerprint"] != rows[1]["fingerprint"]
    assert rows[1]["fingerprint"].endswith("__2")


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
