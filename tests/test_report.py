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
