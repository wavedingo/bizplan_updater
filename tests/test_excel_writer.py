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
