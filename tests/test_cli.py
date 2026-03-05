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
    assert result.exit_code == 0

def test_update_processes_csv(tmp_path, sample_wb, monkeypatch):
    monkeypatch.chdir(tmp_path)
    runner = CliRunner()
    # Copy sample CSV to expected location
    import shutil
    qb_dir = tmp_path / "quickbooks_exports"
    qb_dir.mkdir()
    src = Path("/Users/jrv/Documents/Code/projects/bizplan_updater/quickbooks_exports/TD_Bank-3.csv")
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

def test_report_command_no_data(tmp_path, sample_wb, monkeypatch):
    monkeypatch.chdir(tmp_path)
    runner = CliRunner()
    runner.invoke(cli, ["init", "--workbook", str(sample_wb), "--force"])
    result = runner.invoke(cli, ["report", "--month", "2026-01"])
    # Should run without crashing even with no report file
    assert result.exit_code == 0

def test_map_command_lists_mappings(tmp_path, sample_wb, monkeypatch):
    monkeypatch.chdir(tmp_path)
    runner = CliRunner()
    runner.invoke(cli, ["init", "--workbook", str(sample_wb), "--force"])
    result = runner.invoke(cli, ["map", "--list"])
    assert result.exit_code == 0

def test_h_flag_on_group(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    runner = CliRunner()
    result = runner.invoke(cli, ["-h"])
    assert result.exit_code == 0, result.output
    assert "init" in result.output

def test_h_flag_on_subcommand(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    runner = CliRunner()
    result = runner.invoke(cli, ["map", "-h"])
    assert result.exit_code == 0, result.output
