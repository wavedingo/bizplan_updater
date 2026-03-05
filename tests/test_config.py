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
