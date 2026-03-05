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
