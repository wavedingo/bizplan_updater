import pytest
from unittest.mock import MagicMock, patch
from bizplan.ai_assist import (
    suggest_category_mapping,
    generate_run_summary,
    explain_anomaly,
)

@patch("bizplan.ai_assist.anthropic.Anthropic")
def test_suggest_category_mapping(mock_anthropic_cls):
    mock_client = MagicMock()
    mock_anthropic_cls.return_value = mock_client
    mock_client.messages.create.return_value = MagicMock(
        content=[MagicMock(text='{"sheet": "Marketing Exp", "row_label": "Social Media"}')]
    )
    result = suggest_category_mapping(
        category_path="Advertising & marketing:Social media",
        vendor="Ap Vmo Vimeo",
        existing_mappings={"Advertising & marketing:Website ads": {"sheet": "Marketing Exp", "row_label": "Website Ads"}},
    )
    assert result["sheet"] == "Marketing Exp"
    assert result["row_label"] == "Social Media"

@patch("bizplan.ai_assist.anthropic.Anthropic")
def test_generate_run_summary(mock_anthropic_cls):
    mock_client = MagicMock()
    mock_anthropic_cls.return_value = mock_client
    mock_client.messages.create.return_value = MagicMock(
        content=[MagicMock(text="Subscription revenue was strong this month.")]
    )
    result = generate_run_summary(
        month="2026-01",
        totals={"Sales:Subscription Sales": 8021.0},
        prior_totals={"Sales:Subscription Sales": 7500.0},
    )
    assert isinstance(result, str)
    assert len(result) > 0

@patch("bizplan.ai_assist.anthropic.Anthropic")
def test_explain_anomaly(mock_anthropic_cls):
    mock_client = MagicMock()
    mock_anthropic_cls.return_value = mock_client
    mock_client.messages.create.return_value = MagicMock(
        content=[MagicMock(text="Q4 true-up payment likely.")]
    )
    result = explain_anomaly(
        category_path="Expense:Partner distributions",
        current_total=-15000.0,
        prior_avg=-5000.0,
    )
    assert isinstance(result, str)
