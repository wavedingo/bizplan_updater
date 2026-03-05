import json
import anthropic

_MODEL = "claude-haiku-4-5-20251001"  # Fast and cheap for short classification tasks

def _client() -> anthropic.Anthropic:
    return anthropic.Anthropic()  # Reads ANTHROPIC_API_KEY from environment

def suggest_category_mapping(
    category_path: str,
    vendor: str,
    existing_mappings: dict,
) -> dict | None:
    """Ask Claude to suggest an Excel sheet and row_label for an unknown category.

    Returns {"sheet": ..., "row_label": ...} or None if suggestion fails.
    """
    existing_str = json.dumps(existing_mappings, indent=2)
    prompt = f"""You are helping map QuickBooks categories to Excel model rows.

Unknown category: "{category_path}"
Vendor name: "{vendor}"

Existing mappings for reference:
{existing_str}

Respond with ONLY a JSON object: {{"sheet": "<sheet name>", "row_label": "<row label>"}}
Choose the most appropriate sheet and row_label from the existing mappings, or suggest a new one."""

    try:
        response = _client().messages.create(
            model=_MODEL,
            max_tokens=128,
            messages=[{"role": "user", "content": prompt}],
        )
        text = response.content[0].text.strip()
        return json.loads(text)
    except Exception:
        return None

def generate_run_summary(month: str, totals: dict, prior_totals: dict) -> str:
    """Generate a 2-3 sentence narrative summary of the monthly run."""
    prompt = f"""Summarize the key financial trends for {month} in 2-3 sentences.
Current month totals: {json.dumps(totals)}
Prior month totals: {json.dumps(prior_totals)}
Focus on notable changes, not every line item. Be factual and concise."""

    try:
        response = _client().messages.create(
            model=_MODEL,
            max_tokens=256,
            messages=[{"role": "user", "content": prompt}],
        )
        return response.content[0].text.strip()
    except Exception:
        return "AI summary unavailable."

def explain_anomaly(category_path: str, current_total: float, prior_avg: float) -> str:
    """Return a one-sentence explanation for an anomalous category."""
    ratio = abs(current_total) / abs(prior_avg) if prior_avg != 0 else 0
    prompt = f"""Explain in one sentence why "{category_path}" might be {ratio:.1f}x higher than usual.
Current: ${abs(current_total):,.2f}  Prior avg: ${abs(prior_avg):,.2f}
Be concise and suggest a likely business reason."""

    try:
        response = _client().messages.create(
            model=_MODEL,
            max_tokens=128,
            messages=[{"role": "user", "content": prompt}],
        )
        return response.content[0].text.strip()
    except Exception:
        return f"Anomaly: {ratio:.1f}x prior average — review recommended."
