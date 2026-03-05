import re
from pathlib import Path
import pandas as pd

# Pattern: "Added to:  {Type}: {category_path} {MM/DD/YYYY} ${amount}"
_POSTED_RE = re.compile(
    r"Added to:\s+(\w+):\s+(.+?)\s+\d{2}/\d{2}/\d{4}\s+\$[\d,]+\.?\d*$"
)

def parse_amount(amount_str: str) -> float:
    """Parse amount strings like '-$44.99' or '$1,781.47' to float."""
    clean = amount_str.replace("$", "").replace(",", "").strip()
    return float(clean)

def parse_transaction_posted(posted: str) -> dict:
    """Parse the 'Transaction Posted' column to extract txn_type and category_path."""
    m = _POSTED_RE.match(posted.strip())
    if not m:
        raise ValueError(f"Cannot parse Transaction Posted: {posted!r}")
    return {"txn_type": m.group(1), "category_path": m.group(2).strip()}

def parse_qb_csv(csv_path: Path) -> list[dict]:
    """Parse a QuickBooks CSV export into a list of transaction dicts."""
    df = pd.read_csv(csv_path, dtype=str, keep_default_na=False)
    rows = []
    for _, row in df.iterrows():
        raw_date = row["date"].strip()          # MM/DD/YYYY
        vendor = row["Bank description"].strip()
        raw_amount = row["Amount"].strip()
        posted = row["Transaction Posted"].strip()

        if not posted.startswith("Added to:"):
            continue  # Skip unrecognized rows

        try:
            amount = parse_amount(raw_amount)
            parsed = parse_transaction_posted(posted)
            # Convert date MM/DD/YYYY -> YYYY-MM-DD
            m, d, y = raw_date.split("/")
            date = f"{y}-{m}-{d}"
            category_path = parsed["category_path"]
            fingerprint = f"{date}|{vendor}|{abs(amount):.2f}|{category_path}"
            rows.append({
                "fingerprint": fingerprint,
                "date": date,
                "vendor": vendor,
                "amount": amount,
                "category_path": category_path,
                "txn_type": parsed["txn_type"],
            })
        except (ValueError, KeyError):
            continue  # Skip rows that can't be parsed

    return rows
