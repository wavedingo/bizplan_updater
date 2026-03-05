import sqlite3
from bizplan.db import get_monthly_totals_raw, get_category_mapping

def aggregate_month(conn: sqlite3.Connection, month: str) -> dict[str, float]:
    """Return {category_path: sum_amount} for the given YYYY-MM month.
    Full recomputation from DB — implements Rules 1, 3, 4, 5."""
    return get_monthly_totals_raw(conn, month)

def apply_mappings(conn: sqlite3.Connection, month: str) -> list[dict]:
    """Return aggregated totals enriched with Excel sheet/row targets.
    Only returns categories that have confirmed mappings."""
    totals = aggregate_month(conn, month)
    result = []
    for category_path, amount in totals.items():
        mapping = get_category_mapping(conn, category_path)
        if mapping:
            result.append({
                "category_path": category_path,
                "amount": amount,
                "excel_sheet": mapping["excel_sheet"],
                "excel_row_label": mapping["excel_row_label"],
            })
    return result

def get_unmapped_categories(conn: sqlite3.Connection, month: str) -> list[str]:
    """Return category_paths in month that have no mapping in category_mappings."""
    totals = aggregate_month(conn, month)
    unmapped = []
    for category_path in totals:
        if get_category_mapping(conn, category_path) is None:
            unmapped.append(category_path)
    return unmapped

def detect_anomalies(conn: sqlite3.Connection, month: str,
                     spike_threshold: float = 3.0) -> list[dict]:
    """Flag categories where this month's total is > spike_threshold * prior 2-month avg."""
    current = aggregate_month(conn, month)
    anomalies = []
    y, m = map(int, month.split("-"))

    def prior_month(offset: int) -> str:
        mo = m - offset
        yr = y
        if mo <= 0:
            mo += 12
            yr -= 1
        return f"{yr}-{mo:02d}"

    for category_path, current_total in current.items():
        prior_totals = []
        for offset in (1, 2):
            pm = prior_month(offset)
            pm_totals = get_monthly_totals_raw(conn, pm)
            if category_path in pm_totals:
                prior_totals.append(abs(pm_totals[category_path]))

        if prior_totals:
            avg = sum(prior_totals) / len(prior_totals)
            if avg > 0 and abs(current_total) > spike_threshold * avg:
                anomalies.append({
                    "category_path": category_path,
                    "current_total": current_total,
                    "prior_avg": avg,
                    "ratio": abs(current_total) / avg,
                })
    return anomalies
