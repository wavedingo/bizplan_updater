import sqlite3

# Category paths that represent cash inflows (deposits)
_DEPOSIT_CATEGORIES = frozenset({
    "Sales:Host Read Ad Sales",
    "Sales:Subscription Sales",
    "Sales:Programmatic Ad Revenue",
    "Commissions & fees:Network Ad Revenue Distribution",
})

def reconcile_revenue_cash(conn: sqlite3.Connection, month: str) -> dict:
    """Summarize deposits received vs. expected revenue for the month.

    Returns a dict with total_deposits, by_category, and any timing flags.
    """
    rows = conn.execute(
        """SELECT category_path, SUM(amount) AS total
           FROM transactions
           WHERE strftime('%Y-%m', date) = ?
             AND txn_type = 'Deposit'
           GROUP BY category_path""",
        (month,),
    ).fetchall()

    by_category = {row["category_path"]: row["total"] for row in rows}
    total_deposits = sum(by_category.values())

    flags = []
    for cat, total in by_category.items():
        if cat not in _DEPOSIT_CATEGORIES:
            flags.append(f"Unrecognized deposit category: {cat!r} (${total:,.2f})")

    return {
        "month": month,
        "total_deposits": total_deposits,
        "by_category": by_category,
        "flags": flags,
    }
