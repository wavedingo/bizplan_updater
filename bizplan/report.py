from pathlib import Path
from datetime import datetime

def format_variance_table(current: dict, prior: dict) -> str:
    """Format a Markdown variance table comparing current vs prior month totals."""
    lines = [
        "| Category | Prior | Actual | Variance |",
        "|---|---|---|---|",
    ]
    all_cats = sorted(set(current) | set(prior))
    for cat in all_cats:
        cur = current.get(cat, 0.0)
        prr = prior.get(cat, 0.0)
        var = cur - prr
        direction = "▲" if var > 0 else ("▼" if var < 0 else "—")
        lines.append(
            f"| {cat} | ${abs(prr):,.2f} | ${abs(cur):,.2f} | {direction} ${abs(var):,.2f} |"
        )
    return "\n".join(lines)

def generate_report(
    month: str,
    imported: int,
    new: int,
    cells_written: int,
    workbook_path: str,
    current_totals: dict,
    prior_totals: dict,
    anomalies: list,
    reconcile_result: dict,
    ai_summary: str,
    output_path: Path,
) -> None:
    """Write a Markdown run report to output_path."""
    lines = [
        f"# Business Plan Update — {month}",
        f"",
        f"_Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}_",
        f"",
        f"## Transactions",
        f"- {imported} imported, {imported - new} duplicates skipped, {new} new",
        f"",
        f"## Excel Updates",
        f"- {cells_written} cells written",
        f"- Workbook: `{workbook_path}`",
        f"",
        f"## Variance vs Prior Month",
        f"",
        format_variance_table(current_totals, prior_totals),
        f"",
        f"## Revenue & Cash Reconciliation",
        f"- Total deposits received: ${reconcile_result['total_deposits']:,.2f}",
    ]

    if reconcile_result.get("flags"):
        lines.append("")
        lines.append("**Flags:**")
        for flag in reconcile_result["flags"]:
            lines.append(f"- {flag}")

    if anomalies:
        lines += ["", "## Anomalies Detected"]
        for a in anomalies:
            lines.append(
                f"- **{a['category_path']}**: ${abs(a['current_total']):,.2f} "
                f"({a['ratio']:.1f}x prior avg of ${abs(a['prior_avg']):,.2f})"
            )

    lines += ["", "## AI Summary", ai_summary, ""]

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text("\n".join(lines))
