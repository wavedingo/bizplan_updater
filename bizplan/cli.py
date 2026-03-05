import click
from pathlib import Path
from rich.console import Console
from rich import print as rprint

console = Console()

DB_PATH = Path("data/transactions.db")
MAPPINGS_PATH = Path("config/mappings.yaml")
MODEL_PATH = Path("config/model.yaml")


@click.group()
def cli():
    """bizplan — Business Plan Updater CLI.

    Automates monthly financial model updates from QuickBooks exports.

    Common commands:

      bizplan init      Set up database and discover workbook structure

      bizplan update    Run monthly financial update

      bizplan validate  Check model formula coverage

      bizplan map       Review category mappings

      bizplan report    Print variance report for a past month
    """
    pass


@cli.command()
@click.option("--workbook", "-w", required=True, type=click.Path(exists=True),
              help="Path to the Excel financial model workbook.")
@click.option("--force", is_flag=True,
              help="Re-initialize even if DB already exists.")
def init(workbook: str, force: bool):
    """Initialize the database and discover workbook structure.

    Scans the Excel workbook to identify input sheets and month column
    positions, then writes config/model.yaml. Creates data/transactions.db.
    """
    from bizplan.db import init_db
    from bizplan.excel_writer import scan_workbook_structure
    from bizplan.config import save_model

    wb_path = Path(workbook)

    if DB_PATH.exists() and not force:
        rprint("[yellow]Database already exists. Use --force to reinitialize.[/yellow]")
        return

    rprint(f"[bold]Initializing bizplan...[/bold]")

    # Initialize DB
    init_db(DB_PATH)
    rprint(f"[green]✓[/green] Database created at {DB_PATH}")

    # Scan workbook structure
    structure = scan_workbook_structure(wb_path)
    save_model({"sheets": structure}, MODEL_PATH)
    rprint(f"[green]✓[/green] Discovered {len(structure)} input sheets → {MODEL_PATH}")
    for sheet in structure:
        rprint(f"   • {sheet}")

    rprint("\n[bold green]Ready.[/bold green] Run [bold]bizplan validate[/bold] to check formula coverage.")


@cli.command()
@click.option("--workbook", "-w", required=True, type=click.Path(exists=True),
              help="Path to the Excel financial model workbook.")
def validate(workbook: str):
    """Check model integrity: formula coverage across all months.

    Scans the financial statement sheets (Income Statement, Balance Sheet,
    Cash Flow) for blank cells or formula errors in the month columns.
    Reports issues but does NOT auto-repair — fix manually in Excel, then
    re-run validate to confirm.
    """
    import openpyxl
    from bizplan.excel_writer import LOCKED_SHEETS, _header_to_yyyymm

    wb_path = Path(workbook)
    wb = openpyxl.load_workbook(wb_path, read_only=True, data_only=True)
    issues = []

    for sheet_name in LOCKED_SHEETS:
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        # Find month columns by scanning first 5 rows for date-like headers
        month_cols = []
        for row_idx in range(1, 6):
            for col_idx in range(1, ws.max_column + 1):
                val = ws.cell(row=row_idx, column=col_idx).value
                if val and _header_to_yyyymm(str(val)):
                    month_cols.append((row_idx, col_idx, _header_to_yyyymm(str(val))))
        if not month_cols:
            # Sheet exists but has no month columns yet — skip rather than error
            continue

        # Check key rows (rows 2-10) for formula coverage in month columns
        for hr, col, yyyymm in month_cols[:24]:  # Check up to 24 months
            for data_row in range(hr + 1, min(hr + 10, ws.max_row + 1)):
                val = ws.cell(row=data_row, column=col).value
                if val is None:
                    issues.append(f"{sheet_name} row {data_row} col {col} ({yyyymm}): blank")
                    break  # One issue per column is enough

    wb.close()

    if issues:
        rprint("[bold red]Formula coverage issues detected:[/bold red]")
        for issue in issues:
            rprint(f"  [red]✗[/red] {issue}")
        rprint("\n[yellow]Fix these manually in Excel, then re-run:[/yellow] bizplan validate")
        raise SystemExit(1)
    else:
        rprint("[bold green]✓ Model validation passed.[/bold green]")


@cli.command()
@click.option("--month", "-m", default=None,
              help="Month to update in YYYY-MM format. Defaults to prior calendar month.")
@click.option("--workbook", "-w", required=True, type=click.Path(exists=True),
              help="Path to the Excel financial model workbook.")
@click.option("--no-prompt", is_flag=True,
              help="Skip interactive prompts for unknown categories (useful for testing).")
@click.option("--force", is_flag=True,
              help="Proceed even if validate detects formula issues.")
def update(month: str | None, workbook: str, no_prompt: bool, force: bool):
    """Run the monthly financial update pipeline.

    Steps: detect new QB CSV files, ingest and deduplicate transactions,
    prompt for unknown category mappings, aggregate monthly totals,
    write actuals to the Excel model, reconcile revenue/cash, and
    generate a Markdown variance report.
    """
    from datetime import date
    from bizplan.db import get_conn, insert_transaction, log_run, get_monthly_totals_raw
    from bizplan.ingest import parse_qb_csv
    from bizplan.aggregate import apply_mappings, get_unmapped_categories, detect_anomalies
    from bizplan.categorize import prompt_for_mapping
    from bizplan.excel_writer import scan_workbook_structure, write_monthly_actuals
    from bizplan.reconcile import reconcile_revenue_cash
    from bizplan.report import generate_report
    from bizplan.config import load_mappings, load_model

    if not month:
        today = date.today()
        m = today.month - 1 or 12
        y = today.year if today.month > 1 else today.year - 1
        month = f"{y}-{m:02d}"

    rprint(f"\n[bold]bizplan update — {month}[/bold]\n")

    wb_path = Path(workbook)
    output_dir = Path("output/workbooks")
    report_dir = Path("output/reports")
    output_dir.mkdir(parents=True, exist_ok=True)
    report_dir.mkdir(parents=True, exist_ok=True)

    # Step 1: Find QB CSV files
    qb_dir = Path("quickbooks_exports")
    csv_files = list(qb_dir.glob("*.csv")) if qb_dir.exists() else []
    if not csv_files:
        rprint("[red]No CSV files found in quickbooks_exports/[/red]")
        return

    # Step 2: Ingest
    all_rows = []
    for f in csv_files:
        all_rows.extend(parse_qb_csv(f))
    rprint(f"[cyan]→[/cyan] {len(all_rows)} rows parsed from {len(csv_files)} file(s)")

    # Step 3: Deduplicate and insert
    imported = len(all_rows)
    new_count = 0
    with get_conn(DB_PATH) as conn:
        for row in all_rows:
            if insert_transaction(conn, row):
                new_count += 1
    rprint(f"[cyan]→[/cyan] {new_count} new transactions ({imported - new_count} duplicates skipped)")

    # Step 4: Categorize
    with get_conn(DB_PATH) as conn:
        unmapped = get_unmapped_categories(conn, month)
        if unmapped and not no_prompt:
            rprint(f"\n[yellow]{len(unmapped)} unmapped categories found:[/yellow]")
            try:
                from bizplan.ai_assist import suggest_category_mapping
                existing_mappings = load_mappings()
                # Build vendor map (first vendor seen per category in this month)
                vendor_map: dict[str, str] = {}
                for r in all_rows:
                    if r["date"].startswith(month):
                        vendor_map.setdefault(r["category_path"], r["vendor"])
                for cp in unmapped:
                    ai_suggestion = suggest_category_mapping(cp, vendor_map.get(cp, ""), existing_mappings)
                    with get_conn(DB_PATH) as conn2:
                        prompt_for_mapping(conn2, cp, vendor=vendor_map.get(cp, ""), ai_suggestion=ai_suggestion)
            except Exception as e:
                rprint(f"[yellow]Category prompting error: {e}[/yellow]")
        elif unmapped and no_prompt:
            rprint(f"[dim]{len(unmapped)} unmapped categories skipped (--no-prompt)[/dim]")

    # Step 5 & 6: Aggregate + get prior month for variance
    y_int, m_int = map(int, month.split("-"))
    pm_int = m_int - 1 or 12
    py = y_int if m_int > 1 else y_int - 1
    prior_month = f"{py}-{pm_int:02d}"

    with get_conn(DB_PATH) as conn:
        actuals = apply_mappings(conn, month)
        current_totals = get_monthly_totals_raw(conn, month)
        prior_totals = get_monthly_totals_raw(conn, prior_month)
        anomalies = detect_anomalies(conn, month)

    rprint(f"[cyan]→[/cyan] {len(actuals)} categories mapped to Excel targets")

    # Step 7: Write to Excel
    model_structure = load_model().get("sheets", {})
    if not model_structure:
        model_structure = scan_workbook_structure(wb_path)

    output_wb = output_dir / f"{month}_update.xlsx"
    cells_written = write_monthly_actuals(wb_path, output_wb, actuals, model_structure, month)
    rprint(f"[green]✓[/green] {cells_written} cells written → {output_wb}")

    # Step 8: Reconcile
    with get_conn(DB_PATH) as conn:
        reconcile_result = reconcile_revenue_cash(conn, month)
    rprint(f"[cyan]→[/cyan] Total deposits: ${reconcile_result['total_deposits']:,.2f}")

    # Step 9: AI summary
    try:
        from bizplan.ai_assist import generate_run_summary
        ai_summary = generate_run_summary(month, current_totals, prior_totals)
    except Exception:
        ai_summary = "AI summary unavailable."

    # Step 10: Report
    report_path = report_dir / f"{month}-report.md"
    generate_report(
        month=month,
        imported=imported,
        new=new_count,
        cells_written=cells_written,
        workbook_path=str(output_wb),
        current_totals=current_totals,
        prior_totals=prior_totals,
        anomalies=anomalies,
        reconcile_result=reconcile_result,
        ai_summary=ai_summary,
        output_path=report_path,
    )
    rprint(f"[green]✓[/green] Report → {report_path}")

    # Log run
    with get_conn(DB_PATH) as conn:
        log_run(conn, month, imported, new_count, cells_written, str(output_wb), str(report_path))

    rprint(f"\n[bold green]Update complete.[/bold green]")
    if anomalies:
        rprint(f"[yellow]⚠ {len(anomalies)} anomaly/anomalies detected — see report.[/yellow]")


@cli.command(name="map")
@click.option("--list", "list_mappings", is_flag=True,
              help="List all current category mappings.")
def map_cmd(list_mappings: bool):
    """Review and edit QuickBooks category → Excel row mappings.

    Use --list to see all current mappings.
    Without flags, launches interactive mapping review for any unconfirmed entries.
    """
    from bizplan.db import get_conn
    with get_conn(DB_PATH) as conn:
        rows = conn.execute(
            "SELECT category_path, excel_sheet, excel_row_label, confirmed_by_user "
            "FROM category_mappings ORDER BY excel_sheet, excel_row_label"
        ).fetchall()

    if not rows:
        rprint("[dim]No mappings saved yet. Run bizplan update to add mappings.[/dim]")
        return

    rprint(f"\n[bold]{len(rows)} category mappings:[/bold]\n")
    for row in rows:
        confirmed = "[green]✓[/green]" if row["confirmed_by_user"] else "[yellow]?[/yellow]"
        rprint(f"  {confirmed} [cyan]{row['category_path']}[/cyan]")
        rprint(f"       → {row['excel_sheet']} / {row['excel_row_label']}")


@cli.command(name="report")
@click.option("--month", "-m", required=True, help="Month in YYYY-MM format.")
def report_cmd(month: str):
    """Print the variance report for a past month.

    The report file must already exist (created by bizplan update).
    """
    report_path = Path("output/reports") / f"{month}-report.md"
    if not report_path.exists():
        rprint(f"[dim]No report found for {month}.[/dim] Run: bizplan update --month {month}")
        return
    console.print(report_path.read_text())
