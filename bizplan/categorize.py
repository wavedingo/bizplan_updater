import sqlite3
from rich.prompt import Prompt, Confirm
from rich.console import Console
from rich import print as rprint
from bizplan.db import get_category_mapping, save_category_mapping

console = Console()

def resolve_categories(
    conn: sqlite3.Connection,
    category_paths: list[str],
    vendor_map: dict[str, str],
    prompt: bool = True,
) -> list[str]:
    """For each category_path, ensure a mapping exists in DB.

    If prompt=True, interactively ask user for unknown categories.
    Returns list of still-unmapped category_paths.
    """
    unmapped = []
    for cp in category_paths:
        if get_category_mapping(conn, cp):
            continue
        if prompt:
            vendor = vendor_map.get(cp, "Unknown vendor")
            prompt_for_mapping(conn, cp, vendor=vendor, ai_suggestion=None)
            if not get_category_mapping(conn, cp):
                unmapped.append(cp)
        else:
            unmapped.append(cp)
    return unmapped

def prompt_for_mapping(
    conn: sqlite3.Connection,
    category_path: str,
    vendor: str,
    ai_suggestion: dict | None,
) -> None:
    """Interactively prompt the user to map an unknown category to an Excel location."""
    console.rule(f"[yellow]Unknown Category[/yellow]")
    rprint(f"[bold]Category:[/bold] {category_path}")
    rprint(f"[bold]Vendor:[/bold] {vendor}")

    if ai_suggestion:
        rprint(f"[green]AI Suggestion:[/green] Sheet={ai_suggestion['sheet']!r}, "
               f"Row={ai_suggestion['row_label']!r}")
        if Confirm.ask("Accept AI suggestion?", default=True):
            save_category_mapping(
                conn, category_path,
                ai_suggestion["sheet"], ai_suggestion["row_label"],
                confirmed=True, notes="AI-suggested",
            )
            return

    sheet = Prompt.ask("Excel sheet name")
    row_label = Prompt.ask("Excel row label")
    save_category_mapping(conn, category_path, sheet, row_label, confirmed=True)
    rprint(f"[green]Saved:[/green] {category_path!r} → {sheet!r} / {row_label!r}")
