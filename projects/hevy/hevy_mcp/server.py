from mcp.server.fastmcp import FastMCP

from . import column_mapper, excel_writer, hevy_client, sync_core
from .env import load_env

load_env()

mcp = FastMCP("Hevy Sync")


@mcp.tool()
async def preview_workouts(since_date: str = "") -> str:
    """Show Hevy workouts that would be synced, without writing to Excel.

    Args:
        since_date: Start date in YYYY-MM-DD format. Defaults to day after last Excel entry.
    """
    rows = await sync_core.collect_new_rows(since_date)
    if not rows:
        return "No new workouts found since last Excel entry."
    lines = [
        f"{r['Date']}  {str(r['Type']):<20}  {r['Time']:>3} min"
        f"  Mave:{r['Mave']}  AddCardio:{r['AddCardio']}"
        f"  Rating:{r['Rating'] if r['Rating'] is not None else '-'}"
        f"  {(r['Comments'] or '')[:60]}"
        for r in rows
    ]
    return f"Found {len(rows)} workout(s):\n" + "\n".join(lines)


@mcp.tool()
async def sync_workouts(since_date: str = "") -> str:
    """Fetch Hevy workouts and append new rows to Træning.xlsx.

    Args:
        since_date: Start date in YYYY-MM-DD format. Defaults to day after last Excel entry.
    """
    rows = await sync_core.collect_new_rows(since_date)
    if not rows:
        return "No new workouts to sync — Excel is already up to date."
    count = sync_core.append_new_rows(rows)
    dates = f"{rows[0]['Date']} → {rows[-1]['Date']}" if count > 1 else str(rows[0]["Date"])
    return f"Added {count} workout(s) to Excel ({dates})."


@mcp.tool()
async def add_single_workout(workout_id: str) -> str:
    """Add one specific Hevy workout to Excel by its ID.

    Args:
        workout_id: The Hevy workout ID string.
    """
    workout = await hevy_client.fetch_workout_by_id(workout_id)
    row = column_mapper.workout_to_row(workout)
    excel_writer.append_rows([row])
    return f"Added '{workout.get('title')}' ({row['Date']}, {row['Time']} min) to Excel."


@mcp.tool()
async def fix_synced_rows(confirm: bool = False) -> str:
    """Re-map and correct the rows the old buggy tool wrote (2026-06-14 onward).

    Shows a before/after preview when confirm is False; writes the corrections
    (including restoring the Total formula and setting Claude=1) when confirm is True.
    Requires the Claude column to already exist (run migrate_add_claude_column first).

    Args:
        confirm: Set True to apply the corrections. Default False = preview only.
    """
    changes = await sync_core.collect_fix_changes(dry_run=not confirm)
    if not changes:
        return "No matching rows found to correct."

    lines = []
    for rownum, d, before, after in changes:
        lines.append(
            f"row {rownum} {d}: "
            f"Type {before.get('Type')!r}->{after['Type']!r}  "
            f"Rating {before.get('Rating')!r}->{after['Rating']!r}  "
            f"AddCardio {before.get('AddCardio')!r}->{after['AddCardio']!r}  "
            f"Claude {before.get('Claude')!r}->{after['Claude']!r}"
        )
    header = ("APPLIED" if confirm else "PREVIEW (pass confirm=true to write)")
    return f"{header} — {len(changes)} row(s):\n" + "\n".join(lines)


def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()
