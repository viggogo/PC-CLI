"""Shared orchestration used by both the MCP server and the CLI.

Keeps the Hevy-fetch -> map -> Excel-write pipeline in one place so the two
front ends (server.py tools and cli.py commands) never drift apart.
"""

from datetime import datetime, timedelta
from typing import Optional

from . import column_mapper, excel_writer, hevy_client

# Rows from this date onward were written by the tool; fix operates on them.
FIX_CUTOFF = datetime(2026, 6, 14)


def resolve_since(since_date: str = "") -> Optional[datetime]:
    """Cutoff for fetching. Explicit YYYY-MM-DD wins; otherwise the day of the
    last Excel entry; None means fetch everything."""
    if since_date:
        return datetime.strptime(since_date, "%Y-%m-%d")
    last = excel_writer.get_last_date()
    if last:
        return datetime.combine(last, datetime.min.time())
    return None


async def collect_new_rows(since_date: str = "") -> list[dict]:
    """Mapped rows for workouts newer than the cutoff, oldest first."""
    since = resolve_since(since_date)
    workouts = await hevy_client.fetch_workouts_since(since)
    return [
        column_mapper.workout_to_row(w)
        for w in sorted(workouts, key=lambda x: x["start_time"])
    ]


def append_new_rows(rows: list[dict]) -> int:
    return excel_writer.append_rows(rows)


async def collect_fix_changes(dry_run: bool) -> list:
    """Preview or apply corrections to the tool-written rows (FIX_CUTOFF onward).

    Returns a list of (rownum, date, before, after). Writes only when
    dry_run is False.
    """
    workouts = await hevy_client.fetch_workouts_since(FIX_CUTOFF - timedelta(days=1))
    rows_by_date = {}
    for w in workouts:
        row = column_mapper.workout_to_row(w)
        rows_by_date[row["Date"]] = row
    return excel_writer.overwrite_rows_by_date(rows_by_date, dry_run=dry_run)
