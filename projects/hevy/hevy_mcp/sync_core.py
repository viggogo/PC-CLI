"""Shared orchestration used by both the MCP server and the CLI.

Keeps the Hevy-fetch -> map -> Excel-write pipeline in one place so the two
front ends (server.py tools and cli.py commands) never drift apart.
"""

from datetime import date, datetime, timedelta
from typing import Optional

from . import analysis, column_mapper, excel_writer, hevy_client


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


async def collect_calendar(n_weeks: int, skip_current: bool = False,
                           today: Optional[date] = None) -> list[list[dict]]:
    """The last `n_weeks` Mon-Sun weeks as a calendar grid. Read-only: it never
    opens the spreadsheet.

    `skip_current` drops the in-progress week without pulling an older one in
    to replace it, so the grid holds one week fewer than asked for.
    """
    today = today or date.today()
    if skip_current:
        n_weeks -= 1
        today -= timedelta(weeks=1)
    mondays = analysis.week_starts(n_weeks, today)
    # fetch_workouts_since is exclusive, so step back a day to keep the Monday.
    since = datetime.combine(mondays[0], datetime.min.time()) - timedelta(days=1)
    workouts = await hevy_client.fetch_workouts_since(since)
    return analysis.build_calendar(workouts, n_weeks, today)
