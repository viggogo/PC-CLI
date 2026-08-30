import asyncio
from datetime import date, datetime

import openpyxl
import pytest

from hevy_mcp import sync_core


def _wk(dt, title="push", desc=""):
    return {
        "title": title, "description": desc,
        "start_time": f"{dt}T11:00:00+00:00", "end_time": f"{dt}T12:00:00+00:00",
        "exercises": [{"title": "Bench Press (Barbell)", "sets": [
            {"reps": 8, "weight_kg": 60, "duration_seconds": None, "distance_meters": None}]}],
    }


@pytest.fixture
def tmp_book(tmp_path, monkeypatch):
    path = tmp_path / "book.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Træning"
    ws.append(["Date", "Place", "Type", "Time", "Mave", "AddCardio", "AddCardio2",
               "Total", "Ensamble", "Rating", "Claude", "Comments"])
    ws.append([date(2026, 6, 14), None, "Over", 72, None, None, None,
               72, None, 3, None, "3"])
    wb.save(path)
    monkeypatch.setenv("EXCEL_PATH", str(path))
    monkeypatch.setenv("EXCEL_SHEET", "Træning")
    return path


def test_resolve_since_explicit_date():
    assert sync_core.resolve_since("2026-06-01") == datetime(2026, 6, 1)


def test_resolve_since_from_last_excel_row(tmp_book):
    assert sync_core.resolve_since("") == datetime(2026, 6, 14)


def test_collect_new_rows_sorted_oldest_first_and_mapped(monkeypatch):
    async def fake(since):
        return [_wk("2026-07-02", "Ben"), _wk("2026-07-01", "Skub")]
    monkeypatch.setattr(sync_core.hevy_client, "fetch_workouts_since", fake)
    rows = asyncio.run(sync_core.collect_new_rows("2026-06-30"))
    assert [r["Date"] for r in rows] == [date(2026, 7, 1), date(2026, 7, 2)]
    assert [r["Type"] for r in rows] == ["push", "legs"]  # Skub->push, Ben->legs
    assert all(r["Claude"] == 1 for r in rows)


def test_collect_calendar_ends_on_the_week_containing_today(monkeypatch):
    async def fake(since):
        return []
    monkeypatch.setattr(sync_core.hevy_client, "fetch_workouts_since", fake)

    grid = asyncio.run(sync_core.collect_calendar(2, today=date(2026, 8, 28)))

    assert [grid[0][0]["date"], grid[-1][-1]["date"]] == [
        date(2026, 8, 17), date(2026, 8, 30)]


def test_collect_calendar_skipping_the_current_week_drops_it(monkeypatch):
    async def fake(since):
        return []
    monkeypatch.setattr(sync_core.hevy_client, "fetch_workouts_since", fake)

    grid = asyncio.run(sync_core.collect_calendar(
        5, skip_current=True, today=date(2026, 8, 28)))

    # 5 asked for, the in-progress week removed and not replaced: 4 remain,
    # ending on the last complete Sunday.
    assert len(grid) == 4
    assert [grid[0][0]["date"], grid[-1][-1]["date"]] == [
        date(2026, 7, 27), date(2026, 8, 23)]


def test_collect_calendar_fetches_from_the_first_monday_inclusive(monkeypatch):
    seen = {}

    async def fake(since):
        seen["since"] = since
        return []
    monkeypatch.setattr(sync_core.hevy_client, "fetch_workouts_since", fake)

    asyncio.run(sync_core.collect_calendar(2, today=date(2026, 8, 28)))

    # fetch_workouts_since is exclusive by date, so the cutoff sits one day
    # before the first Monday or that Monday's workouts would be lost.
    assert seen["since"].date() == date(2026, 8, 16)
