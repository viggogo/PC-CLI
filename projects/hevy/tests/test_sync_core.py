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


def test_collect_fix_changes_dry_run_previews_without_writing(monkeypatch, tmp_book):
    async def fake(since):
        return [_wk("2026-06-14", "Over", "3")]
    monkeypatch.setattr(sync_core.hevy_client, "fetch_workouts_since", fake)
    changes = asyncio.run(sync_core.collect_fix_changes(dry_run=True))
    assert len(changes) == 1
    rn, d, before, after = changes[0]
    assert before["Type"] == "Over"
    assert after["Type"] == "upper"
    # dry run: sheet unchanged
    wb = openpyxl.load_workbook(tmp_book)
    assert wb["Træning"].cell(row=2, column=3).value == "Over"
