import datetime as dt

import openpyxl
import pytest

from sengetoej import sheet

from .conftest import SHEET


def test_scan_stops_at_a_run_of_blanks_and_reports_the_first_empty_row():
    values = [dt.datetime(2026, 1, 1), dt.datetime(2026, 1, 15)] + [None] * 300
    dates, first_empty = sheet.scan(values)
    assert dates == [dt.date(2026, 1, 1), dt.date(2026, 1, 15)]
    # Rows 2 and 3 hold the dates, so row 4 is the first empty one.
    assert first_empty == 4


def test_scan_tolerates_a_gap_shorter_than_the_blank_run():
    values = [dt.datetime(2026, 1, 1)] + [None] * 5 + [dt.datetime(2026, 2, 3)] + [None] * 200
    dates, first_empty = sheet.scan(values)
    assert dates == [dt.date(2026, 1, 1), dt.date(2026, 2, 3)]
    assert first_empty == 9


def test_scan_on_an_empty_sheet():
    dates, first_empty = sheet.scan([None] * 200)
    assert dates == []
    assert first_empty == 2


def test_scan_accepts_dates_as_well_as_datetimes():
    dates, _ = sheet.scan([dt.date(2026, 1, 1), None] + [None] * 200)
    assert dates == [dt.date(2026, 1, 1)]


def test_read_dates_ignores_the_prefilled_diff_formulas(pre_migration_book):
    # Column B is filled to row 60; column A stops at row 4. If read_dates
    # keyed off anything but column A it would return far too much.
    assert sheet.read_dates(pre_migration_book, SHEET) == [
        dt.date(2026, 1, 1), dt.date(2026, 1, 15), dt.date(2026, 2, 3)]


def test_read_dates_works_on_the_migrated_shape(post_migration_book):
    assert len(sheet.read_dates(post_migration_book, SHEET)) == 3


def test_read_dates_on_a_sheet_with_no_entries(tmp_path):
    path = tmp_path / "empty.xlsx"
    wb = openpyxl.Workbook()
    wb.active.title = SHEET
    wb.active["A1"] = "Date"
    wb.save(path)
    assert sheet.read_dates(path, SHEET) == []


def test_read_dates_reports_a_missing_sheet(pre_migration_book):
    with pytest.raises(sheet.SheetMissing):
        sheet.read_dates(pre_migration_book, "Findes Ikke")


def test_gaps_gives_none_for_the_first_entry():
    dates = [dt.date(2026, 1, 1), dt.date(2026, 1, 15), dt.date(2026, 2, 3)]
    assert sheet.gaps(dates) == [None, 14, 19]


def test_gaps_on_a_single_entry():
    assert sheet.gaps([dt.date(2026, 1, 1)]) == [None]


def test_gaps_on_nothing():
    assert sheet.gaps([]) == []
