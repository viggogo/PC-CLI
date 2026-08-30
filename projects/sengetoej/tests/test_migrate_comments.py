import openpyxl
import pytest
from openpyxl.worksheet.formula import ArrayFormula

from sengetoej import migrate, sheet

from .conftest import COUNTER_TEXT, SHEET


def test_comment_rows_finds_every_populated_row(pre_migration_book):
    wb, ws = sheet.open_for_write(pre_migration_book, SHEET)
    assert migrate.comment_rows(ws, migrate.COMMENT_COL_BEFORE) == [1, 2, 3, 4]


def test_comment_rows_is_empty_when_the_column_is(pre_migration_book):
    wb, ws = sheet.open_for_write(pre_migration_book, SHEET)
    assert migrate.comment_rows(ws, migrate.COMMENT_COL_AFTER) == []


def test_move_comments_relocates_values_to_column_e(pre_migration_book):
    wb, ws = sheet.open_for_write(pre_migration_book, SHEET)
    moved = migrate.move_comments(ws)
    sheet.save(wb, pre_migration_book)

    ws2 = openpyxl.load_workbook(pre_migration_book)[SHEET]
    assert moved == 4
    assert ws2["E1"].value == "Dage siden sidste skift"
    assert ws2["E3"].value == "Obs: Ferie"
    assert ws2["E4"].value == "Dato indikerer at der er blevet skiftet sengetøj"


def test_move_comments_empties_column_d(pre_migration_book):
    wb, ws = sheet.open_for_write(pre_migration_book, SHEET)
    migrate.move_comments(ws)
    sheet.save(wb, pre_migration_book)

    ws2 = openpyxl.load_workbook(pre_migration_book)[SHEET]
    assert [ws2.cell(r, 4).value for r in range(1, 10)] == [None] * 9


def test_move_comments_does_not_translate_the_counter_formula(pre_migration_book):
    # The whole point: a translated move would make this INDEX(B:B,COUNTA(B:B)).
    wb, ws = sheet.open_for_write(pre_migration_book, SHEET)
    migrate.move_comments(ws)
    sheet.save(wb, pre_migration_book)

    ws2 = openpyxl.load_workbook(pre_migration_book)[SHEET]
    counter = ws2["E2"].value
    assert isinstance(counter, ArrayFormula)
    assert counter.text == COUNTER_TEXT
    assert "A:A" in counter.text
    assert "B:B" not in counter.text


def test_move_comments_rebuilds_the_array_formula_ref(pre_migration_book):
    wb, ws = sheet.open_for_write(pre_migration_book, SHEET)
    migrate.move_comments(ws)
    sheet.save(wb, pre_migration_book)

    ws2 = openpyxl.load_workbook(pre_migration_book)[SHEET]
    assert ws2["E2"].value.ref == "E2"


def test_move_comments_carries_the_styling(pre_migration_book):
    wb, ws = sheet.open_for_write(pre_migration_book, SHEET)
    migrate.move_comments(ws)
    sheet.save(wb, pre_migration_book)

    ws2 = openpyxl.load_workbook(pre_migration_book)[SHEET]
    assert ws2["E2"].font.b is True
    # openpyxl normalises an RGB fill to 8 hex digits on round-trip, so
    # compare the colour, not the string.
    assert ws2["E2"].fill.fgColor.rgb.endswith("FFFF00")


def test_move_comments_moves_the_column_width(pre_migration_book):
    wb, ws = sheet.open_for_write(pre_migration_book, SHEET)
    migrate.move_comments(ws)
    sheet.save(wb, pre_migration_book)

    ws2 = openpyxl.load_workbook(pre_migration_book)[SHEET]
    assert ws2.column_dimensions["E"].width == pytest.approx(20.6)


def test_move_comments_leaves_the_dates_and_formulas_alone(pre_migration_book):
    wb, ws = sheet.open_for_write(pre_migration_book, SHEET)
    migrate.move_comments(ws)
    sheet.save(wb, pre_migration_book)

    ws2 = openpyxl.load_workbook(pre_migration_book)[SHEET]
    assert len(sheet.read_dates(pre_migration_book, SHEET)) == 3
    assert "Table2[[#This Row],[Date]]" in ws2["B2"].value


def test_move_comments_clears_column_d_styling(pre_migration_book):
    """D2 was bold on a yellow fill; after the move it must not be a stray
    coloured block. This pins the required deviation from the brief: reset
    via ws.style = "Normal" rather than borrowing column F's style (which
    would create a phantom cell in F)."""
    wb, ws = sheet.open_for_write(pre_migration_book, SHEET)
    migrate.move_comments(ws)

    # Column F must never have been touched -- ws.cell() on an untouched
    # cell creates it, so checking this on the live worksheet (rather than
    # after a save/reload round-trip, which may silently drop untouched
    # default cells) is the reliable way to catch the phantom-cell hazard.
    assert (1, 6) not in ws._cells and (2, 6) not in ws._cells

    sheet.save(wb, pre_migration_book)

    ws2 = openpyxl.load_workbook(pre_migration_book)[SHEET]
    d2 = ws2["D2"]
    assert d2.font.b is not True
    assert d2.fill.fgColor.rgb != "00FFFF00"
