import openpyxl
import pytest

from sengetoej import migrate, sheet

from .conftest import SHEET, TABLE_BOTTOM


def test_add_cli_column_writes_the_header(pre_migration_book):
    wb, ws = sheet.open_for_write(pre_migration_book, SHEET)
    migrate.add_cli_column(ws)
    sheet.save(wb, pre_migration_book)

    ws2 = openpyxl.load_workbook(pre_migration_book)[SHEET]
    assert ws2["C1"].value == "cli"


def test_add_cli_column_extends_the_ref_and_the_autofilter(pre_migration_book):
    wb, ws = sheet.open_for_write(pre_migration_book, SHEET)
    migrate.add_cli_column(ws)
    sheet.save(wb, pre_migration_book)

    table = openpyxl.load_workbook(pre_migration_book)[SHEET].tables["Table2"]
    assert table.ref == f"A1:C{TABLE_BOTTOM}"
    assert table.autoFilter.ref == f"A1:C{TABLE_BOTTOM}"


def test_add_cli_column_registers_the_table_column(pre_migration_book):
    wb, ws = sheet.open_for_write(pre_migration_book, SHEET)
    migrate.add_cli_column(ws)
    sheet.save(wb, pre_migration_book)

    table = openpyxl.load_workbook(pre_migration_book)[SHEET].tables["Table2"]
    assert [c.name for c in table.tableColumns] == ["Date", "Diff", "cli"]
    assert [c.id for c in table.tableColumns] == [1, 2, 3]


def test_add_cli_column_does_not_backfill_existing_rows(pre_migration_book):
    wb, ws = sheet.open_for_write(pre_migration_book, SHEET)
    migrate.add_cli_column(ws)
    sheet.save(wb, pre_migration_book)

    ws2 = openpyxl.load_workbook(pre_migration_book)[SHEET]
    assert [ws2.cell(r, 3).value for r in (2, 3, 4)] == [None, None, None]


def test_the_saved_workbook_still_parses(pre_migration_book):
    """The nearest a test can get to Excel's own validation."""
    wb, ws = sheet.open_for_write(pre_migration_book, SHEET)
    migrate.move_comments(ws)
    migrate.add_cli_column(ws)
    sheet.save(wb, pre_migration_book)

    ws2 = openpyxl.load_workbook(pre_migration_book)[SHEET]
    assert sheet.has_cli_column(ws2) is True
    assert ws2["E1"].value == "Dage siden sidste skift"
    assert len(sheet.read_dates(pre_migration_book, SHEET)) == 3


def test_plan_lists_both_jobs_before_migration(pre_migration_book):
    wb, ws = sheet.open_for_write(pre_migration_book, SHEET)
    steps = migrate.plan(ws)
    assert len(steps) == 2
    assert any("D -> E" in s for s in steps)
    assert any("cli" in s for s in steps)


def test_plan_is_empty_after_migration(post_migration_book):
    wb, ws = sheet.open_for_write(post_migration_book, SHEET)
    assert migrate.plan(ws) == []


def test_backup_copies_the_file_beside_the_original(pre_migration_book):
    made = migrate.backup(pre_migration_book)
    assert made.exists()
    assert made.parent == pre_migration_book.parent
    assert made.stat().st_size == pre_migration_book.stat().st_size
    assert made.name != pre_migration_book.name


def test_main_is_a_no_op_when_already_migrated(post_migration_book, monkeypatch, capsys):
    monkeypatch.setenv("EXCEL_PATH", str(post_migration_book))
    monkeypatch.setenv("SENGETOEJ_SHEET", SHEET)
    assert migrate.main([]) == 0
    assert "allerede" in capsys.readouterr().out.lower()


def test_main_aborts_when_column_c_holds_something_else(pre_migration_book, monkeypatch, capsys):
    wb, ws = sheet.open_for_write(pre_migration_book, SHEET)
    ws["C1"] = "noget andet"
    sheet.save(wb, pre_migration_book)

    monkeypatch.setenv("EXCEL_PATH", str(pre_migration_book))
    monkeypatch.setenv("SENGETOEJ_SHEET", SHEET)
    assert migrate.main([]) == 1
    assert "C1" in capsys.readouterr().err


def test_main_aborts_when_column_e_is_occupied(pre_migration_book, monkeypatch, capsys):
    wb, ws = sheet.open_for_write(pre_migration_book, SHEET)
    ws["E9"] = "i vejen"
    sheet.save(wb, pre_migration_book)

    monkeypatch.setenv("EXCEL_PATH", str(pre_migration_book))
    monkeypatch.setenv("SENGETOEJ_SHEET", SHEET)
    assert migrate.main([]) == 1
    assert "E" in capsys.readouterr().err


def test_main_declined_changes_nothing(pre_migration_book, monkeypatch):
    monkeypatch.setenv("EXCEL_PATH", str(pre_migration_book))
    monkeypatch.setenv("SENGETOEJ_SHEET", SHEET)
    monkeypatch.setattr("builtins.input", lambda _: "n")

    assert migrate.main([]) == 0
    ws = openpyxl.load_workbook(pre_migration_book)[SHEET]
    assert ws["C1"].value is None
    assert ws["D1"].value == "Dage siden sidste skift"


def test_main_accepted_applies_both_jobs_and_backs_up(pre_migration_book, monkeypatch, capsys):
    monkeypatch.setenv("EXCEL_PATH", str(pre_migration_book))
    monkeypatch.setenv("SENGETOEJ_SHEET", SHEET)
    monkeypatch.setattr("builtins.input", lambda _: "y")

    assert migrate.main([]) == 0

    ws = openpyxl.load_workbook(pre_migration_book)[SHEET]
    assert ws["C1"].value == "cli"
    assert ws["E1"].value == "Dage siden sidste skift"
    assert ws["D1"].value is None
    backups = list(pre_migration_book.parent.glob("pre.*.bak.xlsx"))
    assert len(backups) == 1


def test_main_yes_flag_skips_the_prompt(pre_migration_book, monkeypatch):
    monkeypatch.setenv("EXCEL_PATH", str(pre_migration_book))
    monkeypatch.setenv("SENGETOEJ_SHEET", SHEET)

    def no_input(_):
        raise AssertionError("should not have prompted")

    monkeypatch.setattr("builtins.input", no_input)
    assert migrate.main(["-y"]) == 0
    assert openpyxl.load_workbook(pre_migration_book)[SHEET]["C1"].value == "cli"


def test_main_leaves_the_file_untouched_when_move_comments_raises_mid_loop(
        pre_migration_book, monkeypatch, capsys):
    """The critical safety property carried from Task 5's review: a raise
    from move_comments must abort BEFORE any save, so a half-migrated
    worksheet in memory never gets persisted over the file. move_comments is
    monkeypatched to raise directly -- constructing the raise through real
    column contents always trips the earlier `_blocked` guard instead (it
    refuses to even start once both D and E hold values), so this is the
    only way to exercise a raise that happens *after* the prompt/backup but
    *before* the save. main() must catch it, report on stderr (pointing at
    the backup that was already written), return 1 rather than letting the
    exception escape as a traceback, and -- above all -- never call
    sheet.save() afterwards."""
    def boom(_ws):
        raise RuntimeError("kolonne E har allerede en værdi i række 9")

    monkeypatch.setattr(migrate, "move_comments", boom)

    before_bytes = pre_migration_book.read_bytes()

    monkeypatch.setenv("EXCEL_PATH", str(pre_migration_book))
    monkeypatch.setenv("SENGETOEJ_SHEET", SHEET)
    monkeypatch.setattr("builtins.input", lambda _: "y")

    assert migrate.main([]) == 1
    assert capsys.readouterr().err.strip() != ""
    assert pre_migration_book.read_bytes() == before_bytes
