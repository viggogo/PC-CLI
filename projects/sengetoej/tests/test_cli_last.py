import datetime as dt

import pytest

from sengetoej import cli

from .conftest import SHEET


@pytest.fixture
def book_env(post_migration_book, monkeypatch):
    monkeypatch.setenv("EXCEL_PATH", str(post_migration_book))
    monkeypatch.setenv("SENGETOEJ_SHEET", SHEET)
    return post_migration_book


def test_fmt_date_is_danish():
    assert cli.fmt_date(dt.date(2026, 9, 2)) == "02/09/2026"


@pytest.mark.parametrize("days,expected", [
    (0, "i dag"),
    (1, "i går"),
    (2, "2 dage siden"),
    (57, "57 dage siden"),
    (-3, "om 3 dage"),
])
def test_days_ago_phrase(days, expected):
    assert cli.days_ago_phrase(days) == expected


def test_bare_last_reports_the_final_entry(book_env, monkeypatch, capsys):
    monkeypatch.setattr(cli, "today", lambda: dt.date(2026, 2, 10))
    assert cli.main(["--last"]) == 0
    out = capsys.readouterr().out
    assert "Sidste skift: 03/02/2026" in out
    assert "7 dage siden" in out


def test_no_arguments_behaves_like_bare_last(book_env, monkeypatch, capsys):
    monkeypatch.setattr(cli, "today", lambda: dt.date(2026, 2, 3))
    assert cli.main([]) == 0
    assert "i dag" in capsys.readouterr().out


def test_last_n_lists_oldest_first_with_intervals(book_env, capsys):
    assert cli.main(["--last", "2"]) == 0
    lines = [ln for ln in capsys.readouterr().out.splitlines() if ln.strip()]
    assert len(lines) == 2
    # The older of the two shown, using the entry before the window.
    assert "15/01/2026" in lines[0] and "14 dage" in lines[0]
    assert "03/02/2026" in lines[1] and "19 dage" in lines[1]


def test_last_n_larger_than_the_sheet_shows_everything(book_env, capsys):
    assert cli.main(["--last", "99"]) == 0
    lines = [ln for ln in capsys.readouterr().out.splitlines() if ln.strip()]
    assert len(lines) == 3


def test_the_very_first_entry_has_no_interval(book_env, capsys):
    assert cli.main(["--last", "99"]) == 0
    assert "—" in capsys.readouterr().out.splitlines()[0]


def test_last_zero_is_a_usage_error(book_env, capsys):
    with pytest.raises(SystemExit) as exc:
        cli.main(["--last", "0"])
    assert exc.value.code == 2


def test_last_non_integer_is_a_usage_error(book_env):
    with pytest.raises(SystemExit) as exc:
        cli.main(["--last", "mange"])
    assert exc.value.code == 2


def test_yes_with_last_is_a_usage_error(book_env, capsys):
    assert cli.main(["--last", "-y"]) == 2
    assert "-y" in capsys.readouterr().err


def test_empty_sheet_says_so(tmp_path, monkeypatch, capsys):
    import openpyxl
    path = tmp_path / "empty.xlsx"
    wb = openpyxl.Workbook()
    wb.active.title = SHEET
    wb.active["A1"] = "Date"
    wb.save(path)
    monkeypatch.setenv("EXCEL_PATH", str(path))
    monkeypatch.setenv("SENGETOEJ_SHEET", SHEET)

    assert cli.main(["--last"]) == 0
    assert "Ingen skift registreret." in capsys.readouterr().out


def test_missing_file_is_a_runtime_error(tmp_path, monkeypatch, capsys):
    monkeypatch.setenv("EXCEL_PATH", str(tmp_path / "nope.xlsx"))
    monkeypatch.setenv("SENGETOEJ_SHEET", SHEET)
    assert cli.main(["--last"]) == 1
    assert "findes ikke" in capsys.readouterr().err
