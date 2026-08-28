from datetime import date

import httpx
import pytest

from hevy_mcp import cli, sync_core


@pytest.fixture(autouse=True)
def _env(monkeypatch):
    # Neutralise load_env so the real .env can never leak in and mask a
    # deliberately-missing variable (see test_missing_env_exits).
    monkeypatch.setattr(cli, "load_env", lambda: None)
    monkeypatch.setenv("HEVY_API_KEY", "x")
    monkeypatch.setenv("EXCEL_PATH", "dummy.xlsx")


def _row(d=date(2026, 7, 1), t="push"):
    return {"Date": d, "Place": "", "Type": t, "Time": 50, "Mave": 0,
            "AddCardio": 0, "AddCardio2": 0, "Ensamble": 0, "Rating": 4,
            "Claude": 1, "Comments": "4"}


def test_preview_lists_and_does_not_write(monkeypatch, capsys):
    async def fake_rows(since):
        return [_row()]
    monkeypatch.setattr(sync_core, "collect_new_rows", fake_rows)
    called = []
    monkeypatch.setattr(sync_core, "append_new_rows", lambda rows: called.append(rows))

    cli.main(["preview"])

    out = capsys.readouterr().out
    assert "1 new workout" in out
    assert not called


def test_sync_yes_writes(monkeypatch, capsys):
    async def fake_rows(since):
        return [_row(), _row(date(2026, 7, 2), "pull")]
    monkeypatch.setattr(sync_core, "collect_new_rows", fake_rows)
    written = {}

    def _append(rows):
        written["n"] = len(rows)
        return len(rows)
    monkeypatch.setattr(sync_core, "append_new_rows", _append)

    cli.main(["sync", "-y"])

    assert written["n"] == 2
    assert "Added 2 row" in capsys.readouterr().out


def test_sync_decline_does_not_write(monkeypatch, capsys):
    async def fake_rows(since):
        return [_row()]
    monkeypatch.setattr(sync_core, "collect_new_rows", fake_rows)
    called = []
    monkeypatch.setattr(sync_core, "append_new_rows", lambda rows: called.append(rows))
    monkeypatch.setattr(cli, "_confirm", lambda prompt: False)

    cli.main(["sync"])

    assert not called
    assert "Aborted" in capsys.readouterr().out


def test_no_new_workouts(monkeypatch, capsys):
    async def fake_rows(since):
        return []
    monkeypatch.setattr(sync_core, "collect_new_rows", fake_rows)

    cli.main(["preview"])

    assert "No new workouts" in capsys.readouterr().out


def test_missing_env_exits(monkeypatch):
    monkeypatch.delenv("HEVY_API_KEY", raising=False)
    with pytest.raises(SystemExit):
        cli.main(["preview"])


def _http_error(status: int) -> httpx.HTTPStatusError:
    request = httpx.Request("GET", "https://api.hevyapp.com/v1/workouts")
    response = httpx.Response(status, request=request)
    return httpx.HTTPStatusError("boom", request=request, response=response)


@pytest.mark.parametrize("status,expected", [(401, "API key"), (500, "500")])
def test_http_errors_exit_with_a_message_not_a_traceback(
        monkeypatch, status, expected):
    async def boom(since):
        raise _http_error(status)
    monkeypatch.setattr(sync_core, "collect_new_rows", boom)

    with pytest.raises(SystemExit) as exc:
        cli.main(["preview"])

    assert expected in str(exc.value)


def _stub_calendar(monkeypatch, calls=None):
    from hevy_mcp import analysis

    async def fake(n_weeks, skip_current=False, today=None):
        if calls is not None:
            calls.append((n_weeks, skip_current))
        return analysis.build_calendar([], n_weeks, date(2026, 8, 28))
    monkeypatch.setattr(sync_core, "collect_calendar", fake)


def test_weeks_prints_a_calendar(monkeypatch, capsys):
    _stub_calendar(monkeypatch)

    cli.main(["--weeks", "2"])

    out = capsys.readouterr().out
    assert "Mon" in out and "Sun" in out
    assert "+-----------+" in out


def test_weeks_passes_the_week_count_through(monkeypatch, capsys):
    seen = []
    _stub_calendar(monkeypatch, seen)

    cli.main(["--weeks", "6"])

    assert seen == [(6, False)]


def test_minus_one_asks_for_the_current_week_to_be_dropped(monkeypatch, capsys):
    seen = []
    _stub_calendar(monkeypatch, seen)

    cli.main(["--weeks", "5", "-1"])

    assert seen == [(5, True)]


def test_one_week_minus_the_current_one_leaves_nothing_to_show(monkeypatch):
    _stub_calendar(monkeypatch)

    with pytest.raises(SystemExit) as exc:
        cli.main(["--weeks", "1", "-1"])

    assert exc.value.code == 2


def test_weeks_does_not_need_the_spreadsheet(monkeypatch, capsys):
    monkeypatch.delenv("EXCEL_PATH", raising=False)
    _stub_calendar(monkeypatch)

    cli.main(["--weeks", "2"])

    assert "Mon" in capsys.readouterr().out


@pytest.mark.parametrize("weeks", ["0", "-1", "53"])
def test_weeks_rejects_an_out_of_range_week_count(monkeypatch, weeks):
    _stub_calendar(monkeypatch)

    with pytest.raises(SystemExit) as exc:
        cli.main(["--weeks", weeks])

    assert exc.value.code == 2


@pytest.mark.parametrize("argv", [["--weeks", "abc"], ["--weeks"]])
def test_weeks_requires_an_integer(monkeypatch, argv):
    _stub_calendar(monkeypatch)

    with pytest.raises(SystemExit) as exc:
        cli.main(argv)

    assert exc.value.code == 2


@pytest.mark.parametrize("flag", ["--week", "--wee", "--w"])
def test_an_abbreviation_of_weeks_is_not_accepted(monkeypatch, flag):
    # argparse resolves unambiguous prefixes by default; the flag is --weeks
    # and nothing else, so a near miss must be told rather than guessed at.
    _stub_calendar(monkeypatch)

    with pytest.raises(SystemExit) as exc:
        cli.main([flag, "2"])

    assert exc.value.code == 2


def test_no_command_at_all_is_still_a_usage_error(monkeypatch):
    with pytest.raises(SystemExit) as exc:
        cli.main([])

    assert exc.value.code == 2


def test_locked_spreadsheet_exits_with_a_message(monkeypatch):
    async def boom(since):
        raise PermissionError(13, "in use")
    monkeypatch.setattr(sync_core, "collect_new_rows", boom)

    with pytest.raises(SystemExit) as exc:
        cli.main(["preview"])

    assert "close Excel" in str(exc.value)
