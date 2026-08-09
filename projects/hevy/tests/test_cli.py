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


def test_locked_spreadsheet_exits_with_a_message(monkeypatch):
    async def boom(since):
        raise PermissionError(13, "in use")
    monkeypatch.setattr(sync_core, "collect_new_rows", boom)

    with pytest.raises(SystemExit) as exc:
        cli.main(["preview"])

    assert "close Excel" in str(exc.value)
