import os

from sengetoej import env


def test_defaults_apply_when_nothing_is_set(monkeypatch):
    monkeypatch.delenv("EXCEL_PATH", raising=False)
    monkeypatch.delenv("SENGETOEJ_SHEET", raising=False)
    assert str(env.excel_path()) == env.DEFAULT_EXCEL_PATH
    assert env.sheet_name() == "Sengetøj"


def test_environment_overrides_the_defaults(monkeypatch):
    monkeypatch.setenv("EXCEL_PATH", r"D:\somewhere\Other.xlsx")
    monkeypatch.setenv("SENGETOEJ_SHEET", "Andet")
    assert str(env.excel_path()) == r"D:\somewhere\Other.xlsx"
    assert env.sheet_name() == "Andet"


def test_env_file_sits_next_to_pyproject():
    assert env.ENV_FILE.name == ".env"
    assert (env.PROJECT_ROOT / "pyproject.toml").is_file()
