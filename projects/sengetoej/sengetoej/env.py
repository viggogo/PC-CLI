"""Locate this tool's own .env, whatever directory the command was run from.

A bare load_dotenv() only searches upward from the *current* directory, so
`sengetoej --last` typed in C:\\ would silently find no config. The console
script can be invoked from anywhere, so the .env is addressed by its own
location.

Unlike hevy, both settings have working defaults: this tool holds no secret,
so requiring a .env would be friction with no payoff.
"""

import os
from pathlib import Path

from dotenv import load_dotenv

# sengetoej/env.py -> sengetoej/ -> the project folder holding .env
PROJECT_ROOT = Path(__file__).resolve().parent.parent

ENV_FILE = PROJECT_ROOT / ".env"

DEFAULT_EXCEL_PATH = r"C:\Users\viggo\OneDrive\Privat\Fitness\Træning.xlsx"
DEFAULT_SHEET = "Sengetøj"


def load_env() -> None:
    """Populate os.environ from the tool's .env, then any .env in the CWD.

    load_dotenv never overwrites a variable that is already set, so real
    environment variables win over both files, and this tool's .env wins over
    a stray one in whatever directory you happened to be standing in.
    """
    load_dotenv(ENV_FILE)
    load_dotenv()


def excel_path() -> Path:
    return Path(os.environ.get("EXCEL_PATH") or DEFAULT_EXCEL_PATH)


def sheet_name() -> str:
    return os.environ.get("SENGETOEJ_SHEET") or DEFAULT_SHEET
