"""Locate this tool's own .env, whatever directory the command was run from.

A bare load_dotenv() only searches upward from the *current* directory, so
`hevy sync` typed in C:\\ would silently find no config. The console script can
be invoked from anywhere, so the tool's .env is addressed by its own location.
"""

from pathlib import Path

from dotenv import load_dotenv

# hevy_mcp/env.py -> hevy_mcp/ -> the project folder holding .env and pyproject.toml
PROJECT_ROOT = Path(__file__).resolve().parent.parent

ENV_FILE = PROJECT_ROOT / ".env"


def load_env() -> None:
    """Populate os.environ from the tool's .env, then any .env in the CWD.

    load_dotenv never overwrites a variable that is already set, so real
    environment variables win over both files, and the tool's own .env wins
    over a stray one in whatever directory you happened to be standing in.
    """
    load_dotenv(ENV_FILE)
    load_dotenv()
