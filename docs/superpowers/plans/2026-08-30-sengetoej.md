# sengetoej Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** A CLI that reports when the bed linen was last changed and appends a new change to the `Sengetøj` tab of `Træning.xlsx`.

**Architecture:** A self-contained Python tool under `projects/sengetoej/`, mirroring `projects/hevy`'s skeleton (own `.venv`, `install.ps1`, `bin\*.cmd` on PATH). Three modules: `sheet.py` owns all workbook I/O, `cli.py` owns argument parsing and Danish output, `migrate.py` is a one-time script that restructures the sheet. Reads go through openpyxl's `read_only` mode and stop early; writes append two cells and never touch the table schema.

**Tech Stack:** Python ≥ 3.11, openpyxl ≥ 3.1, python-dotenv ≥ 1.0, pytest ≥ 8.0.

**Spec:** `docs/superpowers/specs/2026-08-30-sengetoej-design.md`

## Global Constraints

- **Confine all changes to `projects/sengetoej/`** and the two doc files. Never modify `projects/hevy/` or `projects/lidaction/` (repo CLAUDE.md: "When working on one tool, confine changes to that tool's folder"). Note `projects/hevy/` currently has unrelated uncommitted changes from another session — leave them alone and never `git add` them.
- **No code is imported from `hevy`.** Patterns are copied; modules are not shared.
- **Tests never open the real workbook.** Every test builds its own file in `tmp_path`. The only contact with `C:\Users\viggo\OneDrive\Privat\Fitness\Træning.xlsx` is Task 9, which is manual and backed up first.
- **Output language is Danish. README and code comments are English.** (hevy's convention.)
- **Exit codes:** `0` success, `1` runtime failure, `2` usage error.
- **`install.ps1` must be ASCII-only** — PowerShell 5.1 decodes a BOM-less script as ANSI and mangles non-ASCII. Every other file is UTF-8 and may contain `æøå`.
- **Danish strings in Python source are fine** but every source file must be read and written as UTF-8.
- Commit after every task. Never use `--no-verify`.

## Verified facts this plan relies on

Measured against the live workbook during planning. Do not re-derive; do not contradict.

| Fact | Value |
|---|---|
| Sheet / table | `Sengetøj` / `Table2`, ref `A1:B1048576`, autofilter `A1:B1048576` |
| Columns today | A `Date` (datetime, format `mm-dd-yy`), B `Diff` (formula), C empty, D comments (outside table, width 20.6), E empty |
| Data rows | 2–73 (72 entries), ascending, last `2026-08-30` |
| Cells actually stored on the sheet | **1,048,659** — the `Diff` formula is materialized in every row |
| `load_workbook()` plain | **6.6 s** |
| `load_workbook(read_only=True)` + column A + early break | **0.01 s** |
| `read_only` full column walk (no break) | 4.0 s — never do this |
| Append round-trip (load + save) | **~15 s**, file stays 8.9 MB, all formulas and the `ArrayFormula` survive |
| References to column D anywhere in the workbook | **none**; no defined names, no conditional formatting, no merged cells |
| `D2` | `ArrayFormula(ref="D2", text="=DAY(ABS(TODAY()-INDEX(A:A,COUNTA(A:A))))")`, bold, fill `FFFFFF00` |
| Comment rows | 1, 2, 4, 29, 36, 48, 50, 59, 70, 71, 72 |

Two consequences that are easy to get wrong:

1. **`ws.max_row` is 1048576 and useless.** So is the table ref. The end of the data can only be found by scanning column A for a run of blanks.
2. **`ws.cell()` does not work in `read_only` mode.** Reads must go through `iter_rows()`.

## File structure

| File | Responsibility |
|---|---|
| `projects/sengetoej/sengetoej/env.py` | Locate `.env`; resolve `EXCEL_PATH` and `SENGETOEJ_SHEET` with defaults |
| `projects/sengetoej/sengetoej/sheet.py` | All workbook I/O: scan column A, detect the `cli` column, append a row |
| `projects/sengetoej/sengetoej/cli.py` | argparse, validation, confirmation prompt, Danish output |
| `projects/sengetoej/sengetoej/migrate.py` | One-time: move comments D→E, add the `cli` column |
| `projects/sengetoej/tests/conftest.py` | Builds pre- and post-migration workbooks in `tmp_path` |
| `projects/sengetoej/install.ps1` | venv, editable install, `.env`, PATH |
| `projects/sengetoej/bin/sengetoej.cmd` | The only file exposed on PATH |

---

### Task 1: Project scaffold and configuration

**Files:**
- Create: `projects/sengetoej/pyproject.toml`, `.gitignore`, `.env.example`, `bin/sengetoej.cmd`
- Create: `projects/sengetoej/sengetoej/__init__.py`, `sengetoej/env.py`
- Test: `projects/sengetoej/tests/__init__.py`, `tests/test_env.py`

**Interfaces:**
- Consumes: nothing.
- Produces: `env.PROJECT_ROOT: Path`, `env.ENV_FILE: Path`, `env.DEFAULT_EXCEL_PATH: str`, `env.DEFAULT_SHEET: str`, `env.load_env() -> None`, `env.excel_path() -> Path`, `env.sheet_name() -> str`.

- [ ] **Step 1: Create the folder skeleton**

```bash
cd "C:/Users/viggo/Git Clone/PC-CLI"
mkdir -p projects/sengetoej/sengetoej projects/sengetoej/tests projects/sengetoej/bin
touch projects/sengetoej/sengetoej/__init__.py projects/sengetoej/tests/__init__.py
```

- [ ] **Step 2: Write `projects/sengetoej/pyproject.toml`**

```toml
[build-system]
requires = ["setuptools>=68"]
build-backend = "setuptools.build_meta"

[project]
name = "sengetoej"
version = "0.1.0"
description = "Read and record bed linen changes in Traening.xlsx from the terminal."
requires-python = ">=3.11"
dependencies = [
    "openpyxl>=3.1",
    "python-dotenv>=1.0",
]

[project.optional-dependencies]
dev = ["pytest>=8.0"]

[project.scripts]
sengetoej = "sengetoej.cli:run"

[tool.setuptools.packages.find]
where = ["."]
include = ["sengetoej*"]
```

- [ ] **Step 3: Write `projects/sengetoej/.gitignore`**

```
.venv/
.env
__pycache__/
*.pyc
*.egg-info/
.pytest_cache/
*.bak
```

- [ ] **Step 4: Write `projects/sengetoej/.env.example`**

```
# install.ps1 copies this file to .env. Both settings already default to the
# values below in code, so the tool works without a .env at all. Override here
# only if the workbook moves or the tab is renamed.
EXCEL_PATH=C:\Users\viggo\OneDrive\Privat\Fitness\Træning.xlsx
SENGETOEJ_SHEET=Sengetøj
```

- [ ] **Step 5: Write `projects/sengetoej/bin/sengetoej.cmd`**

```
@"%~dp0..\.venv\Scripts\sengetoej.exe" %*
```

- [ ] **Step 6: Write the failing test `projects/sengetoej/tests/test_env.py`**

```python
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
```

- [ ] **Step 7: Run the test to verify it fails**

Run from `projects/sengetoej/`: `.\.venv\Scripts\python.exe -m pytest tests/test_env.py -v`

Expected: FAIL — the venv does not exist yet. Create it first:

```powershell
cd "C:\Users\viggo\Git Clone\PC-CLI\projects\sengetoej"
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip --quiet
.\.venv\Scripts\python.exe -m pip install --editable ".[dev]" --quiet
.\.venv\Scripts\python.exe -m pytest tests/test_env.py -v
```

Now expected: FAIL with `ImportError: cannot import name 'env'`.

- [ ] **Step 8: Write `projects/sengetoej/sengetoej/env.py`**

```python
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
```

- [ ] **Step 9: Run the test to verify it passes**

Run: `.\.venv\Scripts\python.exe -m pytest tests/test_env.py -v`
Expected: 3 passed.

- [ ] **Step 10: Commit**

```bash
cd "C:/Users/viggo/Git Clone/PC-CLI"
git add projects/sengetoej/pyproject.toml projects/sengetoej/.gitignore \
        projects/sengetoej/.env.example projects/sengetoej/bin \
        projects/sengetoej/sengetoej projects/sengetoej/tests
git commit -m "feat(sengetoej): scaffold project and configuration"
```

---

### Task 2: Test fixtures shaped like the real sheet

**Files:**
- Create: `projects/sengetoej/tests/conftest.py`
- Test: `projects/sengetoej/tests/test_conftest.py`

**Interfaces:**
- Consumes: nothing.
- Produces: two pytest fixtures. `pre_migration_book(tmp_path) -> Path` — a workbook with `Table2` at `A1:B1048576`, comments in D, empty C. `post_migration_book(tmp_path) -> Path` — the same but `Table2` at `A1:C1048576` with a `cli` column, comments in E, empty D. Both use sheet name `Sengetøj` and hold dates `2026-01-01`, `2026-01-15`, `2026-02-03` in rows 2–4. Also produces the module constant `tests.conftest.SHEET = "Sengetøj"`.

The fixture must reproduce the traps, or the tests prove nothing: an over-wide table ref, `Diff` formulas pre-filled *past* the last date, trailing blank rows, and a styled `ArrayFormula` in the comment column.

- [ ] **Step 1: Write `projects/sengetoej/tests/conftest.py`**

```python
"""Throwaway workbooks shaped like the real Sengetøj sheet.

The real sheet has three properties that break naive code, so the fixtures
reproduce all three:
  * Table2's ref spans the whole column (A1:B1048576), so it says nothing
    about where the data ends.
  * The Diff formula is pre-filled well past the last date, so a scan of
    column B would never terminate.
  * The comment column sits OUTSIDE the table and holds a styled
    ArrayFormula, which a value-only copy would silently mangle.
"""

import datetime as dt

import openpyxl
import pytest
from openpyxl.styles import Font, PatternFill
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.worksheet.table import Table, TableColumn, TableStyleInfo

SHEET = "Sengetøj"

# Deliberately narrower than the real sheet's 1048576 so the tests stay fast.
# Everything that matters is that the ref extends far past the data.
TABLE_BOTTOM = 400
FORMULA_FILL_TO = 60

DATES = [dt.datetime(2026, 1, 1), dt.datetime(2026, 1, 15), dt.datetime(2026, 2, 3)]
COUNTER_TEXT = "=DAY(ABS(TODAY()-INDEX(A:A,COUNTA(A:A))))"


def _diff_formula(row: int) -> str:
    return (f'=IF(ISBLANK(Table2[[#This Row],[Date]]), " ", '
            f'Table2[[#This Row],[Date]]-A{row - 1})')


def _base_sheet(wb):
    ws = wb.active
    ws.title = SHEET
    ws["A1"] = "Date"
    ws["B1"] = "Diff"
    for i, d in enumerate(DATES, start=2):
        ws.cell(i, 1, d).number_format = "mm-dd-yy"
    for r in range(2, FORMULA_FILL_TO + 1):
        ws.cell(r, 2, _diff_formula(r))
    return ws


def _write_comments(ws, col: int):
    """Label, styled array formula, legend, and a per-row remark."""
    ws.cell(1, col, "Dage siden sidste skift")
    cell = ws.cell(2, col)
    cell.value = ArrayFormula(ref=f"{cell.column_letter}2", text=COUNTER_TEXT)
    cell.font = Font(bold=True)
    cell.fill = PatternFill("solid", fgColor="FFFF00")
    ws.cell(3, col, "Obs: Ferie")
    ws.cell(4, col, "Dato indikerer at der er blevet skiftet sengetøj")
    ws.column_dimensions[ws.cell(1, col).column_letter].width = 20.6


def _add_table(ws, ref: str, names: list[str]):
    table = Table(displayName="Table2", ref=ref)
    table.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True)
    ws.add_table(table)
    # add_table derives tableColumns from the header row, so this is a
    # consistency check rather than a mutation.
    assert [c.name for c in table.tableColumns] == names, (
        [c.name for c in table.tableColumns], names)
    return table


@pytest.fixture
def pre_migration_book(tmp_path):
    """Two-column Table2, comments in D, C empty. The state before migrate.py."""
    path = tmp_path / "pre.xlsx"
    wb = openpyxl.Workbook()
    ws = _base_sheet(wb)
    _write_comments(ws, 4)
    _add_table(ws, f"A1:B{TABLE_BOTTOM}", ["Date", "Diff"])
    wb.save(path)
    return path


@pytest.fixture
def post_migration_book(tmp_path):
    """Three-column Table2 with cli in C, comments in E, D empty."""
    path = tmp_path / "post.xlsx"
    wb = openpyxl.Workbook()
    ws = _base_sheet(wb)
    ws["C1"] = "cli"
    _write_comments(ws, 5)
    _add_table(ws, f"A1:C{TABLE_BOTTOM}", ["Date", "Diff", "cli"])
    wb.save(path)
    return path
```

- [ ] **Step 2: Write `projects/sengetoej/tests/test_conftest.py`**

The fixtures are load-bearing for every later task, so they get their own tests.

```python
import openpyxl
from openpyxl.worksheet.formula import ArrayFormula

from .conftest import COUNTER_TEXT, SHEET, TABLE_BOTTOM


def test_pre_migration_book_reproduces_the_traps(pre_migration_book):
    ws = openpyxl.load_workbook(pre_migration_book)[SHEET]
    table = ws.tables["Table2"]

    # The ref says nothing about where the data ends.
    assert table.ref == f"A1:B{TABLE_BOTTOM}"
    assert [c.name for c in table.tableColumns] == ["Date", "Diff"]

    # Diff formulas continue past the last date (row 4).
    assert ws["B40"].value is not None
    assert ws["A40"].value is None

    # C is empty; the comments are in D and outside the table.
    assert ws["C1"].value is None
    assert ws["D1"].value == "Dage siden sidste skift"

    # The counter is a styled array formula, not a plain string.
    counter = ws["D2"].value
    assert isinstance(counter, ArrayFormula)
    assert counter.text == COUNTER_TEXT
    assert counter.ref == "D2"
    assert ws["D2"].font.b is True


def test_post_migration_book_is_the_shape_migration_should_produce(post_migration_book):
    ws = openpyxl.load_workbook(post_migration_book)[SHEET]
    table = ws.tables["Table2"]

    assert table.ref == f"A1:C{TABLE_BOTTOM}"
    assert [c.name for c in table.tableColumns] == ["Date", "Diff", "cli"]
    assert ws["C1"].value == "cli"
    assert ws["D1"].value is None
    assert ws["E1"].value == "Dage siden sidste skift"
    assert isinstance(ws["E2"].value, ArrayFormula)
    assert ws["E2"].value.ref == "E2"
```

- [ ] **Step 3: Run the tests**

Run: `.\.venv\Scripts\python.exe -m pytest tests/test_conftest.py -v`
Expected: 2 passed. If `_add_table`'s assertion fires, the header row and the ref disagree — fix the fixture, not the assertion.

- [ ] **Step 4: Commit**

```bash
git add projects/sengetoej/tests/conftest.py projects/sengetoej/tests/test_conftest.py
git commit -m "test(sengetoej): add workbook fixtures reproducing the real sheet's traps"
```

---

### Task 3: Reading the sheet

**Files:**
- Create: `projects/sengetoej/sengetoej/sheet.py`
- Test: `projects/sengetoej/tests/test_sheet_read.py`

**Interfaces:**
- Consumes: `tests.conftest.SHEET`, both book fixtures.
- Produces:
  - `sheet.BLANK_RUN: int = 100`
  - `sheet.DATE_COL: int = 1`, `sheet.CLI_COL: int = 3`, `sheet.CLI_HEADER: str = "cli"`
  - `sheet.scan(values: Iterable) -> tuple[list[datetime.date], int]` — takes column A values starting at row 2, returns `(dates_ascending, first_empty_row)`
  - `sheet.read_dates(path: Path, sheet_name: str) -> list[datetime.date]` — the fast read-only path
  - `sheet.gaps(dates: list[date]) -> list[int | None]` — days since the previous entry, `None` for the first

- [ ] **Step 1: Write the failing test `projects/sengetoej/tests/test_sheet_read.py`**

```python
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
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `.\.venv\Scripts\python.exe -m pytest tests/test_sheet_read.py -v`
Expected: FAIL — `ImportError: cannot import name 'sheet'`.

- [ ] **Step 3: Write `projects/sengetoej/sengetoej/sheet.py`**

```python
"""All workbook I/O for the Sengetøj sheet.

Two facts about the real sheet drive everything here:

  * Table2's ref is A1:B1048576 and ws.max_row is 1048576, because the Diff
    formula is materialized in every one of those rows. Neither says where
    the data ends. Only a scan of column A does.

  * That also makes a plain load_workbook() cost ~6.6 seconds. Reading goes
    through read_only mode over column A and breaks as soon as it sees a run
    of blanks, which costs ~0.01 seconds. A read_only walk that does NOT
    break early costs ~4 seconds, so the break is not an optimisation — it
    is the difference between a usable tool and an unusable one.
"""

import datetime as dt
from pathlib import Path
from typing import Iterable

import openpyxl

# How many consecutive empty cells in column A mean "the data ended".
# Comfortably more than any gap a hand-kept column could contain.
BLANK_RUN = 100

DATE_COL = 1
CLI_COL = 3
CLI_HEADER = "cli"

# Row 1 is the header, so data starts here.
FIRST_DATA_ROW = 2


class SheetMissing(Exception):
    """The configured tab is not in the workbook."""


def _as_date(value) -> dt.date:
    return value.date() if isinstance(value, dt.datetime) else value


def scan(values: Iterable) -> tuple[list[dt.date], int]:
    """Turn column A values (starting at row 2) into (dates, first_empty_row).

    Stops after BLANK_RUN consecutive blanks, so a caller may hand this a
    lazy iterator over a million rows without paying for them.
    """
    dates: list[dt.date] = []
    last_row = FIRST_DATA_ROW - 1
    blanks = 0
    row = FIRST_DATA_ROW

    for value in values:
        if value is None:
            blanks += 1
            if blanks >= BLANK_RUN:
                break
        else:
            blanks = 0
            dates.append(_as_date(value))
            last_row = row
        row += 1

    return dates, last_row + 1


def read_dates(path: Path, sheet_name: str) -> list[dt.date]:
    """Every recorded change, ascending. Fast, and safe while Excel is open."""
    wb = openpyxl.load_workbook(path, read_only=True)
    try:
        if sheet_name not in wb.sheetnames:
            raise SheetMissing(sheet_name)
        ws = wb[sheet_name]
        # ws.cell() is unsupported in read_only mode; iter_rows is the only way.
        values = (row[0] for row in
                  ws.iter_rows(min_row=FIRST_DATA_ROW, max_col=DATE_COL,
                               values_only=True))
        dates, _ = scan(values)
        return dates
    finally:
        wb.close()


def gaps(dates: list[dt.date]) -> list[int | None]:
    """Days since the previous entry, aligned to `dates`. None for the first."""
    return [None if i == 0 else (d - dates[i - 1]).days
            for i, d in enumerate(dates)]
```

- [ ] **Step 4: Run the test to verify it passes**

Run: `.\.venv\Scripts\python.exe -m pytest tests/test_sheet_read.py -v`
Expected: 11 passed.

- [ ] **Step 5: Commit**

```bash
git add projects/sengetoej/sengetoej/sheet.py projects/sengetoej/tests/test_sheet_read.py
git commit -m "feat(sengetoej): read change dates via a fast read-only column scan"
```

---

### Task 4: Appending a row

**Files:**
- Modify: `projects/sengetoej/sengetoej/sheet.py` (append to the end)
- Test: `projects/sengetoej/tests/test_sheet_write.py`

**Interfaces:**
- Consumes: everything from Task 3.
- Produces:
  - `sheet.open_for_write(path: Path, sheet_name: str) -> tuple[Workbook, Worksheet]`
  - `sheet.dates_and_next_row(ws) -> tuple[list[date], int]`
  - `sheet.has_cli_column(ws) -> bool`
  - `sheet.append_entry(ws, when: date, row: int) -> None`
  - `sheet.save(wb, path: Path) -> None` — raises `sheet.WorkbookLocked`
  - `sheet.WorkbookLocked(Exception)`

- [ ] **Step 1: Write the failing test `projects/sengetoej/tests/test_sheet_write.py`**

```python
import datetime as dt

import openpyxl
import pytest
from openpyxl.worksheet.formula import ArrayFormula

from sengetoej import sheet

from .conftest import COUNTER_TEXT, SHEET


def test_has_cli_column_is_false_before_migration(pre_migration_book):
    wb, ws = sheet.open_for_write(pre_migration_book, SHEET)
    assert sheet.has_cli_column(ws) is False


def test_has_cli_column_is_true_after_migration(post_migration_book):
    wb, ws = sheet.open_for_write(post_migration_book, SHEET)
    assert sheet.has_cli_column(ws) is True


def test_has_cli_column_ignores_case_and_whitespace(post_migration_book):
    wb, ws = sheet.open_for_write(post_migration_book, SHEET)
    ws.cell(1, sheet.CLI_COL, "  CLI ")
    assert sheet.has_cli_column(ws) is True


def test_dates_and_next_row_finds_the_first_empty_row(post_migration_book):
    wb, ws = sheet.open_for_write(post_migration_book, SHEET)
    dates, next_row = sheet.dates_and_next_row(ws)
    assert dates[-1] == dt.date(2026, 2, 3)
    assert next_row == 5  # rows 2-4 hold the three dates


def test_append_writes_the_date_and_the_cli_flag(post_migration_book):
    wb, ws = sheet.open_for_write(post_migration_book, SHEET)
    sheet.append_entry(ws, dt.date(2026, 3, 1), 5)
    sheet.save(wb, post_migration_book)

    ws2 = openpyxl.load_workbook(post_migration_book)[SHEET]
    assert ws2["A5"].value == dt.datetime(2026, 3, 1)
    assert ws2["C5"].value == 1


def test_append_copies_the_number_format_from_the_row_above(post_migration_book):
    wb, ws = sheet.open_for_write(post_migration_book, SHEET)
    sheet.append_entry(ws, dt.date(2026, 3, 1), 5)
    sheet.save(wb, post_migration_book)

    ws2 = openpyxl.load_workbook(post_migration_book)[SHEET]
    assert ws2["A5"].number_format == ws2["A4"].number_format == "mm-dd-yy"


def test_append_leaves_the_diff_formula_alone(post_migration_book):
    wb, ws = sheet.open_for_write(post_migration_book, SHEET)
    before = ws["B5"].value
    sheet.append_entry(ws, dt.date(2026, 3, 1), 5)
    sheet.save(wb, post_migration_book)

    ws2 = openpyxl.load_workbook(post_migration_book)[SHEET]
    assert ws2["B5"].value == before
    assert "Table2[[#This Row],[Date]]" in ws2["B5"].value


def test_append_leaves_the_comment_column_untouched(post_migration_book):
    wb, ws = sheet.open_for_write(post_migration_book, SHEET)
    sheet.append_entry(ws, dt.date(2026, 3, 1), 5)
    sheet.save(wb, post_migration_book)

    ws2 = openpyxl.load_workbook(post_migration_book)[SHEET]
    assert ws2["E1"].value == "Dage siden sidste skift"
    assert isinstance(ws2["E2"].value, ArrayFormula)
    assert ws2["E2"].value.text == COUNTER_TEXT


def test_append_does_not_change_the_table_ref(post_migration_book):
    wb, ws = sheet.open_for_write(post_migration_book, SHEET)
    before = ws.tables["Table2"].ref
    sheet.append_entry(ws, dt.date(2026, 3, 1), 5)
    sheet.save(wb, post_migration_book)

    ws2 = openpyxl.load_workbook(post_migration_book)[SHEET]
    assert ws2.tables["Table2"].ref == before


def test_save_reports_a_locked_workbook(post_migration_book, monkeypatch):
    wb, ws = sheet.open_for_write(post_migration_book, SHEET)

    def boom(*args, **kwargs):
        raise PermissionError(13, "Permission denied")

    monkeypatch.setattr(wb, "save", boom)
    with pytest.raises(sheet.WorkbookLocked):
        sheet.save(wb, post_migration_book)


def test_open_for_write_reports_a_missing_sheet(post_migration_book):
    with pytest.raises(sheet.SheetMissing):
        sheet.open_for_write(post_migration_book, "Findes Ikke")
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `.\.venv\Scripts\python.exe -m pytest tests/test_sheet_write.py -v`
Expected: FAIL with `AttributeError: module 'sengetoej.sheet' has no attribute 'open_for_write'`.

- [ ] **Step 3: Append to `projects/sengetoej/sengetoej/sheet.py`**

```python
class WorkbookLocked(Exception):
    """Excel is holding the file open, so it cannot be saved."""


def open_for_write(path: Path, sheet_name: str):
    """Load the workbook for editing. Costs ~6.5s on the real file.

    read_only mode cannot be used here: it forbids ws.cell() and cannot save.
    """
    wb = openpyxl.load_workbook(path)
    if sheet_name not in wb.sheetnames:
        wb.close()
        raise SheetMissing(sheet_name)
    return wb, wb[sheet_name]


def dates_and_next_row(ws) -> tuple[list[dt.date], int]:
    """The same scan as read_dates, over an already-open worksheet."""
    def column_a():
        row = FIRST_DATA_ROW
        while True:
            yield ws.cell(row, DATE_COL).value
            row += 1

    return scan(column_a())


def has_cli_column(ws) -> bool:
    header = ws.cell(1, CLI_COL).value
    return isinstance(header, str) and header.strip().lower() == CLI_HEADER


def append_entry(ws, when: dt.date, row: int) -> None:
    """Write the date and the cli flag. Two cells, nothing else.

    Column B already holds the Diff formula in this row and every row below,
    and Table2's ref already spans the full column, so neither needs touching.
    """
    target = ws.cell(row, DATE_COL)
    target.value = dt.datetime(when.year, when.month, when.day)
    above = ws.cell(row - 1, DATE_COL)
    if above.value is not None:
        target.number_format = above.number_format
    ws.cell(row, CLI_COL, 1)


def save(wb, path: Path) -> None:
    try:
        wb.save(path)
    except PermissionError as exc:
        raise WorkbookLocked(path) from exc
    finally:
        wb.close()
```

- [ ] **Step 4: Run the test to verify it passes**

Run: `.\.venv\Scripts\python.exe -m pytest tests/test_sheet_write.py -v`
Expected: 11 passed.

- [ ] **Step 5: Run the whole suite**

Run: `.\.venv\Scripts\python.exe -m pytest -q`
Expected: 27 passed.

- [ ] **Step 6: Commit**

```bash
git add projects/sengetoej/sengetoej/sheet.py projects/sengetoej/tests/test_sheet_write.py
git commit -m "feat(sengetoej): append a dated row with the cli flag"
```

---

### Task 5: Migration job 1 — move the comments D to E

**Files:**
- Create: `projects/sengetoej/sengetoej/migrate.py`
- Test: `projects/sengetoej/tests/test_migrate_comments.py`

**Interfaces:**
- Consumes: `sheet.dates_and_next_row`, `sheet.BLANK_RUN`.
- Produces:
  - `migrate.COMMENT_COL_BEFORE: int = 4`, `migrate.COMMENT_COL_AFTER: int = 5`
  - `migrate.comment_rows(ws, col: int) -> list[int]`
  - `migrate.move_comments(ws) -> int` — returns how many cells moved

**Why not `move_range`:** openpyxl's `ws.move_range(..., translate=True)` rewrites relative references by the offset moved, which would turn the counter's `INDEX(A:A,COUNTA(A:A))` into `INDEX(B:B,COUNTA(B:B))` — silently repointing it at the `Diff` column. `translate=False` avoids that but still does not rebuild the `ArrayFormula`'s own `ref`. A cell-by-cell move is explicit about both, and was verified end-to-end during planning.

- [ ] **Step 1: Write the failing test `projects/sengetoej/tests/test_migrate_comments.py`**

```python
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
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `.\.venv\Scripts\python.exe -m pytest tests/test_migrate_comments.py -v`
Expected: FAIL — `ImportError: cannot import name 'migrate'`.

- [ ] **Step 3: Write `projects/sengetoej/sengetoej/migrate.py`**

```python
"""One-time restructuring of the Sengetøj sheet.

Two jobs, one save:

  1. Move the free-text comments one column right, D -> E. Dropping the new
     `cli` column into C would otherwise consume the blank column that
     separates the table from the comments, leaving the table butting
     straight up against them.

  2. Add the `cli` column to Table2 — header, ref, autofilter and
     tableColumns, all of which must agree or Excel calls the file corrupt.

Run once, deliberately, never from the CLI. It takes a backup first.
"""

import datetime as dt
import shutil
from copy import copy
from pathlib import Path

from openpyxl.worksheet.formula import ArrayFormula

from . import sheet

COMMENT_COL_BEFORE = 4  # D
COMMENT_COL_AFTER = 5   # E


def comment_rows(ws, col: int) -> list[int]:
    """Populated rows in `col`, searched to a bound derived from the data.

    The comments include headings above the data and remarks beside it, and
    they are sparse, so this walks a fixed window rather than stopping at the
    first blank.
    """
    _, next_row = sheet.dates_and_next_row(ws)
    return [r for r in range(1, next_row + sheet.BLANK_RUN)
            if ws.cell(r, col).value is not None]


def move_comments(ws) -> int:
    """Move every populated comment one column right. Returns the count.

    Deliberately cell-by-cell rather than ws.move_range: move_range with
    translate=True would rewrite the counter's INDEX(A:A,COUNTA(A:A)) into
    INDEX(B:B,COUNTA(B:B)), and even with translate=False it would not
    rebuild the ArrayFormula's own ref.
    """
    rows = comment_rows(ws, COMMENT_COL_BEFORE)

    for row in rows:
        src = ws.cell(row, COMMENT_COL_BEFORE)
        dst = ws.cell(row, COMMENT_COL_AFTER)

        value = src.value
        if isinstance(value, ArrayFormula):
            # The ref names the cell the formula lives in, so it moves too.
            value = ArrayFormula(ref=f"{dst.column_letter}{row}", text=value.text)

        dst.value = value
        dst._style = copy(src._style)

        src.value = None
        src._style = copy(ws.cell(row, COMMENT_COL_AFTER + 1)._style)

    before = ws.cell(1, COMMENT_COL_BEFORE).column_letter
    after = ws.cell(1, COMMENT_COL_AFTER).column_letter
    if before in ws.column_dimensions:
        ws.column_dimensions[after].width = ws.column_dimensions[before].width
        del ws.column_dimensions[before]

    return len(rows)
```

- [ ] **Step 4: Run the test to verify it passes**

Run: `.\.venv\Scripts\python.exe -m pytest tests/test_migrate_comments.py -v`
Expected: 9 passed.

- [ ] **Step 5: Commit**

```bash
git add projects/sengetoej/sengetoej/migrate.py projects/sengetoej/tests/test_migrate_comments.py
git commit -m "feat(sengetoej): move sheet comments from column D to E"
```

---

### Task 6: Migration job 2 — the cli column, plus the runnable script

**Files:**
- Modify: `projects/sengetoej/sengetoej/migrate.py` (append)
- Test: `projects/sengetoej/tests/test_migrate_column.py`

**Interfaces:**
- Consumes: Task 5's `move_comments`, Task 4's `open_for_write` / `save` / `has_cli_column`.
- Produces:
  - `migrate.TABLE_NAME: str = "Table2"`
  - `migrate.add_cli_column(ws) -> None`
  - `migrate.backup(path: Path) -> Path`
  - `migrate.plan(ws) -> list[str]` — human-readable description of pending changes; empty when nothing to do
  - `migrate.main(argv=None) -> int`

- [ ] **Step 1: Write the failing test `projects/sengetoej/tests/test_migrate_column.py`**

```python
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
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `.\.venv\Scripts\python.exe -m pytest tests/test_migrate_column.py -v`
Expected: FAIL with `AttributeError: module 'sengetoej.migrate' has no attribute 'add_cli_column'`.

- [ ] **Step 3: Append to `projects/sengetoej/sengetoej/migrate.py`**

```python
import argparse
import sys

from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import TableColumn

from .env import excel_path, load_env, sheet_name

TABLE_NAME = "Table2"


def add_cli_column(ws) -> None:
    """Add `cli` to Table2. All four pieces must agree or Excel objects."""
    ws.cell(1, sheet.CLI_COL, sheet.CLI_HEADER)

    table = ws.tables[TABLE_NAME]
    last_col = get_column_letter(sheet.CLI_COL)
    bottom = table.ref.split(":")[1]
    bottom_row = "".join(ch for ch in bottom if ch.isdigit())
    new_ref = f"A1:{last_col}{bottom_row}"

    table.ref = new_ref
    if table.autoFilter is not None:
        table.autoFilter.ref = new_ref

    next_id = max((c.id for c in table.tableColumns), default=0) + 1
    table.tableColumns.append(TableColumn(id=next_id, name=sheet.CLI_HEADER))


def backup(path: Path) -> Path:
    stamp = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    target = path.with_name(f"{path.stem}.{stamp}.bak{path.suffix}")
    shutil.copy2(path, target)
    return target


def plan(ws) -> list[str]:
    """What still needs doing. Empty means the sheet is already migrated."""
    steps = []
    if comment_rows(ws, COMMENT_COL_BEFORE):
        steps.append(f"Flyt kommentarer D -> E "
                     f"({len(comment_rows(ws, COMMENT_COL_BEFORE))} celler)")
    if not sheet.has_cli_column(ws):
        steps.append(f'Tilføj kolonnen "{sheet.CLI_HEADER}" til {TABLE_NAME} (C)')
    return steps


def _blocked(ws) -> str | None:
    """A reason to refuse, or None."""
    header = ws.cell(1, sheet.CLI_COL).value
    if header is not None and not sheet.has_cli_column(ws):
        return (f"C1 indeholder allerede {header!r}, ikke "
                f"{sheet.CLI_HEADER!r}. Migreringen er afbrudt.")
    if comment_rows(ws, COMMENT_COL_BEFORE) and comment_rows(ws, COMMENT_COL_AFTER):
        return ("Kolonne E er ikke tom, og kommentarerne står stadig i D. "
                "Migreringen er afbrudt.")
    return None


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        prog="python -m sengetoej.migrate",
        description="Engangsmigrering: flyt kommentarer D -> E og "
                    'tilføj kolonnen "cli" til Table2.')
    parser.add_argument("-y", "--yes", action="store_true",
                        help="spring bekræftelsen over")
    args = parser.parse_args(argv)

    load_env()
    path, tab = excel_path(), sheet_name()

    try:
        wb, ws = sheet.open_for_write(path, tab)
    except FileNotFoundError:
        print(f"Fejl: filen findes ikke: {path}", file=sys.stderr)
        return 1
    except sheet.SheetMissing:
        print(f"Fejl: arket {tab!r} findes ikke i {path}", file=sys.stderr)
        return 1

    reason = _blocked(ws)
    if reason:
        print(f"Fejl: {reason}", file=sys.stderr)
        wb.close()
        return 1

    steps = plan(ws)
    if not steps:
        print("Arket er allerede migreret. Intet at gøre.")
        wb.close()
        return 0

    print(f"{path}  [{tab}]")
    for step in steps:
        print(f"  - {step}")

    if not args.yes and input("Udfør migreringen? [y/N] ").strip().lower() not in ("y", "yes"):
        print("Afbrudt.")
        wb.close()
        return 0

    made = backup(path)
    print(f"Sikkerhedskopi: {made.name}")

    moved = move_comments(ws)
    if not sheet.has_cli_column(ws):
        add_cli_column(ws)

    try:
        sheet.save(wb, path)
    except sheet.WorkbookLocked:
        print("Fejl: Luk Excel og prøv igen.", file=sys.stderr)
        return 1

    print(f"Færdig. {moved} kommentar(er) flyttet, kolonnen "
          f'"{sheet.CLI_HEADER}" tilføjet.')
    return 0


if __name__ == "__main__":
    sys.exit(main())
```

Move the `import argparse` / `import sys` and the `from .env import ...` lines up to the module's existing import block rather than leaving them mid-file.

- [ ] **Step 4: Run the test to verify it passes**

Run: `.\.venv\Scripts\python.exe -m pytest tests/test_migrate_column.py -v`
Expected: 14 passed.

- [ ] **Step 5: Run the whole suite**

Run: `.\.venv\Scripts\python.exe -m pytest -q`
Expected: 50 passed.

- [ ] **Step 6: Commit**

```bash
git add projects/sengetoej/sengetoej/migrate.py projects/sengetoej/tests/test_migrate_column.py
git commit -m "feat(sengetoej): add the cli column and the runnable migration"
```

---

### Task 7: `--last`

**Files:**
- Create: `projects/sengetoej/sengetoej/cli.py`, `projects/sengetoej/sengetoej/__main__.py`
- Test: `projects/sengetoej/tests/test_cli_last.py`

**Interfaces:**
- Consumes: `sheet.read_dates`, `sheet.gaps`, `env.*`.
- Produces:
  - `cli.fmt_date(d: date) -> str` — `dd/mm/yyyy`
  - `cli.days_ago_phrase(days: int) -> str`
  - `cli.parse_date(text: str) -> date` — raises `cli.BadDate`
  - `cli.BadDate(Exception)`
  - `cli.build_parser() -> argparse.ArgumentParser`
  - `cli.main(argv=None) -> int`
  - `cli.run() -> None` — console-script entry, `sys.exit(main())`

- [ ] **Step 1: Write the failing test `projects/sengetoej/tests/test_cli_last.py`**

```python
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
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `.\.venv\Scripts\python.exe -m pytest tests/test_cli_last.py -v`
Expected: FAIL — `ImportError: cannot import name 'cli'`.

- [ ] **Step 3: Write `projects/sengetoej/sengetoej/cli.py`**

```python
"""Command-line interface for the Sengetøj tab of Træning.xlsx.

Usage:
    sengetoej                    same as --last
    sengetoej --last             sidste skift + dage siden
    sengetoej --last N           de sidste N skift
    sengetoej --new              registrer et skift i dag
    sengetoej --new dd/mm/yyyy   registrer et skift på den dato
    sengetoej --new ... -y       spring bekræftelsen over

Reads EXCEL_PATH and SENGETOEJ_SHEET from the environment or a local .env.
Both have working defaults, so no configuration is required.
"""

import argparse
import datetime as dt
import sys

from . import sheet
from .env import excel_path, load_env, sheet_name

DATE_FORMAT = "%d/%m/%Y"

# Sentinel for a bare `--last`, which takes no value. A real count is >= 1,
# and `--last 0` is rejected by positive_int, so 0 can never arrive any
# other way.
BARE = 0


class BadDate(Exception):
    """The date argument was not dd/mm/yyyy."""


def today() -> dt.date:
    """Indirection so tests can pin the clock."""
    return dt.date.today()


def fmt_date(d: dt.date) -> str:
    return d.strftime(DATE_FORMAT)


def days_ago_phrase(days: int) -> str:
    if days < 0:
        return f"om {abs(days)} dage"
    if days == 0:
        return "i dag"
    if days == 1:
        return "i går"
    return f"{days} dage siden"


def parse_date(text: str) -> dt.date:
    """dd/mm/yyyy, leading zeros optional. One format on purpose: accepting
    ISO as well would reintroduce the day/month ambiguity this avoids."""
    try:
        return dt.datetime.strptime(text.strip(), DATE_FORMAT).date()
    except ValueError:
        raise BadDate(text) from None


def positive_int(text: str) -> int:
    value = int(text)  # argparse turns a ValueError here into exit 2
    if value < 1:
        raise ValueError(text)
    return value


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="sengetoej",
        description="Aflæs og registrer sengetøjsskift.")
    group = parser.add_mutually_exclusive_group()
    group.add_argument("--last", nargs="?", type=positive_int, const=BARE,
                       default=None, metavar="N",
                       help="sidste skift, eller de sidste N skift")
    group.add_argument("--new", nargs="?", const="", default=None,
                       metavar="dd/mm/yyyy",
                       help="registrer et skift (uden dato: i dag)")
    parser.add_argument("-y", "--yes", action="store_true",
                        help="spring bekræftelsen ved --new over")
    return parser


def _load_dates():
    """Returns (dates, None) or (None, exit_code) after printing the error."""
    load_env()
    path, tab = excel_path(), sheet_name()
    try:
        return sheet.read_dates(path, tab), None
    except FileNotFoundError:
        print(f"Fejl: filen findes ikke: {path}", file=sys.stderr)
        return None, 1
    except sheet.SheetMissing:
        print(f"Fejl: arket {tab!r} findes ikke i {path}", file=sys.stderr)
        return None, 1


def cmd_last(count: int) -> int:
    dates, failure = _load_dates()
    if failure is not None:
        return failure

    if not dates:
        print("Ingen skift registreret.")
        return 0

    if count == BARE:
        last = dates[-1]
        print(f"Sidste skift: {fmt_date(last)} "
              f"({days_ago_phrase((today() - last).days)})")
        return 0

    all_gaps = sheet.gaps(dates)
    start = max(0, len(dates) - count)
    for d, gap in zip(dates[start:], all_gaps[start:]):
        interval = "—".rjust(7) if gap is None else f"{gap:>4} dage"
        print(f"  {fmt_date(d)}    {interval}")
    return 0


def main(argv=None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    if args.new is None:
        if args.yes:
            print("Fejl: -y giver kun mening sammen med --new.", file=sys.stderr)
            return 2
        return cmd_last(args.last if args.last is not None else BARE)

    return cmd_new(args.new, args.yes)


def run() -> None:
    sys.exit(main())
```

`cmd_new` arrives in Task 8. To keep this task's suite green, add a placeholder that the next task replaces:

```python
def cmd_new(value: str, assume_yes: bool) -> int:
    raise NotImplementedError  # Task 8
```

- [ ] **Step 4: Write `projects/sengetoej/sengetoej/__main__.py`**

```python
from .cli import run

if __name__ == "__main__":
    run()
```

- [ ] **Step 5: Run the test to verify it passes**

Run: `.\.venv\Scripts\python.exe -m pytest tests/test_cli_last.py -v`
Expected: 16 passed.

- [ ] **Step 6: Commit**

```bash
git add projects/sengetoej/sengetoej/cli.py projects/sengetoej/sengetoej/__main__.py \
        projects/sengetoej/tests/test_cli_last.py
git commit -m "feat(sengetoej): add --last"
```

---

### Task 8: `--new`

**Files:**
- Modify: `projects/sengetoej/sengetoej/cli.py` (replace the `cmd_new` placeholder)
- Test: `projects/sengetoej/tests/test_cli_new.py`

**Interfaces:**
- Consumes: everything from Tasks 4 and 7.
- Produces: a working `cli.cmd_new(value: str, assume_yes: bool) -> int`.

Validation order is fixed, and one test pins it: format → future → duplicate → earlier. A date that is both a duplicate and the last entry must report as a duplicate.

- [ ] **Step 1: Write the failing test `projects/sengetoej/tests/test_cli_new.py`**

```python
import datetime as dt

import openpyxl
import pytest

from sengetoej import cli

from .conftest import SHEET


@pytest.fixture
def book_env(post_migration_book, monkeypatch):
    monkeypatch.setenv("EXCEL_PATH", str(post_migration_book))
    monkeypatch.setenv("SENGETOEJ_SHEET", SHEET)
    monkeypatch.setattr(cli, "today", lambda: dt.date(2026, 2, 10))
    return post_migration_book


def rows(path):
    ws = openpyxl.load_workbook(path)[SHEET]
    return [(ws.cell(r, 1).value, ws.cell(r, 3).value) for r in range(2, 7)]


def test_new_with_a_date_appends_after_confirmation(book_env, monkeypatch, capsys):
    monkeypatch.setattr("builtins.input", lambda _: "y")
    assert cli.main(["--new", "09/02/2026"]) == 0

    assert rows(book_env)[3] == (dt.datetime(2026, 2, 9), 1)
    assert "Tilføjet i række 5." in capsys.readouterr().out


def test_new_without_a_date_uses_today(book_env, monkeypatch):
    monkeypatch.setattr("builtins.input", lambda _: "y")
    assert cli.main(["--new"]) == 0
    assert rows(book_env)[3][0] == dt.datetime(2026, 2, 10)


def test_the_prompt_shows_the_date_and_the_interval(book_env, monkeypatch, capsys):
    seen = {}

    def fake_input(prompt):
        seen["prompt"] = prompt
        return "y"

    monkeypatch.setattr("builtins.input", fake_input)
    cli.main(["--new", "09/02/2026"])
    assert "09/02/2026" in seen["prompt"]
    assert "6 dage siden sidste skift" in seen["prompt"]


def test_declining_writes_nothing_and_exits_zero(book_env, monkeypatch, capsys):
    monkeypatch.setattr("builtins.input", lambda _: "n")
    assert cli.main(["--new", "09/02/2026"]) == 0
    assert rows(book_env)[3] == (None, None)
    assert "Afbrudt." in capsys.readouterr().out


def test_yes_flag_skips_the_prompt(book_env, monkeypatch):
    def no_input(_):
        raise AssertionError("should not have prompted")

    monkeypatch.setattr("builtins.input", no_input)
    assert cli.main(["--new", "09/02/2026", "-y"]) == 0
    assert rows(book_env)[3][0] == dt.datetime(2026, 2, 9)


def test_a_malformed_date_is_a_usage_error(book_env, capsys):
    assert cli.main(["--new", "2026-02-09", "-y"]) == 2
    assert "dd/mm/yyyy" in capsys.readouterr().err
    assert rows(book_env)[3] == (None, None)


def test_leading_zeros_are_optional(book_env, monkeypatch):
    monkeypatch.setattr("builtins.input", lambda _: "y")
    assert cli.main(["--new", "9/2/2026"]) == 0
    assert rows(book_env)[3][0] == dt.datetime(2026, 2, 9)


def test_a_future_date_is_rejected(book_env, capsys):
    assert cli.main(["--new", "01/03/2026", "-y"]) == 2
    assert "fremtiden" in capsys.readouterr().err
    assert rows(book_env)[3] == (None, None)


def test_a_duplicate_date_is_rejected(book_env, capsys):
    assert cli.main(["--new", "03/02/2026", "-y"]) == 2
    err = capsys.readouterr().err
    assert "findes allerede" in err
    assert rows(book_env)[3] == (None, None)


def test_an_earlier_date_is_rejected(book_env, capsys):
    assert cli.main(["--new", "20/01/2026", "-y"]) == 2
    err = capsys.readouterr().err
    assert "ligger før sidste række" in err
    assert "03/02/2026" in err


def test_a_missing_cli_column_refuses_and_points_at_the_migration(
        pre_migration_book, monkeypatch, capsys):
    monkeypatch.setenv("EXCEL_PATH", str(pre_migration_book))
    monkeypatch.setenv("SENGETOEJ_SHEET", SHEET)
    monkeypatch.setattr(cli, "today", lambda: dt.date(2026, 2, 10))

    assert cli.main(["--new", "09/02/2026", "-y"]) == 1
    err = capsys.readouterr().err
    assert "cli" in err
    assert "migrate" in err


def test_a_locked_workbook_is_reported(book_env, monkeypatch, capsys):
    def locked(wb, path):
        raise cli.sheet.WorkbookLocked(path)

    monkeypatch.setattr(cli.sheet, "save", locked)
    assert cli.main(["--new", "09/02/2026", "-y"]) == 1
    assert "Luk Excel" in capsys.readouterr().err


def test_new_on_an_empty_sheet_appends_the_first_row(tmp_path, monkeypatch):
    from openpyxl.worksheet.table import Table
    path = tmp_path / "empty.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = SHEET
    ws["A1"], ws["B1"], ws["C1"] = "Date", "Diff", "cli"
    ws.add_table(Table(displayName="Table2", ref="A1:C400"))
    wb.save(path)

    monkeypatch.setenv("EXCEL_PATH", str(path))
    monkeypatch.setenv("SENGETOEJ_SHEET", SHEET)
    monkeypatch.setattr(cli, "today", lambda: dt.date(2026, 2, 10))

    assert cli.main(["--new", "-y"]) == 0
    ws2 = openpyxl.load_workbook(path)[SHEET]
    assert ws2["A2"].value == dt.datetime(2026, 2, 10)
    assert ws2["C2"].value == 1
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `.\.venv\Scripts\python.exe -m pytest tests/test_cli_new.py -v`
Expected: FAIL with `NotImplementedError`.

- [ ] **Step 3: Replace the `cmd_new` placeholder in `cli.py`**

```python
def cmd_new(value: str, assume_yes: bool) -> int:
    try:
        when = parse_date(value) if value else today()
    except BadDate:
        print(f"Fejl: {value!r} er ikke en gyldig dato. Brug dd/mm/yyyy.",
              file=sys.stderr)
        return 2

    if when > today():
        print(f"Fejl: {fmt_date(when)} ligger i fremtiden.", file=sys.stderr)
        return 2

    load_env()
    path, tab = excel_path(), sheet_name()

    # ~6.5s on the real workbook, and ~8.5s more to save. Say so, or it
    # looks hung.
    print("Åbner regnearket ...")
    try:
        wb, ws = sheet.open_for_write(path, tab)
    except FileNotFoundError:
        print(f"Fejl: filen findes ikke: {path}", file=sys.stderr)
        return 1
    except sheet.SheetMissing:
        print(f"Fejl: arket {tab!r} findes ikke i {path}", file=sys.stderr)
        return 1

    if not sheet.has_cli_column(ws):
        wb.close()
        print(f'Fejl: kolonnen "{sheet.CLI_HEADER}" mangler i Table2.\n'
              f"Kør migreringen først:\n"
              f"  .\\.venv\\Scripts\\python.exe -m sengetoej.migrate",
              file=sys.stderr)
        return 1

    dates, next_row = sheet.dates_and_next_row(ws)

    # Order matters: a date that is both a duplicate and the last entry must
    # report as a duplicate, not as out of order.
    if when in dates:
        wb.close()
        print(f"Fejl: {fmt_date(when)} findes allerede i arket.", file=sys.stderr)
        return 2

    if dates and when < dates[-1]:
        wb.close()
        print(f"Fejl: {fmt_date(when)} ligger før sidste række "
              f"({fmt_date(dates[-1])}).\n"
              f"Arket skal være i stigende rækkefølge — tilføj den i Excel.",
              file=sys.stderr)
        return 2

    if dates:
        interval = f"{(when - dates[-1]).days} dage siden sidste skift"
    else:
        interval = "første registrering"

    if not assume_yes:
        answer = input(f"Tilføj {fmt_date(when)}? ({interval}) [y/N] ")
        if answer.strip().lower() not in ("y", "yes"):
            wb.close()
            print("Afbrudt.")
            return 0

    sheet.append_entry(ws, when, next_row)
    try:
        sheet.save(wb, path)
    except sheet.WorkbookLocked:
        print("Fejl: Luk Excel og prøv igen.", file=sys.stderr)
        return 1

    print(f"Tilføjet i række {next_row}.")
    return 0
```

- [ ] **Step 4: Run the test to verify it passes**

Run: `.\.venv\Scripts\python.exe -m pytest tests/test_cli_new.py -v`
Expected: 14 passed.

- [ ] **Step 5: Run the whole suite**

Run: `.\.venv\Scripts\python.exe -m pytest -q`
Expected: 80 passed.

- [ ] **Step 6: Commit**

```bash
git add projects/sengetoej/sengetoej/cli.py projects/sengetoej/tests/test_cli_new.py
git commit -m "feat(sengetoej): add --new with confirmation and validation"
```

---

### Task 9: Install script, README, and the live migration

**Files:**
- Create: `projects/sengetoej/install.ps1`, `projects/sengetoej/README.md`
- Modify: `projects/README.md` (add the tool to the list, matching the existing format)

This task ends with the only writes to the real workbook in this plan. They are backed up and confirmed.

- [ ] **Step 1: Write `projects/sengetoej/install.ps1`**

ASCII only — PowerShell 5.1 decodes a BOM-less script as ANSI. Adapted from `projects/hevy/install.ps1`, minus the API-key handling.

```powershell
# Creates this tool's virtualenv, installs it, and puts `sengetoej` on your PATH.
# No administrator rights. Safe to re-run.
#
# Only bin\ goes on PATH, never the project folder itself. Two reasons:
#   1. PowerShell searches .ps1 ahead of PATHEXT, so a sengetoej.ps1 sitting
#      next to sengetoej.cmd on PATH would win, then die on the execution policy.
#   2. The project folder also holds install.ps1 -- a generic name that would
#      become a global command in every terminal.
#
# ASCII only on purpose: PowerShell 5.1 decodes a BOM-less script as ANSI, so a
# stray non-ASCII character here would come out mangled.
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$root    = $PSScriptRoot
$venv    = Join-Path $root '.venv'
$venvPy  = Join-Path $venv 'Scripts\python.exe'
$binDir  = Join-Path $root 'bin'
$envFile = Join-Path $root '.env'
$example = Join-Path $root '.env.example'

# --- 1. virtualenv -----------------------------------------------------------
if (Test-Path $venvPy) {
    Write-Host 'Using the existing .venv'
} else {
    Write-Host 'Creating .venv ...'
    python -m venv $venv
    if (-not (Test-Path $venvPy)) {
        throw "venv creation failed -- no interpreter at $venvPy"
    }
}

# --- 2. dependencies ---------------------------------------------------------
# Editable install, so edits to sengetoej\ take effect without reinstalling.
Write-Host 'Installing dependencies (this needs network) ...'
& $venvPy -m pip install --upgrade pip --quiet
if ($LASTEXITCODE -ne 0) { throw 'pip self-upgrade failed' }
& $venvPy -m pip install --editable "$root[dev]" --quiet
if ($LASTEXITCODE -ne 0) { throw 'pip install failed' }

$exe = Join-Path $venv 'Scripts\sengetoej.exe'
if (-not (Test-Path $exe)) { throw "install finished but $exe is missing" }

# --- 3. .env -----------------------------------------------------------------
# Optional here: both settings have working defaults in code.
if (Test-Path $envFile) {
    Write-Host 'Found an existing .env -- left untouched.'
} else {
    Copy-Item $example $envFile
    Write-Host 'Created .env from .env.example.'
}

# --- 4. PATH -----------------------------------------------------------------
# Compare PATH entries case-insensitively, ignoring a trailing backslash.
function Get-NormalizedPathEntry {
    param([string]$Entry)
    return $Entry.TrimEnd('\').ToLowerInvariant()
}

$binKey  = Get-NormalizedPathEntry $binDir
$current = [Environment]::GetEnvironmentVariable('Path', 'User')
if ($null -eq $current) { $current = '' }

# @() keeps this an array even when PATH holds a single entry.
$parts = @($current -split ';' | Where-Object { $_.Trim() -ne '' })
$normalized = @($parts | ForEach-Object { Get-NormalizedPathEntry $_ })

if ($normalized -contains $binKey) {
    Write-Host "Already on your PATH:`n  $binDir"
} else {
    [Environment]::SetEnvironmentVariable('Path', (@($parts + $binDir) -join ';'), 'User')
    Write-Host "Added to your user PATH:`n  $binDir"
}

# --- done --------------------------------------------------------------------
Write-Host ''
Write-Host 'ONE THING LEFT: run the one-time migration, which adds the "cli"'
Write-Host 'column and moves the comments one column right. Close Excel first.'
Write-Host "  $venvPy -m sengetoej.migrate"
Write-Host ''
Write-Host 'Then open a NEW terminal and run:  sengetoej --last'
```

- [ ] **Step 2: Run the installer and confirm the console script exists**

```powershell
cd "C:\Users\viggo\Git Clone\PC-CLI\projects\sengetoej"
.\install.ps1
```

Expected: finishes without throwing, and reports either "Added to your user PATH" or "Already on your PATH".

- [ ] **Step 3: Run the full suite once more against the installed package**

Run: `.\.venv\Scripts\python.exe -m pytest -q`
Expected: 80 passed.

- [ ] **Step 4: Write `projects/sengetoej/README.md`**

Cover, in this order, matching hevy's README structure: purpose; Stack; Install; Usage (the command table and worked examples from the spec); Configuration (the `EXCEL_PATH` / `SENGETOEJ_SHEET` table, noting both default in code); How it works (the column-A scan and why `read_only` with an early break, the two-cell append, why column B and the table ref are never written); Migration (what it changed and that it is one-time); Tests. State plainly that **`--last` works with Excel open, `--new` does not**, and that `--new` takes about 15 seconds because the sheet stores over a million cells.

- [ ] **Step 5: Add the tool to `projects/README.md`**

Read the existing file first and follow its format exactly. Do not restructure it.

- [ ] **Step 6: Commit the docs and installer**

```bash
git add projects/sengetoej/install.ps1 projects/sengetoej/README.md projects/README.md
git commit -m "docs(sengetoej): add README and install script"
```

- [ ] **Step 7: Back up the real workbook by hand**

```powershell
Copy-Item "C:\Users\viggo\OneDrive\Privat\Fitness\Træning.xlsx" `
          "$env:USERPROFILE\Desktop\Traening-before-sengetoej.xlsx"
```

This is belt-and-braces: `migrate.py` takes its own backup. Do it anyway — this is the user's real training history.

- [ ] **Step 8: Confirm Excel is closed, then dry-run the migration by declining**

```powershell
cd "C:\Users\viggo\Git Clone\PC-CLI\projects\sengetoej"
.\.venv\Scripts\python.exe -m sengetoej.migrate
```

Answer `n`. Expected output lists exactly two steps — moving 11 comment cells D → E, and adding the `cli` column — then `Afbrudt.` Confirm the workbook is unchanged before going further.

- [ ] **Step 9: Run the migration for real**

```powershell
.\.venv\Scripts\python.exe -m sengetoej.migrate
```

Answer `y`. Expected: a `Sikkerhedskopi:` line naming the `.bak.xlsx`, then `Færdig.`

- [ ] **Step 10: Verify in Excel — the check no test can make**

Open `Træning.xlsx` in Excel and confirm all five:

1. **No repair prompt** on open.
2. `Table2` covers **three** columns, with `cli` as the header of C.
3. The comments are in **column E**, each beside the date it belongs to — spot-check `Obs: Folkemødet` (2026-06-29), `Obs: Ferie` (2026-07-31) and `Amst` (2026-08-22).
4. The yellow **"Dage siden sidste skift"** counter still shows a plausible number. If it reads a small number that looks like a `Diff` value, the formula was translated — restore the backup and fix `move_comments`.
5. Column D is empty and still separates the table from the comments.

Then **close Excel**.

- [ ] **Step 11: Verify the CLI against the real workbook**

```powershell
sengetoej --last
sengetoej --last 5
```

Expected: the last entry is `30/08/2026`, and `--last 5` prints five rows ending `30/08/2026     8 dage`. Both should return effectively instantly; if either takes seconds, the read is not using `read_only` with an early break.

- [ ] **Step 12: Commit any fixes, then record the outcome**

If steps 8–11 required code changes, commit them:

```bash
git add projects/sengetoej
git commit -m "fix(sengetoej): corrections found during the live migration"
```

Report to the user: how many comment cells moved, that the `cli` column exists, where both backups are, and that the next `sengetoej --new` will be the first row flagged `cli = 1`.

---

## Self-Review

**Spec coverage:**

| Spec section | Task |
|---|---|
| Purpose, scope | 7, 8 |
| Stack, layout, `.env` defaults | 1 |
| `cli` column, `1`/blank convention | 4 (write), 6 (schema) |
| Migration job 1 (comments D→E) | 5 |
| Migration job 2 (schema), backup, prompt, idempotence, abort conditions | 6 |
| Guard: `--new` refuses without the column | 8 |
| Commands, exit codes, mutual exclusion, `-y` misuse | 7 |
| `--last` bare and N, oldest-first, `—`, wording, empty sheet | 7 |
| `--new` validation order, `dd/mm/yyyy`, prompt, decline | 8 |
| Reading (`read_only`, early break, never `Diff`) | 3 |
| Writing (two cells, B and ref untouched, `PermissionError`) | 4 |
| Testing, both fixture variants | 2, and every later task |
| Manual Excel verification | 9 |

No gaps.

**Placeholder scan:** One deliberate placeholder — `cmd_new` in Task 7 raises `NotImplementedError` and is replaced in Task 8, with both the creation and the replacement written out in full. Everything else carries real code.

**Type consistency:** `sheet.scan` returns `(list[date], int)` and is consumed under that shape by `read_dates` and `dates_and_next_row`. `sheet.gaps` returns `list[int | None]`, and Task 7 renders the `None` as `—`. `cli.BARE = 0` is the bare-`--last` sentinel, unreachable from `positive_int`, which rejects anything below 1. `migrate.main` and `cli.main` both take `argv=None` and return `int`. Test counts accumulate 3 → 5 → 16 → 27 → 36 → 50 → 66 → 80.

**Two additions to the spec, flagged for the user:**

1. **`--new` prints `Åbner regnearket ...` before loading.** The spec did not anticipate that the round-trip costs ~15 seconds; silence that long reads as a hang.
2. **`sheet.SheetMissing`** — the spec covers a locked file and a missing column but not a missing or renamed tab, which is reachable through `SENGETOEJ_SHEET`.

## Execution Handoff

Plan complete and saved to `docs/superpowers/plans/2026-08-30-sengetoej.md`. Two execution options:

**1. Subagent-Driven (recommended)** — a fresh subagent per task, review between tasks, fast iteration

**2. Inline Execution** — execute tasks in this session using executing-plans, batch execution with checkpoints

Which approach?
