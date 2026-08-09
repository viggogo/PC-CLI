# Hevy Sync Column Mapping Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make the Hevy→Træning.xlsx sync produce rows that match the user's 650-row hand-kept format — translated categories, numeric rating, cardio registered in AddCardio, a `Claude` provenance column, and a live `Total` formula — then heal the 12 rows the old buggy tool wrote.

**Architecture:** All Hevy→row logic lives in `hevy_mcp/column_mapper.py` (pure, unit-tested). `hevy_mcp/excel_writer.py` knows the physical sheet layout, writes the `Total` formula (never a static value), and appends rows. A one-time `hevy_mcp/migrate_add_claude_column.py` inserts the `Claude` column into the Excel Table and heals formulas. A new `fix_synced_rows` MCP tool re-maps and overwrites the 12 buggy rows.

**Tech Stack:** Python 3.11+, openpyxl 3.1.5, httpx, mcp[cli], pytest (added as dev dependency).

## Global Constraints

- Target file: `C:\Users\viggo\OneDrive\Træning.xlsx`, sheet `Træning`, Excel Table `Table4`.
- Physical column layout **after** the Claude-column migration (1-based):
  `Date=1, Place=2, Type=3, Time=4, Mave=5, AddCardio=6, AddCardio2=7, Total=8, Ensamble=9, Rating=10, Claude=11, Comments=12`.
- `Total` (col 8/H) is a Table calculated column and must **always** be the formula
  `=IF(ISBLANK(Table4[[#This Row],[Date]]), " ", SUM(D{r}:G{r})-IF(E{r} = 1, 1, 0))` for row `{r}` — never a static number.
- `Mave`, `AddCardio`, `AddCardio2`, `Ensamble`, `Claude` are numeric; `Rating` is an `int` or blank.
- `Claude` = `1` for tool-written rows, `0` for the user's manual rows; historical pre-tool rows stay **empty** (the tool never backfills them).
- `Total = Time + AddCardio + AddCardio2`; cardio only goes to `AddCardio` when strength is also present.
- Windows shell is PowerShell; run pytest as `python -m pytest`.
- Commit after every task. Do not commit `__pycache__` or `.bak` files.

---

### Task 1: Test scaffolding and pytest dev dependency

**Files:**
- Modify: `pyproject.toml`
- Create: `tests/__init__.py`
- Create: `tests/conftest.py`
- Create: `.gitignore`

**Interfaces:**
- Consumes: nothing.
- Produces: a runnable `python -m pytest` setup; `tests/conftest.py` exports the fixture `make_workout(**overrides)` returning a Hevy-shaped workout dict, and `strength_ex()` / `cardio_ex()` / `abs_ex()` helpers.

- [ ] **Step 1: Add pytest as a dev dependency**

In `pyproject.toml`, after the `dependencies = [...]` block, add:

```toml
[project.optional-dependencies]
dev = ["pytest>=8.0"]
```

- [ ] **Step 2: Install pytest**

Run: `python -m pip install -e ".[dev]"`
Expected: installs pytest; `python -m pytest --version` prints a version ≥ 8.0.

- [ ] **Step 3: Create `.gitignore`**

```gitignore
__pycache__/
*.pyc
*.bak
*.egg-info/
.pytest_cache/
```

- [ ] **Step 4: Create `tests/__init__.py`**

Empty file:

```python
```

- [ ] **Step 5: Create `tests/conftest.py`**

```python
import pytest


def strength_ex(title="Bench Press (Barbell)", reps=8, weight_kg=60):
    return {"title": title, "sets": [
        {"reps": reps, "weight_kg": weight_kg,
         "duration_seconds": None, "distance_meters": None}]}


def cardio_ex(title="Treadmill", duration_seconds=900, distance_meters=0):
    return {"title": title, "sets": [
        {"reps": None, "weight_kg": None,
         "duration_seconds": duration_seconds, "distance_meters": distance_meters}]}


def abs_ex(title="Plank", duration_seconds=60):
    # A timed, weightless core move: must count as Mave, never cardio.
    return {"title": title, "sets": [
        {"reps": None, "weight_kg": None,
         "duration_seconds": duration_seconds, "distance_meters": None}]}


@pytest.fixture
def make_workout():
    def _make(title="push", description="", exercises=None,
              start="2026-07-04T11:00:00+00:00", end="2026-07-04T12:00:00+00:00"):
        return {
            "title": title,
            "description": description,
            "start_time": start,
            "end_time": end,
            "exercises": exercises if exercises is not None else [strength_ex()],
        }
    return _make
```

- [ ] **Step 6: Verify pytest collects with no errors**

Run: `python -m pytest -q`
Expected: `no tests ran` (exit code 5) — collection succeeds with zero tests. This confirms scaffolding imports cleanly.

- [ ] **Step 7: Commit**

```bash
git add pyproject.toml tests/__init__.py tests/conftest.py .gitignore
git commit -m "test: add pytest scaffolding and workout fixtures"
```

---

### Task 2: column_mapper string helpers (Type, Rating, Place, Ensamble)

**Files:**
- Modify: `hevy_mcp/column_mapper.py`
- Create: `tests/test_column_mapper.py`

**Interfaces:**
- Consumes: nothing.
- Produces:
  - `_map_type(title: str) -> str` — Danish→English map, else verbatim `.strip()`.
  - `_parse_rating(description) -> int | None` — numeric rating or `None`.
  - `_detect_place(description: str) -> str` — one of `"AC"`/`"Center"`/`"Tryg"` or `""`.
  - `_detect_ensamble(description: str) -> int` — `1`/`0`.

- [ ] **Step 1: Write failing tests**

Create `tests/test_column_mapper.py`:

```python
from hevy_mcp import column_mapper as cm


def test_map_type_danish_titles():
    assert cm._map_type("Over") == "upper"
    assert cm._map_type("Ben") == "legs"
    assert cm._map_type("Træk") == "pull"
    assert cm._map_type("Skub") == "push"


def test_map_type_english_passthrough():
    assert cm._map_type("push") == "push"
    assert cm._map_type("full body") == "full body"


def test_map_type_unknown_is_verbatim():
    assert cm._map_type("Crossfit WOD") == "Crossfit WOD"


def test_parse_rating_returns_int():
    assert cm._parse_rating("3") == 3
    assert cm._parse_rating("4 tryg") == 4
    assert cm._parse_rating("2, travlt") == 2
    assert cm._parse_rating("god træning 5") == 5
    assert cm._parse_rating("") is None
    assert cm._parse_rating("lynild") is None
    assert isinstance(cm._parse_rating("4"), int)


def test_detect_place():
    assert cm._detect_place("4, AC") == "AC"
    assert cm._detect_place("Tryg m Peter") == "Tryg"
    assert cm._detect_place("god center dag") == "Center"
    assert cm._detect_place("travlt") == ""
    assert cm._detect_place("") == ""


def test_detect_ensamble():
    assert cm._detect_ensamble("m Silas") == 1
    assert cm._detect_ensamble("med Uwufufu") == 1
    assert cm._detect_ensamble("Tryg m Peter") == 1
    assert cm._detect_ensamble("2, travlt") == 0
    assert cm._detect_ensamble("") == 0
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_column_mapper.py -q`
Expected: FAIL — `_map_type`, `_detect_place`, `_detect_ensamble` do not exist yet (`AttributeError`).

- [ ] **Step 3: Implement the helpers**

Replace the top of `hevy_mcp/column_mapper.py` (imports + existing `_ABS_KEYWORDS` + `_parse_rating`) so the file begins:

```python
import re
from datetime import datetime

# Exercise title keywords that indicate an abs/core movement.
_ABS_KEYWORDS = (
    "crunch", "plank", "sit-up", "sit up", "situp",
    "leg raise", "ab wheel", "ab rollout", "ab roller",
    "hanging knee", "hanging leg", "cable crunch",
    "russian twist", "toe touch", "bicycle crunch",
    "dead bug", "flutter kick", "hollow body",
    "v-up", "v up", "mave", "core",
)

# Danish Hevy workout titles -> the English categories used in the sheet.
_TYPE_MAP = {
    "over": "upper",
    "ben": "legs",
    "træk": "pull",
    "skub": "push",
}

# Known gym names to detect in a note, in priority order. Output is canonical.
_PLACES = ("AC", "Center", "Tryg")

# Danish "with" marker (m / m. / med) followed by a capitalized name.
_ENSAMBLE_RE = re.compile(r"\b(?:med|m\.?)\s+([A-ZÆØÅ][\wÆØÅæøå]+)")


def _map_type(title: str) -> str:
    t = (title or "").strip()
    return _TYPE_MAP.get(t.lower(), t)


def _parse_rating(description):
    if not description:
        return None
    # "4/5" or "8/10"
    m = re.search(r"\b(\d+)\s*/\s*(5|10)\b", description)
    if m:
        return int(m.group(1))
    # Standalone digit 1-10 at start ("4 god træning")
    m = re.match(r"^\s*(\d{1,2})\b", description)
    if m and 1 <= int(m.group(1)) <= 10:
        return int(m.group(1))
    # Standalone digit 1-10 at end ("god træning 4")
    m = re.search(r"\b(\d{1,2})\s*$", description)
    if m and 1 <= int(m.group(1)) <= 10:
        return int(m.group(1))
    return None


def _detect_place(description: str) -> str:
    if not description:
        return ""
    for p in _PLACES:
        if re.search(rf"\b{re.escape(p)}\b", description, re.IGNORECASE):
            return p
    return ""


def _detect_ensamble(description: str) -> int:
    if not description:
        return 0
    return 1 if _ENSAMBLE_RE.search(description) else 0
```

Leave the existing `_has_abs` and `workout_to_row` functions below untouched for now (Task 3 rewrites them).

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_column_mapper.py -q`
Expected: PASS (6 tests).

- [ ] **Step 5: Commit**

```bash
git add hevy_mcp/column_mapper.py tests/test_column_mapper.py
git commit -m "feat: type/rating/place/ensamble helpers in column_mapper"
```

---

### Task 3: column_mapper cardio/abs detection and workout_to_row assembly

**Files:**
- Modify: `hevy_mcp/column_mapper.py`
- Modify: `tests/test_column_mapper.py`

**Interfaces:**
- Consumes: `_map_type`, `_parse_rating`, `_detect_place`, `_detect_ensamble` from Task 2.
- Produces:
  - `_is_cardio_exercise(ex: dict) -> bool` — timed/weightless and not abs.
  - `_cardio_minutes(exercises: list[dict]) -> int`.
  - `_has_abs(exercises: list[dict]) -> bool` (kept).
  - `workout_to_row(workout: dict) -> dict` with keys exactly:
    `Date, Place, Type, Time, Mave, AddCardio, AddCardio2, Ensamble, Rating, Claude, Comments`
    (**no `Total` key** — the writer supplies the formula).

- [ ] **Step 1: Write failing tests**

Append to `tests/test_column_mapper.py`:

```python
from datetime import date
from tests.conftest import strength_ex, cardio_ex, abs_ex


def test_cardio_detection_and_split(make_workout):
    w = make_workout(
        title="Over", description="3",
        exercises=[strength_ex(), cardio_ex(duration_seconds=900)],
        start="2026-07-04T11:26:22+00:00", end="2026-07-04T12:59:52+00:00",
    )  # elapsed ~= 94 min, treadmill 15 min
    row = cm.workout_to_row(w)
    assert row["AddCardio"] == 15
    assert row["Time"] == 94 - 15
    assert row["AddCardio2"] == 0
    assert row["Type"] == "upper"
    assert row["Claude"] == 1
    assert "Total" not in row


def test_cardio_only_session_keeps_time(make_workout):
    w = make_workout(
        title="cardio", description="",
        exercises=[cardio_ex(duration_seconds=1800)],
        start="2026-07-04T11:00:00+00:00", end="2026-07-04T11:30:00+00:00",
    )  # elapsed 30, all cardio
    row = cm.workout_to_row(w)
    assert row["Time"] == 30
    assert row["AddCardio"] == 0


def test_abs_is_mave_not_cardio(make_workout):
    w = make_workout(
        title="push", exercises=[strength_ex(), abs_ex()],
        start="2026-07-04T11:00:00+00:00", end="2026-07-04T12:00:00+00:00",
    )
    row = cm.workout_to_row(w)
    assert row["Mave"] == 1
    assert row["AddCardio"] == 0   # plank must NOT be counted as cardio
    assert row["Time"] == 60


def test_row_shape_and_types(make_workout):
    w = make_workout(title="Skub", description="4 m Silas, AC",
                     exercises=[strength_ex()])
    row = cm.workout_to_row(w)
    assert row["Date"] == date(2026, 7, 4)
    assert row["Type"] == "push"
    assert row["Rating"] == 4
    assert isinstance(row["Rating"], int)
    assert row["Ensamble"] == 1
    assert row["Place"] == "AC"
    assert row["Comments"] == "4 m Silas, AC"
    assert set(row.keys()) == {
        "Date", "Place", "Type", "Time", "Mave", "AddCardio",
        "AddCardio2", "Ensamble", "Rating", "Claude", "Comments"}
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_column_mapper.py -q`
Expected: FAIL — current `workout_to_row` still returns a `Total` key and a `Mave` based on old logic, `_is_cardio_exercise` missing.

- [ ] **Step 3: Replace `_has_abs` and `workout_to_row`**

In `hevy_mcp/column_mapper.py`, replace the existing `_has_abs` function and the entire `workout_to_row` function with:

```python
def _has_abs(exercises: list[dict]) -> bool:
    for ex in exercises:
        title = ex.get("title", "").lower()
        if any(kw in title for kw in _ABS_KEYWORDS):
            return True
    return False


def _is_cardio_exercise(ex: dict) -> bool:
    title = ex.get("title", "").lower()
    if any(kw in title for kw in _ABS_KEYWORDS):
        return False  # abs/core -> Mave, never cardio (even when timed)
    for s in ex.get("sets", []):
        timed = s.get("duration_seconds") is not None or s.get("distance_meters") is not None
        no_load = s.get("reps") is None and s.get("weight_kg") is None
        if timed and no_load:
            return True
    return False


def _cardio_minutes(exercises: list[dict]) -> int:
    total_sec = 0
    for ex in exercises:
        if _is_cardio_exercise(ex):
            for s in ex.get("sets", []):
                total_sec += s.get("duration_seconds") or 0
    return round(total_sec / 60)


def workout_to_row(workout: dict) -> dict:
    start = datetime.fromisoformat(workout["start_time"].replace("Z", "+00:00"))
    end = datetime.fromisoformat(workout["end_time"].replace("Z", "+00:00"))
    elapsed = max(1, round((end - start).total_seconds() / 60))

    exercises = workout.get("exercises", [])
    description = (workout.get("description") or "").strip()

    has_cardio = any(_is_cardio_exercise(ex) for ex in exercises)
    has_strength = any(not _is_cardio_exercise(ex) for ex in exercises)

    if has_cardio and has_strength:
        add_cardio = _cardio_minutes(exercises)
        time = max(1, elapsed - add_cardio)
    else:
        add_cardio = 0
        time = elapsed

    return {
        "Date": start.date(),
        "Place": _detect_place(description),
        "Type": _map_type(workout.get("title")),
        "Time": time,
        "Mave": 1 if _has_abs(exercises) else 0,
        "AddCardio": add_cardio,
        "AddCardio2": 0,
        "Ensamble": _detect_ensamble(description),
        "Rating": _parse_rating(description),
        "Claude": 1,
        "Comments": description,
    }
```

- [ ] **Step 4: Run the full mapper test file**

Run: `python -m pytest tests/test_column_mapper.py -q`
Expected: PASS (10 tests).

- [ ] **Step 5: Commit**

```bash
git add hevy_mcp/column_mapper.py tests/test_column_mapper.py
git commit -m "feat: cardio/abs detection and Total-less row assembly"
```

---

### Task 4: excel_writer — new layout, Total formula, Claude column

**Files:**
- Modify: `hevy_mcp/excel_writer.py`
- Create: `tests/test_excel_writer.py`

**Interfaces:**
- Consumes: row dicts from `column_mapper.workout_to_row` (Task 3).
- Produces:
  - `COLUMN_INDEX: dict[str, int]` (name → 1-based column, per Global Constraints).
  - `total_formula(r: int) -> str`.
  - `append_rows(rows: list[dict]) -> int` — writes value columns + `Total` formula, returns count.
  - `get_last_date()` and `_find_first_empty_row(ws)` unchanged in behavior.

- [ ] **Step 1: Write failing tests**

Create `tests/test_excel_writer.py`. These build a tiny local workbook (no dependency on the real file) and point `EXCEL_PATH` at it:

```python
import os
from datetime import date

import openpyxl
import pytest

from hevy_mcp import excel_writer as ew


@pytest.fixture
def tmp_book(tmp_path, monkeypatch):
    path = tmp_path / "book.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Træning"
    headers = ["Date", "Place", "Type", "Time", "Mave", "AddCardio",
               "AddCardio2", "Total", "Ensamble", "Rating", "Claude", "Comments"]
    ws.append(headers)
    ws.append([date(2026, 5, 1), "AC", "push", 50, 0, 0, 0,
               "=SUM(D2:G2)", 0, 3, 1, "seed"])
    wb.save(path)
    monkeypatch.setenv("EXCEL_PATH", str(path))
    monkeypatch.setenv("EXCEL_SHEET", "Træning")
    return path


def test_total_formula_text():
    assert ew.total_formula(5) == (
        '=IF(ISBLANK(Table4[[#This Row],[Date]]), " ", '
        'SUM(D5:G5)-IF(E5 = 1, 1, 0))')


def test_append_writes_values_claude_and_total_formula(tmp_book):
    row = {
        "Date": date(2026, 7, 4), "Place": "AC", "Type": "upper", "Time": 79,
        "Mave": 1, "AddCardio": 15, "AddCardio2": 0, "Ensamble": 0,
        "Rating": 3, "Claude": 1, "Comments": "3",
    }
    n = ew.append_rows([row])
    assert n == 1
    wb = openpyxl.load_workbook(tmp_book)
    ws = wb["Træning"]
    r = 3  # header + seed + this
    assert ws.cell(row=r, column=ew.COLUMN_INDEX["Type"]).value == "upper"
    assert ws.cell(row=r, column=ew.COLUMN_INDEX["Time"]).value == 79
    assert ws.cell(row=r, column=ew.COLUMN_INDEX["AddCardio"]).value == 15
    assert ws.cell(row=r, column=ew.COLUMN_INDEX["Rating"]).value == 3
    assert ws.cell(row=r, column=ew.COLUMN_INDEX["Claude"]).value == 1
    assert ws.cell(row=r, column=ew.COLUMN_INDEX["Comments"]).value == "3"
    # Total is the formula, not a number:
    assert ws.cell(row=r, column=ew.COLUMN_INDEX["Total"]).value == ew.total_formula(r)


def test_append_blank_rating_is_empty(tmp_book):
    row = {
        "Date": date(2026, 7, 5), "Place": "", "Type": "pull", "Time": 40,
        "Mave": 0, "AddCardio": 0, "AddCardio2": 0, "Ensamble": 0,
        "Rating": None, "Claude": 1, "Comments": "",
    }
    ew.append_rows([row])
    wb = openpyxl.load_workbook(tmp_book)
    ws = wb["Træning"]
    r = 3
    assert ws.cell(row=r, column=ew.COLUMN_INDEX["Rating"]).value is None
    assert ws.cell(row=r, column=ew.COLUMN_INDEX["Place"]).value is None
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_excel_writer.py -q`
Expected: FAIL — `total_formula` and `COLUMN_INDEX` do not exist; `append_rows` still uses the old `COLUMNS` list without `Claude`.

- [ ] **Step 3: Rewrite `hevy_mcp/excel_writer.py`**

Replace the whole file with:

```python
import os
from datetime import date, datetime
from typing import Optional

import openpyxl

# Physical layout AFTER the Claude-column migration (1-based indices).
COLUMN_INDEX = {
    "Date": 1, "Place": 2, "Type": 3, "Time": 4, "Mave": 5,
    "AddCardio": 6, "AddCardio2": 7, "Total": 8, "Ensamble": 9,
    "Rating": 10, "Claude": 11, "Comments": 12,
}

# Value columns the mapper supplies. Total is a formula, not a value.
_VALUE_COLUMNS = [
    "Date", "Place", "Type", "Time", "Mave", "AddCardio", "AddCardio2",
    "Ensamble", "Rating", "Claude", "Comments",
]


def total_formula(r: int) -> str:
    return (f'=IF(ISBLANK(Table4[[#This Row],[Date]]), " ", '
            f'SUM(D{r}:G{r})-IF(E{r} = 1, 1, 0))')


def _open() -> tuple[openpyxl.Workbook, object]:
    path = os.environ["EXCEL_PATH"]
    sheet = os.environ.get("EXCEL_SHEET", "Træning")
    wb = openpyxl.load_workbook(path)
    return wb, wb[sheet]


def get_last_date() -> Optional[date]:
    """Return the date in the last non-empty row of the sheet."""
    _, ws = _open()
    last_date: Optional[date] = None
    for row in ws.iter_rows(min_row=2, max_col=1, values_only=True):
        val = row[0]
        if val is None:
            continue
        if isinstance(val, datetime):
            last_date = val.date()
        elif isinstance(val, date):
            last_date = val
    return last_date


def _find_first_empty_row(ws) -> int:
    """First row index (1-based) with no Date value."""
    for i in range(2, ws.max_row + 2):
        if ws.cell(row=i, column=1).value is None:
            return i
    return ws.max_row + 1


def append_rows(rows: list[dict]) -> int:
    """Append row dicts to the Træning sheet. Returns count written."""
    path = os.environ["EXCEL_PATH"]
    sheet = os.environ.get("EXCEL_SHEET", "Træning")
    wb = openpyxl.load_workbook(path)
    ws = wb[sheet]

    first_empty = _find_first_empty_row(ws)

    for idx, row_data in enumerate(rows):
        r = first_empty + idx
        for col_name in _VALUE_COLUMNS:
            value = row_data.get(col_name)
            ws.cell(row=r, column=COLUMN_INDEX[col_name],
                    value=value if value not in ("", None) else None)
        ws.cell(row=r, column=COLUMN_INDEX["Total"], value=total_formula(r))

    wb.save(path)
    return len(rows)
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_excel_writer.py -q`
Expected: PASS (3 tests).

- [ ] **Step 5: Commit**

```bash
git add hevy_mcp/excel_writer.py tests/test_excel_writer.py
git commit -m "feat: excel_writer writes Claude column and Total formula"
```

---

### Task 5: One-time Excel migration — insert Claude column, heal formulas

**Files:**
- Create: `hevy_mcp/migrate_add_claude_column.py`
- Create: `tests/test_migration.py`

**Interfaces:**
- Consumes: nothing from other tasks.
- Produces: `migrate(path: str, sheet: str = "Træning") -> str` — backs up, inserts the `Claude` column before Comments, repairs the shifted side-calc, expands `Table4`, restores the `Total` formula on every dated row, saves, and returns the backup path.

- [ ] **Step 1: Write the failing test**

Create `tests/test_migration.py`. It builds a small synthetic workbook that mirrors
the real structure (Table `Table4` ending in column K, a `Total` formula column, a
hand-typed static `Total`, two tool rows, and a side-calc in column M). This keeps
the test fast and independent of the private file. (The same migration was also
validated against a copy of the real workbook during planning.)

```python
from datetime import datetime

import openpyxl
from openpyxl.worksheet.table import Table, TableColumn

from hevy_mcp.migrate_add_claude_column import migrate

EXPECTED_COLS = ["Date", "Place", "Type", "Time", "Mave", "AddCardio",
                 "AddCardio2", "Total", "Ensamble", "Rating", "Claude", "Comments"]


def _hf(r):
    return (f'=IF(ISBLANK(Table4[[#This Row],[Date]]), " ", '
            f'SUM(D{r}:G{r})-IF(E{r} = 1, 1, 0))')


def _build(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Træning"
    headers = ["Date", "Place", "Type", "Time", "Mave", "AddCardio",
               "AddCardio2", "Total", "Ensamble", "Rating", "Comments"]
    ws.append(headers)
    # row 2: MANUAL row with a hand-typed static Total -> must be preserved.
    ws.append([datetime(2025, 8, 16), "AC", "legs", 14, 1, 10, 0, 24, 0, 3, None])
    # row 3: manual row that already holds the formula.
    ws.append([datetime(2025, 9, 1), "AC", "push", 50, 0, 0, 0, None, 0, 4, None])
    ws.cell(3, 8).value = _hf(3)
    # rows 4-5: TOOL rows (date >= cutoff) with clobbered static Totals.
    ws.append([datetime(2026, 6, 14), None, "Over", 72, None, None, None, 72, None, 3, "3"])
    ws.append([datetime(2026, 7, 4), None, "Over", 94, 0, None, None, 94, None, 3, "3"])
    cols = [TableColumn(id=i + 1, name=h) for i, h in enumerate(headers)]
    ws.add_table(Table(displayName="Table4", ref="A1:K5", tableColumns=cols))
    # Side-calc in column M (13) with self-references, like the real sheet.
    ws.cell(3, 13).value = 100          # M3 (static)
    ws.cell(4, 13).value = "=M3/2"      # M4
    ws.cell(5, 13).value = "=7*M4"      # M5
    wb.save(path)


def test_migration_adds_claude_and_heals_tool_rows(tmp_path):
    path = str(tmp_path / "book.xlsx")
    _build(path)

    migrate(path)

    wb = openpyxl.load_workbook(path)
    ws = wb["Træning"]

    # Claude inserted before Comments.
    assert ws.cell(row=1, column=11).value == "Claude"
    assert ws.cell(row=1, column=12).value == "Comments"

    # Table4 expanded to column L, Claude between Rating and Comments.
    tbl = ws.tables["Table4"]
    assert tbl.ref == "A1:L5"
    assert [c.name for c in tbl.tableColumns] == EXPECTED_COLS

    # Tool rows (>= 2026-06-14) healed to the Total formula.
    assert str(ws.cell(row=4, column=8).value).startswith("=IF(ISBLANK")
    assert str(ws.cell(row=5, column=8).value).startswith("=IF(ISBLANK")

    # MANUAL row left exactly as the user made it (static Total 24 preserved).
    assert ws.cell(row=2, column=8).value == 24

    # Side-calc moved M->N and its references were repaired.
    assert ws.cell(row=4, column=14).value == "=N3/2"
    assert ws.cell(row=5, column=14).value == "=7*N4"
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `python -m pytest tests/test_migration.py -q`
Expected: FAIL — `hevy_mcp.migrate_add_claude_column` does not exist (`ModuleNotFoundError`).

- [ ] **Step 3: Implement the migration**

Create `hevy_mcp/migrate_add_claude_column.py`:

```python
import os
import re
import shutil
from datetime import date, datetime

import openpyxl
from openpyxl.worksheet.table import TableColumn

# Only rows from this date onward were written by the (buggy) tool. Manual rows
# before it — including any hand-typed static Total — must never be touched.
_TOOL_CUTOFF = date(2026, 6, 14)

_TOTAL = ('=IF(ISBLANK(Table4[[#This Row],[Date]]), " ", '
          'SUM(D{r}:G{r})-IF(E{r} = 1, 1, 0))')


def _backup(path: str) -> str:
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    bak = f"{path}.{ts}.bak"
    shutil.copy2(path, bak)
    return bak


def _last_data_row(ws) -> int:
    """Last row with a Date. Bounded scan (the Table spans the whole sheet, so
    ws.max_row is ~1M and must not be iterated)."""
    last, gap, r = 1, 0, 2
    while gap < 50:
        if ws.cell(row=r, column=1).value is not None:
            last, gap = r, 0
        else:
            gap += 1
        r += 1
    return last


def _as_date(v):
    return v.date() if isinstance(v, datetime) else v


def migrate(path: str, sheet: str = "Træning") -> str:
    """Insert the Claude column before Comments and heal the tool rows' formulas.

    Returns the path of the timestamped backup made before writing.
    """
    bak = _backup(path)
    wb = openpyxl.load_workbook(path)
    ws = wb[sheet]
    last = _last_data_row(ws)

    # 1. Insert a blank column at position 11 (before Comments). This shifts
    #    Comments 11->12, the empty buffer 12->13, and the side-calc 13->14.
    ws.insert_cols(11)
    ws.cell(row=1, column=11, value="Claude")

    # 2. Repair the side-calc: its cells moved from column M to column N, but
    #    openpyxl does not rewrite formula references. Point M-refs at N.
    for r in range(2, last + 1):
        cell = ws.cell(row=r, column=14)
        v = cell.value
        if isinstance(v, str) and v.startswith("=") and re.search(r"\bM\d+", v):
            cell.value = re.sub(r"\bM(\d+)", r"N\1", v)

    # 3. Expand Table4 (A1:K.... -> A1:L....) and add a Claude table column
    #    positioned between Rating and Comments.
    tbl = ws.tables["Table4"]
    tbl.ref = re.sub(r":K(\d+)", r":L\1", tbl.ref)
    new_id = max(tc.id for tc in tbl.tableColumns) + 1
    tbl.tableColumns.insert(len(tbl.tableColumns) - 1,
                            TableColumn(id=new_id, name="Claude"))

    # 4. Restore the Total (H) formula ONLY on the tool's rows (date >= cutoff),
    #    which the old tool overwrote with static numbers. Manual rows — even one
    #    with a hand-typed static Total — are left untouched.
    for r in range(2, last + 1):
        d = _as_date(ws.cell(row=r, column=1).value)
        if isinstance(d, date) and d >= _TOOL_CUTOFF:
            ws.cell(row=r, column=8).value = _TOTAL.format(r=r)

    wb.save(path)
    return bak


if __name__ == "__main__":
    target = os.environ["EXCEL_PATH"]
    backup = migrate(target)
    print(f"Migrated {target} (backup: {backup})")
```

> **Note:** on the real workbook, `insert_cols` + `save` take ~2–3 minutes because
> openpyxl processes the Table's full 1,048,576-row span. This is expected for the
> one-time run; the synthetic unit test above is instant.

- [ ] **Step 4: Run the test to verify it passes**

Run: `python -m pytest tests/test_migration.py -q`
Expected: PASS (1 test, runs instantly on the synthetic workbook).

- [ ] **Step 5: Commit**

```bash
git add hevy_mcp/migrate_add_claude_column.py tests/test_migration.py
git commit -m "feat: one-time migration adds Claude column and heals formulas"
```

---

### Task 6: `fix_synced_rows` MCP tool — correct the 12 buggy rows

**Files:**
- Modify: `hevy_mcp/excel_writer.py`
- Modify: `hevy_mcp/server.py`
- Modify: `tests/test_excel_writer.py`

**Interfaces:**
- Consumes: `column_mapper.workout_to_row` (Task 3), `hevy_client.fetch_workouts_since` (existing), `COLUMN_INDEX`/`total_formula` (Task 4).
- Produces:
  - `excel_writer.find_row_by_date(ws, d: date) -> int | None`.
  - `excel_writer.overwrite_rows_by_date(rows_by_date: dict[date, dict], dry_run: bool) -> list[tuple[int, date, dict, dict]]` — returns `(rownum, date, before, after)` for each matched row; writes only when `dry_run` is False.
  - `server.fix_synced_rows(confirm: bool = False) -> str` — preview when `confirm=False`, apply when `confirm=True`.

- [ ] **Step 1: Write failing tests for the writer helpers**

Append to `tests/test_excel_writer.py`:

```python
def test_overwrite_rows_by_date_preview_then_write(tmp_book):
    # Seed a "buggy" row: static Total, string Rating, Danish Type.
    wb = openpyxl.load_workbook(tmp_book)
    ws = wb["Træning"]
    ws.append([date(2026, 6, 28), None, "Skub", 67, "x", None, None,
               67, None, "4", None, "4 m Silas"])
    wb.save(tmp_book)

    after = {
        "Date": date(2026, 6, 28), "Place": "", "Type": "push", "Time": 67,
        "Mave": 0, "AddCardio": 0, "AddCardio2": 0, "Ensamble": 1,
        "Rating": 4, "Claude": 1, "Comments": "4 m Silas",
    }

    preview = ew.overwrite_rows_by_date({date(2026, 6, 28): after}, dry_run=True)
    assert len(preview) == 1
    rownum, d, before, aft = preview[0]
    assert d == date(2026, 6, 28)
    assert before["Type"] == "Skub"          # unchanged on dry run
    assert aft["Type"] == "push"

    # Nothing written yet.
    wb = openpyxl.load_workbook(tmp_book)
    assert wb["Træning"].cell(row=rownum, column=ew.COLUMN_INDEX["Type"]).value == "Skub"

    ew.overwrite_rows_by_date({date(2026, 6, 28): after}, dry_run=False)
    wb = openpyxl.load_workbook(tmp_book)
    ws = wb["Træning"]
    assert ws.cell(row=rownum, column=ew.COLUMN_INDEX["Type"]).value == "push"
    assert ws.cell(row=rownum, column=ew.COLUMN_INDEX["Rating"]).value == 4
    assert ws.cell(row=rownum, column=ew.COLUMN_INDEX["Claude"]).value == 1
    assert ws.cell(row=rownum, column=ew.COLUMN_INDEX["Total"]).value == ew.total_formula(rownum)
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `python -m pytest tests/test_excel_writer.py -q`
Expected: FAIL — `overwrite_rows_by_date` / `find_row_by_date` do not exist.

- [ ] **Step 3: Add the writer helpers**

Append to `hevy_mcp/excel_writer.py`:

```python
def find_row_by_date(ws, d: date) -> Optional[int]:
    for i in range(2, ws.max_row + 1):
        val = ws.cell(row=i, column=1).value
        if isinstance(val, datetime):
            val = val.date()
        if val == d:
            return i
    return None


def _read_row(ws, r: int) -> dict:
    out = {}
    for name, col in COLUMN_INDEX.items():
        out[name] = ws.cell(row=r, column=col).value
    return out


def overwrite_rows_by_date(rows_by_date: dict, dry_run: bool) -> list:
    """Overwrite existing rows matched by Date. Returns (row, date, before, after)."""
    path = os.environ["EXCEL_PATH"]
    sheet = os.environ.get("EXCEL_SHEET", "Træning")
    wb = openpyxl.load_workbook(path)
    ws = wb[sheet]

    changes = []
    for d, after in sorted(rows_by_date.items()):
        r = find_row_by_date(ws, d)
        if r is None:
            continue
        before = _read_row(ws, r)
        changes.append((r, d, before, after))
        if not dry_run:
            for col_name in _VALUE_COLUMNS:
                value = after.get(col_name)
                ws.cell(row=r, column=COLUMN_INDEX[col_name],
                        value=value if value not in ("", None) else None)
            ws.cell(row=r, column=COLUMN_INDEX["Total"], value=total_formula(r))

    if not dry_run:
        wb.save(path)
    return changes
```

- [ ] **Step 4: Run the writer test to verify it passes**

Run: `python -m pytest tests/test_excel_writer.py -q`
Expected: PASS (4 tests).

- [ ] **Step 5: Add the `fix_synced_rows` MCP tool**

In `hevy_mcp/server.py`, update the imports at the top:

```python
from datetime import datetime, timedelta
```

Then add this tool after `add_single_workout`:

```python
@mcp.tool()
async def fix_synced_rows(confirm: bool = False) -> str:
    """Re-map and correct the rows the old buggy tool wrote (2026-06-14 onward).

    Shows a before/after preview when confirm is False; writes the corrections
    (including restoring the Total formula and setting Claude=1) when confirm is True.
    Requires the Claude column to already exist (run migrate_add_claude_column first).

    Args:
        confirm: Set True to apply the corrections. Default False = preview only.
    """
    cutoff = datetime(2026, 6, 14)
    workouts = await hevy_client.fetch_workouts_since(cutoff - timedelta(days=1))
    rows_by_date = {}
    for w in workouts:
        row = column_mapper.workout_to_row(w)
        rows_by_date[row["Date"]] = row

    changes = excel_writer.overwrite_rows_by_date(rows_by_date, dry_run=not confirm)
    if not changes:
        return "No matching rows found to correct."

    lines = []
    for rownum, d, before, after in changes:
        lines.append(
            f"row {rownum} {d}: "
            f"Type {before.get('Type')!r}->{after['Type']!r}  "
            f"Rating {before.get('Rating')!r}->{after['Rating']!r}  "
            f"AddCardio {before.get('AddCardio')!r}->{after['AddCardio']!r}  "
            f"Claude {before.get('Claude')!r}->{after['Claude']!r}"
        )
    header = ("APPLIED" if confirm else "PREVIEW (pass confirm=true to write)")
    return f"{header} — {len(changes)} row(s):\n" + "\n".join(lines)
```

- [ ] **Step 6: Verify the server imports and tools load**

Run: `python -c "import hevy_mcp.server as s; print([t for t in ['preview_workouts','sync_workouts','add_single_workout','fix_synced_rows']])"`
Expected: prints the four tool names with no import error.

- [ ] **Step 7: Run the full test suite**

Run: `python -m pytest -q`
Expected: PASS (all tests; migration test may SKIP if the workbook is absent).

- [ ] **Step 8: Commit**

```bash
git add hevy_mcp/excel_writer.py hevy_mcp/server.py tests/test_excel_writer.py
git commit -m "feat: fix_synced_rows tool corrects the 12 legacy rows"
```

---

## Manual run-book (after all tasks pass, executed by the user with approval)

These steps touch the real `Træning.xlsx` and Hevy data; run them only when the
suite is green and the user approves each write.

1. **Close Excel** so the file is not locked.
2. **Migrate once:** `python -m hevy_mcp.migrate_add_claude_column`
   (reads `EXCEL_PATH`; writes a `.bak` backup, inserts the `Claude` column,
   heals the `Total` formula on the tool rows only). Takes ~2–3 minutes on the
   real file. Open the file and eyeball: `Claude` header before Comments, `Total`
   values recomputing, side-calc still correct, and your manual rows unchanged.
3. **Preview the row corrections:** call the `fix_synced_rows` MCP tool with
   `confirm=false`; review the before/after list.
4. **Apply:** call `fix_synced_rows` with `confirm=true`.
5. **Future syncs:** `sync_workouts` now writes `Claude=1`, the `Total` formula,
   translated `Type`, numeric `Rating`, and cardio in `AddCardio` automatically.

---

## Self-Review

**Spec coverage:**
- Type translation → Task 2 (`_map_type`) + Task 3 assembly. ✓
- Numeric Rating → Task 2 (`_parse_rating` returns `int`), asserted in Task 3. ✓
- Cardio → AddCardio, abs → Mave, cardio-only day → Task 3. ✓
- Place / Ensamble detection → Task 2. ✓
- AddCardio2 = 0, Comments verbatim → Task 3. ✓
- Total stays a formula → Task 4 (`append`) + Task 5 (heal) + Task 6 (overwrite). ✓
- Claude column (1/tool, 0/user, empty history) → Task 4 (write 1), Task 5 (insert, leaves history empty), Task 6 (set 1). ✓
- Excel table migration + side-calc repair + backup → Task 5. ✓
- Manual rows never touched (heal gated to date ≥ 2026-06-14; row 467's hand-typed
  static Total preserved) → Task 5 (`_TOOL_CUTOFF`), asserted in the migration test. ✓
- Correct the 12 rows with preview/approve → Task 6. ✓

**Placeholder scan:** No TBD/TODO; every code step shows full code. ✓

**Type consistency:** `COLUMN_INDEX`, `total_formula`, `overwrite_rows_by_date`,
`find_row_by_date`, `workout_to_row` (no `Total` key, has `Claude`) are used with
identical names/signatures across Tasks 3–6. ✓
