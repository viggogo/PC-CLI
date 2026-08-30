# sengetoej — design

**Date:** 2026-08-30
**Status:** Approved
**Project:** `projects/sengetoej/`

## Purpose

A CLI to record and look up bed linen changes — when the sheets were last
changed, how long ago that was, and to log a new change — without opening
`Træning.xlsx` and typing a date into the `Sengetøj` tab by hand.

## Scope

In scope:

- Report the last change and how many days ago it was.
- Report the last N changes with the interval between each.
- Append a new change, dated today or on a given date, marked as tool-written.
- Add the `cli` column to `Table2`, once, via a separate migration.
- Show usage.
- Install itself onto the user's PATH.

Explicitly out of scope (decided, not deferred):

- **No back-dating.** Appending only. Inserting a row in date order would shift
  the free-text notes in column D — which sit *outside* `Table2` — out of
  alignment with their dates, silently corrupting them. The user was asked and
  chose append-only.
- **No note column.** Column D is neither read nor written. The user was asked
  and chose the plain date + interval list.
- **No summary statistics.** No average interval, no streaks. Same decision.
- **No calendar view.** Changes are ~20 days apart, so a `hevy --weeks`-style
  grid would be nearly all empty cells.
- **No editing or deleting** existing rows.
- **No MCP server.** hevy has one; this tool does not need one.

## Workbook facts

Verified against the live workbook during design
(`C:\Users\viggo\OneDrive\Privat\Fitness\Træning.xlsx`):

| Fact | Value |
|---|---|
| Sheet name | `Sengetøj` (with `ø`) |
| Table name | `Table2` |
| Table ref | `A1:B1048576` — the whole column, not just the data |
| Autofilter ref | `A1:B1048576` — tracks the table ref |
| Table style | `TableStyleLight9`, banded rows |
| Column A | `Date`, a real `datetime`, number format `mm-dd-yy` |
| Column B | `Diff`, a formula (see below) |
| Column C | **Entirely empty**, rows 1–80 verified. Width 10.7, not hidden |
| Column D | Free text, **outside** `Table2` |
| Data rows | 2 through 73 as of this design (72 entries) |
| First entry | 2023-10-09 |
| Last entry | 2026-08-30 |
| Order | Strictly ascending by date |

The `Diff` formula, present in every row of column B:

```
=IF(ISBLANK(Table2[[#This Row],[Date]]), " ", Table2[[#This Row],[Date]]-A<prev>)
```

**It is pre-filled all the way down.** Verified present at rows 74, 76, 100,
200, 1000 and 5000 — every one of them empty of data. Two consequences drive
the whole design:

1. `--new` never writes column B. The gap count computes itself the moment a
   date lands in column A.
2. `ws.max_row` is 1048576 and therefore useless for finding the end of the
   data. So is `Table2`'s ref. The end of the data can only be found by
   scanning column A.

Column D holds a label in D1, an array formula in D2
(`=DAY(ABS(TODAY()-INDEX(A:A,COUNTA(A:A))))`), a legend in D4, and scattered
per-row remarks further down (`NZ og CA`, `Juleferie`, `Obs: Ferie`, `Amst`…).
Because D is outside the table, its rows are aligned to the dates by position
alone — nothing but append-only writing keeps that alignment true.

## Stack

Python ≥ 3.11 in a virtualenv inside the tool's own folder, mirroring
`projects/hevy`. Dependencies: `openpyxl`, `python-dotenv`. `pytest` for tests.

Chosen over PowerShell because PowerShell has no clean xlsx library — it would
need Excel COM automation (requires Excel installed, fails when the workbook is
open, slow) or a gallery module. openpyxl is already proven against this exact
workbook: `hevy sync` round-trips it regularly without disturbing `Table4` or
its formulas.

Chosen over folding into `hevy` because CLAUDE.md requires tools to be
independent and self-contained, and bed linen has nothing to do with Hevy
beyond sharing a file.

## Layout

```
projects/sengetoej/
├── README.md            purpose, stack, install, usage
├── pyproject.toml       openpyxl + python-dotenv; pytest as a dev extra
├── install.ps1          .venv, editable install, .env, PATH → bin\
├── .env.example
├── .gitignore           .venv, .env, __pycache__, *.egg-info
├── bin/
│   └── sengetoej.cmd    the only PATH-exposed file
├── sengetoej/
│   ├── __init__.py
│   ├── __main__.py
│   ├── cli.py           argparse, confirmation prompt, Danish output
│   ├── env.py           EXCEL_PATH / SENGETOEJ_SHEET
│   ├── migrate.py       one-time: add the cli column to Table2
│   └── sheet.py         read entries · append a date
└── tests/
    ├── __init__.py
    ├── conftest.py      builds a throwaway workbook in tmp_path
    ├── test_migrate.py
    ├── test_sheet.py
    └── test_cli.py
```

The folder is ASCII `sengetoej`, per CLAUDE.md's lowercase-kebab-case rule. Only
the *sheet* name keeps its `ø`, as a configuration default.

No code is imported from `hevy`. The overlap is roughly thirty lines of `.env`
loading; duplicating it is cheaper than coupling two independent tools through a
shared workbook helper.

Only `bin\sengetoej.cmd` goes on the PATH, never the project folder — PowerShell
resolves `.ps1` from PATH ahead of PATHEXT, so exposing the folder would make
`install.ps1` a global command. Same reasoning as hevy.

## Configuration

Read from `.env` in the tool's folder, or from real environment variables, which
win. The `.env` is located from the package's own path, not the working
directory, so the command works from anywhere.

| Variable | Default | Meaning |
|---|---|---|
| `EXCEL_PATH` | `C:\Users\viggo\OneDrive\Privat\Fitness\Træning.xlsx` | the workbook |
| `SENGETOEJ_SHEET` | `Sengetøj` | the tab |

**Deviation from hevy:** both have working built-in defaults, so the tool runs
correctly the moment it is installed. hevy requires a `.env` because it needs an
API key; this tool has no secret, so requiring one would be friction with no
payoff. `.env.example` still ships, and `.env` is still git-ignored.

## The `cli` column

`Table2` gains a third column, `cli`, recording whether a row was written by this
tool or typed by hand. It mirrors hevy's `Claude` column on `Table4`, including
its convention: **`1` on tool-written rows, blank on hand-typed ones** — not `0`.
Blank is what the 72 existing rows already say, so the migration writes no values
at all, and a blank cell keeps meaning "not mine" forever without backfilling.

The header is lowercase `cli`, as the user wrote it, though its siblings are
`Date` and `Diff`. Header text is not load-bearing anywhere except the guard
below, which compares case-insensitively.

### Migration

Adding the column is **not** just writing `C1`. Excel will report the workbook as
corrupt if the pieces disagree, so all four must change together:

1. `C1` ← `cli`
2. `Table2.ref` → `A1:C1048576`
3. `Table2.autoFilter.ref` → `A1:C1048576`
4. A `TableColumn(id=3, name="cli")` appended to `Table2.tableColumns`

This happens once, in `migrate.py`, run deliberately — never from `--new`. The
script:

- **Refuses to run twice.** If `C1` already reads `cli`, it reports that and
  exits 0, changing nothing.
- **Refuses to run on unexpected data.** If `C1` is non-empty and is not `cli`,
  it stops rather than overwrite whatever is there.
- **Writes a timestamped backup** beside the workbook before saving, since this
  is the one operation in the tool that rewrites table structure rather than
  cell values.
- **Shows what it will change and asks `[y/N]`** before touching anything.

`TableStyleLight9`'s banded rows will extend over column C once the ref grows.
That is expected, not a defect.

Because openpyxl's table-schema handling is the least-exercised part of this
design, the migration is **verified by opening the workbook in Excel afterwards**
and confirming the table reads as three columns with no repair prompt. That check
is a step in the implementation plan, not something the tests can cover.

### Guard

`--new` never repairs the schema. Before writing it checks that `C1` reads `cli`,
and if not:

```
Fejl: kolonnen "cli" mangler i Table2.
Kør migreringen først:
  .\.venv\Scripts\python.exe -m sengetoej.migrate
```

The migration is deliberately *not* exposed as a `sengetoej` subcommand or flag.
Reaching for the venv's Python is friction on purpose: it is a once-ever
operation on the structure of a live workbook, and it should not sit one typo
away from the command used daily.

Exit 1, nothing written. `--last` reads only column A, so it is unaffected and
keeps working whether or not the migration has run.

## Commands

```
sengetoej                    same as --last
sengetoej --last             sidste skift + dage siden
sengetoej --last N           de sidste N skift
sengetoej --new              registrer et skift i dag
sengetoej --new dd/mm/yyyy   registrer et skift på den dato
  -y, --yes                  spring bekræftelsen over
sengetoej --help             vis brug
```

`--last` and `--new` are mutually exclusive. Bare `sengetoej` is `--last`.
`-y` alongside `--last` is a usage error rather than a silent no-op — it can only
mean the user thought they were writing.

Output is in Danish. The README is in English, as hevy's is.

Exit codes, matching hevy: `0` success, `1` runtime failure, `2` usage error.

### `--last`

```
> sengetoej --last
Sidste skift: 30/08/2026 (i dag)

> sengetoej --last 5
  05/06/2026    14 dage
  29/06/2026    24 dage
  31/07/2026    32 dage
  22/08/2026    22 dage
  30/08/2026     8 dage
```

Rows are printed oldest-first, so the most recent change is the line nearest the
prompt. The interval on each row is the gap from the row above it — that is, from
the previous change. The oldest row shown prints its true interval, computed from
the entry before it even when that entry falls outside the window; the very first
entry in the sheet has no predecessor and prints `—`.

Days-ago wording for bare `--last`: `i dag` for 0, `i går` for 1, `N dage siden`
otherwise. A negative value is impossible given ascending order, but if the sheet
somehow holds a future last entry, it prints `om N dage`.

`N` must be a positive integer; `0`, a negative, or a non-integer is a usage
error (exit 2). An `N` larger than the number of entries shows every entry
without complaint. An entirely empty sheet prints `Ingen skift registreret.` and
exits 0.

### `--new`

Writes **two cells** in the first empty row: the date into column A, and `1` into
column C. Nothing in the output changes — the `cli` flag is for reading the sheet
in Excel, and `--last` stays the plain date-and-interval list.

```
> sengetoej --new 02/09/2026
Tilføj 02/09/2026? (3 dage siden sidste skift) [y/N] y
Tilføjet i række 74.

> sengetoej --new
Tilføj 30/08/2026? (i dag, 8 dage siden sidste skift) [y/N] n
Afbrudt.
```

Declining exits 0 and writes nothing. `-y` / `--yes` skips the prompt.

The date argument is `dd/mm/yyyy`. Leading zeros are optional, so `2/9/2026`
parses. No other format is accepted — `02/09/2026` and `2026-09-02` in the same
tool would invite exactly the day/month ambiguity this format choice avoids.

Rejected before any write:

| Condition | Exit |
|---|---|
| Malformed date | 2 |
| Date already present in the sheet | 2 |
| Date earlier than the last entry | 2 |
| Date in the future | 2 |
| `cli` column missing (migration not run) | 1 |

Checked in that order, so a date that is both a duplicate and the last entry is
reported as a duplicate, not as out of order.

The ordering rule exists because `Diff` subtracts the row above; an out-of-order
date would produce a negative gap and a sheet that no longer means what it says.
The message names the offending date and the last entry, and points at Excel:

```
Fejl: 22/08/2026 ligger før sidste række (30/08/2026).
Arket skal være i stigende rækkefølge — tilføj den i Excel.
```

The future-date rule is a typo guard, not a workbook constraint. It was proposed
during design and accepted.

## How it works

### Reading

`--last` never reads the `Diff` column. It reads the dates from column A and
subtracts them in Python.

Reading the formula's cached value would require `data_only=True`, which returns
`None` for any cell Excel has not recalculated and saved — a new row written by
this tool would read back as `None` until the workbook is next opened. Computing
the gaps ourselves also means `--last` works **while Excel has the file open**,
the same way `hevy --weeks` does.

`sheet.py` finds the end of the data by walking column A lazily from row 2 and
stopping after a run of consecutive blanks — 100 of them, comfortably more than
any gap that could appear in a hand-kept column, and cheap because the walk is
lazy and stops long before the millionth row. Neither `ws.max_row` nor
`Table2`'s ref can be used, for the reasons under **Workbook facts**.

### Writing

`--new` opens the workbook, writes the date into column A of the first empty row
and `1` into column C, copies the number format from the previous date cell so
the new row renders like its neighbours, and saves.

It writes nothing else:

- Column B already holds the `Diff` formula in the target row and every row
  below it. Writing it would be redundant, and writing it *wrong* would be worse.
- After the migration, `Table2`'s ref spans `A1:C1048576`, so the table range
  still needs no extension per row.
- Column D is never touched, and append-only writing keeps its rows aligned to
  their dates.
- The table schema is never altered here. That is `migrate.py`'s job alone.

If Excel holds the file, `wb.save()` raises `PermissionError`. That is caught and
reported as `Luk Excel og prøv igen.`, exit 1 — the workbook is not modified.
This mirrors hevy, which reports the same condition rather than failing halfway.

## Testing

pytest, entirely offline. **No test opens the real spreadsheet.**

`conftest.py` builds a miniature workbook in `tmp_path` shaped like the real one:
a `Table2` with an over-wide ref, `Diff` formulas pre-filled past the end of the
data, trailing blank rows, an empty column C, and a column D outside the table.
The traps the design exists to handle are therefore present in the fixture. Two
variants are needed — pre-migration (two columns) and post-migration (three) —
so the migration and the guard can both be tested.

Cases to cover:

- **sheet:** end-of-data scan past pre-filled formulas and trailing blanks; scan
  on an empty sheet; entries read in order; append lands in the first empty row;
  append writes `1` to column C; append leaves column B's formula intact; append
  leaves column D untouched; append copies the number format; duplicate
  detected; out-of-order detected.
- **migrate:** header, table ref, autofilter ref and `tableColumns` all updated
  together; existing rows left blank in column C, not backfilled; re-running is
  a no-op that exits 0; a non-empty, non-`cli` `C1` aborts; the backup file is
  written; declining the prompt changes nothing. The saved workbook is reopened
  and re-parsed within the test to prove openpyxl can still read the table it
  just wrote — the nearest a test can get to Excel's own validation.
- **cli:** bare `--last` wording for `i dag`, `i går` and `N dage siden`;
  `--last N` list formatting; the oldest shown row using the entry before the
  window; `—` for the sheet's first entry; `N` greater than the entry count;
  `N` of 0 or a non-integer as exit 2; empty sheet; `--new` accepted at the
  prompt; `--new` declined writes nothing and exits 0; `-y` skips the prompt;
  bad date format; duplicate date; earlier date; future date; missing `cli`
  column as exit 1 pointing at the migration; `PermissionError` surfacing as
  exit 1 with the Excel message.

## Open questions

None.
