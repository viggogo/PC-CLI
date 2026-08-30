# sengetoej

Read and record bed linen changes on the `Sengetøj` tab of `Træning.xlsx` —
from the terminal, instead of typing each date in by hand.

## Stack

Python >= 3.11 (installed here against 3.13), in a virtualenv inside this
folder. Dependencies: `openpyxl`, `python-dotenv`. `pytest` for tests.

## Install

```powershell
.\install.ps1
```

Creates `.venv`, installs this package into it (editable, so edits to
`sengetoej\` take effect immediately), copies `.env.example` -> `.env`, and
adds this folder's `bin` subfolder to your user `PATH`. Safe to re-run; it
never overwrites an existing `.env`.

Neither setting in `.env` is required — both default to working values in
code (see [Configuration](#configuration)) — so the tool runs right after
install with no further setup. Then open a new terminal, and run the
one-time migration described under [Migration](#migration) before the first
`--new`.

Only `bin\sengetoej.cmd` is exposed, never the project folder itself —
PowerShell resolves `.ps1` from `PATH` ahead of `PATHEXT`, so `install.ps1`
would otherwise become a global command.

## Usage

```
sengetoej                    Same as --last.
sengetoej --last             Last change, and how long ago.
sengetoej --last N           The last N changes.
sengetoej --new              Record a change, dated today.
sengetoej --new dd/mm/yyyy   Record a change on that date.
sengetoej --new ... -y       Skip the confirmation prompt.
sengetoej --help             Show usage.
```

`--last` reads the sheet and never writes; `--new` is the only command that
writes, and appends a single row.

**`--last` works with Excel open. `--new` does not** — `--new` has to save,
and Excel holds an exclusive lock on the file while it's open. Close Excel
before `--new`; the tool reports the lock rather than failing halfway
through, so nothing is left half-written.

Worked examples:

```
> sengetoej --last
Sidste skift: 22/08/2026 (8 dage siden)

> sengetoej --last 3
  15/08/2026       —
  20/08/2026       5 dage
  22/08/2026       2 dage

> sengetoej --new
Åbner regnearket ...
Tilføj 30/08/2026? (8 dage siden sidste skift) [y/N] y
Tilføjet i række 214.

> sengetoej --new 15/08/2026
Åbner regnearket ...
Fejl: 15/08/2026 findes allerede i arket.
```

`--last` is effectively instant (~0.01 s). `--new` takes roughly 15 seconds
— about 6.5 s to load the workbook and 8.5 s to save it — because the sheet
stores over a million cells (the `Diff` formula is materialised in every
row, all the way to the sheet's last row). `--new` prints
`Åbner regnearket ...` before it starts, so that wait doesn't read as a
hang.

Exit codes: `0` success, `1` runtime failure (missing file, missing sheet,
workbook locked), `2` usage error (bad date, future date, duplicate date,
out-of-order date, `-y` without `--new`).

## Configuration

Read from `.env` in this folder, or from real environment variables (which
win):

| Variable | Value |
|---|---|
| `EXCEL_PATH` | `C:\Users\viggo\OneDrive\Privat\Fitness\Træning.xlsx` (optional — this is the default) |
| `SENGETOEJ_SHEET` | `Sengetøj` (optional — this is the default) |

Unlike `hevy`, this tool holds no secret, so a `.env` is never required —
both settings already work out of the box.

`.env` is git-ignored. It is located from the package's own path, not the
current directory, so `sengetoej` works from anywhere. A `.env` in the
directory you happen to be standing in is also read, but this folder's copy
wins.

## How it works

The sheet stores one row per change: column A holds the date, column B a
`Diff` formula (days since the previous row), and column C the `cli` flag.
`Table2`'s ref spans the entire column — `A1:C1048576` — because the `Diff`
formula is materialised in every one of those rows, so neither the ref nor
`ws.max_row` says where the real data actually ends. Only a scan of column A
does.

That also makes a plain `load_workbook()` cost ~6.6 s just to open. Reading
(`sheet.read_dates`) goes through `read_only` mode over column A alone and
breaks out as soon as it sees a run of 100 consecutive blank cells — that
early break is not an optimisation, it's the difference between a usable
tool and an unusable one; a `read_only` walk that does *not* break early
still costs ~4 s just to reach the end of the data.

Writing (`sheet.append_entry`) is a two-cell append: the date into column A,
and `1` into column C. Column B's `Diff` formula is already present in every
row below the last real one — it's a formula, not a value, so it doesn't
need writing — and `Table2`'s ref already spans the full column, so it's
never extended either. `open_for_write` cannot use `read_only` mode, since
that mode forbids `ws.cell()` and can't save; this is also the ~6.5 s /
~8.5 s cost `--new` pays that `--last` doesn't.

## Migration

A one-time, standalone script, run once by hand — deliberately **not**
exposed as a `sengetoej` subcommand, so it can't be triggered by accident:

```powershell
.\.venv\Scripts\python.exe -m sengetoej.migrate
```

It does two jobs in one save:

1. **Moves the free-text comment column D -> E.** The sheet used a blank
   column D as a visual spacer between `Table2` and a column of hand-typed
   comments. Adding `cli` in C would otherwise push into that spacer, so the
   comments move one column right first, restoring the gap.
2. **Adds the `cli` column to `Table2`** in column C — header, table ref,
   autofilter and `tableColumns` all have to agree, or Excel refuses to open
   the file and calls it corrupt.

It takes its own backup before writing (`<name>.<timestamp>.bak.xlsx` next
to the workbook), prints the exact steps it's about to take, and asks for
confirmation before touching anything. Running it again on an already
migrated sheet reports there's nothing to do.

## The `cli` column

Column C is `1` on any row `sengetoej --new` wrote, and blank on a row typed
by hand in Excel — the same convention as `hevy`'s `Claude` column, so a
glance at the sheet shows which rows came from the tool.

## Tests

```powershell
.\.venv\Scripts\python.exe -m pytest -q
```

99 tests, entirely offline — every test builds a small throwaway workbook in
a temp folder. **Nothing here ever touches the real spreadsheet.**
