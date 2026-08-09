# hevy

Pull workouts from the Hevy app and append them to `Træning.xlsx` — from the
terminal, instead of typing each session in by hand.

The same three actions are also exposed as MCP tools, so Claude can run them
directly. See [MCP server](#mcp-server).

## Stack

Python ≥ 3.11 (installed here against 3.13), in a virtualenv inside this folder.
Dependencies: `httpx`, `openpyxl`, `python-dotenv`, `mcp[cli]`. `pytest` for tests.

## Install

```powershell
.\install.ps1
```

Creates `.venv`, installs this package into it (editable, so edits to `hevy_mcp\`
take effect immediately), copies `.env.example` → `.env`, and adds this folder's
`bin` subfolder to your user `PATH`. Safe to re-run; it never overwrites an
existing `.env`.

Then put your Hevy API key in `.env` (Hevy app → Settings → Developer / API) and
open a new terminal.

Only `bin\hevy.cmd` is exposed, never the project folder itself — PowerShell
resolves `.ps1` from `PATH` ahead of `PATHEXT`, so `install.ps1` would otherwise
become a global command.

## Usage

```
hevy preview          Show new workouts as a table. Never writes, never asks.
hevy sync             Show the same table, then ask before adding them.
hevy sync -y          Add without asking.
hevy fix              Show before → after for tool-written rows, then ask.
hevy fix -y           Apply the corrections without asking.
hevy --help           Show usage. Also: hevy sync --help, etc.
```

Options:

```
--since YYYY-MM-DD    (preview, sync) Override the start date.
                      Default = everything after your last Excel row.
-y, --yes             (sync, fix) Skip the confirmation prompt.
```

**Close Excel before any command that writes** (`sync`, `fix`). Excel holds a lock
on the file; the tool reports this rather than failing halfway.

Typical run:

```
> hevy preview
2 new workout(s):
  2026-08-02  push         62 min  AC:0  Mave:1  Ens:0  R:4  AC      4 god træning
  2026-08-05  legs         71 min  AC:12 Mave:0  Ens:1  R:5  Center  5 med Anders

> hevy sync
2 new workout(s):
  2026-08-02  push         62 min  AC:0  Mave:1  Ens:0  R:4  AC      4 god træning
  2026-08-05  legs         71 min  AC:12 Mave:0  Ens:1  R:5  Center  5 med Anders

Add 2 workout(s) to Excel? [y/N] y
Added 2 row(s).
```

Exit codes: `0` success, `1` runtime failure, `2` usage error.

## Configuration

Read from `.env` in this folder, or from real environment variables (which win):

| Variable | Value |
|---|---|
| `HEVY_API_KEY` | your Hevy API key |
| `EXCEL_PATH` | `C:\Users\viggo\OneDrive\Privat\Fitness\Træning.xlsx` |
| `EXCEL_SHEET` | `Træning` (optional — this is the default) |

`.env` is git-ignored, so the key is never committed.

The `.env` is located from the package's own path, not the current directory, so
`hevy sync` works from anywhere. A `.env` in the directory you happen to be
standing in is also read, but this folder's copy wins.

## How it works

`hevy_client` pages through `api.hevyapp.com/v1/workouts` (newest first) and stops
as soon as it reaches a workout on or before the cutoff — by default the date of
the last row in the sheet.

`column_mapper` turns each workout into a row. It is pure and unit-tested:

- Danish workout titles map to the sheet's categories — `over`→`upper`,
  `ben`→`legs`, `træk`→`pull`, `skub`→`push`; anything else passes through.
- **Rating** is dug out of the description: `4/5`, `8/10`, or a bare 1–10 at the
  start or end (`4 god træning`).
- **Place** is matched against `AC`, `Center`, `Tryg`.
- **Ensamble** is 1 when the description says `med <Name>` / `m. <Name>`.
- **Mave** is 1 when any exercise looks like core work (plank, crunch, leg raise…).
  Core work is never counted as cardio, even though it is timed and weightless.
- **AddCardio** collects minutes from timed, load-free exercises, and those minutes
  are subtracted from `Time` so the two don't double-count.
- **Claude** is set to 1, marking the row as tool-written rather than hand-typed.

`excel_writer` knows the physical layout of `Table4` (columns `A`–`L`: Date, Place,
Type, Time, Mave, AddCardio, AddCardio2, Total, Ensamble, Rating, Claude,
Comments). **`Total` is written as a formula, never as a value**, so the column
keeps recalculating the way the hand-kept rows do:

```
=IF(ISBLANK(Table4[[#This Row],[Date]]), " ", SUM(Dr:Gr)-IF(Er = 1, 1, 0))
```

`hevy fix` re-fetches and re-maps rows from **2026-06-14 onward**
(`sync_core.FIX_CUTOFF`) and overwrites them, matched by date. That cutoff is the
first day the tool wrote rows — earlier rows are hand-typed and are never touched.

## MCP server

`hevy_mcp.server` exposes `preview_workouts`, `sync_workouts`, `add_single_workout`
and `fix_synced_rows` over MCP, sharing the same `sync_core` pipeline as the CLI.

`mcp` is pinned to `<2` in `pyproject.toml`: version 2.0 replaced `FastMCP` with
`MCPServer`, and `server.py` is written against the 1.x API. Porting it is an open
job — until then, unpinning breaks the server (the CLI is unaffected).

It is **not registered anywhere** right now. To wire it into Claude Code:

```powershell
claude mcp add hevy -- "C:\Users\viggo\Git Clone\PC-CLI\projects\hevy\.venv\Scripts\hevy-mcp.exe"
```

## Tests

```powershell
.\.venv\Scripts\python.exe -m pytest -q
```

27 tests, all offline — the Hevy API is stubbed and the Excel tests build a tiny
throwaway workbook in a temp folder. **Nothing touches your real spreadsheet.**

## Origin

Imported from the `Hevy` repo's local working copy at commit `8d09193`
(branch `hevy-cli`). The GitHub remote `viggogo/Hevy` is 16 commits behind that
and predates both the CLI and the `Claude` column — do not treat it as the source.

The one-time migration that inserted the `Claude` column into `Table4` has already
been run against the live workbook, and its script was deliberately left behind in
the old repo.
