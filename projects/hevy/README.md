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
hevy --weeks N        Show the last N weeks of training as a calendar.
hevy --weeks N -1     Same, but drop the current, unfinished week.
hevy --help           Show usage. Also: hevy sync --help, etc.
```

Options:

```
--since YYYY-MM-DD    (preview, sync) Override the start date.
                      Default = everything after your last Excel row.
-y, --yes             (sync) Skip the confirmation prompt.
--weeks N             Whole number of weeks, 1–52. Required — there is no
                      default. Reads the API only; never opens the workbook.
-1                    (--weeks) Drop the current, unfinished week. Nothing
                      older replaces it, so you get N-1 weeks. `--weeks 1 -1`
                      is a usage error — it would leave nothing to show.
```

**Close Excel before `sync`**, the one command that writes. Excel holds a lock on
the file; the tool reports this rather than failing halfway. `sync` only ever
appends new rows — nothing already in the sheet is touched.

Typical run — the two workouts were noted in Hevy as `4, AC, god træning` and
`5, Center, med Anders`:

```
> hevy preview
2 new workout(s):
  2026-08-02  push         62 min  AC:0  Mave:1  Ens:0  R:4  AC      god træning
  2026-08-05  legs         71 min  AC:12 Mave:0  Ens:1  R:5  Center  med Anders

> hevy sync
2 new workout(s):
  2026-08-02  push         62 min  AC:0  Mave:1  Ens:0  R:4  AC      god træning
  2026-08-05  legs         71 min  AC:12 Mave:0  Ens:1  R:5  Center  med Anders

Add 2 workout(s) to Excel? [y/N] y
Added 2 row(s).
```

The calendar:

```
> hevy --weeks 2
              Mon         Tue         Wed         Thu         Fri         Sat         Sun
         +-----------+-----------+-----------+-----------+-----------+-----------+-----------+
 Aug     | 17        | 18        | 19        | 20        | 21        | 22        | 23        |
 17-23   | skub      |           | ben       |           | over      |           |           |
         | 4/6    AC |           | 6/6  Cent |           | 3/6    AC |           |           |
         | 4210 kg   |           | 6090 kg   |           | 3820 kg   |           |           |
         +-----------+-----------+-----------+-----------+-----------+-----------+-----------+
 Aug     | 24        | 25        | 26        | 27        | 28        | 29        | 30        |
 24-30   |           | traek     |           | ben       |           |           | over      |
         |           | 4/6    AC |           | 5/6    AC |           |           | 5/6  Tryg |
         |           | 5100 kg   |           | 17340 kg  |           |           | 4650 kg   |
         |           |           |           | cardio    |           |           |           |
         |           |           |           | -         |           |           |           |
         |           |           |           | 0 kg      |           |           |           |
         +-----------+-----------+-----------+-----------+-----------+-----------+-----------+
 2 weeks: 3.0 training days/week, 3.6 hours/week
```

Exit codes: `0` success, `1` runtime failure, `2` usage error.

## The note convention

The description you write on a workout in Hevy fills three columns, read by
position:

```
<rating>, <place>, <comment>
```

| Note in Hevy | Rating | Place | Comments |
|---|---|---|---|
| `4, AC, godt tempo` | 4 | `AC` | `godt tempo` |
| `4, sats, tungt i dag` | 4 | `sats` | `tungt i dag` |
| `4, ac, fint` | 4 | `AC` | `fint` |
| `4, AC` | 4 | `AC` | *(blank)* |
| `4` — also `4/5`, `8/10` | 4 | *(blank)* | *(blank)* |
| `4, AC, godt, men træt` | 4 | `AC` | `godt, men træt` |
| `god træning, lidt træt` | *(blank)* | *(blank)* | `god træning, lidt træt` |
| *(empty)* | *(blank)* | *(blank)* | *(blank)* |

The convention only engages when the **first field is a bare rating** — `4`,
`4/5`, `4/6` or `8/10`, value 1–10 and nothing else in that field. Any other note
is a comment in full, commas included; that is what stops ordinary prose like
`god træning, lidt træt` from writing `lidt træt` into the Place column.

**Place is free-form**: type any gym and it lands in the sheet as written. The
three names the sheet already uses — `AC`, `Center`, `Tryg` — are recognised
case-insensitively and normalised to that spelling, so `ac` still reads `AC`.

Everything from the third field on is the comment, so commas inside it survive.
The rating and place are *not* repeated in the comment — they have their own
columns.

## Configuration

Read from `.env` in this folder, or from real environment variables (which win):

| Variable | Value |
|---|---|
| `HEVY_API_KEY` | your Hevy API key |
| `EXCEL_PATH` | `C:\Users\viggo\OneDrive\Privat\Fitness\Træning.xlsx` |
| `EXCEL_SHEET` | `Træning` (optional — this is the default) |
| `HEVY_BODYWEIGHT_KG` | your bodyweight, e.g. `83` — `--weeks` needs it to match the app |

`HEVY_BODYWEIGHT_KG` exists because Hevy counts pull-ups, chin-ups, dips and
handstand push-ups as lifting your whole bodyweight, but the API sends only the
plate you hung off yourself — and has no endpoint for bodyweight at all. Without
it a pull day comes out thousands of kg short. Leave it unset and those sets
count as their added weight alone.

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
- **Rating, Place and Comments** come from the workout note, written as three
  comma-separated fields in that order — see [The note convention](#the-note-convention).
- **Ensamble** is 1 when the note says `med <Name>` / `m. <Name>`, wherever in the
  note that sits.
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

`analysis` builds the calendar behind `--weeks`. The window is always whole
Mon–Sun weeks, the last of which is the week containing today — so days later
this week show as empty cells rather than being cut off. Each day-cell carries
the workout's **name** (the raw Hevy title, not the sheet's category), **rating**
and **place** — both read through `column_mapper`, so the calendar can never
disagree with `preview` — and **volume**, which is new here. It is printed
without a thousands separator, since in a 9-character cell a space reads as two
numbers and a comma reads as a Danish decimal point.

Volume follows Hevy's own definition, so the cell matches what the app shows on
the workout: load × reps over every set, warm-ups included. Load is the logged
weight, except on the four movements Hevy treats as lifting all of you — pull-up,
chin-up, dip, handstand push-up — where `HEVY_BODYWEIGHT_KG` is added to it
(subtracted, for an `(Assisted)` variant). Nothing else gets bodyweight: a
weighted sit-up counts its plate alone, and `Bench Dip` is excluded by name
despite the "dip". Timed and rep-only sets carry no load, so a pure cardio
session still reads `0 kg`.

One limitation: Hevy uses your bodyweight *as measured at the time of each
workout*, while `.env` holds one current value. Over a 2–12 week window the drift
is negligible, but volumes for much older workouts will read slightly off — and
the API exposes no per-workout weights to fix that with.

Ratings render on the sheet's **1–6** scale (`4/6`); a rating above 6 can only
have come from an explicit `n/10` in the description, so that scale is kept as
written. An unrated session shows `-`.

The footer under the grid averages the window. **Training days per week** counts
distinct days you trained, so a day with two sessions still counts once.
**Hours per week** sums whole sessions, `end_time − start_time`, cardio included
— time in the gym, not the sheet's cardio-adjusted `Time` column. Rest weeks stay
in the denominator; dropping them would flatter the average. Since the current
week is usually unfinished, it drags both figures down until the week fills in —
`-1` removes it.

Two sessions on one day stack inside the cell, and each week is only as tall as
its own busiest day. Names longer than 9 characters are truncated to the cell
width. `--weeks` never opens the workbook, so it works with Excel still open.

`excel_writer` only ever **appends**. There is no command that rewrites an
existing row — an earlier `hevy fix`, which re-mapped rows from 2026-06-14 onward
in place, has been removed. Rows already in the sheet keep whatever they were
written with, including notes that predate the comma convention; to correct one,
edit it in Excel.

## MCP server

`hevy_mcp.server` exposes `preview_workouts`, `sync_workouts` and
`add_single_workout` over MCP, sharing the same `sync_core` pipeline as the CLI.

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

93 tests, all offline — the Hevy API is stubbed and the Excel tests build a tiny
throwaway workbook in a temp folder. **Nothing touches your real spreadsheet.**

## Origin

Imported from the `Hevy` repo's local working copy at commit `8d09193`
(branch `hevy-cli`). The GitHub remote `viggogo/Hevy` is 16 commits behind that
and predates both the CLI and the `Claude` column — do not treat it as the source.

The one-time migration that inserted the `Claude` column into `Table4` has already
been run against the live workbook, and its script was deliberately left behind in
the old repo.
