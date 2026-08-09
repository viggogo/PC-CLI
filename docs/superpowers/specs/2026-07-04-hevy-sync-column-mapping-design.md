# Hevy → Træning.xlsx Sync: Column Mapping Improvements

**Date:** 2026-07-04
**Status:** Approved design, pending implementation plan

## Problem

The existing Hevy MCP server (`hevy_mcp/`) appends workout rows to `Træning.xlsx`
(sheet `Træning`), but the mapping has gaps that make synced rows diverge from the
~650 rows the user has maintained by hand:

- Danish workout titles from Hevy (`Over`, `Ben`, `Træk`, `Skub`) are written
  verbatim instead of the English categories the sheet uses (`upper`, `legs`,
  `pull`, `push`).
- `Rating` is sometimes stored as a string instead of a number.
- `Mave` is sometimes `'x'` instead of numeric `1`/`0`.
- Cardio performed inside a workout (e.g. 15 min treadmill) is not registered in
  the `AddCardio` column.
- `AddCardio` / `AddCardio2` were left blank instead of `0`, and `Total` did not
  account for cardio.
- 12 already-synced rows (654–665, dates 2026-06-14 → 2026-07-04) were written by
  the buggy logic and need correcting.

## Goals

1. Translate Danish workout categories into the English equivalents already used
   in the sheet.
2. Keep `Rating` a numeric variable.
3. Register cardio (e.g. treadmill minutes) in the `AddCardio` column when it is
   *added onto* a strength workout.
4. Preserve the established style of the existing ~650 rows for every column,
   including keeping `Total` a live table formula (never a static value).
5. Add a `Claude` provenance column (before Comments) flagging tool-written rows
   (`1`) vs manual rows (`0`), empty for pre-tool history.
6. Correct the 12 rows the buggy tool wrote, without touching any row the user
   entered manually.

## Verified facts (from data inspection)

- **`Total = Time + AddCardio + AddCardio2`** holds in all 664 historical rows.
  `Time` is the strength-only duration; `AddCardio` is added on top.
- The latest workout ("Over", 2026-07-04) contains a **Treadmill** exercise with
  `duration_seconds: 900` (15 min) and null `reps`/`weight_kg` — the detectable
  cardio signal.
- Across the user's entire Hevy history there are only 21 unique exercises; only
  **Treadmill** has the cardio set-shape. The only ab move is **"Sit Up"**, which
  is rep-based (not timed), so it is already safe from cardio misclassification.
- Dedicated cardio/running days historically have `AddCardio = 0` and
  `Time = full duration` — cardio only lands in `AddCardio` when it accompanies
  strength work.
- Rows ≤ 653 (through 2026-05-24) are the user's manual entries; rows 654–665 are
  the buggy tool's output.
- **The sheet is an Excel Table, `Table4`, spanning `A1:K…`** (columns Date →
  Comments), covering the full sheet height so appended rows sit inside the table.
- **`Total` (column H) is a table calculated column:**
  `=IF(ISBLANK(Table4[[#This Row],[Date]])," ",SUM(D:G)-IF(E=1,1,0))`, present in
  rows 2–819 — **except** the 12 buggy rows, where the old tool overwrote the
  formula with a static number. Total must stay a formula, never a written value.
- The only other formulas are **M615–618**: the user's personal side-calc
  (`DAYS(…)`, a ratio, `×7`), sitting outside the table with self-references like
  `=M616/M615`.

## Architecture

- All mapping logic stays in **`hevy_mcp/column_mapper.py`** — a rewritten
  `workout_to_row(workout) -> dict` plus small private helpers. This is the single
  source of truth for how a Hevy workout becomes a sheet row.
- **`excel_writer.py`** gains: (a) knowledge of the new `Claude` column and its
  position, (b) logic to write the `Total` (H) **formula** rather than a static
  value for any row it creates, and (c) a one-time structural migration that
  inserts the `Claude` column (see below).
- `hevy_client.py` (API access) and the existing `server.py` tools
  (`preview_workouts`, `sync_workouts`, `add_single_workout`) keep their surface;
  they call the improved mapper/writer.
- A new MCP tool **`fix_synced_rows`** handles the one-time correction of rows
  654–665. It re-fetches the relevant workouts, re-maps them with the new logic,
  shows a before/after preview, and writes only after user approval — reusing the
  same mapper so corrected rows and future rows are consistent in style. It also
  **restores the `Total` formula** in those 12 rows.

## Excel structure changes (one-time migration)

Adding the `Claude` column and healing the sheet is a structural change, done once
with safeguards:

1. **Back up** `Træning.xlsx` (timestamped copy) before any structural write.
2. **Insert** a new column at position K (immediately left of Comments). Header
   `K1 = "Claude"`. Comments shifts K→L; the empty column L→M; the side-calc M→N.
3. **Repair** the shifted side-calc: rewrite its self-referencing formulas from the
   `M…` column letters to their new `N…` positions (e.g. `=M616/M615` →
   `=N616/N615`). Absolute references like `A2` and the `Table4` name are unchanged.
4. **Expand `Table4`** from `A1:K…` to `A1:L…` and insert a `Claude` table column
   between `Rating` and `Comments`.
5. **Restore** the `Total` formula in the 12 clobbered H cells (654–665).
6. **Verify** after save: reopen the workbook and confirm the H formula is intact
   across all rows, `Table4` is valid with the new column, the side-calc still
   evaluates, and no training data shifted out of alignment.

Because `openpyxl.insert_cols` does not adjust formula references, table ranges, or
defined names, steps 3–5 are performed explicitly rather than relied upon.

## Column specification

| Column | Rule |
|---|---|
| **Date** | `start_time` → date. Unchanged. |
| **Place** | Case-insensitive scan of the description for a known gym name (`AC`, `Center`, `Tryg`) → use the matched name. If none found → **blank** (empty cell). |
| **Type** | Danish→English map: `Over→upper`, `Ben→legs`, `Træk→pull`, `Skub→push`. Titles that are already English categories (`push`, `pull`, `legs`, `cardio`, `running`, `full body`) pass through. Any title not in the map passes through **unchanged / verbatim**. |
| **Time** | Strength minutes = `elapsed − cardio_minutes`. If the session is cardio-only, `Time = elapsed`. |
| **Mave** | `1` if any ab/core exercise is present, else `0`. Always numeric. |
| **AddCardio** | Cardio minutes, **only when the session also contains strength**. For a cardio-only session → `0`. |
| **AddCardio2** | Always `0`. |
| **Total** | **Not written as a value.** The tool writes/preserves the table calculated-column formula in H so Excel computes it (`SUM(D:G) − (Mave==1 ? 1 : 0)`, i.e. `Time + AddCardio + AddCardio2`). |
| **Ensamble** | `1` if the description contains a Danish "with" marker (`m`, `m.`, or `med`) immediately followed by a capitalized name (e.g. `m Silas`, `med Uwufufu`, `Tryg m Peter`), else `0`. |
| **Rating** | Numeric rating extracted from the description (`4 tryg` → `4`, `2, travlt` → `2`, `3` → `3`). Stored as a Python `int`. |
| **Claude** | Provenance flag. `1` = row written by the sync tool; `0` = row entered manually by the user. Rows predating the tool's first additions (the ~653 historical manual rows) stay **empty**. The tool only ever writes `1` (on rows it creates); the user sets `0` on their own future manual rows. Column positioned immediately **before** Comments; header text `Claude`. |
| **Comments** | The full raw description string, verbatim (including the rating digit and any place/partner words). |

## Cardio and abs detection

**Cardio exercise:** a Hevy exercise whose sets are time/distance-based with
**no weight and no reps** — i.e. at least one set has `duration_seconds` (or
`distance_meters`) populated while `weight_kg` and `reps` are null. Example:
Treadmill. Cardio minutes for the workout = `sum(duration_seconds) / 60`, rounded.

**Ab/core exercise:** a Hevy exercise whose title matches the core-keyword list
(`crunch`, `plank`, `sit-up`/`sit up`/`situp`, `leg raise`, `ab wheel`,
`ab rollout`/`ab roller`, `hanging knee`/`hanging leg`, `cable crunch`,
`russian twist`, `toe touch`, `bicycle crunch`, `dead bug`, `flutter kick`,
`hollow body`, `v-up`/`v up`, `mave`, `core`). Ab exercises:

- set `Mave = 1`, and
- are **excluded** from cardio detection even if they are timed (e.g. a plank has
  `duration_seconds` and no weight but must not count as cardio).

**Cardio-only session:** if *every* exercise in the workout is cardio (no strength
exercise present), the session is treated as a cardio day — `Time = elapsed`,
`AddCardio = 0`. Cardio minutes only move to `AddCardio` when strength work is also
present in the same session.

## Correcting the 12 buggy rows (654–665)

- New `fix_synced_rows` MCP tool:
  1. Re-fetches Hevy workouts dated 2026-06-14 onward.
  2. Re-maps each with the new `workout_to_row` logic.
  3. Prints a **before/after preview of all 12 rows** (current sheet values vs new
     values).
  4. Overwrites rows 654–665 **only after the user approves**, including:
     - all mapped column values (Type, Rating as `int`, Mave/AddCardio/AddCardio2
       numeric, Ensamble, Place, Comments),
     - **restoring the `Total` (H) formula** these rows lost,
     - setting **`Claude = 1`** (they were tool-created).
- The tool never touches rows ≤ 653 (the user's manual entries), which keep an
  **empty** `Claude` cell. Only the tool-added rows in the 2026-06-14 → 2026-07-04
  range are rewritten, in place.

## Testing

Unit tests against `column_mapper.workout_to_row` using fixture workout dicts that
cover:

- Danish title mapping (`Over→upper`, `Ben→legs`, `Træk→pull`, `Skub→push`) and
  verbatim pass-through of an unknown title.
- Treadmill-plus-strength: `Time = elapsed − 15`, `AddCardio = 15` (Total left to
  the formula).
- Cardio-only session: `Time = elapsed`, `AddCardio = 0`.
- Plank present: `Mave = 1`, not counted as cardio.
- Rating extraction to `int` from `"4 tryg"`, `"2, travlt"`, `"3"`.
- Place detection (`Tryg` in note → `Place = Tryg`; no keyword → blank).
- Ensamble "m/med + name" pattern (`m Silas` → `1`; note with no partner → `0`).
- Mapper sets `Claude = 1` on every row it produces.

A separate integration-style check covers the Excel structural migration on a
**copy** of the workbook: after inserting the `Claude` column, assert the H formula
is intact for all data rows, `Table4` ref is `A1:L…` with a `Claude` column, the
side-calc formulas now reference `N…` and still evaluate, and Comments data landed
in column L unshifted.

## Out of scope

- No change to how workouts are fetched or paginated (`hevy_client.py`).
- No new `AddCardio2` semantics beyond defaulting to `0`.
- No automatic place inference beyond the `AC` / `Center` / `Tryg` keyword list.
- The tool does not backfill `Claude = 0` on historical manual rows; they stay
  empty. The user marks their own future manual rows.
