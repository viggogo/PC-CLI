"""Command-line interface for syncing Hevy workouts to Træning.xlsx.

Usage:
    hevy preview            show new workouts, write nothing
    hevy sync               show new workouts, then ask before adding
    hevy sync -y            add without asking
    hevy --weeks 2          show the last 2 weeks of training as a calendar
    hevy --weeks 5 -1       the same, without the current unfinished week

Reads HEVY_API_KEY and EXCEL_PATH from the environment or a local .env file.
`--weeks` is read-only and needs HEVY_API_KEY alone.
"""

import argparse
import asyncio
import os
import sys

import httpx

from . import analysis, sync_core
from .env import ENV_FILE, load_env


def _fmt_row(r: dict) -> str:
    rating = r["Rating"] if r["Rating"] is not None else "-"
    return (f"{r['Date']}  {str(r['Type'])[:10]:<10}  {r['Time']:>3} min  "
            f"AC:{r['AddCardio']}  Mave:{r['Mave']}  Ens:{r['Ensamble']}  "
            f"R:{rating}  {r['Place'] or '':<6}  {r['Comments'] or ''}")


def _print_rows(rows: list) -> None:
    for r in rows:
        print("  " + _fmt_row(r))


def _confirm(prompt: str) -> bool:
    return input(prompt).strip().lower() in ("y", "yes")


def _require_env(*names: str) -> None:
    missing = [v for v in names if not os.environ.get(v)]
    if missing:
        sys.exit(f"Missing env var(s): {', '.join(missing)}. "
                 f"Set them, or fill them in here: {ENV_FILE}")


def cmd_preview(args) -> None:
    rows = asyncio.run(sync_core.collect_new_rows(args.since))
    if not rows:
        print("No new workouts.")
        return
    print(f"{len(rows)} new workout(s):")
    _print_rows(rows)


def cmd_sync(args) -> None:
    rows = asyncio.run(sync_core.collect_new_rows(args.since))
    if not rows:
        print("No new workouts to sync.")
        return
    print(f"{len(rows)} new workout(s):")
    _print_rows(rows)
    if not args.yes and not _confirm(f"\nAdd {len(rows)} workout(s) to Excel? [y/N] "):
        print("Aborted — nothing written.")
        return
    n = sync_core.append_new_rows(rows)
    print(f"Added {n} row(s).")


MAX_WEEKS = 52


def cmd_weeks(weeks: int, skip_current: bool = False) -> None:
    grid = asyncio.run(sync_core.collect_calendar(weeks, skip_current))
    print(analysis.render_calendar(grid))


def build_parser() -> argparse.ArgumentParser:
    # allow_abbrev=False: argparse otherwise accepts any unambiguous prefix, so
    # `--week`, `--wee` and `--w` would all silently mean `--weeks`.
    p = argparse.ArgumentParser(
        prog="hevy", description="Sync Hevy workouts to Træning.xlsx",
        allow_abbrev=False)
    p.add_argument("--weeks", type=int, metavar="N",
                   help=f"show the last N weeks (1-{MAX_WEEKS}) of training "
                        f"as a calendar")
    p.add_argument("-1", dest="skip_current", action="store_true",
                   help="with --weeks: drop the current, unfinished week")
    # Not required: `hevy --weeks 2` carries no subcommand. main() still
    # rejects an invocation that names neither.
    sub = p.add_subparsers(dest="command")

    pv = sub.add_parser("preview", help="show new workouts without writing")
    pv.add_argument("--since", default="",
                    help="YYYY-MM-DD (default: after last Excel row)")
    pv.set_defaults(func=cmd_preview)

    sy = sub.add_parser("sync", help="add new workouts (asks first)")
    sy.add_argument("--since", default="",
                    help="YYYY-MM-DD (default: after last Excel row)")
    sy.add_argument("-y", "--yes", action="store_true", help="skip confirmation")
    sy.set_defaults(func=cmd_sync)

    return p


def main(argv=None) -> None:
    load_env()
    parser = build_parser()
    args = parser.parse_args(argv)

    if args.weeks is not None:
        if not 1 <= args.weeks <= MAX_WEEKS:
            parser.error(f"--weeks takes a whole number between 1 "
                         f"and {MAX_WEEKS}, got {args.weeks}")
        if args.skip_current and args.weeks == 1:
            parser.error("--weeks 1 -1 leaves no weeks to show")
    elif args.skip_current:
        parser.error("-1 only means anything alongside --weeks N")
    elif args.command is None:
        parser.error("give a command (preview, sync) or --weeks N")

    try:
        if args.weeks is not None:
            _require_env("HEVY_API_KEY")   # reads the API only, never Excel
            cmd_weeks(args.weeks, args.skip_current)
            return
        _require_env("HEVY_API_KEY", "EXCEL_PATH")
        args.func(args)
    except httpx.HTTPStatusError as e:
        # 401 is the common one: the key is still the .env.example placeholder.
        if e.response.status_code == 401:
            sys.exit(f"Hevy rejected the API key (401). Check HEVY_API_KEY in "
                     f"{ENV_FILE}")
        sys.exit(f"Hevy API error {e.response.status_code} for {e.request.url}")
    except FileNotFoundError as e:
        sys.exit(f"Cannot open the spreadsheet: {e.filename}\n"
                 f"Check EXCEL_PATH in {ENV_FILE}")
    except PermissionError:
        sys.exit(f"The spreadsheet is locked — close Excel and try again.\n"
                 f"  {os.environ.get('EXCEL_PATH')}")


if __name__ == "__main__":
    main()
