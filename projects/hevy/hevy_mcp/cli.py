"""Command-line interface for syncing Hevy workouts to Træning.xlsx.

Usage:
    hevy preview            show new workouts, write nothing
    hevy sync               show new workouts, then ask before adding
    hevy sync -y            add without asking
    hevy fix                show corrections for recent tool rows, then ask

Reads HEVY_API_KEY and EXCEL_PATH from the environment or a local .env file.
"""

import argparse
import asyncio
import os
import sys

import httpx

from . import sync_core
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


def _require_env() -> None:
    missing = [v for v in ("HEVY_API_KEY", "EXCEL_PATH") if not os.environ.get(v)]
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


def cmd_fix(args) -> None:
    changes = asyncio.run(sync_core.collect_fix_changes(dry_run=True))
    if not changes:
        print("No rows to fix.")
        return
    print(f"{len(changes)} row(s) would change:")
    for rn, d, before, after in changes:
        print(f"  row {rn} {d}: {before.get('Type')!r}->{after['Type']!r}  "
              f"R {before.get('Rating')!r}->{after['Rating']!r}  "
              f"AC {before.get('AddCardio')!r}->{after['AddCardio']!r}")
    if not args.yes and not _confirm("\nApply these corrections? [y/N] "):
        print("Aborted — nothing written.")
        return
    asyncio.run(sync_core.collect_fix_changes(dry_run=False))
    print("Applied.")


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="hevy", description="Sync Hevy workouts to Træning.xlsx")
    sub = p.add_subparsers(dest="command", required=True)

    pv = sub.add_parser("preview", help="show new workouts without writing")
    pv.add_argument("--since", default="",
                    help="YYYY-MM-DD (default: after last Excel row)")
    pv.set_defaults(func=cmd_preview)

    sy = sub.add_parser("sync", help="add new workouts (asks first)")
    sy.add_argument("--since", default="",
                    help="YYYY-MM-DD (default: after last Excel row)")
    sy.add_argument("-y", "--yes", action="store_true", help="skip confirmation")
    sy.set_defaults(func=cmd_sync)

    fx = sub.add_parser("fix", help="correct the recent tool-written rows")
    fx.add_argument("-y", "--yes", action="store_true", help="skip confirmation")
    fx.set_defaults(func=cmd_fix)

    return p


def main(argv=None) -> None:
    load_env()
    args = build_parser().parse_args(argv)
    _require_env()
    try:
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
