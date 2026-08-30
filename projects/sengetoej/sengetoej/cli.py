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


def cmd_new(value: str, assume_yes: bool) -> int:
    raise NotImplementedError  # Task 8


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
