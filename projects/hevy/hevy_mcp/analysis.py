"""Read-only calendar view of recent training, for `hevy --analysis N`."""

import os
from datetime import date, datetime, timedelta

from . import column_mapper

# Hevy adds your bodyweight to a set's load only for movements that use 100% of
# it, and their list is exactly these four. Everything else counts the logged
# weight alone -- a weighted sit-up scores its plate and nothing more.
_BODYWEIGHT_KEYWORDS = ("pull up", "chin up", "dip", "handstand push up")

# "Bench Dip" matches "dip" but rests on a bench, so it carries no bodyweight.
_NOT_BODYWEIGHT = ("bench dip",)


def week_starts(n_weeks: int, today: date) -> list[date]:
    """The Mondays of the last `n_weeks` weeks, oldest first. The final one is
    the Monday of the week `today` falls in, so the grid ends on the present."""
    if n_weeks < 1:
        raise ValueError(f"weeks must be 1 or more, got {n_weeks}")
    this_monday = today - timedelta(days=today.weekday())
    return [this_monday - timedelta(weeks=n)
            for n in range(n_weeks - 1, -1, -1)]


def bodyweight_kg() -> float:
    """Your bodyweight, for the exercises Hevy loads with it. The Hevy API has
    no endpoint for it, so it comes from .env. Unset or unreadable means 0,
    which leaves bodyweight out of the total rather than guessing a number."""
    try:
        return float(os.environ.get("HEVY_BODYWEIGHT_KG") or 0)
    except ValueError:
        return 0.0


def _is_bodyweight_exercise(title: str) -> bool:
    t = (title or "").lower()
    if any(kw in t for kw in _NOT_BODYWEIGHT):
        return False
    return any(kw in t for kw in _BODYWEIGHT_KEYWORDS)


def workout_volume(workout: dict, bodyweight: float = 0.0) -> float:
    """Total kg moved, the way Hevy counts it: load x reps over every set,
    warm-ups included.

    Load is the logged weight, except on the four movements that lift all of
    you -- pull-ups, chin-ups, dips, handstand push-ups -- where Hevy adds your
    bodyweight to it. The API sends only the *added* plate for those, so
    without this a bodyweight-only set scores zero and a pull day comes out
    thousands of kg light. Assisted variants take the assistance off instead.
    Timed and rep-only sets have no reps or no weight, so they add nothing.
    """
    total = 0.0
    for ex in workout.get("exercises", []):
        title = ex.get("title", "")
        carries_bodyweight = _is_bodyweight_exercise(title)
        assisted = carries_bodyweight and "assisted" in title.lower()
        base = bodyweight if carries_bodyweight else 0.0
        for s in ex.get("sets", []):
            weight = s.get("weight_kg") or 0
            load = max(0.0, base - weight) if assisted else base + weight
            total += load * (s.get("reps") or 0)
    return total


def _start(workout: dict) -> datetime:
    return datetime.fromisoformat(workout["start_time"].replace("Z", "+00:00"))


def _session(workout: dict) -> dict:
    """The four facts the calendar shows. Rating and place come from the same
    mapper `preview` and `sync` use, so the views never disagree. The name is
    the raw Hevy title, not the sheet's category."""
    row = column_mapper.workout_to_row(workout)
    end = datetime.fromisoformat(workout["end_time"].replace("Z", "+00:00"))
    return {
        "name": workout.get("title") or "",
        "rating": row["Rating"],
        "volume": workout_volume(workout, bodyweight_kg()),
        "place": row["Place"],
        # Whole session, cardio included: the summary measures time in the gym,
        # not the sheet's cardio-adjusted Time column.
        "minutes": max(1, round((end - _start(workout)).total_seconds() / 60)),
    }


def build_calendar(workouts: list[dict], n_weeks: int,
                   today: date) -> list[list[dict]]:
    """A grid of `n_weeks` rows by 7 day-cells, Monday first, ending on the week
    containing `today`. Each cell is {"date", "sessions"}; a rest day has no
    sessions. Workouts falling outside the window are ignored."""
    mondays = week_starts(n_weeks, today)
    grid = [[{"date": monday + timedelta(days=i), "sessions": []}
             for i in range(7)] for monday in mondays]

    by_date = {cell["date"]: cell for week in grid for cell in week}
    for w in sorted(workouts, key=_start):
        cell = by_date.get(_start(w).date())
        if cell is not None:
            cell["sessions"].append(_session(w))
    return grid


def summarize(grid: list[list[dict]]) -> dict:
    """Weekly averages over the whole grid. Rest weeks stay in the denominator,
    and a day trained twice is still one training day."""
    weeks = len(grid)
    days = sum(1 for week in grid for cell in week if cell["sessions"])
    minutes = sum(s["minutes"] for week in grid for cell in week
                  for s in cell["sessions"])
    return {
        "weeks": weeks,
        "days": days,
        "minutes": minutes,
        "days_per_week": days / weeks,
        "hours_per_week": minutes / 60 / weeks,
    }


# --- rendering --------------------------------------------------------------

DAYS = ("Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun")
CELL = 11          # printable width of one day-cell, borders excluded
TEXT = CELL - 2    # usable characters once the cell's padding is removed
GUTTER = 9         # left label column, wide enough for " Aug/Sep "


def _fmt_rating(rating) -> str:
    """The sheet rates 1-6. Anything above that can only have come from an
    explicit "n/10" in the description, so keep that scale as written."""
    if rating is None:
        return "-"
    return f"{rating}/6" if rating <= 6 else f"{rating}/10"


def _fmt_volume(volume: float) -> str:
    """No thousands separator: in a 9-character cell a space reads as two
    numbers, and a comma reads as a Danish decimal point."""
    return f"{int(round(volume))} kg"


def _cell_lines(session: dict) -> list[str]:
    return [
        session["name"],
        f"{_fmt_rating(session['rating']):<5}{session['place'][:4]:>4}",
        _fmt_volume(session["volume"]),
    ]


def _row(label: str, cells: list[str]) -> str:
    body = "|".join(f" {c[:TEXT]:<{TEXT}} " for c in cells)
    return f" {label}".ljust(GUTTER) + f"|{body}|"


def _week_label(week: list[dict]) -> tuple[str, str]:
    """Month (or months, for a week that straddles one) and the day range."""
    first, last = week[0]["date"], week[6]["date"]
    months = first.strftime("%b")
    if last.month != first.month:
        months += "/" + last.strftime("%b")
    return months, f"{first.day:02d}-{last.day:02d}"


def render_calendar(grid: list[list[dict]]) -> str:
    """Draw the grid as a Mon-Sun calendar. Every week gets the same height, so
    the rows stay aligned whether or not anything was trained in them."""
    border = " " * GUTTER + "+" + "+".join("-" * CELL for _ in DAYS) + "+"
    header = " " * (GUTTER + 1) + " ".join(d.center(CELL) for d in DAYS)
    lines = [header.ljust(len(border)), border]

    for week in grid:
        # Each week is only as tall as its own busiest day, so a rest week
        # doesn't inherit blank rows from a week with two-a-days.
        slots = max(1, max(len(c["sessions"]) for c in week))
        months, day_range = _week_label(week)
        lines.append(_row(months, [f"{c['date'].day:02d}" for c in week]))
        labels = [day_range] + [""] * (3 * slots - 1)
        for slot in range(slots):
            blocks = [_cell_lines(c["sessions"][slot])
                      if slot < len(c["sessions"]) else ["", "", ""]
                      for c in week]
            for i in range(3):
                lines.append(_row(labels[slot * 3 + i],
                                  [b[i] for b in blocks]))
        lines.append(border)

    s = summarize(grid)
    plural = "" if s["weeks"] == 1 else "s"
    lines.append(f" {s['weeks']} week{plural}: "
                 f"{s['days_per_week']:.1f} training days/week, "
                 f"{s['hours_per_week']:.1f} hours/week")
    return "\n".join(lines)
