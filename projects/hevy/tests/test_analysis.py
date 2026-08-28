from datetime import date, timedelta

import pytest

from hevy_mcp import analysis

from .conftest import cardio_ex, strength_ex


def _sets(*pairs):
    """Build a strength exercise from (reps, weight_kg) pairs."""
    return {"title": "Squat (Barbell)", "sets": [
        {"reps": r, "weight_kg": w, "duration_seconds": None,
         "distance_meters": None} for r, w in pairs]}


def test_volume_is_weight_times_reps(make_workout):
    w = make_workout(exercises=[strength_ex(reps=8, weight_kg=60)])

    assert analysis.workout_volume(w) == 480


def test_volume_sums_every_set_of_every_exercise(make_workout):
    w = make_workout(exercises=[
        _sets((5, 100), (5, 100), (3, 120)),
        strength_ex(reps=10, weight_kg=40),
    ])

    assert analysis.workout_volume(w) == 500 + 500 + 360 + 400


def test_bodyweight_sets_add_no_volume(make_workout):
    w = make_workout(exercises=[_sets((12, None), (12, 0))])

    assert analysis.workout_volume(w) == 0


def test_cardio_sets_add_no_volume(make_workout):
    w = make_workout(exercises=[cardio_ex()])

    assert analysis.workout_volume(w) == 0


def test_workout_without_exercises_has_zero_volume(make_workout):
    w = make_workout(exercises=[])

    assert analysis.workout_volume(w) == 0


# --- week_starts ------------------------------------------------------------

def test_weeks_end_with_the_monday_of_this_week():
    # Fri 28 Aug 2026 sits in the week starting Mon 24 Aug.
    assert analysis.week_starts(2, date(2026, 8, 28)) == [
        date(2026, 8, 17), date(2026, 8, 24)]


def test_a_monday_belongs_to_its_own_week_not_the_previous_one():
    assert analysis.week_starts(1, date(2026, 8, 24)) == [date(2026, 8, 24)]


def test_a_sunday_belongs_to_the_week_that_started_six_days_earlier():
    assert analysis.week_starts(1, date(2026, 8, 30)) == [date(2026, 8, 24)]


def test_weeks_run_backwards_across_a_year_boundary():
    # Sun 4 Jan 2026 is in the week starting Mon 29 Dec 2025.
    assert analysis.week_starts(3, date(2026, 1, 4)) == [
        date(2025, 12, 15), date(2025, 12, 22), date(2025, 12, 29)]


@pytest.mark.parametrize("n", [0, -1])
def test_fewer_than_one_week_is_rejected(n):
    with pytest.raises(ValueError):
        analysis.week_starts(n, date(2026, 8, 28))


# --- build_calendar ---------------------------------------------------------

TODAY = date(2026, 8, 28)  # Friday; the 2-week window is 17 Aug .. 30 Aug


def _at(day: str, hour: int = 11):
    return f"{day}T{hour:02d}:00:00+00:00"


def test_grid_is_one_row_per_week_of_seven_days():
    grid = analysis.build_calendar([], 3, TODAY)

    assert len(grid) == 3
    assert all(len(week) == 7 for week in grid)


def test_cells_carry_consecutive_dates_starting_on_monday():
    grid = analysis.build_calendar([], 2, TODAY)

    assert [c["date"] for c in grid[0]] == [
        date(2026, 8, d) for d in range(17, 24)]
    assert grid[1][0]["date"] == date(2026, 8, 24)
    assert grid[1][6]["date"] == date(2026, 8, 30)


def test_a_workout_lands_on_its_own_day_with_its_details(make_workout):
    w = make_workout(title="ben", description="4/5 AC",
                     exercises=[strength_ex(reps=5, weight_kg=100)],
                     start=_at("2026-08-19"), end=_at("2026-08-19", 12))

    grid = analysis.build_calendar([w], 2, TODAY)

    wednesday = grid[0][2]
    assert wednesday["date"] == date(2026, 8, 19)
    assert wednesday["sessions"] == [
        {"name": "ben", "rating": 4, "volume": 500, "place": "AC",
         "minutes": 60}]


def test_a_session_carries_its_whole_length_in_minutes(make_workout):
    # 09:00 -> 10:45, cardio included: the summary counts time in the gym.
    w = make_workout(start=_at("2026-08-19", 9),
                     end="2026-08-19T10:45:00+00:00")

    grid = analysis.build_calendar([w], 2, TODAY)

    assert grid[0][2]["sessions"][0]["minutes"] == 105


def test_days_without_training_have_no_sessions(make_workout):
    w = make_workout(start=_at("2026-08-19"), end=_at("2026-08-19", 12))

    grid = analysis.build_calendar([w], 2, TODAY)

    assert grid[0][0]["sessions"] == []
    assert [c["date"] for c in grid[0] if c["sessions"]] == [date(2026, 8, 19)]


def test_two_workouts_on_one_day_are_both_kept_in_time_order(make_workout):
    evening = make_workout(title="ben", start=_at("2026-08-19", 18),
                           end=_at("2026-08-19", 19))
    morning = make_workout(title="skub", start=_at("2026-08-19", 7),
                           end=_at("2026-08-19", 8))

    grid = analysis.build_calendar([evening, morning], 2, TODAY)

    assert [s["name"] for s in grid[0][2]["sessions"]] == ["skub", "ben"]


def test_workouts_outside_the_window_are_dropped(make_workout):
    before = make_workout(start=_at("2026-08-16"), end=_at("2026-08-16", 12))
    after = make_workout(start=_at("2026-08-31"), end=_at("2026-08-31", 12))

    grid = analysis.build_calendar([before, after], 2, TODAY)

    assert all(cell["sessions"] == [] for week in grid for cell in week)


def test_a_workout_without_a_rating_or_place_still_appears(make_workout):
    w = make_workout(title="over", description="",
                     start=_at("2026-08-25"), end=_at("2026-08-25", 12))

    grid = analysis.build_calendar([w], 2, TODAY)

    session = grid[1][1]["sessions"][0]
    assert session["rating"] is None
    assert session["place"] == ""


# --- grid helpers -----------------------------------------------------------

def _day(d: date, *sessions):
    return {"date": d, "sessions": list(sessions)}


def _session(name="ben", rating=4, volume=6090, place="Center", minutes=60):
    return {"name": name, "rating": rating, "volume": volume, "place": place,
            "minutes": minutes}


def _weeks(*specs):
    """Consecutive weeks from Mon 24 Aug 2026. Each spec maps a weekday index
    to that day's sessions."""
    return [[_day(date(2026, 8, 24) + timedelta(weeks=w, days=i),
                  *spec.get(i, []))
             for i in range(7)]
            for w, spec in enumerate(specs)]


def _one_week(*sessions_by_index):
    """A single week (Mon 24 Aug 2026) with sessions placed by weekday index."""
    placed = dict(sessions_by_index)
    return [[_day(date(2026, 8, 24) + timedelta(days=i), *placed.get(i, []))
             for i in range(7)]]


def _cells(line: str) -> list[str]:
    """The seven day-cells of a rendered row, Monday first."""
    return line.split("|")[1:8]


def _row_containing(out: str, text: str) -> str:
    return next(ln for ln in out.splitlines() if text in ln)


# --- summarize --------------------------------------------------------------

def test_summary_counts_training_days_and_total_minutes():
    grid = _weeks({1: [_session(minutes=60)], 3: [_session(minutes=90)]})

    s = analysis.summarize(grid)

    assert (s["weeks"], s["days"], s["minutes"]) == (1, 2, 150)


def test_a_two_a_day_counts_as_one_day_but_both_sessions_of_time():
    grid = _weeks({1: [_session(minutes=60), _session(minutes=30)]})

    s = analysis.summarize(grid)

    assert s["days"] == 1
    assert s["minutes"] == 90


def test_averages_divide_by_every_week_including_rest_weeks():
    grid = _weeks({1: [_session(minutes=60)], 3: [_session(minutes=60)]}, {})

    s = analysis.summarize(grid)

    assert s["days_per_week"] == 1.0
    assert s["hours_per_week"] == 1.0


def test_a_grid_with_no_training_averages_zero():
    s = analysis.summarize(_weeks({}, {}))

    assert s["days_per_week"] == 0
    assert s["hours_per_week"] == 0


# --- render_calendar --------------------------------------------------------

def test_header_names_the_days_monday_first():
    out = analysis.render_calendar(_one_week())

    header = out.splitlines()[0]
    assert header.split() == ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]


def test_every_grid_line_is_the_same_width_so_the_columns_line_up():
    out = analysis.render_calendar(_one_week((1, [_session()])))

    widths = {len(ln) for ln in out.splitlines()[:-1]}  # last line is the footer
    assert len(widths) == 1


def test_the_calendar_ends_with_the_weekly_averages():
    grid = _weeks({1: [_session(minutes=60)], 3: [_session(minutes=120)]}, {})

    footer = analysis.render_calendar(grid).splitlines()[-1]

    assert "2 weeks" in footer
    assert "1.0 training days/week" in footer   # 2 days over 2 weeks
    assert "1.5 hours/week" in footer           # 180 min over 2 weeks


def test_a_single_week_footer_reads_in_the_singular():
    footer = analysis.render_calendar(_weeks({})).splitlines()[-1]

    assert "1 week:" in footer and "1 weeks" not in footer


def test_the_footer_starts_under_the_grid_not_inside_it():
    out = analysis.render_calendar(_one_week((1, [_session()])))
    lines = out.splitlines()

    assert lines[-2].strip().startswith("+")   # grid closed first
    assert "|" not in lines[-1]


def test_each_day_number_sits_in_its_own_cell():
    out = analysis.render_calendar(_one_week())

    numbers = _row_containing(out, "24")
    assert [c.strip() for c in _cells(numbers)] == [
        str(d) for d in range(24, 31)]


def test_a_session_renders_name_rating_place_and_volume_in_its_own_cell():
    out = analysis.render_calendar(_one_week((1, [_session()])))

    assert _cells(_row_containing(out, "ben"))[1].strip() == "ben"
    rating = _cells(_row_containing(out, "4/6"))[1].strip()
    assert rating.startswith("4/6") and rating.endswith("Cent")
    assert _cells(_row_containing(out, "kg"))[1].strip() == "6090 kg"


def test_a_large_volume_carries_no_thousands_separator():
    # A separator reads as two numbers in a narrow cell, so there isn't one.
    out = analysis.render_calendar(_one_week((1, [_session(volume=17340)])))

    assert _cells(_row_containing(out, "kg"))[1].strip() == "17340 kg"


def test_rest_days_render_as_blank_cells():
    out = analysis.render_calendar(_one_week((1, [_session()])))

    for row in ("ben", "4/6", "kg"):
        cells = _cells(_row_containing(out, row))
        assert [c.strip() for c in cells] == ["", cells[1].strip(), "",
                                              "", "", "", ""]


def test_two_sessions_on_one_day_stack_in_the_same_cell():
    out = analysis.render_calendar(
        _one_week((1, [_session(name="skub"), _session(name="ben")])))

    assert _cells(_row_containing(out, "skub"))[1].strip() == "skub"
    assert _cells(_row_containing(out, "ben"))[1].strip() == "ben"


def test_a_long_name_is_truncated_to_the_cell_width():
    out = analysis.render_calendar(
        _one_week((1, [_session(name="Fuldkrop styrke og kondition")])))

    name = _cells(_row_containing(out, "Fuldkrop"))[1].strip()
    assert name == "Fuldkrop "[:9].strip()
    assert len(name) <= 9


def test_a_missing_rating_renders_as_a_dash():
    # "AC" only ever appears on the rating/place line, unlike "-" which the
    # border rows are full of.
    out = analysis.render_calendar(
        _one_week((1, [_session(rating=None, place="AC")])))

    cell = _cells(_row_containing(out, "AC"))[1].strip()
    assert cell.startswith("-") and cell.endswith("AC")


def test_the_top_of_the_scale_renders_as_six_of_six():
    # The sheet rates 1-6, so a 6 is full marks, not a 10-scale score.
    out = analysis.render_calendar(_one_week((1, [_session(rating=6)])))

    assert "6/6" in _row_containing(out, "6/6")


def test_a_ten_scale_rating_keeps_its_scale():
    out = analysis.render_calendar(_one_week((1, [_session(rating=8)])))

    assert "8/10" in _row_containing(out, "8/10")


def _week_blocks(out: str) -> list[int]:
    """Line counts of each week's block, i.e. the runs between border rows."""
    blocks, run = [], 0
    for line in out.splitlines():
        if line.strip().startswith("+"):
            if run:
                blocks.append(run)
            run = 0
        else:
            run += 1
    return blocks[1:] if blocks else blocks  # drop the header run


def test_a_quiet_week_is_not_padded_to_a_busy_weeks_height():
    quiet = [_day(date(2026, 8, 17) + timedelta(days=i)) for i in range(7)]
    quiet[0]["sessions"] = [_session()]
    busy = [_day(date(2026, 8, 24) + timedelta(days=i)) for i in range(7)]
    busy[3]["sessions"] = [_session(name="ben"), _session(name="cardio")]

    blocks = _week_blocks(analysis.render_calendar([quiet, busy]))

    assert blocks[0] < blocks[1]


def test_an_empty_week_still_renders_a_full_height_row():
    empty = analysis.render_calendar(_one_week())
    trained = analysis.render_calendar(_one_week((1, [_session()])))

    assert len(empty.splitlines()) == len(trained.splitlines())
