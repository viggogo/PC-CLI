from hevy_mcp import column_mapper as cm


def test_map_type_danish_titles():
    assert cm._map_type("Over") == "upper"
    assert cm._map_type("Ben") == "legs"
    assert cm._map_type("Træk") == "pull"
    assert cm._map_type("Skub") == "push"


def test_map_type_english_passthrough():
    assert cm._map_type("push") == "push"
    assert cm._map_type("full body") == "full body"


def test_map_type_unknown_is_verbatim():
    assert cm._map_type("Crossfit WOD") == "Crossfit WOD"


def test_parse_rating_returns_int():
    assert cm._parse_rating("3") == 3
    assert cm._parse_rating("4 tryg") == 4
    assert cm._parse_rating("2, travlt") == 2
    assert cm._parse_rating("god træning 5") == 5
    assert cm._parse_rating("") is None
    assert cm._parse_rating("lynild") is None
    assert isinstance(cm._parse_rating("4"), int)


def test_detect_place():
    assert cm._detect_place("4, AC") == "AC"
    assert cm._detect_place("Tryg m Peter") == "Tryg"
    assert cm._detect_place("god center dag") == "Center"
    assert cm._detect_place("travlt") == ""
    assert cm._detect_place("") == ""


def test_detect_ensamble():
    assert cm._detect_ensamble("m Silas") == 1
    assert cm._detect_ensamble("med Uwufufu") == 1
    assert cm._detect_ensamble("Tryg m Peter") == 1
    assert cm._detect_ensamble("2, travlt") == 0
    assert cm._detect_ensamble("") == 0


from datetime import date
from tests.conftest import strength_ex, cardio_ex, abs_ex


def test_cardio_detection_and_split(make_workout):
    w = make_workout(
        title="Over", description="3",
        exercises=[strength_ex(), cardio_ex(duration_seconds=900)],
        start="2026-07-04T11:26:22+00:00", end="2026-07-04T12:59:52+00:00",
    )  # elapsed ~= 94 min, treadmill 15 min
    row = cm.workout_to_row(w)
    assert row["AddCardio"] == 15
    assert row["Time"] == 94 - 15
    assert row["AddCardio2"] == 0
    assert row["Type"] == "upper"
    assert row["Claude"] == 1
    assert "Total" not in row


def test_cardio_only_session_keeps_time(make_workout):
    w = make_workout(
        title="cardio", description="",
        exercises=[cardio_ex(duration_seconds=1800)],
        start="2026-07-04T11:00:00+00:00", end="2026-07-04T11:30:00+00:00",
    )  # elapsed 30, all cardio
    row = cm.workout_to_row(w)
    assert row["Time"] == 30
    assert row["AddCardio"] == 0


def test_abs_is_mave_not_cardio(make_workout):
    w = make_workout(
        title="push", exercises=[strength_ex(), abs_ex()],
        start="2026-07-04T11:00:00+00:00", end="2026-07-04T12:00:00+00:00",
    )
    row = cm.workout_to_row(w)
    assert row["Mave"] == 1
    assert row["AddCardio"] == 0   # plank must NOT be counted as cardio
    assert row["Time"] == 60


def test_row_shape_and_types(make_workout):
    w = make_workout(title="Skub", description="4 m Silas, AC",
                     exercises=[strength_ex()])
    row = cm.workout_to_row(w)
    assert row["Date"] == date(2026, 7, 4)
    assert row["Type"] == "push"
    assert row["Rating"] == 4
    assert isinstance(row["Rating"], int)
    assert row["Ensamble"] == 1
    assert row["Place"] == "AC"
    assert row["Comments"] == "4 m Silas, AC"
    assert set(row.keys()) == {
        "Date", "Place", "Type", "Time", "Mave", "AddCardio",
        "AddCardio2", "Ensamble", "Rating", "Claude", "Comments"}
