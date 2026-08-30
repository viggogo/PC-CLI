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


def test_parse_note_splits_rating_place_comment():
    assert cm._parse_note("4, AC, godt tempo") == (4, "AC", "godt tempo")
    # An unknown gym is kept exactly as typed, not blanked.
    assert cm._parse_note("4, sats, tungt i dag") == (4, "sats", "tungt i dag")
    assert isinstance(cm._parse_note("4, AC, godt tempo")[0], int)


def test_parse_note_canonicalises_known_places():
    assert cm._parse_note("3, ac, fint")[1] == "AC"
    assert cm._parse_note("3, center, fint")[1] == "Center"
    assert cm._parse_note("3, TRYG, fint")[1] == "Tryg"


def test_parse_note_partial_fields():
    assert cm._parse_note("4, AC") == (4, "AC", "")
    assert cm._parse_note("4") == (4, "", "")
    assert cm._parse_note("4/5") == (4, "", "")
    assert cm._parse_note("8/10") == (8, "", "")
    assert cm._parse_note("") == (None, "", "")


def test_parse_note_keeps_commas_inside_the_comment():
    assert cm._parse_note("4, AC, godt, men træt") == (4, "AC", "godt, men træt")


def test_parse_note_without_a_leading_rating_is_all_comment():
    # Prose with a comma must never name a gym.
    assert cm._parse_note("god træning, lidt træt") == (
        None, "", "god træning, lidt træt")
    assert cm._parse_note("Tryg m Peter") == (None, "", "Tryg m Peter")
    assert cm._parse_note("lynild") == (None, "", "lynild")
    assert cm._parse_note("4 m Silas, AC") == (None, "", "4 m Silas, AC")
    assert cm._parse_note("0, AC") == (None, "", "0, AC")   # out of the 1-10 range


def test_canonical_place():
    assert cm._canonical_place("ac") == "AC"
    assert cm._canonical_place("Tryg") == "Tryg"
    assert cm._canonical_place("hjemme") == "hjemme"
    assert cm._canonical_place("") == ""


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
    w = make_workout(title="Skub", description="4, AC, m Silas",
                     exercises=[strength_ex()])
    row = cm.workout_to_row(w)
    assert row["Date"] == date(2026, 7, 4)
    assert row["Type"] == "push"
    assert row["Rating"] == 4
    assert isinstance(row["Rating"], int)
    assert row["Ensamble"] == 1
    assert row["Place"] == "AC"
    assert row["Comments"] == "m Silas"
    assert set(row.keys()) == {
        "Date", "Place", "Type", "Time", "Mave", "AddCardio",
        "AddCardio2", "Ensamble", "Rating", "Claude", "Comments"}
