import re
from datetime import datetime

# Exercise title keywords that indicate an abs/core movement.
_ABS_KEYWORDS = (
    "crunch", "plank", "sit-up", "sit up", "situp",
    "leg raise", "ab wheel", "ab rollout", "ab roller",
    "hanging knee", "hanging leg", "cable crunch",
    "russian twist", "toe touch", "bicycle crunch",
    "dead bug", "flutter kick", "hollow body",
    "v-up", "v up", "mave", "core",
)

# Danish Hevy workout titles -> the English categories used in the sheet.
_TYPE_MAP = {
    "over": "upper",
    "ben": "legs",
    "træk": "pull",
    "skub": "push",
}

# Gym names with a fixed spelling in the sheet. Anything else is kept verbatim.
_PLACES = ("AC", "Center", "Tryg")

# Danish "with" marker (m / m. / med) followed by a capitalized name.
_ENSAMBLE_RE = re.compile(r"\b(?:med|m\.?)\s+([A-ZÆØÅ][\wÆØÅæøå]+)")

# The anchors are the point: field 1 must be a rating and nothing else. Without
# them, prose that happens to contain a comma would put its tail in Place.
_RATING_FIELD = re.compile(r"^(\d{1,2})(?:\s*/\s*(?:5|6|10))?$")


def _map_type(title: str) -> str:
    t = (title or "").strip()
    return _TYPE_MAP.get(t.lower(), t)


def _canonical_place(text: str) -> str:
    """Known gyms get their sheet spelling; anything else is kept as typed."""
    for p in _PLACES:
        if text.lower() == p.lower():
            return p
    return text


def _parse_note(description: str) -> tuple:
    """Split a note written as "rating, place, comment" into its three parts.

    The convention only engages when the first field is a bare rating. Any other
    note — no commas, or prose that merely contains one — is a comment in full,
    which is what stops "god træning, lidt træt" from naming a gym.
    """
    if not description:
        return None, "", ""

    fields = [f.strip() for f in description.split(",")]
    m = _RATING_FIELD.match(fields[0])
    if not m or not 1 <= int(m.group(1)) <= 10:
        return None, "", description

    rating = int(m.group(1))
    place = _canonical_place(fields[1]) if len(fields) > 1 else ""
    # Rejoined, so a comma inside the comment survives the split.
    comment = ", ".join(fields[2:])
    return rating, place, comment


def _detect_ensamble(description: str) -> int:
    if not description:
        return 0
    return 1 if _ENSAMBLE_RE.search(description) else 0


def _has_abs(exercises: list[dict]) -> bool:
    for ex in exercises:
        title = ex.get("title", "").lower()
        if any(kw in title for kw in _ABS_KEYWORDS):
            return True
    return False


def _is_cardio_exercise(ex: dict) -> bool:
    title = ex.get("title", "").lower()
    if any(kw in title for kw in _ABS_KEYWORDS):
        return False  # abs/core -> Mave, never cardio (even when timed)
    for s in ex.get("sets", []):
        timed = s.get("duration_seconds") is not None or s.get("distance_meters") is not None
        no_load = s.get("reps") is None and s.get("weight_kg") is None
        if timed and no_load:
            return True
    return False


def _cardio_minutes(exercises: list[dict]) -> int:
    total_sec = 0
    for ex in exercises:
        if _is_cardio_exercise(ex):
            for s in ex.get("sets", []):
                total_sec += s.get("duration_seconds") or 0
    return round(total_sec / 60)


def workout_to_row(workout: dict) -> dict:
    start = datetime.fromisoformat(workout["start_time"].replace("Z", "+00:00"))
    end = datetime.fromisoformat(workout["end_time"].replace("Z", "+00:00"))
    elapsed = max(1, round((end - start).total_seconds() / 60))

    exercises = workout.get("exercises", [])
    description = (workout.get("description") or "").strip()
    rating, place, comment = _parse_note(description)

    has_cardio = any(_is_cardio_exercise(ex) for ex in exercises)
    has_strength = any(not _is_cardio_exercise(ex) for ex in exercises)

    if has_cardio and has_strength:
        add_cardio = _cardio_minutes(exercises)
        time = max(1, elapsed - add_cardio)
    else:
        add_cardio = 0
        time = elapsed

    return {
        "Date": start.date(),
        "Place": place,
        "Type": _map_type(workout.get("title")),
        "Time": time,
        "Mave": 1 if _has_abs(exercises) else 0,
        "AddCardio": add_cardio,
        "AddCardio2": 0,
        "Ensamble": _detect_ensamble(description),
        "Rating": rating,
        "Claude": 1,
        "Comments": comment,
    }
