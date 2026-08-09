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

# Known gym names to detect in a note, in priority order. Output is canonical.
_PLACES = ("AC", "Center", "Tryg")

# Danish "with" marker (m / m. / med) followed by a capitalized name.
_ENSAMBLE_RE = re.compile(r"\b(?:med|m\.?)\s+([A-ZÆØÅ][\wÆØÅæøå]+)")


def _map_type(title: str) -> str:
    t = (title or "").strip()
    return _TYPE_MAP.get(t.lower(), t)


def _parse_rating(description):
    if not description:
        return None
    # "4/5" or "8/10"
    m = re.search(r"\b(\d+)\s*/\s*(5|10)\b", description)
    if m:
        return int(m.group(1))
    # Standalone digit 1-10 at start of string ("4 god træning")
    m = re.match(r"^\s*(\d{1,2})\b", description)
    if m and 1 <= int(m.group(1)) <= 10:
        return int(m.group(1))
    # Standalone digit 1-10 at end of string ("god træning 4")
    m = re.search(r"\b(\d{1,2})\s*$", description)
    if m and 1 <= int(m.group(1)) <= 10:
        return int(m.group(1))
    return None


def _detect_place(description: str) -> str:
    if not description:
        return ""
    for p in _PLACES:
        if re.search(rf"\b{re.escape(p)}\b", description, re.IGNORECASE):
            return p
    return ""


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
        "Place": _detect_place(description),
        "Type": _map_type(workout.get("title")),
        "Time": time,
        "Mave": 1 if _has_abs(exercises) else 0,
        "AddCardio": add_cardio,
        "AddCardio2": 0,
        "Ensamble": _detect_ensamble(description),
        "Rating": _parse_rating(description),
        "Claude": 1,
        "Comments": description,
    }
