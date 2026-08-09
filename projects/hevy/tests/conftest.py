import pytest


def strength_ex(title="Bench Press (Barbell)", reps=8, weight_kg=60):
    return {"title": title, "sets": [
        {"reps": reps, "weight_kg": weight_kg,
         "duration_seconds": None, "distance_meters": None}]}


def cardio_ex(title="Treadmill", duration_seconds=900, distance_meters=0):
    return {"title": title, "sets": [
        {"reps": None, "weight_kg": None,
         "duration_seconds": duration_seconds, "distance_meters": distance_meters}]}


def abs_ex(title="Plank", duration_seconds=60):
    # A timed, weightless core move: must count as Mave, never cardio.
    return {"title": title, "sets": [
        {"reps": None, "weight_kg": None,
         "duration_seconds": duration_seconds, "distance_meters": None}]}


@pytest.fixture
def make_workout():
    def _make(title="push", description="", exercises=None,
              start="2026-07-04T11:00:00+00:00", end="2026-07-04T12:00:00+00:00"):
        return {
            "title": title,
            "description": description,
            "start_time": start,
            "end_time": end,
            "exercises": exercises if exercises is not None else [strength_ex()],
        }
    return _make
