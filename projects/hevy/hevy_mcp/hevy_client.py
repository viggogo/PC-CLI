import os
from datetime import datetime
from typing import Optional

import httpx

HEVY_BASE = "https://api.hevyapp.com/v1"


def _headers() -> dict:
    return {"api-key": os.environ["HEVY_API_KEY"]}


def _parse_dt(s: str) -> datetime:
    return datetime.fromisoformat(s.replace("Z", "+00:00"))


async def fetch_workouts_since(since: Optional[datetime] = None) -> list[dict]:
    """
    Return all workouts newer than `since` (exclusive by date).
    Paginates through the Hevy API (newest-first) and stops early once
    workouts older than `since` are encountered.
    """
    workouts: list[dict] = []
    page = 1

    async with httpx.AsyncClient(timeout=30) as client:
        while True:
            resp = await client.get(
                f"{HEVY_BASE}/workouts",
                headers=_headers(),
                params={"page": page, "pageSize": 10},
            )
            resp.raise_for_status()
            data = resp.json()
            batch: list[dict] = data.get("workouts", [])

            if not batch:
                break

            for w in batch:
                start = _parse_dt(w["start_time"])
                if since is not None and start.date() <= since.date():
                    return workouts
                workouts.append(w)

            # The Hevy API returns the total page count as "page_count".
            page_count = data.get("page_count") or data.get("pageCount") or 1
            if page >= page_count:
                break
            page += 1

    return workouts


async def fetch_workout_by_id(workout_id: str) -> dict:
    async with httpx.AsyncClient(timeout=30) as client:
        resp = await client.get(
            f"{HEVY_BASE}/workouts/{workout_id}",
            headers=_headers(),
        )
        resp.raise_for_status()
        return resp.json()
