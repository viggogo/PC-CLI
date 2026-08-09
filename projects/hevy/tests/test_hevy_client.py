import asyncio

from hevy_mcp import hevy_client


class _FakeResp:
    def __init__(self, payload):
        self._payload = payload

    def raise_for_status(self):
        pass

    def json(self):
        return self._payload


class _FakeClient:
    """Stand-in for httpx.AsyncClient that serves canned pages."""

    def __init__(self, pages):
        self._pages = pages

    async def __aenter__(self):
        return self

    async def __aexit__(self, *exc):
        return False

    async def get(self, url, headers=None, params=None):
        page = params["page"]
        return _FakeResp(self._pages[page - 1])


def _w(dt):
    return {"start_time": f"{dt}T10:00:00Z", "end_time": f"{dt}T11:00:00Z"}


def test_pagination_spans_all_pages(monkeypatch):
    # page_count=2: newest page then an older page. Bug read "pageCount" and
    # stopped after page 1; the fix reads "page_count".
    pages = [
        {"page": 1, "page_count": 2, "workouts": [_w("2026-06-20"), _w("2026-06-18")]},
        {"page": 2, "page_count": 2, "workouts": [_w("2026-06-16"), _w("2026-06-14")]},
    ]
    monkeypatch.setattr(hevy_client.httpx, "AsyncClient", lambda *a, **k: _FakeClient(pages))
    monkeypatch.setenv("HEVY_API_KEY", "x")

    out = asyncio.run(hevy_client.fetch_workouts_since(None))
    assert [w["start_time"][:10] for w in out] == [
        "2026-06-20", "2026-06-18", "2026-06-16", "2026-06-14"]
