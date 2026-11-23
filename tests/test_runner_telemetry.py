"""The summary must describe what the run did, not what the meter noticed.

Run 45218527700 reported ``network_requests: 0`` and ``provider_health: []``
for lanes whose adapters had read their hosts: providers given ``fetch=None``
build their own ``HttpClient``, which the pipeline's injected meter never sees,
and ``_report()`` never carried the health rows it had already stored. Both
readings said "this lane asked for nothing", which was false.
"""

from __future__ import annotations

from typing import Any, Dict

import pytest

from ingestion import http
from ingestion.http import HttpClient
from ingestion.models import CoverageState, League, ProviderHealth
from ingestion.pipeline import Pipeline, PipelineConfig
from ingestion.store import InMemoryStore


class _Resp:
    def __init__(self, body: str = "{}", status: int = 200) -> None:
        self.status_code = status
        self.headers: Dict[str, str] = {}
        self.text = body
        self.content = body.encode("utf-8")


class _Session:
    def __init__(self, status: int = 200) -> None:
        self.status = status
        self.calls = 0

    def get(self, url: str, headers: Any = None, timeout: Any = None) -> _Resp:
        self.calls += 1
        return _Resp(status=self.status)


def test_the_transport_meter_counts_reads_the_injected_meter_never_sees():
    http.reset_transport_meter()
    client = HttpClient(session=_Session(), retries=0)

    client.get_text("https://example.test/a")
    client.get_text("https://example.test/b")

    meter = http.transport_meter()
    assert meter["attempts"] == 2
    assert meter["responses"] == 2
    assert meter["errors"] == 0


def test_a_permanent_failure_is_counted_as_a_response_not_a_hidden_failure():
    http.reset_transport_meter()
    client = HttpClient(session=_Session(status=404), retries=0)

    with pytest.raises(http.HttpError):
        client.get_text("https://example.test/missing")

    meter = http.transport_meter()
    assert meter["attempts"] == 1
    assert meter["responses"] == 1


def test_the_report_names_both_meters():
    pipeline = Pipeline(
        PipelineConfig(league="JAPAN_B_LEAGUE", date_yyyymmdd="20261124"),
        InMemoryStore(),
    )
    report = pipeline._report("write", 0, {})

    acquisition = report["acquisition"]
    assert acquisition["network_requests"] == 0  # injected fetch: never used
    assert "transport_attempts" in acquisition
    assert "meter" in acquisition


def test_health_written_by_a_lane_reaches_the_printed_summary():
    pipeline = Pipeline(
        PipelineConfig(league="JAPAN_B_LEAGUE", date_yyyymmdd="20261124"),
        InMemoryStore(),
    )
    pipeline._record_health(
        ProviderHealth(
            provider="bleague",
            league=League.JAPAN_B_LEAGUE,
            last_attempt_utc="2026-11-23T18:37:00Z",
            state=CoverageState.MISSING,
            coverage={"games": 0, "slate_games": 10, "off_date_dropped": 10},
            detail="url=https://www.bleague.jp/x requests=4",
        )
    )

    report = pipeline._report("write", 0, {})

    assert report["provider_health"] == [
        {
            "provider": "bleague",
            "league": "JAPAN_B_LEAGUE",
            "last_attempt_utc": "2026-11-23T18:37:00Z",
            "last_success_utc": None,
            "state": "missing",
            "coverage": {"games": 0, "slate_games": 10, "off_date_dropped": 10},
            "quota": {},
            "detail": "url=https://www.bleague.jp/x requests=4",
        }
    ]
    # ... and the store still holds it, so the two never diverge.
    stored = pipeline.store.get_provider_health("bleague", League.JAPAN_B_LEAGUE)
    assert stored is not None and stored.coverage["off_date_dropped"] == 10
