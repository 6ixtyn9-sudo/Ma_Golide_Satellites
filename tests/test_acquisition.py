"""Durable public-payload cache, cooldown, budget, and redaction tests."""

from __future__ import annotations

from datetime import datetime, timedelta, timezone
import json

import pytest

from ingestion.acquisition import (
    AcquisitionBudgetExceeded,
    AcquisitionClient,
    CachedProviderFailure,
)


def test_durable_cache_preserves_original_observation_and_redacts_url(tmp_path):
    now = [datetime(2026, 11, 7, 18, tzinfo=timezone.utc)]
    calls = []

    def fetch(url):
        calls.append(url)
        return {"events": [{"id": "final-1"}]}

    client = AcquisitionClient(
        fetch,
        cache_dir=str(tmp_path),
        now_fn=lambda: now[0].strftime("%Y-%m-%dT%H:%M:%SZ"),
    )
    url = "https://site.api.espn.com/scoreboard?dates=20261106&apiKey=supersecret"
    first = client.get_json(url, ttl_seconds=86400)
    assert first == {"events": [{"id": "final-1"}]}
    assert client.last_observed_utc == "2026-11-07T18:00:00Z"
    now[0] += timedelta(hours=1)
    assert client.get_json(url, ttl_seconds=86400) == first
    assert calls == [url]
    assert client.cache_hits == 1
    assert client.last_observed_utc == "2026-11-07T18:00:00Z"

    cache_files = list(tmp_path.glob("*.json"))
    assert len(cache_files) == 1
    raw = cache_files[0].read_text()
    assert "supersecret" not in raw
    assert "apiKey" not in raw
    document = json.loads(raw)
    assert document["observed_utc"] == "2026-11-07T18:00:00Z"
    telemetry = json.dumps(client.telemetry())
    assert "supersecret" not in telemetry
    assert "apiKey" not in telemetry

    now[0] += timedelta(hours=25)
    client.get_json(url, ttl_seconds=86400)
    assert calls == [url, url]


def test_negative_cache_cools_down_failures_without_persisting_messages(tmp_path):
    now = [datetime(2026, 11, 7, 18, tzinfo=timezone.utc)]
    calls = []

    def fetch(_url):
        calls.append(1)
        raise ValueError("provider returned token=never-write-this-message")

    client = AcquisitionClient(
        fetch,
        cache_dir=str(tmp_path),
        now_fn=lambda: now[0].strftime("%Y-%m-%dT%H:%M:%SZ"),
    )
    url = "https://site.api.espn.com/summary?event=1&token=never-write-this-message"
    with pytest.raises(ValueError):
        client.get_json(url)
    with pytest.raises(CachedProviderFailure):
        client.get_json(url)
    assert len(calls) == 1
    assert client.negative_cache_hits == 1
    raw = next(tmp_path.glob("*.json")).read_text()
    assert "never-write-this-message" not in raw
    assert "ValueError" in raw

    now[0] += timedelta(seconds=901)
    with pytest.raises(ValueError):
        client.get_json(url)
    assert len(calls) == 2


def test_network_budget_counts_actual_requests_not_cache_hits(tmp_path):
    calls = []

    def fetch(url):
        calls.append(url)
        return {"url": url}

    client = AcquisitionClient(fetch, cache_dir=str(tmp_path), request_budget=1)
    client.get_json("https://example.invalid/one")
    client.get_json("https://example.invalid/one")
    assert len(calls) == 1
    with pytest.raises(AcquisitionBudgetExceeded):
        client.get_json("https://example.invalid/two")
    assert client.telemetry()["network_requests"] == 1
    assert client.telemetry()["recent_attempts"][-1]["outcome"] == "budget_exceeded"


def test_fixture_reads_do_not_consume_network_budget_or_claim_live_requests():
    calls = []

    def fetch(url):
        calls.append(url)
        return {"fixture": True}

    fetch.magolide_acquisition_source = "fixture"
    client = AcquisitionClient(fetch, request_budget=1)
    assert client("https://example.invalid/one") == {"fixture": True}
    assert client("https://example.invalid/two") == {"fixture": True}
    telemetry = client.telemetry()
    assert telemetry["network_requests"] == 0
    assert telemetry["fixture_requests"] == 2
    assert [row["outcome"] for row in telemetry["recent_attempts"]] == [
        "fixture_success",
        "fixture_success",
    ]


def test_uncacheable_payload_is_served_but_never_persisted_or_reused(tmp_path):
    now = [datetime(2026, 11, 18, 6, 11, tzinfo=timezone.utc)]
    payloads = [{"state": "in"}, {"state": "post"}, {"state": "post"}]
    calls = []

    def fetch(_url):
        calls.append(1)
        return payloads[len(calls) - 1]

    client = AcquisitionClient(
        fetch,
        cache_dir=str(tmp_path),
        now_fn=lambda: now[0].strftime("%Y-%m-%dT%H:%M:%SZ"),
    )
    settled = lambda payload: payload.get("state") == "post"  # noqa: E731
    url = "https://site.api.espn.com/scoreboard?dates=20261117"

    assert client.get_json(url, cacheable=settled) == {"state": "in"}
    assert list(tmp_path.glob("*.json")) == []
    assert client.telemetry()["recent_attempts"][-1]["outcome"] == "network_success_not_cached"

    now[0] += timedelta(hours=3)
    assert client.get_json(url, cacheable=settled) == {"state": "post"}
    assert len(calls) == 2
    now[0] += timedelta(hours=3)
    assert client.get_json(url, cacheable=settled) == {"state": "post"}
    assert len(calls) == 2  # settled payload now served from cache
    assert client.unsettled_cache_skips == 1


def test_legacy_unsettled_cache_entry_is_ignored_when_predicate_given(tmp_path):
    calls = []

    def fetch(_url):
        calls.append(1)
        return {"state": "in"} if len(calls) == 1 else {"state": "post"}

    url = "https://site.api.espn.com/scoreboard?dates=20261117"
    legacy = AcquisitionClient(fetch, cache_dir=str(tmp_path))
    legacy.get_json(url)  # pre-fix behaviour: live payload cached for 24h
    assert len(list(tmp_path.glob("*.json"))) == 1

    fixed = AcquisitionClient(fetch, cache_dir=str(tmp_path))
    settled = lambda payload: payload.get("state") == "post"  # noqa: E731
    assert fixed.get_json(url, cacheable=settled) == {"state": "post"}
    assert len(calls) == 2
    assert fixed.cache_hits == 0


def test_cache_entries_with_the_retired_schema_tag_are_still_served(tmp_path):
    from ingestion.acquisition import CACHE_SCHEMA

    calls = []

    def fetch(_url):
        calls.append(1)
        return {"events": []}

    url = "https://site.api.espn.com/scoreboard?dates=20261117"
    AcquisitionClient(fetch, cache_dir=str(tmp_path)).get_json(url)
    entry = next(tmp_path.glob("*.json"))
    document = json.loads(entry.read_text())
    assert document["schema_version"] == CACHE_SCHEMA
    document["schema_version"] = CACHE_SCHEMA + ".v" + "1"
    entry.write_text(json.dumps(document))

    client = AcquisitionClient(fetch, cache_dir=str(tmp_path))
    assert client.get_json(url) == {"events": []}
    assert client.cache_hits == 1 and len(calls) == 1
