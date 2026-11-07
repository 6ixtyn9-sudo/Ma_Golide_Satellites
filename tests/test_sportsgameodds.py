"""Tests for ingestion/providers/sportsgameodds.py.

Fixtures under tests/fixtures/sportsgameodds/ are DOCUMENTATION_ONLY
synthetic responses authored from the documented SportsGameOdds getEvents
schema (see each file's "_note"); they are NOT captured API responses and no
coverage claim is live-verified. No network is used: a fake
``fetch_factory(provider)`` serves the fixtures.

Binding semantics proven here:
- Documented oddIDs map to exactly their own periods ('-1q-'→Q1 …
  '-4q-'→Q4, '-game-'→FULL_GAME); Q2/Q3 absence is never backfilled
  (partial quarters are allowed).
- The total ``line`` comes from a line/points field; a market without it
  yields ``line=None`` (NOT usable for total forecasts) and smoke counts it
  under "not_usable".
- American prices converted via books.american_to_decimal; books come from
  ``bookOdds`` keys.
- Halves/OT/other bet types are never relabelled as quarters.
- Live status → pre_match=False; books.py pre-match filters exclude it.
- EuroLeague raises ValueError (coverage unproven, never fabricated).
- smoke() captures /account/usage under quota["usage"].
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Callable, Dict, List, Tuple

import pytest

from ingestion import books
from ingestion.models import Market, Period, payload_hash
from ingestion.providers import sportsgameodds

FIXTURES = Path(__file__).parent / "fixtures" / "sportsgameodds"
API_KEY = "sgo-test-key-not-real"
OBSERVED = "2026-11-06T22:00:00Z"
EVENTS_KEY = "/events?leagueID=NBA"


def _load(name: str) -> Any:
    return json.loads((FIXTURES / name).read_text(encoding="utf-8"))


def _factory(
    mapping: Dict[str, Any]
) -> Tuple[Callable[[str], Callable[[str], Any]], List[str]]:
    """fetch_factory(provider) -> FetchFn serving fixtures by URL substring."""
    calls: List[str] = []

    def fetch(url: str) -> Any:
        calls.append(url)
        matches = [key for key in mapping if key in url]
        assert matches, f"no fixture mapped for URL: {url}"
        value = mapping[max(matches, key=len)]
        if isinstance(value, tuple):
            payload, headers = value
            return {"json": payload, "headers": headers}
        return value

    def factory(provider: str) -> Callable[[str], Any]:
        assert provider == sportsgameodds.PROVIDER
        return fetch

    return factory, calls


@pytest.fixture(autouse=True)
def _api_key(monkeypatch):
    monkeypatch.setenv(sportsgameodds.ENV_KEY, API_KEY)
    yield


def _default_mapping() -> Dict[str, Any]:
    return {
        EVENTS_KEY: _load("events_nba.json"),
        "/account/usage": _load("usage.json"),
    }


def _totals(snapshots, period: Period, book: str = "draftkings"):
    return [
        s
        for s in snapshots
        if s.market is Market.TOTAL and s.period is period and s.book == book
    ]


def _all_totals(snapshots, period: Period):
    return [s for s in snapshots if s.market is Market.TOTAL and s.period is period]


# ---------------------------------------------------------------------------
# Enabled / disabled
# ---------------------------------------------------------------------------


def test_disabled_without_key(monkeypatch):
    """Missing SPORTSGAMEODDS_API_KEY: fetch_odds returns [] and smoke
    reports DOCUMENTATION_ONLY — the keyless pipeline must never raise."""
    monkeypatch.delenv(sportsgameodds.ENV_KEY, raising=False)
    factory, calls = _factory(_default_mapping())
    assert sportsgameodds.fetch_odds("NBA", "20261107", factory) == []
    result = sportsgameodds.smoke("NBA", factory)
    assert result["enabled"] is False
    assert result["verification"] == "DOCUMENTATION_ONLY"
    assert calls == []


# ---------------------------------------------------------------------------
# oddID -> period mapping
# ---------------------------------------------------------------------------


def test_maps_game_and_quarter_oddids():
    """points-all-game-ou-* → FULL_GAME 232.5; points-all-1q-ou-* → Q1 58.5
    with over+under; points-all-4q-ou-over → Q4 56.5 over only (no under
    side exists); Q2/Q3 are absent from the payload and from the output."""
    factory, _calls = _factory(_default_mapping())
    snapshots = sportsgameodds.fetch_odds("NBA", "20261107", factory)
    assert snapshots

    full = _totals(snapshots, Period.FULL_GAME)
    assert len(full) == 1
    assert full[0].line == pytest.approx(232.5)
    assert full[0].line > 0
    assert full[0].line != full[0].over_price_decimal  # line, not a price
    assert full[0].line != full[0].under_price_decimal

    q1 = _totals(snapshots, Period.Q1)
    assert len(q1) == 1
    assert q1[0].line == pytest.approx(58.5)
    assert q1[0].over_price_decimal is not None
    assert q1[0].under_price_decimal is not None

    q4 = _totals(snapshots, Period.Q4)
    assert len(q4) == 1
    assert q4[0].line == pytest.approx(56.5)
    assert q4[0].over_price_decimal is not None
    assert q4[0].under_price_decimal is None  # over side only in the payload

    # No backfill: Q2/Q3 never appear because the payload lacks them.
    periods_seen = {s.period for s in snapshots}
    assert Period.Q2 not in periods_seen
    assert Period.Q3 not in periods_seen
    assert all(s.provider == "sportsgameodds" for s in snapshots)
    assert all(s.pre_match is True for s in snapshots)  # status "upcoming"


def test_unmapped_period_oddids_skipped_with_gap():
    """Halves ('1h'), a fifth quarter and moneyline oddIDs are NOT totals in
    the documented grammar: skipped with gaps, never relabelled as quarters
    (or anything else)."""
    payload = {
        "data": [
            {
                "eventID": "sgo-evt-x",
                "status": "upcoming",
                "markets": {
                    "points-all-1h-ou-over": {
                        "line": 116.5,
                        "bookOdds": {"draftkings": {"price": -110}},
                    },
                    "points-all-1h-ou-under": {
                        "line": 116.5,
                        "bookOdds": {"draftkings": {"price": -110}},
                    },
                    "points-home-game-ml-home": {
                        "bookOdds": {"draftkings": {"price": -108}}
                    },
                },
            }
        ]
    }
    factory, _calls = _factory({EVENTS_KEY: payload})
    snapshots = sportsgameodds.fetch_odds("NBA", "20261107", factory)
    assert snapshots == []
    assert not any(
        s.period in (Period.Q1, Period.Q2, Period.Q3, Period.Q4, Period.H1)
        for s in snapshots
    )
    gaps = sportsgameodds.last_gaps()
    assert any(g.required_field == "oddid" and "1h" in g.detail for g in gaps)
    assert any("ml-home" in g.detail for g in gaps)


# ---------------------------------------------------------------------------
# Line semantics
# ---------------------------------------------------------------------------


def test_no_line_means_unusable_total():
    """Market present but no line/points field: the snapshot is emitted with
    line=None, is NOT usable for total forecasts (the posted line is required),
    and smoke counts it under "not_usable"."""
    factory, _calls = _factory(_default_mapping())
    snapshots = sportsgameodds.fetch_odds("NBA", "20261107", factory)
    assert snapshots  # sanity: events_nba.json carries proper lines

    factory, calls = _factory(
        {EVENTS_KEY: _load("events_no_line.json"), "/account/usage": _load("usage.json")}
    )
    snapshots = sportsgameodds.fetch_odds("NBA", "20261107", factory)
    assert len(snapshots) == 1
    assert snapshots[0].market is Market.TOTAL
    assert snapshots[0].line is None
    assert snapshots[0].over_price_decimal == books.american_to_decimal(-110)

    result = sportsgameodds.smoke("NBA", factory, date_yyyymmdd="20261107")
    assert result["not_usable"] == 1
    assert result["full_game_totals_present"] == 0
    gaps = sportsgameodds.last_gaps()
    assert any(g.required_field == "total_line_full_game" for g in gaps)


def test_partial_quarters_allowed():
    """Coverage is reported per quarter separately: Q1 1/1, Q2 0, Q3 0,
    Q4 1/1 — and absence of Q2/Q3 snapshots is the no-backfill assertion."""
    factory, calls = _factory(_default_mapping())
    result = sportsgameodds.smoke("NBA", factory, date_yyyymmdd="20261107")
    assert result["verification"] == "DOCUMENTATION_ONLY"
    assert result["events_checked"] == 1
    assert result["quarters"] == {
        "Q1": {"present": 1, "total": 1},
        "Q2": {"present": 0, "total": 1},
        "Q3": {"present": 0, "total": 1},
        "Q4": {"present": 1, "total": 1},
    }
    assert result["full_game_totals_present"] == 1
    assert result["pre_match_vs_live"] == {"pre_match": 1, "not_pre_match": 0}
    assert any("/account/usage" in u for u in calls)

    # And fetch_odds itself emits no Q2/Q3 snapshot to backfill with.
    snapshots = sportsgameodds.fetch_odds("NBA", "20261107", factory)
    assert not any(s.period in (Period.Q2, Period.Q3) for s in snapshots)


# ---------------------------------------------------------------------------
# Prices / books
# ---------------------------------------------------------------------------


def test_price_conversion_and_book_selection():
    """bookOdds draftkings -105 → decimal 1 + 100/105; the book slug is set
    on the snapshot; provider is sportsgameodds."""
    factory, _calls = _factory(_default_mapping())
    snapshots = sportsgameodds.fetch_odds("NBA", "20261107", factory)
    q1 = _totals(snapshots, Period.Q1)[0]
    assert q1.book == "draftkings"
    assert q1.provider == "sportsgameodds"
    assert q1.over_price_decimal == pytest.approx(1.9523809523809523, abs=1e-9)
    assert q1.over_price_decimal == books.american_to_decimal(-105)
    assert q1.under_price_decimal == books.american_to_decimal(-115)
    assert q1.market is Market.TOTAL

    # Books come from bookOdds keys: full-game has draftkings + fanduel.
    assert {s.book for s in _all_totals(snapshots, Period.FULL_GAME)} == {
        "draftkings",
        "fanduel",
    }


# ---------------------------------------------------------------------------
# Pre-match semantics
# ---------------------------------------------------------------------------


def test_live_status_not_pre_match():
    """status 'in-progress' → pre_match=False (never mark live as
    pre-match); books.select_total_snapshots(pre_match_only=True) excludes
    the snapshots."""
    factory, _calls = _factory({EVENTS_KEY: _load("events_live.json")})
    snapshots = sportsgameodds.fetch_odds("NBA", "20261107", factory)
    assert snapshots
    assert all(s.pre_match is False for s in snapshots)

    config = books.BookConfig(provider="sportsgameodds")  # pre_match_only=True
    assert books.select_total_snapshots(snapshots, config, Period.FULL_GAME) == []
    assert books.select_moneyline_snapshots(snapshots, config) == []
    allow_live = books.BookConfig(provider="sportsgameodds", pre_match_only=False)
    assert books.select_total_snapshots(snapshots, allow_live, Period.FULL_GAME)


# ---------------------------------------------------------------------------
# League gate / provenance / usage quota
# ---------------------------------------------------------------------------


def test_euroleague_not_enabled():
    """EuroLeague coverage is unproven: raise instead of fabricating lines."""
    factory, _calls = _factory(_default_mapping())
    with pytest.raises(ValueError, match="not enabled for this league yet"):
        sportsgameodds.fetch_odds("EUROLEAGUE", "20261107", factory)
    with pytest.raises(ValueError, match="not enabled for this league yet"):
        sportsgameodds.smoke("EuroLeague", factory)


def test_observed_utc_param_and_payload_hash():
    """observed_utc comes from the injectable now_utc param; payload_hash
    hashes the RAW event dict."""
    factory, _calls = _factory(_default_mapping())
    snapshots = sportsgameodds.fetch_odds("NBA", "20261107", factory, now_utc=OBSERVED)
    assert snapshots
    assert all(s.observed_utc == OBSERVED for s in snapshots)
    event = _load("events_nba.json")["data"][0]
    assert all(s.payload_hash == payload_hash(event) for s in snapshots)
    assert all(s.main_line is True for s in snapshots)  # documented main oddIDs


def test_smoke_reports_usage_quota():
    """smoke() fetches GET /account/usage and reports the payload verbatim
    under quota['usage'] (free tier: 2,500 objects / rolling 30 days)."""
    factory, calls = _factory(_default_mapping())
    result = sportsgameodds.smoke("NBA", factory, date_yyyymmdd="20261107")
    assert result["enabled"] is True
    assert result["quota"]["usage"]["objectsRemaining"] == 2380
    assert result["quota"]["usage"]["objectsUsed"] == 120
    assert any("/account/usage" in u for u in calls)
