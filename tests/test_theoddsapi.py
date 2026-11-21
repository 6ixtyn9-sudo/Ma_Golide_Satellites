"""Tests for ingestion/providers/theoddsapi.py.

Fixtures under tests/fixtures/theoddsapi/ are DOCUMENTATION_ONLY synthetic
responses authored from The Odds API v4 documented schemas (see each file's
"_note"; the top-level ARRAY fixture odds_list_full_game.json carries its
synthetic note here and in the quarter fixture's "_note"). They are NOT
captured API responses and no coverage claim is live-verified. No network is
used: a fake ``fetch_factory(provider)`` serves the fixtures.

Binding semantics proven here:
- h2h outcomes map to home/away by NAME match against home_team/away_team
  (never array order); an outcome matching neither team is skipped with a
  gap, never guessed.
- Routine acquisition returns only full-game h2h/totals/spreads; period payloads
  remain parseable but are never fetched automatically by the free-tier path.
- Decimal prices are stored AS-IS — oddsFormat=decimal is requested and
  american_to_decimal must NOT be applied.
- The total ``line`` is the outcome "point", never a price.
- ``book`` is the bookmaker KEY.
- Quota: parse_quota normalises x-requests-remaining/used; smoke reports it
  via the documented fetch envelope convention.
- EuroLeague raises ValueError (no verified sport key; never fabricated).
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Callable, Dict, List, Tuple

import pytest

from ingestion import books
from ingestion.models import Market, Period
from ingestion.providers import theoddsapi

FIXTURES = Path(__file__).parent / "fixtures" / "theoddsapi"
API_KEY = "oddsapi-test-key-not-real"
OBSERVED = "2026-11-06T22:00:00Z"
FUTURE_DATE = "20991107"  # fixture commence_time 2099-11-07: pre-match for any clock
LIST_KEY = "/basketball_nba/odds/?"
PER_EVENT_KEY = "events/odds-evt-1/odds"


def _load(name: str) -> Any:
    return json.loads((FIXTURES / name).read_text(encoding="utf-8"))


def _factory(mapping: Dict[str, Any]) -> Tuple[Callable[[str], Callable[[str], Any]], List[str]]:
    """fetch_factory(provider) -> FetchFn serving fixtures by URL substring.

    A ``(payload, headers)`` mapping value is served through the documented
    envelope convention ``{"json": ..., "headers": ...}`` so quota capture
    can be exercised.
    """
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
        assert provider == theoddsapi.PROVIDER
        return fetch

    return factory, calls


def _default_mapping() -> Dict[str, Any]:
    return {
        LIST_KEY: _load("odds_list_full_game.json"),
        PER_EVENT_KEY: _load("odds_event_quarters.json"),
    }


@pytest.fixture(autouse=True)
def _api_key(monkeypatch):
    monkeypatch.setenv(theoddsapi.ENV_KEY, API_KEY)
    yield


def _fetch(factory):
    return theoddsapi.fetch_odds("NBA", FUTURE_DATE, factory)


# ---------------------------------------------------------------------------
# Enabled / disabled
# ---------------------------------------------------------------------------


def test_disabled_without_key(monkeypatch):
    """Missing THE_ODDS_API_KEY: fetch_odds returns [] and smoke reports
    DOCUMENTATION_ONLY — the keyless pipeline must never raise or fetch."""
    monkeypatch.delenv(theoddsapi.ENV_KEY, raising=False)
    factory, calls = _factory(_default_mapping())
    assert theoddsapi.fetch_odds("NBA", FUTURE_DATE, factory) == []
    result = theoddsapi.smoke("NBA", factory)
    assert result["enabled"] is False
    assert result["verification"] == "DOCUMENTATION_ONLY"
    assert calls == []


# ---------------------------------------------------------------------------
# h2h / totals mapping
# ---------------------------------------------------------------------------


def test_h2h_and_totals_mapping():
    """h2h outcomes arrive away-first but must map by NAME (never order);
    totals Outcomes carry point 232.5 as the LINE VALUE; book is the
    bookmaker key 'draftkings' (not the display title)."""
    factory, calls = _factory(_default_mapping())
    snapshots = _fetch(factory)
    assert snapshots

    moneyline = next(s for s in snapshots if s.market is Market.MONEYLINE)
    assert moneyline.provider == "theoddsapi"
    assert moneyline.game_id == "odds-evt-1"
    assert moneyline.book == "draftkings"
    assert moneyline.period is Period.FULL_GAME
    assert moneyline.pre_match is True  # commence_time 2099-11-07 is future
    assert moneyline.home_price_decimal == 1.85
    assert moneyline.away_price_decimal == 2.05

    total = next(s for s in snapshots if s.market is Market.TOTAL and s.period is Period.FULL_GAME)
    assert total.book == "draftkings"
    assert total.line == pytest.approx(232.5)  # the point, never a price
    assert total.line != total.over_price_decimal != total.under_price_decimal
    assert total.over_price_decimal == 1.9
    assert total.under_price_decimal == 1.92

    # Off-date event in the list is skipped (date narrowing on commence_time).
    assert not any(s.game_id == "odds-evt-offdate" for s in snapshots)
    assert any(g.required_field == "date_window" for g in theoddsapi.last_gaps())
    assert len(calls) == 1  # routine permits never fan out into per-event calls


def test_decimal_prices_stored_as_is():
    """oddsFormat=decimal is requested: 1.85 stays exactly 1.85 — applying
    american_to_decimal would corrupt already-decimal prices."""
    factory, _calls = _factory(_default_mapping())
    snapshots = _fetch(factory)
    moneyline = next(s for s in snapshots if s.market is Market.MONEYLINE)
    assert moneyline.home_price_decimal == 1.85
    assert moneyline.away_price_decimal == 2.05
    # 1.85 interpreted as American would corrupt to 1.0185 — never happens.
    assert moneyline.home_price_decimal != pytest.approx(1.0185, abs=1e-9)
    total = next(s for s in snapshots if s.market is Market.TOTAL and s.period is Period.FULL_GAME)
    assert total.over_price_decimal == 1.9
    assert total.under_price_decimal == 1.92


def test_parse_events_list_payload():
    """parse_events accepts the documented top-level ARRAY list payload and
    maps every event's h2h/totals (the list endpoint's full-game markets)."""
    snapshots = theoddsapi.parse_events(
        _load("odds_list_full_game.json"), OBSERVED, now_utc=OBSERVED
    )
    # Only odds-evt-1 yields snapshots: the off-date event has an empty
    # bookmakers array (and is date-filtered out of fetch_odds anyway).
    assert {s.game_id for s in snapshots} == {"odds-evt-1"}
    assert {s.market for s in snapshots} == {Market.MONEYLINE, Market.TOTAL}
    assert all(s.period is Period.FULL_GAME for s in snapshots)
    assert not any(s.period in (Period.Q1, Period.Q2, Period.Q3, Period.Q4) for s in snapshots)


def test_home_away_name_mismatch_raises_or_skips():
    """An h2h outcome matching NEITHER team is skipped with a gap; the other
    outcome is still mapped. The side is never guessed."""
    payload = _load("odds_event_name_mismatch.json")
    snapshots = theoddsapi.parse_event(payload, OBSERVED, now_utc=OBSERVED)
    moneyline = next(s for s in snapshots if s.market is Market.MONEYLINE)
    assert moneyline.home_price_decimal == 1.85  # "Atlanta Hawks" matched
    assert moneyline.away_price_decimal is None  # "Mystery Team" NOT guessed
    gaps = theoddsapi.last_gaps()
    assert any(g.required_field == "h2h_outcome" and "Mystery Team" in g.detail for g in gaps)


# ---------------------------------------------------------------------------
# Quarter markets (per-event endpoint only)
# ---------------------------------------------------------------------------


def test_period_payloads_remain_parseable_but_routine_fetch_never_fans_out():
    """The parser retains native-market semantics without charging per event."""
    factory, calls = _factory(_default_mapping())
    routine = _fetch(factory)
    assert all(snapshot.period is Period.FULL_GAME for snapshot in routine)
    assert len(calls) == 1
    assert not any(PER_EVENT_KEY in url for url in calls)

    snapshots = theoddsapi.parse_event(
        _load("odds_event_quarters.json"), OBSERVED, now_utc=OBSERVED
    )
    quarters = {
        s.period: s
        for s in snapshots
        if s.market is Market.TOTAL and s.period is not Period.FULL_GAME
    }
    assert set(quarters) == {Period.Q1, Period.Q2, Period.Q3, Period.Q4}
    assert quarters[Period.Q1].line == pytest.approx(58.5)
    assert quarters[Period.Q4].under_price_decimal == 1.95

    # A partial documented payload remains partial; the parser never fills a
    # missing market or infers a period line.
    partial = json.loads(json.dumps(_load("odds_event_quarters.json")))
    partial["bookmakers"][0]["markets"] = [
        market for market in partial["bookmakers"][0]["markets"] if market["key"] == "totals_q1"
    ]
    snapshots = theoddsapi.parse_event(partial, OBSERVED, now_utc=OBSERVED)
    quarter_periods = {s.period for s in snapshots if s.market is Market.TOTAL}
    assert quarter_periods == {Period.Q1}


# ---------------------------------------------------------------------------
# Quota
# ---------------------------------------------------------------------------


def test_quota_headers_captured():
    """parse_quota normalises x-requests-remaining/used (case-insensitive,
    ints) and smoke reports them via the fetch envelope convention."""
    quota = theoddsapi.parse_quota({"x-requests-remaining": "43", "x-requests-used": "57"})
    assert quota == {"x-requests-remaining": 43, "x-requests-used": 57}
    assert theoddsapi.parse_quota({}) == {}
    assert theoddsapi.parse_quota({"X-Requests-Remaining": "9"}) == {"x-requests-remaining": 9}

    headers = {"x-requests-remaining": "43", "x-requests-used": "57"}
    factory, _calls = _factory(
        {
            LIST_KEY: (_load("odds_list_full_game.json"), headers),
            PER_EVENT_KEY: _load("odds_event_quarters.json"),
        }
    )
    result = theoddsapi.smoke("NBA", factory, date_yyyymmdd=FUTURE_DATE)
    assert result["quota"]["x-requests-remaining"] == 43
    assert result["quota"]["x-requests-used"] == 57


# ---------------------------------------------------------------------------
# League gate / request shape / smoke
# ---------------------------------------------------------------------------


def test_euroleague_not_enabled():
    """No EuroLeague sport key is assumed to exist: raise instead of
    fabricating coverage."""
    factory, _calls = _factory(_default_mapping())
    with pytest.raises(ValueError, match="not enabled for this league yet"):
        theoddsapi.fetch_odds("EUROLEAGUE", FUTURE_DATE, factory)
    with pytest.raises(ValueError, match="not enabled for this league yet"):
        theoddsapi.smoke("EuroLeague", factory)


def test_routine_request_is_documented_date_bounded_full_game_only():
    factory, calls = _factory(_default_mapping())
    _fetch(factory)

    assert len(calls) == 1
    list_url = calls[0]
    assert list_url.startswith("https://api.the-odds-api.com/v4/sports/basketball_nba/odds/?")
    assert f"apiKey={API_KEY}" in list_url  # documented query-param auth
    assert "regions=us" in list_url
    assert "markets=h2h%2Ctotals%2Cspreads" in list_url
    assert "bookmakers=draftkings" in list_url
    assert "oddsFormat=decimal" in list_url
    assert "dateFormat=iso" in list_url
    assert "commenceTimeFrom=2099-11-07T00%3A00%3A00Z" in list_url
    assert "commenceTimeTo=2099-11-07T23%3A59%3A59Z" in list_url
    assert PER_EVENT_KEY not in list_url


def test_enforced_missing_or_wrong_permit_makes_no_network_call(monkeypatch):
    monkeypatch.setenv(theoddsapi.REQUIRE_PERMIT_ENV, "true")
    factory, calls = _factory(_default_mapping())
    assert _fetch(factory) == []
    assert calls == []
    assert any(g.required_field == "provider_credit_permit" for g in theoddsapi.last_gaps())

    permit = {
        "schema_version": "magolide.theoddsapi_permit",
        "permit_id": "permit-1",
        "run_id": "run-1",
        "league": "NCAA_W",
        "target_date": "2099-11-07",
        "issued_utc": OBSERVED,
        "maximum_credits": 3,
        "markets": ["h2h", "totals", "spreads"],
    }
    monkeypatch.setenv(theoddsapi.PERMIT_ENV, json.dumps(permit))
    assert _fetch(factory) == []
    assert calls == []
    assert any("another league" in g.detail for g in theoddsapi.last_gaps())

    permit["league"] = "NBA"
    monkeypatch.setenv(theoddsapi.PERMIT_ENV, json.dumps(permit))
    # Permit-gated execution ignores an inherited bookmaker override. Otherwise
    # an up-to-ten-book group could exceed the reservation's three-credit cost.
    monkeypatch.setenv(theoddsapi.ENV_KEY.replace("KEY", "BOOKMAKERS"), "draftkings,fanduel")
    monkeypatch.setenv(theoddsapi.ENV_KEY.replace("KEY", "REGIONS"), "us,uk")
    assert _fetch(factory)
    assert len(calls) == 1
    assert "bookmakers=draftkings" in calls[0]
    assert "regions=us" in calls[0]
    assert "fanduel" not in calls[0]
    assert "uk" not in calls[0]


def test_started_event_not_pre_match():
    """commence_time in the past → pre_match=False (never mark live data as
    pre-match); the books.py pre-match filter excludes the snapshot."""
    started = {
        "id": "odds-evt-old",
        "sport_key": "basketball_nba",
        "commence_time": "2001-11-07T00:30:00Z",
        "home_team": "Atlanta Hawks",
        "away_team": "Denver Nuggets",
        "bookmakers": [
            {
                "key": "draftkings",
                "title": "DraftKings",
                "markets": [
                    {
                        "key": "totals",
                        "outcomes": [
                            {"name": "Over", "price": 1.9, "point": 232.5},
                            {"name": "Under", "price": 1.92, "point": 232.5},
                        ],
                    }
                ],
            }
        ],
    }
    snapshots = theoddsapi.parse_event(started, OBSERVED, now_utc=OBSERVED)
    assert snapshots
    assert all(s.pre_match is False for s in snapshots)
    config = books.BookConfig(provider="theoddsapi")  # pre_match_only=True
    assert books.select_total_snapshots(snapshots, config, Period.FULL_GAME) == []


def test_smoke_reports_full_game_only_counts_and_quota():
    """Routine smoke exposes full-game evidence and no charged period fan-out."""
    factory, _calls = _factory(_default_mapping())
    result = theoddsapi.smoke("NBA", factory, date_yyyymmdd=FUTURE_DATE, now_utc=OBSERVED)
    assert result["enabled"] is True
    assert result["verification"] == "DOCUMENTATION_ONLY"
    assert result["sport_key"] == "basketball_nba"
    assert result["events_checked"] == 1
    assert result["full_game_totals_present"] == 1
    assert result["quarters"] == {
        "Q1": {"present": 0, "total": 1},
        "Q2": {"present": 0, "total": 1},
        "Q3": {"present": 0, "total": 1},
        "Q4": {"present": 0, "total": 1},
    }
    assert result["pre_match_vs_live"] == {"pre_match": 1, "not_pre_match": 0}
    assert result["not_usable"] == 0


def test_quarter_spread_and_first_half_1x2_are_exact_facts():
    payload = {
        "id": "odds-extended",
        "commence_time": "2099-11-07T00:30:00Z",
        "home_team": "Atlanta Hawks",
        "away_team": "Denver Nuggets",
        "bookmakers": [
            {
                "key": "draftkings",
                "markets": [
                    {
                        "key": "spreads_q1",
                        "last_update": "2026-11-06T21:59:00Z",
                        "outcomes": [
                            {"name": "Atlanta Hawks", "price": 1.9, "point": -2.5},
                            {"name": "Denver Nuggets", "price": 1.92, "point": 2.5},
                        ],
                    },
                    {
                        "key": "h2h_h1",
                        "outcomes": [
                            {"name": "Denver Nuggets", "price": 2.1},
                            {"name": "Draw", "price": 11.0},
                            {"name": "Atlanta Hawks", "price": 1.85},
                        ],
                    },
                ],
            }
        ],
    }
    snapshots = theoddsapi.parse_event(payload, OBSERVED, now_utc=OBSERVED)
    spread = next(row for row in snapshots if row.market is Market.SPREAD)
    assert spread.period is Period.Q1
    assert spread.line == -2.5  # home-oriented, never inferred from a price
    assert spread.home_price_decimal == 1.9
    assert spread.away_price_decimal == 1.92
    first_half = next(
        row for row in snapshots if row.market is Market.MONEYLINE and row.period is Period.H1
    )
    assert first_half.home_price_decimal == 1.85
    assert first_half.away_price_decimal == 2.1
    assert first_half.draw_price_decimal == 11.0


def test_contradictory_spread_points_fail_closed():
    payload = {
        "id": "odds-bad-spread",
        "commence_time": "2099-11-07T00:30:00Z",
        "home_team": "Atlanta Hawks",
        "away_team": "Denver Nuggets",
        "bookmakers": [
            {
                "key": "draftkings",
                "markets": [
                    {
                        "key": "spreads_q1",
                        "outcomes": [
                            {"name": "Atlanta Hawks", "price": 1.9, "point": -2.5},
                            {"name": "Denver Nuggets", "price": 1.9, "point": 3.5},
                        ],
                    }
                ],
            }
        ],
    }
    theoddsapi._reset_gaps()
    assert theoddsapi.parse_event(payload, OBSERVED, now_utc=OBSERVED) == []
    assert any(gap.required_field == "spread_line_q1" for gap in theoddsapi.last_gaps())
