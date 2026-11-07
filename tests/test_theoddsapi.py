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
- totals → FULL_GAME; totals_q1..q4 → exactly Q1..Q4 (quarters come from
  the documented per-event odds endpoint).
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


def _factory(
    mapping: Dict[str, Any]
) -> Tuple[Callable[[str], Callable[[str], Any]], List[str]]:
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

    total = next(
        s for s in snapshots if s.market is Market.TOTAL and s.period is Period.FULL_GAME
    )
    assert total.book == "draftkings"
    assert total.line == pytest.approx(232.5)  # the point, never a price
    assert total.line != total.over_price_decimal != total.under_price_decimal
    assert total.over_price_decimal == 1.9
    assert total.under_price_decimal == 1.92

    # Off-date event in the list is skipped (date narrowing on commence_time).
    assert not any(s.game_id == "odds-evt-offdate" for s in snapshots)
    assert any(g.required_field == "date_window" for g in theoddsapi.last_gaps())
    assert calls  # per-event quarter calls were made for the on-date event


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
    total = next(
        s for s in snapshots if s.market is Market.TOTAL and s.period is Period.FULL_GAME
    )
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
    assert any(
        g.required_field == "h2h_outcome" and "Mystery Team" in g.detail for g in gaps
    )


# ---------------------------------------------------------------------------
# Quarter markets (per-event endpoint only)
# ---------------------------------------------------------------------------


def test_quarter_markets_from_per_event_odds():
    """totals_q1..q4 from the per-event odds payload map to exactly
    Q1..Q4; only the quarters present in the payload are emitted."""
    factory, calls = _factory(_default_mapping())
    snapshots = _fetch(factory)

    quarters = {
        s.period: s
        for s in snapshots
        if s.market is Market.TOTAL and s.period is not Period.FULL_GAME
    }
    assert set(quarters) == {Period.Q1, Period.Q2, Period.Q3, Period.Q4}
    assert quarters[Period.Q1].line == pytest.approx(58.5)
    assert quarters[Period.Q2].line == pytest.approx(57.5)
    assert quarters[Period.Q3].line == pytest.approx(56.5)
    assert quarters[Period.Q4].line == pytest.approx(60.0)
    assert quarters[Period.Q1].over_price_decimal == 1.87
    assert quarters[Period.Q4].under_price_decimal == 1.95
    assert all(s.book == "draftkings" for s in quarters.values())

    # Quarters came from the documented per-event endpoint, not the list one.
    assert any(PER_EVENT_KEY in u for u in calls)
    assert all(LIST_KEY not in u for u in calls if PER_EVENT_KEY in u)

    # Partial payload: only totals_q1 present → only a Q1 snapshot appears.
    partial = json.loads(json.dumps(_load("odds_event_quarters.json")))
    partial["bookmakers"][0]["markets"] = [
        market
        for market in partial["bookmakers"][0]["markets"]
        if market["key"] == "totals_q1"
    ]
    snapshots = theoddsapi.parse_event(partial, OBSERVED, now_utc=OBSERVED)
    quarter_periods = {
        s.period for s in snapshots if s.market is Market.TOTAL
    }
    assert quarter_periods == {Period.Q1}


# ---------------------------------------------------------------------------
# Quota
# ---------------------------------------------------------------------------


def test_quota_headers_captured():
    """parse_quota normalises x-requests-remaining/used (case-insensitive,
    ints) and smoke reports them via the fetch envelope convention."""
    quota = theoddsapi.parse_quota(
        {"x-requests-remaining": "43", "x-requests-used": "57"}
    )
    assert quota == {"x-requests-remaining": 43, "x-requests-used": 57}
    assert theoddsapi.parse_quota({}) == {}
    assert theoddsapi.parse_quota({"X-Requests-Remaining": "9"}) == {
        "x-requests-remaining": 9
    }

    headers = {"x-requests-remaining": "43", "x-requests-used": "57"}
    factory, _calls = _factory(
        {LIST_KEY: (_load("odds_list_full_game.json"), headers), PER_EVENT_KEY: _load("odds_event_quarters.json")}
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


def test_request_urls_use_documented_params():
    """List endpoint: /v4/sports/basketball_nba/odds/? with
    markets=h2h,totals, regions=us, bookmakers=draftkings, oddsFormat=decimal
    and the apiKey query param. Per-event endpoint: markets=totals_q1..q4."""
    factory, calls = _factory(_default_mapping())
    _fetch(factory)

    list_url = next(u for u in calls if LIST_KEY in u)
    assert list_url.startswith(
        "https://api.the-odds-api.com/v4/sports/basketball_nba/odds/?"
    )
    assert f"apiKey={API_KEY}" in list_url  # documented query-param auth
    assert "regions=us" in list_url
    assert "markets=h2h%2Ctotals" in list_url
    assert "bookmakers=draftkings" in list_url
    assert "oddsFormat=decimal" in list_url
    assert "dateFormat=iso" in list_url

    per_event = next(u for u in calls if PER_EVENT_KEY in u)
    assert per_event.startswith(
        "https://api.the-odds-api.com/v4/sports/basketball_nba/events/odds-evt-1/odds?"
    )
    assert "markets=totals_q1%2Ctotals_q2%2Ctotals_q3%2Ctotals_q4" in per_event
    assert "oddsFormat=decimal" in per_event


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


def test_smoke_reports_counts_and_quota():
    """smoke() reports per-quarter coverage, pre-match vs started counts and
    the sport_key used for the NBA (verified at enablement)."""
    factory, _calls = _factory(_default_mapping())
    result = theoddsapi.smoke("NBA", factory, date_yyyymmdd=FUTURE_DATE, now_utc=OBSERVED)
    assert result["enabled"] is True
    assert result["verification"] == "DOCUMENTATION_ONLY"
    assert result["sport_key"] == "basketball_nba"
    assert result["events_checked"] == 1
    assert result["full_game_totals_present"] == 1
    assert result["quarters"] == {
        "Q1": {"present": 1, "total": 1},
        "Q2": {"present": 1, "total": 1},
        "Q3": {"present": 1, "total": 1},
        "Q4": {"present": 1, "total": 1},
    }
    assert result["pre_match_vs_live"] == {"pre_match": 1, "not_pre_match": 0}
    assert result["not_usable"] == 0
