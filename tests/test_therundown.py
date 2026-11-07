"""Tests for ingestion/providers/therundown.py — documented API v2 contract.

The v2 payloads below are DOCUMENTATION_ONLY synthetic responses authored
from the documented TheRundown v2 schema (``events → markets →
participants → lines → prices``); they are NOT captured API responses and no
coverage claim is live-verified. No network is used and NO fixture files
are needed: a fake ``fetch_factory(provider)`` serves the payloads by URL
substring, exercising both documented calls (``/api/v2/sports`` and
``/api/v2/sports/{id}/events/{YYYY-MM-DD}``).

Binding semantics proven here:
- Auth: the key reaches the ``X-TheRundown-Key`` header (and the production
  fetch factory attaches it via ``HttpClient(default_headers=...)``), and
  the secret NEVER appears in any requested URL.
- The NBA sport id is RESOLVED from ``/sports`` — there is no hardcoded id
  (the module has no ``NBA_SPORT_ID`` constant; a stub uses id 4, and a
  reference payload without NBA raises instead of guessing).
- Event dates are ``YYYY-MM-DD`` (``YYYYMMDD`` input is normalised).
- Query parameters are the documented ``market_ids=1,2,3`` and
  ``affiliate_ids=...``; the spread ("2") is requested but unmapped.
- American prices converted via books.american_to_decimal; the total
  ``line`` is the LINE VALUE, never a price.
- The adapter has NO code path that can emit a Period.Q1..Q4 snapshot: a
  quarter market id and a quarter-period line both produce explicit gaps
  and no snapshot, the parser asserts every built snapshot is FULL_GAME,
  and books.py additionally filters the provider out of quarter cells
  (defense in depth).
- ``event_status_phase != "pre-game"`` → ``pre_match=False`` (final games
  are never pre-match); books.py pre-match filters exclude them.
- EuroLeague raises ValueError (coverage unproven, never fabricated).
- Quota: ``parse_quota`` normalises X-Datapoints headers delivered through
  the documented fetch envelope convention (``{"json": ..., "headers": ...}``).
"""

from __future__ import annotations

from dataclasses import replace
from typing import Any, Callable, Dict, List, Tuple

import pytest

from ingestion import books
from ingestion.models import BookSnapshot, Market, Period, payload_hash
from ingestion.providers import therundown

API_KEY = "tdn-test-key-not-real"
OBSERVED = "2026-11-06T22:00:00Z"
SPORT_ID = 4  # deliberately NOT the old hardcoded 2 — proves /sports resolution
EVENT_ID = "11879001"
LINE_TS = "2026-11-06T21:30:00Z"
HOME = "Denver Nuggets"
AWAY = "Golden State Warriors"

SPORTS_URL = "https://therundown.io/api/v2/sports"
EVENTS_URL = (
    f"https://therundown.io/api/v2/sports/{SPORT_ID}/events/2026-11-07"
    "?market_ids=1,2,3&affiliate_ids=23"
)

_V2_SPORTS = {
    "sports": [
        {"id": 1, "abbreviation": "NFL"},
        {"id": SPORT_ID, "abbreviation": "NBA"},
        {"id": 7, "abbreviation": "MLB"},
    ]
}


def _moneyline_market(affiliate: int = 23) -> Dict[str, Any]:
    base = {"affiliate_id": affiliate, "period_id": 1, "last_modified": LINE_TS}
    return {
        "id": 1,
        "name": "moneyline",
        "participants": [
            {"side": "home", "is_home": True,
             "lines": [dict(base, prices=[{"type": "home", "price": -150}])]},
            {"side": "away", "is_home": False,
             "lines": [dict(base, prices=[{"type": "away", "price": 130}])]},
        ],
    }


def _total_market(affiliate: int = 23) -> Dict[str, Any]:
    base = {"affiliate_id": affiliate, "period_id": 1, "total": 232.0, "last_modified": LINE_TS}
    return {
        "id": 3,
        "name": "total",
        "participants": [
            {"side": "over", "lines": [dict(base, prices=[{"type": "over", "price": -110}])]},
            {"side": "under", "lines": [dict(base, prices=[{"type": "under", "price": -110}])]},
        ],
    }


def _v2_event(
    markets: List[Dict[str, Any]] = None, phase: str = "pre-game", affiliate: int = 23
) -> Dict[str, Any]:
    return {
        "event_id": EVENT_ID,
        "event_date": "2026-11-07",
        "event_status_phase": phase,
        "participants": [
            {"full_name": HOME, "is_home": True},
            {"full_name": AWAY, "is_home": False},
        ],
        "markets": markets if markets is not None else [_moneyline_market(affiliate), _total_market(affiliate)],
    }


def _factory(mapping: Dict[str, Any]) -> Tuple[Callable[[str], Callable[[str], Any]], List[str]]:
    """fetch_factory(provider) -> FetchFn serving payloads by URL substring.

    A ``(payload, headers)`` mapping value is served through the documented
    envelope convention ``{"json": ..., "headers": ...}`` so quota capture
    can be exercised. Longest matching key wins (``/events/`` beats
    ``/sports`` on the events URL).
    """
    calls: List[str] = []

    def fetch(url: str) -> Any:
        calls.append(url)
        matches = [key for key in mapping if key in url]
        assert matches, f"no payload mapped for URL: {url}"
        value = mapping[max(matches, key=len)]
        if isinstance(value, tuple):
            payload, headers = value
            return {"json": payload, "headers": headers}
        return value

    def factory(provider: str) -> Callable[[str], Any]:
        assert provider == therundown.PROVIDER
        return fetch

    return factory, calls


def _nba_mapping(events_payload: Any, headers: Dict[str, str] = None, sports: Any = None) -> Dict[str, Any]:
    events_value = (events_payload, headers) if headers else events_payload
    return {"/sports": sports if sports is not None else _V2_SPORTS, "/events/": events_value}


@pytest.fixture(autouse=True)
def _api_key(monkeypatch):
    monkeypatch.setenv(therundown.ENV_KEY, API_KEY)
    yield


# ---------------------------------------------------------------------------
# Enabled / disabled / auth
# ---------------------------------------------------------------------------


def test_disabled_without_key(monkeypatch):
    """Missing THERUNDOWN_API_KEY: fetch_odds returns [] and smoke reports
    DOCUMENTATION_ONLY — the keyless pipeline must never raise or fetch."""
    monkeypatch.delenv(therundown.ENV_KEY, raising=False)
    factory, calls = _factory(_nba_mapping({"events": []}))
    assert therundown.fetch_odds("NBA", "20261107", factory) == []
    result = therundown.smoke("NBA", factory)
    assert result["enabled"] is False
    assert result["verification"] == "DOCUMENTATION_ONLY"
    assert calls == []  # never even builds/sends a request without a key


def test_auth_reaches_header_and_secret_never_in_urls():
    """v2 auth contract: exactly two documented calls, the secret appears in
    NO URL (header-only auth), and auth_headers() carries X-TheRundown-Key."""
    factory, calls = _factory(_nba_mapping({"events": [_v2_event()]}))
    snapshots = therundown.fetch_odds("NBA", "20261107", factory)
    assert snapshots
    assert calls == [SPORTS_URL, EVENTS_URL]
    for url in calls:
        assert API_KEY not in url
        assert "access_token" not in url and "apiKey" not in url and "token=" not in url
    assert therundown.auth_headers() == {"X-TheRundown-Key": API_KEY}
    assert therundown.AUTH_HEADER == "X-TheRundown-Key"


def test_production_fetch_factory_attaches_key_as_default_header():
    """pipeline._default_fetch_factory("therundown") wires the key into the
    HttpClient default headers (X-TheRundown-Key) — never into a URL builder."""
    from ingestion.http import HttpClient
    from ingestion.pipeline import _default_fetch_factory

    getter = _default_fetch_factory("therundown")
    client = getter.__self__
    assert isinstance(client, HttpClient)
    assert client.default_headers.get("X-TheRundown-Key") == API_KEY
    # No header VALUE is a URL, and the key travels nowhere else by default.
    assert all("http" not in str(value) for value in client.default_headers.values())


# ---------------------------------------------------------------------------
# NBA sport id resolution (no hardcoded id)
# ---------------------------------------------------------------------------


def test_nba_sport_id_resolved_from_sports_endpoint():
    """The NBA id comes from GET /api/v2/sports (stub says 4 — not the old
    hardcoded 2), and the module exposes no NBA_SPORT_ID constant at all."""
    assert not hasattr(therundown, "NBA_SPORT_ID")
    factory, calls = _factory(_nba_mapping({"events": [_v2_event()]}))
    therundown.fetch_odds("NBA", "20261107", factory)
    assert calls[0] == SPORTS_URL
    assert any(f"/sports/{SPORT_ID}/events/" in u for u in calls)
    assert not any("/sports/2/" in u for u in calls)


def test_nba_unresolvable_raises_without_guessing():
    """A /sports payload without NBA raises ValueError and NO events call is
    made — the adapter refuses to guess or hardcode a sport id."""
    factory, calls = _factory(_nba_mapping({"events": [_v2_event()]}, sports={"sports": [{"id": 1, "abbreviation": "NFL"}]}))
    with pytest.raises(ValueError, match="refusing to hardcode or guess"):
        therundown.fetch_odds("NBA", "20261107", factory)
    assert calls == [SPORTS_URL]
    gaps = therundown.last_gaps()
    assert any(g.required_field == "sport_id" for g in gaps)


# ---------------------------------------------------------------------------
# URL shape: dates, market_ids, affiliate_ids
# ---------------------------------------------------------------------------


def test_event_dates_are_iso_and_market_params_documented():
    """Both accepted input formats produce the documented YYYY-MM-DD URL, and
    the query carries exactly market_ids=1,2,3 and affiliate_ids=23."""
    factory, calls = _factory(_nba_mapping({"events": [_v2_event()]}))
    therundown.fetch_odds("NBA", "2026-11-07", factory)
    therundown.fetch_odds("NBA", "20261107", factory)
    assert calls == [SPORTS_URL, EVENTS_URL, SPORTS_URL, EVENTS_URL]
    with pytest.raises(ValueError, match="slate date"):
        therundown.fetch_odds("NBA", "Nov 7", _factory(_nba_mapping({"events": []}))[0])


def test_affiliate_override_switches_param_and_book(monkeypatch):
    """THERUNDOWN_AFFILIATE_ID=22 (BetMGM) → affiliate_ids=22 and, for a
    payload whose lines carry affiliate_id 22, book BetMGM. (A line-level
    affiliate_id is authoritative when present — a line from 23 IS a
    FanDuel line even if a request disagreed.)"""
    monkeypatch.setenv("THERUNDOWN_AFFILIATE_ID", "22")
    factory, calls = _factory(_nba_mapping({"events": [_v2_event(affiliate=22)]}))
    snapshots = therundown.fetch_odds("NBA", "20261107", factory)
    assert calls == [SPORTS_URL, EVENTS_URL.replace("affiliate_ids=23", "affiliate_ids=22")]
    assert {s.book for s in snapshots} == {"BetMGM"}


# ---------------------------------------------------------------------------
# Moneyline / full-game total normalisation
# ---------------------------------------------------------------------------


def test_parses_moneyline_and_full_game_total_to_decimal():
    """FanDuel (affiliate 23) pre-game event: American -150/+130 moneyline
    and a 232.0 total with -110/-110 prices become decimal snapshots."""
    factory, _calls = _factory(_nba_mapping({"events": [_v2_event()]}))
    snapshots = therundown.fetch_odds("NBA", "20261107", factory)
    assert snapshots

    moneyline = next(s for s in snapshots if s.market is Market.MONEYLINE)
    assert moneyline.provider == "therundown"
    assert moneyline.game_id == EVENT_ID
    assert moneyline.book == "FanDuel"
    assert moneyline.period is Period.FULL_GAME
    assert moneyline.pre_match is True
    # American -150 -> 1 + 100/150
    assert moneyline.home_price_decimal == pytest.approx(1.6666666666666667, abs=1e-9)
    assert moneyline.home_price_decimal == books.american_to_decimal(-150)
    # American +130 -> 1 + 130/100
    assert moneyline.away_price_decimal == pytest.approx(2.3, abs=1e-9)
    assert moneyline.away_price_decimal == books.american_to_decimal(130)

    total = next(s for s in snapshots if s.market is Market.TOTAL)
    assert total.period is Period.FULL_GAME
    assert total.line == pytest.approx(232.0)  # the LINE VALUE, never a price
    assert total.line > 0
    assert total.line != total.over_price_decimal  # line, not a price
    assert total.line != total.under_price_decimal
    # American -110 -> 1 + 100/110 (both sides)
    assert total.over_price_decimal == pytest.approx(1.9090909090909092, abs=1e-9)
    assert total.under_price_decimal == pytest.approx(1.9090909090909092, abs=1e-9)
    assert total.over_price_decimal == books.american_to_decimal(-110)
    assert total.under_price_decimal == books.american_to_decimal(-110)


def test_observed_utc_param_and_payload_hash():
    """observed_utc comes from the injectable now_utc param (fetch time) and
    payload_hash hashes the RAW event dict."""
    event = _v2_event()
    factory, _calls = _factory(_nba_mapping({"events": [event]}))
    snapshots = therundown.fetch_odds("NBA", "20261107", factory, now_utc=OBSERVED)
    assert snapshots
    assert all(s.observed_utc == OBSERVED for s in snapshots)
    assert all(s.payload_hash == payload_hash(event) for s in snapshots)
    assert all(s.main_line is True for s in snapshots)
    assert all(s.source_modified_utc == LINE_TS for s in snapshots)


# ---------------------------------------------------------------------------
# Quarter prohibition (hard rule: no quarter code path exists)
# ---------------------------------------------------------------------------


def test_never_emits_quarter_periods():
    """A quarter market id (26) AND a quarter-period line (period_id 4) inside
    the total market both produce explicit gaps and NO Q1..Q4 snapshot; the
    hand-crafted 58.5 quarter line never surfaces anywhere."""
    quarter_market = {
        "id": 26,
        "name": "1st quarter total",
        "participants": [
            {"side": "over",
             "lines": [{"affiliate_id": 23, "period_id": 4, "total": 58.5,
                        "prices": [{"type": "over", "price": -105}]}]},
        ],
    }
    poisoned_total = _total_market()
    poisoned_total["participants"][0]["lines"].append(
        {"affiliate_id": 23, "period_id": 4, "total": 58.5,
         "prices": [{"type": "over", "price": -105}]}
    )
    factory, _calls = _factory(_nba_mapping({"events": [_v2_event([_moneyline_market(), quarter_market, poisoned_total])]}))
    snapshots = therundown.fetch_odds("NBA", "20261107", factory)
    quarter_periods = {Period.Q1, Period.Q2, Period.Q3, Period.Q4}
    assert snapshots
    assert all(s.period is Period.FULL_GAME for s in snapshots)
    assert all(s.period not in quarter_periods for s in snapshots)
    assert not any(s.line == pytest.approx(58.5) for s in snapshots)
    # Both poison sources are visible as explicit gaps, never silently mapped.
    gaps = therundown.last_gaps()
    assert any("26" in g.detail and "quarter" in g.detail.lower() for g in gaps)
    assert any("4" in g.detail and "not full-game" in g.detail for g in gaps)

    # Belt and braces: the runtime assertion raises if a Q* snapshot is ever
    # handed to the guard, and the snapshot factory refuses quarter periods.
    bad = BookSnapshot(
        game_id="g",
        provider="therundown",
        book="FanDuel",
        market=Market.TOTAL,
        period=Period.Q1,
        observed_utc=OBSERVED,
        line=58.5,
    )
    with pytest.raises(ValueError, match="quarter"):
        therundown._assert_full_game_only([bad])
    with pytest.raises(ValueError, match="full-game only"):
        therundown._snapshot(
            "g", "FanDuel", Market.TOTAL, Period.Q1, OBSERVED, True, {}, None, line=58.5
        )


def test_quarter_block_guard_via_books():
    """Defense in depth: a hand-crafted therundown Q1 BookSnapshot is
    filtered out of quarter cells by books.select_total_snapshots, while the
    provider's FULL_GAME snapshots still pass."""
    quarter = BookSnapshot(
        game_id="g1",
        provider="therundown",
        book="FanDuel",
        market=Market.TOTAL,
        period=Period.Q1,
        observed_utc=OBSERVED,
        line=58.5,
        over_price_decimal=1.9090909090909092,
        under_price_decimal=1.9090909090909092,
        pre_match=True,
        main_line=True,
    )
    config = books.BookConfig(provider="therundown")
    assert books.select_total_snapshots([quarter], config, Period.Q1) == []
    assert books.select_total_snapshots([quarter], config) == []

    full_game = replace(quarter, period=Period.FULL_GAME, line=232.0)
    assert books.select_total_snapshots([full_game], config, Period.FULL_GAME) == [full_game]
    assert "therundown" in books.PROVIDERS_BLOCKED_FROM_QUARTERS


# ---------------------------------------------------------------------------
# Spread market + pre-match semantics
# ---------------------------------------------------------------------------


def test_spread_requested_but_unmapped_with_gap():
    """market_ids=1,2,3 requests the spread too; the adapter maps only
    moneyline/total and records the spread as an explicit UNSUPPORTED gap."""
    spread_market = {
        "id": 2,
        "name": "spread",
        "participants": [
            {"side": "home",
             "lines": [{"affiliate_id": 23, "period_id": 1, "spread": -4.5,
                        "prices": [{"type": "home", "price": -110}]}]},
        ],
    }
    factory, _calls = _factory(_nba_mapping({"events": [_v2_event([_moneyline_market(), spread_market, _total_market()])]}))
    snapshots = therundown.fetch_odds("NBA", "20261107", factory)
    assert {s.market for s in snapshots} == {Market.MONEYLINE, Market.TOTAL}
    assert any(g.required_field == "markets" and "'2'" in g.detail for g in therundown.last_gaps())


def test_final_game_not_pre_match():
    """event_status_phase 'final' -> pre_match=False (never pre-match), and
    the books.py pre-match filters exclude the snapshots downstream."""
    factory, _calls = _factory(_nba_mapping({"events": [_v2_event(phase="final")]}))
    snapshots = therundown.fetch_odds("NBA", "20261107", factory)
    assert snapshots
    assert all(s.pre_match is False for s in snapshots)
    config = books.BookConfig(provider="therundown")
    assert books.select_total_snapshots(snapshots, config, Period.FULL_GAME) == []
    assert books.select_moneyline_snapshots(snapshots, config) == []
    # The same snapshots are usable when live data is explicitly allowed.
    allow_live = books.BookConfig(provider="therundown", pre_match_only=False)
    assert books.select_total_snapshots(snapshots, allow_live, Period.FULL_GAME)


def test_ambiguous_teams_are_skipped_not_guessed():
    """Event-level participants without resolvable home/away sides → the
    event is skipped with a gap (orientation is never guessed)."""
    event = _v2_event()
    event["participants"] = [{"full_name": HOME}, {"full_name": AWAY}]
    factory, _calls = _factory(_nba_mapping({"events": [event]}))
    assert therundown.fetch_odds("NBA", "20261107", factory) == []
    assert any(g.required_field == "teams" for g in therundown.last_gaps())


# ---------------------------------------------------------------------------
# League gates
# ---------------------------------------------------------------------------


def test_euroleague_not_enabled():
    """EuroLeague coverage is unproven for every bookmaker adapter: raise
    instead of fabricating lines."""
    factory, _calls = _factory(_nba_mapping({"events": [_v2_event()]}))
    with pytest.raises(ValueError, match="not enabled for this league yet"):
        therundown.fetch_odds("EUROLEAGUE", "20261107", factory)
    with pytest.raises(ValueError, match="not enabled for this league yet"):
        therundown.smoke("euroleague", factory)  # case-insensitive gate
    with pytest.raises(ValueError, match="not enabled for this league yet"):
        therundown.fetch_odds("EuroLeague", "20261107", factory)


# ---------------------------------------------------------------------------
# Quota / smoke
# ---------------------------------------------------------------------------


def test_parse_quota_reads_datapoints_headers():
    """X-Datapoints* headers normalise case-insensitively to a quota dict
    with int values; unrelated headers are ignored."""
    quota = therundown.parse_quota(
        {
            "X-Datapoints": "8",
            "x-datapoints-remaining": "18992",
            "X-Datapoints-Used": "1008",
            "Content-Type": "application/json",
        }
    )
    assert quota == {
        "x-datapoints": 8,
        "x-datapoints-remaining": 18992,
        "x-datapoints-used": 1008,
    }
    assert therundown.parse_quota({}) == {}


def test_smoke_reports_envelope_quota_sport_id_and_counts():
    """smoke() makes both documented calls, unwraps the {"json", "headers"}
    envelope, reports the resolved sport id and structurally zero quarters."""
    headers = {"X-Datapoints": "10", "X-Datapoints-Remaining": "18990"}
    factory, calls = _factory(_nba_mapping({"events": [_v2_event()]}, headers=headers))
    result = therundown.smoke("NBA", factory, date_yyyymmdd="20261107", now_utc=OBSERVED)
    assert calls == [SPORTS_URL, EVENTS_URL]
    assert result["enabled"] is True
    assert result["verification"] == "DOCUMENTATION_ONLY"
    assert result["sport_id"] == SPORT_ID
    assert result["date"] == "2026-11-07"
    assert result["affiliate_id"] == 23
    assert result["events_checked"] == 1
    assert result["full_game_totals_present"] == 1
    assert result["quarters"] == {
        "Q1": {"present": 0, "total": 1},
        "Q2": {"present": 0, "total": 1},
        "Q3": {"present": 0, "total": 1},
        "Q4": {"present": 0, "total": 1},
    }
    assert result["pre_match_vs_live"] == {"pre_match": 1, "not_pre_match": 0}
    assert result["quota"] == {"x-datapoints": 10, "x-datapoints-remaining": 18990}
    # Documented cost formula: 2 * events + prices (2 ml + 2 total prices).
    assert result["expected_datapoints_cost"] == 2 * 1 + 4
