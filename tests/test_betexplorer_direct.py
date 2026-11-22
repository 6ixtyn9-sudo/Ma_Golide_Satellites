"""Hermetic contracts for the bounded direct BetExplorer named-book fallback."""

from __future__ import annotations

import json

from ingestion.http import HttpError
from ingestion.league_registry import league_spec, price_evidence_policy
from ingestion.models import (
    BookSnapshot,
    Game,
    GameStatus,
    League,
    Market,
    OperationalValidation,
    Period,
    PriceProvenance,
    SourcePermission,
)
from ingestion.pipeline import Pipeline, PipelineConfig
from ingestion.providers import betexplorer_direct, theoddsapi
from ingestion.store import InMemoryStore

OBSERVED = "2026-11-20T12:00:00Z"
DATE = "20261120"
EVENT_ID = "AbCdEf12"

BOARD = f'''<table class="table-main">
<tr data-dt="20,11,2026,20,00">
  <td class="table-main__tt"><a href="/basketball/usa/nba/los-angeles-lakers-boston-celtics/{EVENT_ID}/">
    <span>Los Angeles Lakers</span> - <span>Boston Celtics</span>
  </a></td>
</tr>
<tr data-dt="20,11,2026,20,00">
  <td class="table-main__tt"><a href="/football/usa/example/ignored/ZZZZ9999/">
    <span>Wrong Sport</span> - <span>Ignored</span>
  </a></td>
</tr>
</table>'''

# Basketball H/A is the two-way full-game moneyline surface.  These use the
# compact and all-book comparison shapes so the parser can prove it ignores the
# surrounding aggregate price widgets and unnamed rows.
HA_ODDS = json.dumps(
    {
        "odds": '''
        <div class="oddsComparisonAll__average_text" data-odd="1.64"></div>
        <div class="oddsComparisonAll__average_text" data-odd="2.31"></div>
        <table><tbody id="best-odds-0">
          <tr>
            <td><a class="in-bookmaker-logo-link" href="/bookmaker/365/">Bet365</a></td>
            <td data-odd="1.70"></td><td data-odd="2.20"></td>
          </tr>
          <tr>
            <td data-odd="1.75"></td><td data-odd="2.12"></td>
          </tr>
        </tbody></table>
        <div data-all-handicap="0">
          <div class="oddsComparisonAll__rowBookie">
            <a class="in-bookmaker-logo-link" href="/bookmaker/365/">Bet365</a>
            <div data-odd="1.72"></div><div data-odd="2.18"></div>
          </div>
          <div class="oddsComparisonAll__rowBookie">
            <span class="in-bookmaker-logo" title="Pinnacle"></span>
            <div data-odd="1.74"></div><div data-odd="2.16"></div>
          </div>
        </div>
        '''
    }
)

# O/U records the total threshold and the price order directly from the source:
# first column Over, second column Under.  The active source navigation picks
# the only main line; the 142.5 row remains present but non-main.
OU_ODDS = json.dumps(
    {
        "odds": '''
        <ul class="bestOddsComparison">
          <li id="142.50">142.5</li>
          <li id="143.50" class="oddsComparison__activeSubLi">143.5</li>
        </ul>
        <table data-handicap="142.50"><tbody>
          <tr><td><span class="in-bookmaker-logo" title="BetMGM.us"></span></td>
          <td data-odd="1.91"></td><td data-odd="1.91"></td></tr>
        </tbody></table>
        <table data-handicap="143.50"><tbody>
          <tr><td><span class="in-bookmaker-logo" title="BetMGM.us"></span></td>
          <td data-odd="1.83"></td><td data-odd="2.00"></td></tr>
          <tr><td><span class="in-bookmaker-logo" title="Stake.com"></span></td>
          <td data-odd="1.85"></td><td data-odd="1.95"></td></tr>
          <tr><td>unnamed</td><td data-odd="1.90"></td><td data-odd="1.90"></td></tr>
        </tbody></table>
        '''
    }
)

# AH uses the same two source columns as H/A: first price is home at the
# source's home-oriented handicap, second price is away.
AH_ODDS = json.dumps(
    {
        "odds": '''
        <ul class="bestOddsComparison">
          <li id="-0.50" class="oddsComparison__activeSubLi">-0.5</li>
          <li id="-1.50">-1.5</li>
        </ul>
        <table data-handicap="-0.50"><tbody>
          <tr><td><a class="in-bookmaker-logo-link" href="/bookmaker/365/">Bet365</a></td>
          <td data-odd="1.92"></td><td data-odd="1.90"></td></tr>
        </tbody></table>
        <table data-handicap="-1.50"><tbody>
          <tr><td><span class="in-bookmaker-logo" title="Pinnacle"></span></td>
          <td data-odd="2.15"></td><td data-odd="1.70"></td></tr>
        </tbody></table>
        '''
    }
)

# Captured basketball response shape: the visible active-nav id is a stale
# cross-sport template value and not one of the source's actual table lines.
# Such a marker must never manufacture a main line.
STALE_NAV_OU = json.dumps(
    {
        "odds": '''
        <ul class="bestOddsComparison">
          <li id="2.50" class="oddsComparison__activeSubLi">2.5</li>
        </ul>
        <table data-handicap="133.50"><tbody>
          <tr><td><span class="in-bookmaker-logo" title="BetMGM.us"></span></td>
          <td data-odd="1.91"></td><td data-odd="1.91"></td></tr>
        </tbody></table>
        '''
    }
)
STALE_NAV_AH = json.dumps(
    {
        "odds": '''
        <ul class="bestOddsComparison">
          <li id="-0.50" class="oddsComparison__activeSubLi">-0.5</li>
        </ul>
        <table data-handicap="-14.50"><tbody>
          <tr><td><span class="in-bookmaker-logo" title="BetMGM.us"></span></td>
          <td data-odd="1.91"></td><td data-odd="1.91"></td></tr>
        </tbody></table>
        <table data-handicap="-13.50"><tbody>
          <tr><td><span class="in-bookmaker-logo" title="Stake.com"></span></td>
          <td data-odd="1.91"></td><td data-odd="1.91"></td></tr>
        </tbody></table>
        '''
    }
)


def _game() -> Game:
    return Game(
        game_id="nba-lal-bos",
        league=League.NBA,
        season="2026-27",
        home_team_id="lal",
        away_team_id="bos",
        home_name="Los Angeles Lakers",
        away_name="Boston Celtics",
        scheduled_utc="2026-11-20T19:00:00Z",
        status=GameStatus.SCHEDULED,
    )


def _fetch(calls: list[str]):
    def fetch(url: str) -> str:
        calls.append(url)
        if "/basketball/results/" in url:
            return BOARD
        if f"/match-odds/{EVENT_ID}/1/ha/bestOdds/" in url:
            return HA_ODDS
        if f"/match-odds/{EVENT_ID}/1/ou/bestOdds/" in url:
            return OU_ODDS
        if f"/match-odds/{EVENT_ID}/1/ah/bestOdds/" in url:
            return AH_ODDS
        raise AssertionError(f"unexpected direct BetExplorer URL: {url}")

    return fetch


def test_day_board_and_named_row_parsers_refuse_other_sports_averages_and_unnamed_rows():
    fixtures = betexplorer_direct.parse_day_board(BOARD)
    assert fixtures == [
        {
            "event_id": EVENT_ID,
            "home_name": "Los Angeles Lakers",
            "away_name": "Boston Celtics",
            "date": "2026-11-20",
            "match_url": (
                "https://www.betexplorer.com/basketball/usa/nba/"
                "los-angeles-lakers-boston-celtics/AbCdEf12/"
            ),
        }
    ]

    moneyline_rows = betexplorer_direct.parse_named_moneyline_rows(json.loads(HA_ODDS)["odds"])
    assert moneyline_rows == [
        {"book": "Bet365", "home_price_decimal": 1.72, "away_price_decimal": 2.18},
        {"book": "Pinnacle", "home_price_decimal": 1.74, "away_price_decimal": 2.16},
    ]
    assert betexplorer_direct.parse_named_line_rows(json.loads(OU_ODDS)["odds"]) == [
        {
            "book": "BetMGM.us",
            "line": 142.5,
            "first_price_decimal": 1.91,
            "second_price_decimal": 1.91,
        },
        {
            "book": "BetMGM.us",
            "line": 143.5,
            "first_price_decimal": 1.83,
            "second_price_decimal": 2.0,
        },
        {
            "book": "Stake.com",
            "line": 143.5,
            "first_price_decimal": 1.85,
            "second_price_decimal": 1.95,
        },
    ]
    assert betexplorer_direct._source_active_line(json.loads(OU_ODDS)["odds"]) == 143.5
    assert betexplorer_direct._source_active_line(json.loads(AH_ODDS)["odds"]) == -0.5


def test_stale_cross_sport_active_nav_id_cannot_invent_a_basketball_main_line(monkeypatch):
    ou_html = json.loads(STALE_NAV_OU)["odds"]
    ah_html = json.loads(STALE_NAV_AH)["odds"]
    assert betexplorer_direct._response_lines(ou_html) == {133.5}
    assert betexplorer_direct._response_lines(ah_html) == {-14.5, -13.5}
    assert betexplorer_direct._source_active_line(ou_html) is None
    assert betexplorer_direct._source_active_line(ah_html) is None

    calls: list[str] = []
    monkeypatch.setattr(betexplorer_direct, "MIN_REQUEST_INTERVAL_SECONDS", 0.0)

    def fetch(url: str) -> str:
        calls.append(url)
        if "/basketball/results/" in url:
            return BOARD
        if f"/match-odds/{EVENT_ID}/1/ha/bestOdds/" in url:
            return HA_ODDS
        if f"/match-odds/{EVENT_ID}/1/ou/bestOdds/" in url:
            return STALE_NAV_OU
        if f"/match-odds/{EVENT_ID}/1/ah/bestOdds/" in url:
            return STALE_NAV_AH
        raise AssertionError(f"unexpected direct BetExplorer URL: {url}")

    snapshots = betexplorer_direct.fetch_odds(
        "NBA", DATE, now_utc=OBSERVED, games=[_game()], text_fetch=fetch
    )
    assert not any(
        snapshot.main_line for snapshot in snapshots if snapshot.market in {Market.TOTAL, Market.SPREAD}
    )
    receipt = betexplorer_direct.last_receipt()["fixture_receipts"][0]
    assert [market["main_line"] for market in receipt["markets"]] == [None, None, None]
    assert all("/1x2/" not in call for call in calls)


def test_direct_fallback_emits_exact_named_moneyline_total_and_spread_rows_with_receipt(monkeypatch):
    calls: list[str] = []
    monkeypatch.setattr(betexplorer_direct, "MIN_REQUEST_INTERVAL_SECONDS", 0.0)
    snapshots = betexplorer_direct.fetch_odds(
        "NBA", DATE, now_utc=OBSERVED, games=[_game()], text_fetch=_fetch(calls)
    )

    assert len(snapshots) == 7
    moneylines = [snapshot for snapshot in snapshots if snapshot.market is Market.MONEYLINE]
    totals = [snapshot for snapshot in snapshots if snapshot.market is Market.TOTAL]
    spreads = [snapshot for snapshot in snapshots if snapshot.market is Market.SPREAD]
    assert {(snapshot.book, snapshot.home_price_decimal, snapshot.away_price_decimal) for snapshot in moneylines} == {
        ("Bet365", 1.72, 2.18),
        ("Pinnacle", 1.74, 2.16),
    }
    assert {(snapshot.book, snapshot.line, snapshot.over_price_decimal, snapshot.under_price_decimal) for snapshot in totals} == {
        ("BetMGM.us", 142.5, 1.91, 1.91),
        ("BetMGM.us", 143.5, 1.83, 2.0),
        ("Stake.com", 143.5, 1.85, 1.95),
    }
    assert {(snapshot.book, snapshot.line, snapshot.home_price_decimal, snapshot.away_price_decimal) for snapshot in spreads} == {
        ("Bet365", -0.5, 1.92, 1.9),
        ("Pinnacle", -1.5, 2.15, 1.7),
    }
    assert all(snapshot.period is Period.FULL_GAME for snapshot in snapshots)
    assert all(snapshot.provider == "betexplorer_direct" for snapshot in snapshots)
    assert all(snapshot.price_provenance is PriceProvenance.NAMED_BOOK for snapshot in snapshots)
    assert all(snapshot.source_permission is SourcePermission.UNRESOLVED for snapshot in snapshots)
    assert all(snapshot.operational_validation is OperationalValidation.PROBATION for snapshot in snapshots)
    assert {snapshot.line for snapshot in totals if snapshot.main_line} == {143.5}
    assert {snapshot.line for snapshot in spreads if snapshot.main_line} == {-0.5}

    receipt = betexplorer_direct.last_receipt()
    assert receipt["status"] == "success"
    assert receipt["requests"] == 4
    assert receipt["fixtures_matched"] == 1
    assert receipt["snapshots_emitted"] == 7
    fixture = receipt["fixture_receipts"][0]
    assert fixture["named_books"] == ["Bet365", "BetMGM.us", "Pinnacle", "Stake.com"]
    assert [(market["market"], market["tab"]) for market in fixture["markets"]] == [
        ("moneyline", "ha"),
        ("total", "ou"),
        ("spread", "ah"),
    ]
    assert all(market["url"] in calls for market in fixture["markets"])


def test_direct_fallback_recovers_the_page_supplied_ha_comparison_parameter(monkeypatch):
    calls: list[str] = []
    monkeypatch.setattr(betexplorer_direct, "MIN_REQUEST_INTERVAL_SECONDS", 0.0)

    def fetch(url: str) -> str:
        calls.append(url)
        if "/basketball/results/" in url:
            return BOARD
        if f"/match-odds/{EVENT_ID}/1/ha/bestOdds/" in url:
            return json.dumps({"odds": "<div>no named rows at page one</div>"})
        if url.endswith(f"/{EVENT_ID}/"):
            return "match_load_tabs('basketball', 'ha', '', '', '2')"
        if f"/match-odds/{EVENT_ID}/2/ha/bestOdds/" in url:
            return HA_ODDS
        if f"/match-odds/{EVENT_ID}/2/ou/bestOdds/" in url:
            return OU_ODDS
        if f"/match-odds/{EVENT_ID}/2/ah/bestOdds/" in url:
            return AH_ODDS
        raise AssertionError(f"unexpected direct BetExplorer URL: {url}")

    snapshots = betexplorer_direct.fetch_odds(
        "NBA", DATE, now_utc=OBSERVED, games=[_game()], text_fetch=fetch
    )

    assert len(snapshots) == 7
    receipt = betexplorer_direct.last_receipt()
    assert receipt["requests"] == 6
    assert receipt["fixture_receipts"][0]["comparison_page_param"] == "2"
    assert receipt["fixture_receipts"][0]["match_page_sha256"] is not None


def test_direct_fallback_recovers_page_from_response_nav_without_match_page(monkeypatch):
    calls: list[str] = []
    monkeypatch.setattr(betexplorer_direct, "MIN_REQUEST_INTERVAL_SECONDS", 0.0)

    def fetch(url: str) -> str:
        calls.append(url)
        if "/basketball/results/" in url:
            return BOARD
        if f"/match-odds/{EVENT_ID}/1/ha/bestOdds/" in url:
            return json.dumps(
                {
                    "odds": (
                        "<li onclick=\"match_change_bettype_best(2, 'AbCdEf12', 'ha', "
                        "'basketball', false, this, 'en', 1);\">H/A</li>"
                    )
                }
            )
        if f"/match-odds/{EVENT_ID}/2/ha/bestOdds/" in url:
            return HA_ODDS
        if f"/match-odds/{EVENT_ID}/2/ou/bestOdds/" in url:
            return OU_ODDS
        if f"/match-odds/{EVENT_ID}/2/ah/bestOdds/" in url:
            return AH_ODDS
        raise AssertionError(f"unexpected direct BetExplorer URL: {url}")

    snapshots = betexplorer_direct.fetch_odds(
        "NBA", DATE, now_utc=OBSERVED, games=[_game()], text_fetch=fetch
    )

    assert len(snapshots) == 7
    assert len(calls) == 5
    assert not any(url.endswith(f"/{EVENT_ID}/") for url in calls)
    assert betexplorer_direct.last_receipt()["fixture_receipts"][0]["comparison_page_param"] == "2"


def test_direct_fallback_stops_the_run_after_a_source_refusal(monkeypatch):
    calls: list[str] = []
    monkeypatch.setattr(betexplorer_direct, "MIN_REQUEST_INTERVAL_SECONDS", 0.0)

    def fetch(url: str) -> str:
        calls.append(url)
        if "/basketball/results/" in url:
            return BOARD
        raise HttpError(429, url, "rate limited")

    assert betexplorer_direct.fetch_odds(
        "NBA", DATE, now_utc=OBSERVED, games=[_game()], text_fetch=fetch
    ) == []
    receipt = betexplorer_direct.last_receipt()
    assert receipt["status"] == "http_429_cooldown"
    assert receipt["cooldown"] is True
    assert receipt["requests"] == 2
    assert len(calls) == 2
    assert any(gap.required_field == "source_access" for gap in betexplorer_direct.last_gaps())


def test_direct_fallback_keeps_partial_prices_and_stops_at_request_cap(monkeypatch):
    calls: list[str] = []
    monkeypatch.setattr(betexplorer_direct, "MIN_REQUEST_INTERVAL_SECONDS", 0.0)
    monkeypatch.setattr(betexplorer_direct, "MAX_REQUESTS_PER_RUN", 3)

    snapshots = betexplorer_direct.fetch_odds(
        "NBA", DATE, now_utc=OBSERVED, games=[_game()], text_fetch=_fetch(calls)
    )

    assert {snapshot.market for snapshot in snapshots} == {Market.MONEYLINE, Market.TOTAL}
    receipt = betexplorer_direct.last_receipt()
    assert receipt["status"] == "request_cap_reached"
    assert receipt["requests"] == 3
    assert len(calls) == 3
    assert any(gap.required_field == "request_budget" for gap in betexplorer_direct.last_gaps())


def test_direct_fallback_stops_after_a_challenge_surface(monkeypatch):
    calls: list[str] = []
    monkeypatch.setattr(betexplorer_direct, "MIN_REQUEST_INTERVAL_SECONDS", 0.0)

    def fetch(url: str) -> str:
        calls.append(url)
        if "/basketball/results/" in url:
            return BOARD
        return "<html><title>Attention Required</title><div>Checking your browser</div></html>"

    assert betexplorer_direct.fetch_odds(
        "NBA", DATE, now_utc=OBSERVED, games=[_game()], text_fetch=fetch
    ) == []
    receipt = betexplorer_direct.last_receipt()
    assert receipt["status"] == "source_challenge"
    assert receipt["cooldown"] is True
    assert receipt["requests"] == 2
    assert len(calls) == 2
    assert any(gap.required_field == "source_access" for gap in betexplorer_direct.last_gaps())


def test_direct_fallback_does_not_fetch_odds_for_an_unmatched_or_past_fixture(monkeypatch):
    calls: list[str] = []
    monkeypatch.setattr(betexplorer_direct, "MIN_REQUEST_INTERVAL_SECONDS", 0.0)
    wrong_orientation = _game()
    wrong_orientation.home_name = "Boston Celtics"
    wrong_orientation.away_name = "Los Angeles Lakers"

    assert betexplorer_direct.fetch_odds(
        "NBA", DATE, now_utc=OBSERVED, games=[wrong_orientation], text_fetch=_fetch(calls)
    ) == []
    assert len(calls) == 1, "only the dated board is read before identity refusal"
    assert betexplorer_direct.last_receipt()["identity_rejections"] == {
        "ordered_team_pair_not_found_on_board_date": 1
    }
    diagnostic = betexplorer_direct.last_receipt()["candidate_fixture_diagnostics"][0]
    assert diagnostic["game_id"] == wrong_orientation.game_id
    assert diagnostic["reason"] == "ordered_team_pair_not_found_on_board_date"

    calls.clear()
    past = _game()
    past.scheduled_utc = "2026-11-20T11:59:00Z"
    assert betexplorer_direct.fetch_odds(
        "NBA", DATE, now_utc=OBSERVED, games=[past], text_fetch=_fetch(calls)
    ) == []
    assert calls == []
    assert betexplorer_direct.last_receipt()["status"] == "no_future_candidate_games"


def test_direct_fallback_fetches_each_required_paris_board_date(monkeypatch):
    game = _game()
    game.scheduled_utc = "2026-11-20T23:30:00Z"  # 2026-11-21 00:30 in Paris
    next_day_board = BOARD.replace('data-dt="20,11,2026,20,00"', 'data-dt="21,11,2026,00,30"')
    calls: list[str] = []
    monkeypatch.setattr(betexplorer_direct, "MIN_REQUEST_INTERVAL_SECONDS", 0.0)

    def fetch(url: str) -> str:
        calls.append(url)
        if "/basketball/results/" in url:
            assert "year=2026&month=11&day=21" in url
            return next_day_board
        return _fetch([])(url)

    snapshots = betexplorer_direct.fetch_odds(
        "NBA", DATE, now_utc=OBSERVED, games=[game], text_fetch=fetch
    )
    assert snapshots
    receipt = betexplorer_direct.last_receipt()
    assert [board["date"] for board in receipt["board_dates"]] == ["2026-11-21"]
    assert receipt["candidate_fixture_diagnostics"] == [
        {
            "game_id": game.game_id,
            "scheduled_utc": game.scheduled_utc,
            "board_date": "2026-11-21",
            "home": game.home_name,
            "away": game.away_name,
            "status": "matched",
            "reason": "exact_ordered_pair_on_paris_board_date",
        }
    ]


def test_registry_exposes_direct_fallback_for_every_league_and_pipeline_keeps_receipt(monkeypatch):
    assert "betexplorer_direct" in league_spec(League.NBA).book_providers
    assert "betexplorer_direct" in league_spec(League.EUROLEAGUE_W).book_providers
    policy = price_evidence_policy(League.NBA_G_LEAGUE, "betexplorer_direct")
    assert policy.price_provenance is PriceProvenance.NAMED_BOOK
    assert policy.source_permission is SourcePermission.UNRESOLVED
    assert policy.operational_validation is OperationalValidation.PROBATION

    calls: list[str] = []
    monkeypatch.setattr(theoddsapi, "is_enabled", lambda: False)
    monkeypatch.setattr(betexplorer_direct, "MIN_REQUEST_INTERVAL_SECONDS", 0.0)
    monkeypatch.setattr(
        betexplorer_direct,
        "_text_fetch",
        lambda _factory, _provided: _fetch(calls),
    )
    store = InMemoryStore()
    game = _game()
    store.upsert_game(game)
    pipeline = Pipeline(
        PipelineConfig(league="NBA", date_yyyymmdd=DATE, fixture_now_utc=OBSERVED),
        store,
        fetch=lambda _url: {},
    )
    pipeline._reset_stage()
    pipeline._acquire_book_lines([game])

    snapshots = store.get_book_snapshots(game.game_id)
    assert len(snapshots) == 7
    assert {snapshot.market for snapshot in snapshots} == {
        Market.MONEYLINE,
        Market.TOTAL,
        Market.SPREAD,
    }
    status = pipeline.book_provider_status["betexplorer_direct"]
    assert status["status"] == "fetched"
    assert status["snapshots_mapped"] == 7
    assert status["receipt"]["fixtures_matched"] == 1


def test_pipeline_skips_direct_request_only_when_all_named_core_markets_exist(monkeypatch):
    monkeypatch.setattr(theoddsapi, "is_enabled", lambda: False)
    store = InMemoryStore()
    game = _game()
    store.upsert_game(game)
    shared = {
        "game_id": game.game_id,
        "provider": "espn",
        "book": "DraftKings",
        "period": Period.FULL_GAME,
        "observed_utc": "2026-11-20T11:00:00Z",
        "scheduled_utc": game.scheduled_utc,
    }
    store.upsert_book_snapshot(
        BookSnapshot(
            market=Market.MONEYLINE,
            home_price_decimal=1.70,
            away_price_decimal=2.20,
            **shared,
        )
    )
    store.upsert_book_snapshot(
        BookSnapshot(
            market=Market.TOTAL,
            line=221.5,
            over_price_decimal=1.91,
            under_price_decimal=1.91,
            main_line=False,
            **shared,
        )
    )
    store.upsert_book_snapshot(
        BookSnapshot(
            market=Market.SPREAD,
            line=-4.5,
            home_price_decimal=1.91,
            away_price_decimal=1.91,
            main_line=False,
            **shared,
        )
    )
    monkeypatch.setattr(
        betexplorer_direct,
        "fetch_odds",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(AssertionError("direct request")),
    )
    pipeline = Pipeline(
        PipelineConfig(league="NBA", date_yyyymmdd=DATE, fixture_now_utc=OBSERVED),
        store,
        fetch=lambda _url: {},
    )
    pipeline._reset_stage()
    pipeline._acquire_book_lines([game])

    status = pipeline.book_provider_status["betexplorer_direct"]
    assert status["status"] == "fallback_not_needed"
    assert status["snapshots_fetched"] == 0
