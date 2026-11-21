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

ODDS = json.dumps(
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
        if f"/match-odds/{EVENT_ID}/1/1x2/bestOdds/" in url:
            return ODDS
        raise AssertionError(f"unexpected direct BetExplorer URL: {url}")

    return fetch


def test_day_board_and_named_row_parsers_refuse_other_sports_and_averages():
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

    rows = betexplorer_direct.parse_named_moneyline_rows(json.loads(ODDS)["odds"])
    assert rows == [
        {"book": "Bet365", "home_price_decimal": 1.72, "away_price_decimal": 2.18},
        {"book": "Pinnacle", "home_price_decimal": 1.74, "away_price_decimal": 2.16},
    ]


def test_direct_fallback_emits_only_explicit_named_book_rows_with_a_receipt(monkeypatch):
    calls: list[str] = []
    monkeypatch.setattr(betexplorer_direct, "MIN_REQUEST_INTERVAL_SECONDS", 0.0)
    snapshots = betexplorer_direct.fetch_odds(
        "NBA",
        DATE,
        now_utc=OBSERVED,
        games=[_game()],
        text_fetch=_fetch(calls),
    )

    assert len(snapshots) == 2
    by_book = {snapshot.book: snapshot for snapshot in snapshots}
    assert set(by_book) == {"Bet365", "Pinnacle"}
    assert (
        by_book["Bet365"].home_price_decimal,
        by_book["Bet365"].away_price_decimal,
    ) == (1.72, 2.18)
    assert by_book["Pinnacle"].scheduled_utc == "2026-11-20T19:00:00Z"
    assert all(snapshot.provider == "betexplorer_direct" for snapshot in snapshots)
    assert all(snapshot.price_provenance is PriceProvenance.NAMED_BOOK for snapshot in snapshots)
    assert all(snapshot.source_permission is SourcePermission.UNRESOLVED for snapshot in snapshots)
    assert all(snapshot.operational_validation is OperationalValidation.PROBATION for snapshot in snapshots)

    receipt = betexplorer_direct.last_receipt()
    assert receipt["status"] == "success"
    assert receipt["requests"] == 2
    assert receipt["fixtures_matched"] == 1
    assert receipt["snapshots_emitted"] == 2
    assert receipt["fixture_receipts"][0]["odds_url"] in calls
    assert receipt["fixture_receipts"][0]["named_books"] == ["Bet365", "Pinnacle"]


def test_direct_fallback_recovers_the_page_supplied_comparison_parameter(monkeypatch):
    calls: list[str] = []
    monkeypatch.setattr(betexplorer_direct, "MIN_REQUEST_INTERVAL_SECONDS", 0.0)

    def fetch(url: str) -> str:
        calls.append(url)
        if "/basketball/results/" in url:
            return BOARD
        if f"/match-odds/{EVENT_ID}/1/1x2/bestOdds/" in url:
            return json.dumps({"odds": "<div>no named rows at page one</div>"})
        if url.endswith(f"/{EVENT_ID}/"):
            return "match_load_tabs('basketball', '1x2', '', '', '2')"
        if f"/match-odds/{EVENT_ID}/2/1x2/bestOdds/" in url:
            return ODDS
        raise AssertionError(f"unexpected direct BetExplorer URL: {url}")

    snapshots = betexplorer_direct.fetch_odds(
        "NBA", DATE, now_utc=OBSERVED, games=[_game()], text_fetch=fetch
    )

    assert {snapshot.book for snapshot in snapshots} == {"Bet365", "Pinnacle"}
    receipt = betexplorer_direct.last_receipt()
    assert receipt["requests"] == 4
    assert receipt["fixture_receipts"][0]["comparison_page_param"] == "2"
    assert receipt["fixture_receipts"][0]["match_page_sha256"] is not None


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
        "no_unique_ordered_team_pair_on_dated_board": 1
    }

    calls.clear()
    past = _game()
    past.scheduled_utc = "2026-11-20T11:59:00Z"
    assert betexplorer_direct.fetch_odds(
        "NBA", DATE, now_utc=OBSERVED, games=[past], text_fetch=_fetch(calls)
    ) == []
    assert calls == []
    assert betexplorer_direct.last_receipt()["status"] == "no_future_candidate_games"


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
    assert {snapshot.book for snapshot in snapshots} == {"Bet365", "Pinnacle"}
    status = pipeline.book_provider_status["betexplorer_direct"]
    assert status["status"] == "fetched"
    assert status["snapshots_mapped"] == 2
    assert status["receipt"]["fixtures_matched"] == 1


def test_pipeline_skips_direct_request_when_a_named_moneyline_already_exists(monkeypatch):
    monkeypatch.setattr(theoddsapi, "is_enabled", lambda: False)
    store = InMemoryStore()
    game = _game()
    store.upsert_game(game)
    store.upsert_book_snapshot(
        BookSnapshot(
            game_id=game.game_id,
            provider="espn",
            book="DraftKings",
            market=Market.MONEYLINE,
            period=Period.FULL_GAME,
            observed_utc="2026-11-20T11:00:00Z",
            home_price_decimal=1.70,
            away_price_decimal=2.20,
            scheduled_utc=game.scheduled_utc,
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
