"""Cross-layer NCAA half semantics and multi-league routing tests."""

from __future__ import annotations

from copy import deepcopy
from types import SimpleNamespace

import pytest

from forecasting.service import MarketForecastService
from forecasting.types import Forecast, _period_scores, grade_forecast
from ingestion.models import (
    ActualResult, BookSnapshot, FinalStatus, Game, League,
    Market, Period, PeriodKind, Prediction,
)
from ingestion.pipeline import Pipeline, PipelineConfig
from ingestion.providers import espn, isportsapi, sportradar, theoddsapi
from ingestion.store import InMemoryStore


def _event(league=League.NCAA_M):
    return {
        "id": "ncaa-1",
        "season": {"year": 2027},
        "competitions": [{
            "date": "2026-11-08T00:00:00Z",
            "status": {"type": {"state": "pre"}},
            "competitors": [
                {"homeAway": "home", "team": {"id": "100", "displayName": "Home University"}, "score": "0"},
                {"homeAway": "away", "team": {"id": "200", "displayName": "Away College"}, "score": "0"},
            ],
        }],
    }


def _summary():
    return {
        "header": {
            "season": {"year": 2027},
            "competitions": [{
                "date": "2026-11-08T00:00:00Z",
                "status": {"type": {"state": "post"}},
                "competitors": [
                    {
                        "homeAway": "home", "team": {"id": "100", "displayName": "Home University"},
                        "score": "95", "linescores": [{"value": 40}, {"value": 45}, {"value": 10}],
                    },
                    {
                        "homeAway": "away", "team": {"id": "200", "displayName": "Away College"},
                        "score": "83", "linescores": [{"value": 35}, {"value": 40}, {"value": 8}],
                    },
                ],
            }],
        }
    }


@pytest.mark.parametrize("league,slug", [
    (League.NCAA_M, "mens-college-basketball"),
    (League.NCAA_W, "womens-college-basketball"),
])
def test_ncaa_discovery_uses_division_one_league_endpoint(league, slug):
    urls = []
    def fetch(url):
        urls.append(url)
        return {"events": [_event(league)]}
    games = espn.discover_games("20261107", fetch, league)
    assert games[0].league is league
    assert f"/basketball/{slug}/scoreboard" in urls[0]
    assert "groups=50" in urls[0]
    assert "limit=500" in urls[0]


def test_ncaa_halves_and_overtime_reconcile_and_grade_natively():
    scheduled = espn.discover_games(
        "20261107", lambda _url: {"events": [_event()]}, League.NCAA_M
    )[0]
    final = espn.fetch_game_result(
        "ncaa-1", lambda _url: deepcopy(_summary()), scheduled, League.NCAA_M
    )
    actual = final.actual
    assert actual.period_kind is PeriodKind.HALVES
    assert actual.home_quarters == [40, 45]
    assert actual.away_quarters == [35, 40]
    assert actual.overtime_periods == 1
    assert actual.reconciles() == (True, "")
    assert final.final_status() is FinalStatus.AOT
    assert _period_scores(final, Period.H1) == (40, 35)
    assert _period_scores(final, Period.H2) == (45, 40)
    assert _period_scores(final, Period.Q1) is None

    forecast = Forecast(
        game_id=final.game_id, league="NCAA_M", scheduled_utc=final.scheduled_utc,
        home_name=final.home_name, away_name=final.away_name,
        market=Market.TOTAL.value, period=Period.H2.value, selection="over",
        raw_probability=0.6, alternative_probabilities={"over": 0.6, "under": 0.4},
        source="test", line=84,
    )
    assert grade_forecast(forecast, final) == "win"
    forecast.period = Period.Q1.value
    assert grade_forecast(forecast, final) is None
    forecast.market = Market.HIGHEST_QUARTER.value
    forecast.selection = "Q1"
    assert grade_forecast(forecast, final) is None


def test_half_result_requires_exactly_two_values():
    actual = ActualResult(
        period_kind=PeriodKind.HALVES,
        home_quarters=[40, 45, 10], away_quarters=[35, 40, 8],
        home_overtime=10, away_overtime=8,
        home_final=95, away_final=83, overtime_periods=1,
    )
    assert actual.reconciles() == (False, "expected exactly two half values per team")


def _outcome(name, price, point=None):
    value = {"name": name, "price": price}
    if point is not None:
        value["point"] = point
    return value


def _bookmaker(markets):
    return {"key": "draftkings", "title": "DraftKings", "markets": markets}


def test_ncaa_womens_odds_request_halves_not_quarters(monkeypatch):
    monkeypatch.setenv("THE_ODDS_API_KEY", "fixture-key")
    event = {
        "id": "w-ncaa-1", "home_team": "Home University", "away_team": "Away College",
        "commence_time": "2026-11-08T00:00:00Z",
        "bookmakers": [_bookmaker([
            {"key": "h2h", "outcomes": [_outcome("Home University", 1.8), _outcome("Away College", 2.1)]},
            {"key": "totals", "outcomes": [_outcome("Over", 1.9, 140.5), _outcome("Under", 1.9, 140.5)]},
        ])],
    }
    period_event = deepcopy(event)
    period_event["bookmakers"] = [_bookmaker([
        {"key": "totals_h1", "outcomes": [_outcome("Over", 1.9, 68.5), _outcome("Under", 1.9, 68.5)]},
        {"key": "totals_h2", "outcomes": [_outcome("Over", 1.9, 71.5), _outcome("Under", 1.9, 71.5)]},
        {"key": "spreads_h1", "outcomes": [_outcome("Home University", 1.9, -2.5), _outcome("Away College", 1.9, 2.5)]},
        {"key": "h2h_h2", "outcomes": [_outcome("Home University", 1.8), _outcome("Away College", 2.1)]},
    ])]
    urls = []
    def fetch(url):
        urls.append(url)
        if "/events/w-ncaa-1/odds" in url:
            return deepcopy(period_event)
        return [deepcopy(event)]

    snapshots = theoddsapi.fetch_odds(
        "NCAA_W", "20261107", lambda _provider: fetch,
        now_utc="2026-11-07T12:00:00Z", tz_name="America/New_York",
    )
    assert "basketball_wncaab" in urls[0]
    assert "totals_h1" in urls[1]
    assert "totals_h2" in urls[1]
    assert "totals_q1" not in urls[1]
    assert {snapshot.period for snapshot in snapshots} == {
        Period.FULL_GAME, Period.H1, Period.H2
    }
    assert not any(snapshot.period in {Period.Q1, Period.Q2, Period.Q3, Period.Q4} for snapshot in snapshots)


def test_forecast_generation_uses_ncaa_halves_and_skips_highest_quarter():
    game = Game(
        game_id="ncaa-forecast", league=League.NCAA_W, season="2027",
        home_team_id="1", away_team_id="2", home_name="Home", away_name="Away",
        scheduled_utc="2026-11-08T00:00:00Z",
    )
    books = [
        BookSnapshot(
            game_id=game.game_id, provider="theoddsapi", book="book",
            market=market, period=period, observed_utc="2026-11-07T12:00:00Z",
            line=50.0, home_price_decimal=1.9, away_price_decimal=1.9,
            over_price_decimal=1.9, under_price_decimal=1.9,
        )
        for market, period in (
            (Market.TOTAL, Period.FULL_GAME), (Market.TOTAL, Period.H1),
            (Market.TOTAL, Period.H2), (Market.SPREAD, Period.H1),
            (Market.SPREAD, Period.H2),
        )
    ]
    captured = []
    store = SimpleNamespace(is_book_snapshot_frozen=lambda *args: True)
    service = MarketForecastService(store, League.NCAA_W)
    service.model.available_history = lambda *args, **kwargs: []
    service._prediction = lambda *args, **kwargs: (None, False)
    service._books = lambda *args, **kwargs: books
    service.model.period_winner = lambda *args, **kwargs: []
    service.model.lines = lambda _game, _history, snapshots, *args: captured.extend(snapshots) or []
    service.model.highest_quarter = lambda *args: (_ for _ in ()).throw(AssertionError("NCAA HQ forbidden"))
    service._generate(game, [], as_of_utc="2026-11-07T13:00:00Z", historical=False)
    assert {(snapshot.market, snapshot.period) for snapshot in captured} == {
        (Market.TOTAL, Period.FULL_GAME), (Market.TOTAL, Period.H1),
        (Market.TOTAL, Period.H2), (Market.SPREAD, Period.H1),
        (Market.SPREAD, Period.H2),
    }


def test_probability_only_cba_prediction_produces_moneyline_without_fake_scores():
    game = Game(
        game_id="cba-forecast", league=League.CBA, season="26-27",
        home_team_id="10", away_team_id="20", home_name="Beijing Ducks",
        away_name="Shanghai Sharks", scheduled_utc="2026-11-08T11:00:00Z",
    )
    prediction = Prediction(
        game_id=game.game_id, provider="sportradar", observed_utc="2026-11-07T08:00:00Z",
        model="sportradar-pre-match-2way", prob_home=0.55, prob_away=0.45,
    )
    store = InMemoryStore()
    store.upsert_game(game)
    store.upsert_prediction(prediction, frozen=True)
    service = MarketForecastService(store, League.CBA)
    forecasts = service._generate(
        game, [], as_of_utc="2026-11-07T09:00:00Z", historical=False
    )
    assert len(forecasts) == 1
    assert forecasts[0].market == Market.MONEYLINE.value
    assert forecasts[0].selection == "home"
    assert forecasts[0].expected_total is None
    assert forecasts[0].prediction_provider == "sportradar"


def test_cba_pipeline_separates_sportradar_predictions_from_isportsapi_books(monkeypatch):
    game = Game(
        game_id="cba-pipeline", league=League.CBA, season="26-27",
        home_team_id="10", away_team_id="20", home_name="Beijing Ducks",
        away_name="Shanghai Sharks", scheduled_utc="2026-11-07T11:35:00Z",
        provider_ids={"isportsapi": "cba-pipeline"},
    )
    prediction = Prediction(
        game_id=game.game_id, provider="sportradar",
        observed_utc="2026-11-07T07:50:00Z", prob_home=0.61, prob_away=0.39,
        model="sportradar-pre-match-2way", source_revision="sr:sport_event:1",
    )
    book = BookSnapshot(
        game_id=game.game_id, provider="isportsapi", provider_event_id=game.game_id,
        book="Bet365", market=Market.MONEYLINE, period=Period.FULL_GAME,
        observed_utc="2026-11-07T07:55:00Z", home_price_decimal=1.8,
        away_price_decimal=2.1,
    )
    monkeypatch.setattr(isportsapi, "discover_games", lambda *_args: [game])
    monkeypatch.setattr(isportsapi, "is_enabled", lambda: True)
    monkeypatch.setattr(sportradar, "is_enabled", lambda: True)
    probability_calls = []
    book_calls = []

    def probabilities(date, games, fetch, **kwargs):
        probability_calls.append((date, games, kwargs))
        return {game.game_id: prediction}

    def odds(league, date, fetch_factory, **kwargs):
        book_calls.append((league, date, kwargs))
        return [book]

    monkeypatch.setattr(sportradar, "fetch_cba_probabilities", probabilities)
    monkeypatch.setattr(isportsapi, "fetch_odds", odds)
    store = InMemoryStore()
    pipeline = Pipeline(
        PipelineConfig(
            league="CBA", date_yyyymmdd="20261107",
            fixture_now_utc="2026-11-07T08:00:00Z",
        ),
        store,
        fetch=lambda _url, **_kwargs: pytest.fail("unexpected direct fetch"),
    )
    report = pipeline.pregame()

    assert probability_calls[0][0] == "20261107"
    assert probability_calls[0][1] == [game]
    assert book_calls == [("CBA", "20261107", {})]
    assert report["predictions_selected"] == {"sportradar": 1}
    assert report["rows"][0]["prediction_provider"] == "sportradar"
    assert [item.provider for item in store.get_predictions(game.game_id)] == ["sportradar"]
    assert [item.provider for item in store.get_book_snapshots(game.game_id)] == ["isportsapi"]
    assert not hasattr(isportsapi, "fetch_prediction")


def test_ncaa_pipeline_uses_espn_and_native_half_books(monkeypatch):
    game = Game(
        game_id="ncaa-pipeline", league=League.NCAA_W, season="2027",
        home_team_id="100", away_team_id="200", home_name="Home University",
        away_name="Away College", scheduled_utc="2026-11-08T00:00:00Z",
        provider_ids={"espn": "ncaa-pipeline"},
    )
    prediction = Prediction(
        game_id=game.game_id, provider="espn", observed_utc="2026-11-07T11:00:00Z",
        prob_home=0.57, prob_away=0.43,
    )
    books = [
        BookSnapshot(
            game_id=game.game_id, provider="theoddsapi", provider_event_id="odds-ncaa-pipeline",
            book="DraftKings", market=Market.TOTAL, period=period,
            observed_utc="2026-11-07T11:30:00Z", line=line,
            over_price_decimal=1.9, under_price_decimal=1.9,
            home_name=game.home_name, away_name=game.away_name,
            scheduled_utc=game.scheduled_utc,
        )
        for period, line in ((Period.H1, 68.5), (Period.H2, 71.5))
    ]
    monkeypatch.setattr(espn, "discover_games", lambda *_args: [game])
    monkeypatch.setattr(espn, "fetch_bpi", lambda *_args, **_kwargs: prediction)
    monkeypatch.setattr(theoddsapi, "is_enabled", lambda: True)
    calls = []

    def odds(league, date, fetch_factory, **kwargs):
        calls.append((league, date, kwargs))
        return books

    monkeypatch.setattr(theoddsapi, "fetch_odds", odds)
    monkeypatch.setattr(isportsapi, "is_enabled", lambda: pytest.fail("CBA provider called"))
    store = InMemoryStore()
    report = Pipeline(
        PipelineConfig(
            league="NCAA_W", date_yyyymmdd="20261107",
            fixture_now_utc="2026-11-07T12:00:00Z",
        ),
        store,
        fetch=lambda _url, **_kwargs: pytest.fail("unexpected direct fetch"),
    ).pregame()

    assert calls == [("NCAA_W", "20261107", {"tz_name": "America/New_York"})]
    assert report["rows"][0]["prediction_provider"] == "espn"
    assert report["rows"][0]["quarters"] == {}
    assert report["rows"][0]["halves"] == {"H1": 68.5, "H2": 71.5}
    assert report["rows"][0]["native_periods"] == {"H1": 68.5, "H2": 71.5}


def test_cba_and_ncaa_results_use_their_native_provider(monkeypatch):
    cba_final = Game(
        game_id="cba-final", league=League.CBA, season="26-27",
        home_team_id="10", away_team_id="20", home_name="Home", away_name="Away",
        scheduled_utc="2026-11-07T11:35:00Z",
    )
    ncaa_final = Game(
        game_id="ncaa-final", league=League.NCAA_M, season="2027",
        home_team_id="100", away_team_id="200", home_name="Home", away_name="Away",
        scheduled_utc="2026-11-07T23:00:00Z",
    )
    monkeypatch.setattr(isportsapi, "fetch_recent_results", lambda *_args: [cba_final])
    monkeypatch.setattr(espn, "fetch_recent_results", lambda *_args: [ncaa_final])
    cba_pipeline = Pipeline(
        PipelineConfig(league="CBA", date_yyyymmdd="20261107"), InMemoryStore()
    )
    ncaa_pipeline = Pipeline(
        PipelineConfig(league="NCAA_M", date_yyyymmdd="20261107"), InMemoryStore()
    )
    assert cba_pipeline._fetch_finals() == [cba_final]
    assert ncaa_pipeline._fetch_finals() == [ncaa_final]
