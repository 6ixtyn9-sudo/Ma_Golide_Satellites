"""The prediction wall, broken by the lane-local model instead of a feed.

Run 45250804751: Italy LBA logged six fixtures and a fresh named full-game
moneyline on every one, and still reported ``games_with_any_pre_tip_prediction:
0``. The cause was structural -- the quota-scoped external feed is scoped to
NBA/EUROLEAGUE and 200 SHARED calls a month cannot be widened to ~20 lanes.

``MarketForecastService.history_prediction`` closes that gap without a
credential: it turns the reconciled finals a lane already stores into one
empirical full-game moneyline, complete by ``is_complete``. When it cannot, it
declines *by name* with the numbers behind the decline -- never an empty dict.

These tests are hermetic: synthetic finals only, no network, no fixtures.
"""

from __future__ import annotations

from datetime import datetime, timedelta, timezone

from ingestion.models import ActualResult, Game, GameStatus, League, PeriodKind
from ingestion.pipeline import Pipeline, PipelineConfig
from ingestion.predictions import is_complete
from ingestion.store import InMemoryStore

LEAGUE = League.ITALY_LBA
TEAMS = ("Alpha Milano", "Beta Bologna")


def _final(day: int, *, home_margin: int) -> Game:
    """One reconciled final on ``2026-10-01 + day``, four quarters each."""

    if home_margin >= 0:
        home_quarters = [24.0, 24.0, 24.0, 24.0]
        away_quarters = [22.0, 22.0, 22.0, 22.0]
    else:
        home_quarters = [21.0, 21.0, 21.0, 22.0]
        away_quarters = [22.0, 22.0, 23.0, 23.0]
    scheduled = datetime(2026, 10, 1, 19, 0, tzinfo=timezone.utc) + timedelta(days=day)
    return Game(
        game_id=f"LBA-HIST-{day:03d}",
        league=LEAGUE,
        season="2026",
        home_team_id="ALPHA",
        away_team_id="BETA",
        home_name=TEAMS[0],
        away_name=TEAMS[1],
        scheduled_utc=scheduled.strftime("%Y-%m-%dT%H:%M:%SZ"),
        status=GameStatus.FINAL,
        provider_ids={"lba": f"hist-{day}"},
        actual=ActualResult(
            period_kind=PeriodKind.QUARTERS,
            home_quarters=home_quarters,
            away_quarters=away_quarters,
            home_final=sum(home_quarters),
            away_final=sum(away_quarters),
        ),
    )


def _history(count: int, *, start: int = 0) -> list[Game]:
    """``count`` walk-forward finals: two home wins for every home loss."""

    return [
        _final(start + index, home_margin=8 if index % 3 else -5)
        for index in range(count)
    ]


def _fixture(day: int = 61) -> Game:
    scheduled = datetime(2026, 10, 1, 19, 0, tzinfo=timezone.utc) + timedelta(days=day)
    return Game(
        game_id="LBA-FIXTURE-1",
        league=LEAGUE,
        season="2026",
        home_team_id="ALPHA",
        away_team_id="BETA",
        home_name=TEAMS[0],
        away_name=TEAMS[1],
        scheduled_utc=scheduled.strftime("%Y-%m-%dT%H:%M:%SZ"),
        status=GameStatus.SCHEDULED,
        provider_ids={"lba": "fixture-1"},
    )


def _pipeline(store: InMemoryStore) -> Pipeline:
    config = PipelineConfig(
        league=LEAGUE.value,
        date_yyyymmdd="20261130",
        fixture_now_utc="2026-11-30T12:00:00Z",
    )
    return Pipeline(config, store, allow_writes=False)


def test_stored_history_produces_a_complete_pre_tip_candidate():
    store = InMemoryStore()
    for game in _history(44):
        store.upsert_game(game)
    fixture = _fixture()
    pipeline = _pipeline(store)

    selections = pipeline._acquire_predictions([fixture])

    selection = selections[fixture.game_id]
    assert selection.selected is not None, "44 reconciled finals must be enough"
    assert selection.selected.provider == "history_model"
    assert is_complete(selection.selected)
    assert selection.selected.pre_tip is True
    assert selection.selected.expected_total is not None
    # The bundle is a model expectation, split arithmetically, not a book line.
    assert abs(
        selection.selected.expected_home_score
        + selection.selected.expected_away_score
        - selection.selected.expected_total
    ) < 1e-9
    # And it reached the store, so the paper pipeline can freeze it.
    assert [p.provider for p in store.get_predictions(fixture.game_id)] == ["history_model"]
    status = pipeline.prediction_provider_status["history_model"]
    assert status["finals_in_store"] == 44
    assert status["candidates_returned"] == 1
    assert status["decline_reasons"] == {}
    assert [g for g in pipeline.gaps if g.required_field == "pre_tip_prediction"] == []


def test_an_empty_store_declines_by_name_instead_of_returning_nothing():
    store = InMemoryStore()
    fixture = _fixture()
    pipeline = _pipeline(store)

    selections = pipeline._acquire_predictions([fixture])

    assert selections[fixture.game_id].selected is None
    gap = next(g for g in pipeline.gaps if g.required_field == "pre_tip_prediction")
    assert gap.provider == "history_model"
    assert "no_reconciled_finals_in_store" in gap.detail
    status = pipeline.prediction_provider_status["history_model"]
    assert status["finals_in_store"] == 0
    assert status["candidates_returned"] == 0
    assert status["decline_reasons"] == {"no_reconciled_finals_in_store (finals_in_store=0)": 1}
    # The lane still holds a source: the slate-level "never" gap must not fire.
    assert [g for g in pipeline.gaps if "no prediction source is wired" in str(g.detail)] == []


def test_thin_history_declines_with_the_numbers_behind_it():
    store = InMemoryStore()
    for game in _history(10):
        store.upsert_game(game)
    fixture = _fixture()
    pipeline = _pipeline(store)

    selections = pipeline._acquire_predictions([fixture])

    assert selections[fixture.game_id].selected is None
    gap = next(g for g in pipeline.gaps if g.required_field == "pre_tip_prediction")
    assert "model_declined_insufficient_history" in gap.detail
    assert "available_before_tip=10" in gap.detail
    assert "min_history_games=20" in gap.detail


def test_finals_after_the_tip_are_not_history():
    """No look-ahead: a store full of later games answers nothing."""

    store = InMemoryStore()
    for game in _history(44, start=70):
        store.upsert_game(game)
    fixture = _fixture()
    pipeline = _pipeline(store)

    selections = pipeline._acquire_predictions([fixture])

    assert selections[fixture.game_id].selected is None
    gap = next(g for g in pipeline.gaps if g.required_field == "pre_tip_prediction")
    assert "no_history_before_tip" in gap.detail
    assert "finals_in_store=44" in gap.detail


def test_espn_lanes_keep_their_own_provider_surface():
    store = InMemoryStore()
    for game in _history(44):
        store.upsert_game(game)
    ncaa = Pipeline(
        PipelineConfig(
            league="NCAA_M",
            date_yyyymmdd="20261130",
            fixture_now_utc="2026-11-30T12:00:00Z",
        ),
        store,
        allow_writes=False,
    )
    fixture = Game(
        game_id="NCAA-1",
        league=League.NCAA_M,
        season="2026",
        home_team_id="H",
        away_team_id="A",
        home_name="Home",
        away_name="Away",
        scheduled_utc="2026-12-01T19:00:00Z",
        status=GameStatus.SCHEDULED,
        provider_ids={},
    )

    candidates = ncaa._history_model_candidates_for_games([fixture])

    assert candidates == {fixture.game_id: []}
    assert "history_model" not in ncaa.prediction_provider_status
    assert "history_model" not in ncaa._prediction_sources_for(League.NCAA_M)


def test_a_game_another_source_answers_gets_no_decline_noise():
    """EUROLEAGUE already has a feed; the local model must not nag per game."""

    from ingestion.models import Prediction

    store = InMemoryStore()
    fixture = Game(
        game_id="E2026-999",
        league=League.EUROLEAGUE,
        season="2026",
        home_team_id="RMB",
        away_team_id="FCB",
        home_name="Real Madrid",
        away_name="Barcelona",
        scheduled_utc="2026-12-01T19:45:00Z",
        status=GameStatus.SCHEDULED,
        provider_ids={"euroleague": "999", "euroleague_season": "E2026"},
    )
    store.upsert_game(fixture)
    store.upsert_prediction(
        Prediction(
            game_id=fixture.game_id,
            provider="predixsport",
            observed_utc="2026-11-30T10:00:00Z",
            prob_home=0.55,
            prob_away=0.45,
            expected_home_score=84.0,
            expected_away_score=81.0,
            expected_total=165.0,
        )
    )
    pipeline = Pipeline(
        PipelineConfig(
            league="EUROLEAGUE",
            date_yyyymmdd="20261130",
            fixture_now_utc="2026-11-30T12:00:00Z",
        ),
        store,
        allow_writes=False,
    )

    selections = pipeline._acquire_predictions([fixture])

    assert selections[fixture.game_id].selected.provider == "predixsport"
    # The local model still declined (no finals in store), and says so in its
    # receipt, but it does not file a gap against a game already answered.
    status = pipeline.prediction_provider_status["history_model"]
    assert status["candidates_returned"] == 0
    assert status["decline_reasons"] == {"no_reconciled_finals_in_store (finals_in_store=0)": 1}
    assert [g for g in pipeline.gaps if g.provider == "history_model"] == []


def test_the_market_service_reads_the_stored_candidate_back():
    """Wiring check: the paper surface sees exactly what the pipeline stored.

    Producing a prediction is not enough -- the reporting/paper path must pick
    it up as a candidate. ``history_model`` is outside
    ``DEFAULT_PROVIDER_ORDER``, so this pins that unordered providers are still
    admitted rather than silently dropped.
    """

    from forecasting.service import MarketForecastService

    store = InMemoryStore()
    for game in _history(44):
        store.upsert_game(game)
    fixture = _fixture()
    store.upsert_game(fixture)
    pipeline = _pipeline(store)
    selection = pipeline._acquire_predictions([fixture])[fixture.game_id]

    service = MarketForecastService(store, LEAGUE)
    candidates = service._predictions(
        fixture, as_of_utc="2026-11-30T12:30:00Z", historical=False
    )

    assert [pred.provider for pred, _frozen in candidates] == ["history_model"]
    assert candidates[0][0].observed_utc == selection.selected.observed_utc
    assert candidates[0][0].prob_home == selection.selected.prob_home
