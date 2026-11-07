"""End-to-end headless pipeline tests on committed fixtures (no network)."""

from __future__ import annotations

import json
import os

import pytest

from ingestion.books import BookConfig
from ingestion.models import (
    ActualResult,
    BookSnapshot,
    Game,
    GameStatus,
    League,
    Market,
    Period,
    PeriodKind,
    Prediction,
)
from ingestion.pipeline import Pipeline, PipelineConfig
from ingestion.store import InMemoryStore

FIXTURES = os.path.join(os.path.dirname(__file__), "fixtures")


def _nba_fetch(
    scoreboard="espn/scoreboard_synth_golden.json",
    predictor="espn/predictor_401909938.json",
    hardwood_path="hardwood/game_forecasts.json",
):
    mapping = {
        "scoreboard": scoreboard,
        "competitions/401909938/predictor": predictor,
        "summary?event=401909938": "espn/summary_401909938_context.json",
        "game_forecasts": hardwood_path,
    }

    def fetch(url):
        for needle, relative in mapping.items():
            if needle in url:
                with open(os.path.join(FIXTURES, relative)) as handle:
                    return json.load(handle)
        from ingestion.http import HttpError

        raise HttpError(404, url, "no fixture mapped")

    return fetch


def _config(**changes):
    values = dict(
        league="NBA",
        date_yyyymmdd="20261107",
        book=BookConfig(provider="none"),
        fixture_now_utc="2026-11-07T18:00:00Z",
    )
    values.update(changes)
    return PipelineConfig(**values)


def _final_game() -> Game:
    return Game(
        game_id="401909938",
        league=League.NBA,
        season="2026",
        home_team_id="11",
        away_team_id="7",
        home_name="Indiana Pacers",
        away_name="Denver Nuggets",
        scheduled_utc="2026-11-07T22:00:00Z",
        status=GameStatus.FINAL,
        actual=ActualResult(
            period_kind=PeriodKind.QUARTERS,
            home_quarters=[28.0, 27.0, 29.0, 26.0],
            away_quarters=[30.0, 31.0, 28.0, 31.0],
            home_final=110.0,
            away_final=120.0,
            overtime_periods=0,
            source="espn",
        ),
    )


def test_pregame_dry_run_returns_canonical_prediction_and_book_gaps():
    store = InMemoryStore()
    pipeline = Pipeline(_config(), store, fetch=_nba_fetch(), allow_writes=False)
    report = pipeline.pregame()

    assert report["fixtures_found"] == 1
    assert report["mode"] == "dry_run"
    assert report["persistence"]["system_of_record"] == "ephemeral"
    row = report["rows"][0]
    assert row["game_id"] == "401909938"
    assert row["prediction_provider"] == "espn"
    assert row["prediction"]["prob_home"] == pytest.approx(0.23062)
    assert row["prediction"]["prob_away"] == pytest.approx(0.76938)
    assert row["prediction"]["expected_total"] == pytest.approx(230.586)
    assert "prob_pct" not in row
    assert row["game_id"] == "401909938"
    assert row["full_game_line"] is None
    assert all(value is None for value in row["quarters"].values())
    assert row["book_evidence"]["full_game"] is None
    assert any(gap["required_field"].startswith("book_") for gap in report["gaps"])
    assert len(store.get_predictions("401909938")) == 1


def test_pregame_uses_only_fresh_exact_book_evidence():
    store = InMemoryStore()
    store.upsert_book_snapshot(
        BookSnapshot(
            game_id="401909938",
            provider="sportsgameodds",
            book="draftkings",
            market=Market.TOTAL,
            period=Period.Q1,
            observed_utc="2026-11-07T17:00:00Z",
            line=58.5,
            over_price_decimal=1.95,
            under_price_decimal=1.95,
        )
    )
    store.upsert_book_snapshot(
        BookSnapshot(
            game_id="401909938",
            provider="sportsgameodds",
            book="draftkings",
            market=Market.TOTAL,
            period=Period.Q2,
            observed_utc="2026-11-05T17:00:00Z",  # stale at the 24-hour boundary
            line=999.0,
        )
    )
    config = _config(
        book=BookConfig(provider="sportsgameodds", book="draftkings")
    )
    report = Pipeline(config, store, fetch=_nba_fetch()).pregame()
    row = report["rows"][0]
    assert row["quarters"]["Q1"] == 58.5
    assert row["book_evidence"]["Q1"]["provider"] == "sportsgameodds"
    assert row["quarters"]["Q2"] is None
    assert row["book_evidence"]["Q2"] is None


def test_bounded_backfill_acquires_reconciled_history_without_learning_claim():
    store = InMemoryStore()
    calls = []

    def fetch(url):
        calls.append(url)
        assert "dates=20261106" in url
        with open(os.path.join(FIXTURES, "espn/scoreboard_20261105.json")) as handle:
            return json.load(handle)

    pipeline = Pipeline(
        _config(history_days=1, history_request_budget=1), store, fetch=fetch
    )
    report = pipeline.backfill()
    assert report["backfill"]["requested_days"] == 1
    assert report["backfill"]["learning_or_calibration_claimed"] is False
    assert report["backfill"]["dates"] == [
        {
            "date": "20261106",
            "status": "acquired",
            "finals_found": 2,
            "finals_reconciled": 2,
            "canonical_persistence": "dry_run_not_persisted",
        }
    ]
    assert report["finals_written"] == 2
    assert len(store.list_games(league=League.NBA)) == 2
    assert all(game.status is GameStatus.FINAL for game in store.list_games())
    assert len(calls) == 1
    assert report["acquisition"]["network_requests"] == 1


def test_backfill_budget_is_explicitly_bounded():
    store = InMemoryStore()
    calls = []

    def fetch(_url):
        calls.append(1)
        return {"events": []}

    report = Pipeline(
        _config(history_days=5, history_request_budget=2), store, fetch=fetch
    ).backfill()
    assert len(calls) == 2
    assert len(report["backfill"]["dates"]) == 2
    assert any(
        gap["required_field"] == "historical_backfill_budget"
        for gap in report["gaps"]
    )


def test_pipeline_stage_gaps_do_not_accumulate_between_commands():
    store = InMemoryStore()
    pipeline = Pipeline(_config(), store, fetch=_nba_fetch())
    first = pipeline.pregame()
    first_gaps = len(first["gaps"])
    second = pipeline.pregame()
    assert first_gaps > 0
    assert len(second["gaps"]) == first_gaps
    coverage = pipeline.coverage()
    assert coverage["gaps"] == []


def test_freeze_then_results_uses_frozen_pre_tip_prediction_and_price():
    store = InMemoryStore()
    scheduled = _final_game()
    scheduled.status = GameStatus.SCHEDULED
    scheduled.actual = None
    store.upsert_game(scheduled)
    store.upsert_prediction(
        Prediction(
            game_id=scheduled.game_id,
            provider="espn",
            observed_utc="2026-11-07T18:00:00Z",
            prob_home=0.23062,
            prob_away=0.76938,
            expected_home_score=110.54,
            expected_away_score=120.046,
            expected_total=230.586,
            expected_margin=-9.506,
        )
    )
    store.upsert_prediction(
        Prediction(
            game_id=scheduled.game_id,
            provider="espn",
            observed_utc="2026-11-07T23:30:00Z",
            prob_home=0.5,
            prob_away=0.5,
            expected_home_score=100,
            expected_away_score=100,
            expected_total=200,
            expected_margin=0,
            pre_tip=False,
        )
    )
    store.upsert_book_snapshot(
        BookSnapshot(
            game_id=scheduled.game_id,
            provider="sportsgameodds",
            book="draftkings",
            market=Market.MONEYLINE,
            period=Period.FULL_GAME,
            observed_utc="2026-11-07T18:00:00Z",
            home_price_decimal=4.0,
            away_price_decimal=1.25,
        )
    )

    class OfflinePipeline(Pipeline):
        def _fetch_finals(self):
            return [_final_game()]

    config = _config(
        fixture_now_utc="2026-11-08T00:00:00Z",
        book=BookConfig(provider="sportsgameodds", book="draftkings"),
    )
    pipeline = OfflinePipeline(config, store)
    frozen = pipeline.freeze()
    assert frozen["book_snapshots_frozen"] == 1
    report = pipeline.results()
    assert report["finals_written"] == 1
    assert report["frozen_predictions_used"] == 1
    row = report["rows"][0]
    assert row["prediction"]["prob_away"] == 0.76938
    assert row["prediction_frozen"] is True
    assert row["home_final"] == 110
    assert row["away_final"] == 120
    assert row["odds"] == 1.25
    assert row["odds_evidence"]["book"] == "draftkings"


def test_results_quarantines_unreconciled_final():
    game = _final_game()
    game.actual.home_quarters = [28.0, 27.0, 29.0, 0.0]

    class OfflinePipeline(Pipeline):
        def _fetch_finals(self):
            return [game]

    report = OfflinePipeline(_config(), InMemoryStore()).results()
    assert report["finals_written"] == 0
    assert report["rows"] == []
    assert report["quarantined"][0]["reason"].startswith("sum")


def test_coverage_counts_only_reconciled_finals_and_real_quarter_lines():
    store = InMemoryStore()
    pipeline = Pipeline(_config(), store, fetch=_nba_fetch())
    pipeline.pregame()
    unreconciled = _final_game()
    unreconciled.game_id = "bad-final"
    unreconciled.actual.away_final = 999
    store.upsert_game(unreconciled)
    store.upsert_game(_final_game())
    store.upsert_book_snapshot(
        BookSnapshot(
            game_id="401909938",
            provider="sportsgameodds",
            book="draftkings",
            market=Market.TOTAL,
            period=Period.Q1,
            observed_utc="2026-11-07T17:00:00Z",
            line=58.5,
        )
    )
    report = pipeline.coverage()
    assert report["games_in_store"] == 2
    assert report["games_with_prediction"] == 1
    assert report["games_with_full_game_line"] == 0
    assert report["quarter_lines"] == {"Q1": 1, "Q2": 0, "Q3": 0, "Q4": 0}
    assert report["finals_reconciled"] == 1
