from __future__ import annotations

from datetime import datetime, timedelta, timezone
import itertools
import math

import pytest

from forecasting.accas import build_accas
from forecasting.harness import CalibrationConfig, apply_profiles, build_profiles
from forecasting.model import ForecastConfig, HistoryModel
from forecasting.service import MarketForecastService
from forecasting.types import BookEvidence, Forecast, grade_forecast
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
from ingestion.store import InMemoryStore

UTC = timezone.utc
AS_OF = datetime(2026, 1, 30, 20, tzinfo=UTC)


def iso(value: datetime) -> str:
    return value.isoformat().replace("+00:00", "Z")


def _game(game_id: str, tip: datetime, final: bool, index: int) -> Game:
    home_quarters = [28, 27, 29, 26]
    away_quarters = [22 + index % 2, 23, 21, 24]
    actual = None
    if final:
        actual = ActualResult(
            period_kind=PeriodKind.QUARTERS,
            home_quarters=home_quarters,
            away_quarters=away_quarters,
            home_final=sum(home_quarters),
            away_final=sum(away_quarters),
            source="official-test",
        )
    return Game(
        game_id=game_id,
        league=League.NBA,
        season="2026",
        home_team_id=f"h{index % 4}",
        away_team_id=f"a{index % 4}",
        home_name=f"Home {index % 4}",
        away_name=f"Away {index % 4}",
        scheduled_utc=iso(tip),
        status=GameStatus.FINAL if final else GameStatus.SCHEDULED,
        actual=actual,
    )


def _snapshot(
    game_id: str,
    market: Market,
    period: Period,
    observed: datetime,
    **values,
) -> BookSnapshot:
    return BookSnapshot(
        game_id=game_id,
        provider="theoddsapi",
        book="testbook",
        market=market,
        period=period,
        observed_utc=iso(observed),
        **values,
    )


def populated_store(*, finals: int = 12, upcoming: int = 3) -> InMemoryStore:
    store = InMemoryStore()
    start = datetime(2026, 1, 1, 22, tzinfo=UTC)
    for index in range(finals + upcoming):
        final = index < finals
        tip = (
            start + timedelta(days=2 * index)
            if final
            else AS_OF + timedelta(hours=2 + 3 * (index - finals))
        )
        game_id = f"g{index}"
        game = _game(game_id, tip, final, index)
        store.upsert_game(game)
        observed = tip - timedelta(hours=2) if final else AS_OF - timedelta(hours=2)
        store.upsert_prediction(
            Prediction(
                game_id=game_id,
                provider="espn",
                observed_utc=iso(observed),
                prob_home=0.65,
                prob_away=0.35,
                expected_home_score=110,
                expected_away_score=90,
                expected_margin=20,
                expected_total=200,
            ),
            frozen=final,
        )
        book_observed = (
            tip - timedelta(hours=1) if final else AS_OF - timedelta(hours=1)
        )
        snapshots = [
            _snapshot(
                game_id,
                Market.MONEYLINE,
                Period.FULL_GAME,
                book_observed,
                home_price_decimal=1.7,
                away_price_decimal=2.2,
            ),
            _snapshot(
                game_id,
                Market.MONEYLINE,
                Period.H1,
                book_observed,
                home_price_decimal=1.8,
                away_price_decimal=2.1,
                draw_price_decimal=12.0,
            ),
            _snapshot(
                game_id,
                Market.TOTAL,
                Period.FULL_GAME,
                book_observed,
                line=195.5,
                over_price_decimal=1.91,
                under_price_decimal=1.91,
            ),
        ]
        for period in (Period.Q1, Period.Q2, Period.Q3, Period.Q4):
            snapshots.extend(
                [
                    _snapshot(
                        game_id,
                        Market.TOTAL,
                        period,
                        book_observed,
                        line=45.5,
                        over_price_decimal=1.91,
                        under_price_decimal=1.91,
                    ),
                    _snapshot(
                        game_id,
                        Market.SPREAD,
                        period,
                        book_observed,
                        line=-1.5,
                        home_price_decimal=1.91,
                        away_price_decimal=1.91,
                    ),
                ]
            )
        for snapshot in snapshots:
            store.upsert_book_snapshot(snapshot, frozen=final)
    return store


def service(store: InMemoryStore) -> MarketForecastService:
    return MarketForecastService(
        store,
        League.NBA,
        forecast_config=ForecastConfig(
            min_history_games=3,
            min_residual_games=3,
            training_window=50,
            residual_window=30,
        ),
        calibration_config=CalibrationConfig(min_samples=3, prior_strength=2),
    )


def test_all_old_markets_are_forecast_and_accas_use_calibrated_probabilities():
    report = service(populated_store()).build_report(
        as_of_utc=iso(AS_OF), acca_legs=2, acca_limit=4
    )
    current = [
        row for row in report["forecasts"] if row["game_id"] == "g12"
    ]
    markets = {(row["period"], row["market"]) for row in current}
    assert ("full_game", "moneyline") in markets
    assert ("H1", "moneyline") in markets
    assert ("full_game", "total") in markets
    assert ("full_game", "highest_quarter") in markets
    full_game_moneyline_sources = {
        row["model"]
        for row in current
        if row["market"] == "moneyline" and row["period"] == "full_game"
    }
    assert full_game_moneyline_sources == {
        "provider:espn",
        "team_history_empirical_v1",
    }
    for period in ("Q1", "Q2", "Q3", "Q4"):
        assert (period, "total") in markets
        assert (period, "spread") in markets

    assert report["walk_forward"]["method"] == "prequential_by_scheduled_tip"
    assert report["walk_forward"]["graded_forecasts"] > 0
    assert report["walk_forward"]["segments"]
    assert report["walk_forward"]["raw_metrics"]["brier_score"] is not None
    accas = report["accas"]["results"]
    assert accas
    probabilities = [row["estimated_joint_win_probability"] for row in accas]
    assert probabilities == sorted(probabilities, reverse=True)
    for acca in accas:
        assert len({leg["game_id"] for leg in acca["legs"]}) == 2
        assert all(leg["book"]["selected_decimal_price"] for leg in acca["legs"])
        assert acca["ranking_metric"] == "calibrated_joint_win_probability"
        assert acca["ranking_status"] == "highest_estimated_not_guaranteed"


def test_full_game_history_moneyline_earns_calibration_without_provider_history():
    store = populated_store(finals=12, upcoming=3)
    # This reproduces the live coverage gap: exact current books exist, but no
    # frozen provider-prediction archive exists from which to build calibration.
    store.predictions = {}
    store.frozen_predictions = {}
    # Remove every other priced market so a returned acca must use the new
    # full-game history surface rather than passing through H1/totals/spreads.
    for game_id, snapshots in list(store.books.items()):
        kept = {
            key: snapshot
            for key, snapshot in snapshots.items()
            if snapshot.market is Market.MONEYLINE
            and snapshot.period is Period.FULL_GAME
        }
        store.books[game_id] = kept
        store.frozen_books[game_id] = {
            key: value
            for key, value in store.frozen_books.get(game_id, {}).items()
            if key in kept
        }

    report = service(store).build_report(
        as_of_utc=iso(AS_OF), acca_legs=2, acca_limit=4
    )
    rows = [
        row
        for row in report["forecasts"]
        if row["market"] == "moneyline"
        and row["period"] == "full_game"
        and row["model"] == "team_history_empirical_v1"
    ]
    assert len(rows) == 3
    for row in rows:
        assert row["prediction"] is None
        assert row["training_games"] >= 3
        assert row["residual_games"] >= 3
        assert row["calibration"]["n"] >= 3
        assert row["calibration"]["sufficient_for_accas"] is True
        assert row["book"]["selected_decimal_price"] == 1.7
        assert row["eligibility"]["acca_eligible"] is True
        assert row["eligibility"]["reasons"] == []

    segment = next(
        row
        for row in report["walk_forward"]["segments"]
        if row["league"] == "NBA"
        and row["market"] == "moneyline"
        and row["period"] == "full_game"
        and row["source"] == "team_history_empirical_v1"
    )
    assert segment["n"] >= 3
    assert report["accas"]["results"]
    assert report["accas"]["eligible_games"] == 3
    for acca in report["accas"]["results"]:
        for leg in acca["legs"]:
            assert leg["market"] == "moneyline"
            assert leg["period"] == "full_game"
            assert leg["model"] == "team_history_empirical_v1"


def test_expected_total_never_becomes_a_book_line():
    store = populated_store(upcoming=1)
    store.books["g12"] = {}
    store.frozen_books["g12"] = {}
    report = service(store).build_report(as_of_utc=iso(AS_OF))
    current = [row for row in report["forecasts"] if row["game_id"] == "g12"]
    assert not any(row["market"] == "total" for row in current)
    moneyline = next(row for row in current if row["market"] == "moneyline" and row["period"] == "full_game")
    assert moneyline["book"] is None
    assert "no_exact_bookmaker_market" in moneyline["eligibility"]["reasons"]


def test_quarter_line_is_not_synthesized_for_other_periods():
    store = populated_store(upcoming=1)
    finals = store.list_games(league=League.NBA, status=GameStatus.FINAL)
    game = store.get_game("g12")
    book = _snapshot(
        game.game_id,
        Market.TOTAL,
        Period.Q1,
        AS_OF - timedelta(hours=1),
        line=45.5,
        over_price_decimal=1.91,
        under_price_decimal=1.91,
    )
    rows = HistoryModel(
        ForecastConfig(
            min_history_games=3,
            min_residual_games=3,
            training_window=50,
            residual_window=30,
        )
    ).lines(game, finals, [book], {})
    assert rows
    assert {row.period for row in rows} == {Period.Q1.value}


def test_history_cutoff_excludes_future_and_not_yet_available_finals():
    store = populated_store(finals=4, upcoming=1)
    model = HistoryModel(ForecastConfig(result_available_after_hours=4))
    games = store.list_games(league=League.NBA)
    target = store.get_game("g4")
    history = model.available_history(
        games, scheduled_utc=target.scheduled_utc, as_of_utc=iso(AS_OF)
    )
    assert history
    for game in history:
        assert (
            datetime.fromisoformat(game.scheduled_utc.replace("Z", "+00:00"))
            <= datetime.fromisoformat(target.scheduled_utc.replace("Z", "+00:00"))
            - timedelta(hours=4)
        )


def _forecast(game_id: str, probability: float, *, outcome=None) -> Forecast:
    return Forecast(
        game_id=game_id,
        league="NBA",
        scheduled_utc="2026-01-01T22:00:00Z",
        home_name="Home",
        away_name="Away",
        market="moneyline",
        period="full_game",
        selection="home",
        raw_probability=probability,
        alternative_probabilities={"home": probability, "away": 1 - probability},
        source="provider:test",
        book=BookEvidence(
            provider="test",
            book="book",
            observed_utc="2026-01-01T20:00:00Z",
            source_modified_utc=None,
            pre_match=True,
            main_line=True,
            frozen=True,
            prices={"home": 1.8},
        ),
        outcome=outcome,
    )


def test_market_accuracy_adjusts_confidence_and_eligibility():
    records = [
        _forecast(f"g{index}", 0.8, outcome="loss") for index in range(10)
    ]
    config = CalibrationConfig(min_samples=5, prior_strength=2)
    profiles = build_profiles(records, config)
    profile = next(iter(profiles.values()))
    assert profile.observed_win_rate == 0
    assert profile.adjustment < 0
    calibrated = apply_profiles(
        [_forecast("new", 0.8)], profiles, config
    )[0]
    assert calibrated.calibrated_probability < 0.8
    assert not calibrated.eligibility["acca_eligible"]


def test_acca_search_matches_brute_force_optimum():
    probabilities = [0.91, 0.55, 0.82, 0.74, 0.63, 0.88]
    forecasts = []
    for index, probability in enumerate(probabilities):
        row = _forecast(f"g{index}", probability)
        row.calibrated_probability = probability
        row.eligibility = {"acca_eligible": True}
        forecasts.append(row)
    actual = build_accas(forecasts, legs=3, limit=5)
    expected = []
    for combo in itertools.combinations(forecasts, 3):
        expected.append((math.prod(row.calibrated_probability for row in combo), combo))
    expected.sort(key=lambda item: item[0], reverse=True)
    assert len(actual) == 5
    for result, (probability, _) in zip(actual, expected[:5]):
        assert result["estimated_joint_win_probability"] == pytest.approx(
            probability, abs=1e-7
        )


def test_invalid_calibrated_probabilities_never_reach_acca_logarithms():
    valid = [_forecast("g0", 0.8), _forecast("g1", 0.7)]
    for row in valid:
        row.calibrated_probability = row.raw_probability
        row.eligibility = {"acca_eligible": True}
    invalid = _forecast("g2", 0.99)
    invalid.calibrated_probability = 0.0
    invalid.eligibility = {"acca_eligible": True}
    result = build_accas([*valid, invalid], legs=2, limit=1)
    assert len(result) == 1
    assert {leg["game_id"] for leg in result[0]["legs"]} == {"g0", "g1"}


def test_push_and_highest_quarter_tie_are_graded_explicitly():
    game = _game("result", datetime(2026, 1, 1, 22, tzinfo=UTC), True, 0)
    total = _forecast("result", 0.5)
    total.market = "total"
    total.period = "Q1"
    total.selection = "over"
    total.line = 50.0
    assert grade_forecast(total, game) == "push"

    highest = _forecast("result", 0.3)
    highest.market = "highest_quarter"
    highest.selection = "Q1"
    game.actual.away_quarters[0] = 28
    game.actual.away_quarters[1] = 29  # Q1 and Q2 both total 56.
    game.actual.away_final = sum(game.actual.away_quarters)
    assert grade_forecast(highest, game) == "win"
