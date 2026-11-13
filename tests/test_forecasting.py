from __future__ import annotations

from datetime import datetime, timedelta, timezone
import itertools
import math

import pytest

from forecasting.accas import build_accas
from forecasting.harness import CalibrationConfig, apply_profiles, build_profiles
from forecasting.model import ForecastConfig, HistoryModel
import forecasting.service as forecasting_service_module
from forecasting.service import MarketForecastService
from forecasting.types import BookEvidence, Forecast, grade_forecast, opposite_selection
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


def test_opposite_selection_only_returns_unique_outcomes():
    assert opposite_selection("moneyline", "full_game", "home") == "away"
    assert opposite_selection("moneyline", "H1", "home") is None
    assert opposite_selection("moneyline", "full_game", "home", has_draw=True) is None
    assert opposite_selection("spread", "Q1", "home") == "away"
    assert opposite_selection("total", "full_game", "over") == "under"
    assert opposite_selection("highest_quarter", "full_game", "Q1") is None


def test_forecast_serialization_keeps_two_sided_snapshot_and_scope():
    snapshot = _snapshot(
        "quote-round-trip",
        Market.TOTAL,
        Period.FULL_GAME,
        AS_OF - timedelta(hours=1),
        line=220.5,
        over_price_decimal=1.91,
        under_price_decimal=2.02,
    )
    evidence = BookEvidence.from_snapshot(snapshot, frozen=True)
    forecast = Forecast(
        game_id="quote-round-trip",
        league="NBA",
        scheduled_utc=iso(AS_OF + timedelta(hours=2)),
        home_name="Home",
        away_name="Away",
        market=Market.TOTAL.value,
        period=Period.FULL_GAME.value,
        selection="over",
        raw_probability=0.6,
        alternative_probabilities={"over": 0.6, "under": 0.4},
        source="test-model",
        line=220.5,
        book=evidence,
    )
    row = forecast.to_dict()

    assert forecast.opposite_decimal_price == 2.02
    assert row["book"]["side_prices"] == {
        "home": None,
        "away": None,
        "draw": None,
        "over": 1.91,
        "under": 2.02,
    }
    assert row["book"]["game_id"] == "quote-round-trip"
    assert row["book"]["market"] == "total"
    assert row["book"]["period"] == "full_game"
    assert row["book"]["line"] == 220.5
    assert evidence.price_for("home") is None


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
                model="espn-bpi",
                model_version="adapter:espn-bpi-v1",
                produced_by="ESPN BPI",
                prob_home=0.65,
                prob_away=0.35,
                expected_home_score=110,
                expected_away_score=90,
                expected_margin=20,
                expected_total=200,
            ),
            frozen=final,
        )
        book_observed = tip - timedelta(hours=1) if final else AS_OF - timedelta(hours=1)
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
            _snapshot(
                game_id,
                Market.SPREAD,
                Period.FULL_GAME,
                book_observed,
                line=-1.5,
                home_price_decimal=1.91,
                away_price_decimal=1.91,
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


def service(
    store: InMemoryStore,
    *,
    result_available_after_hours: int = 4,
    min_calibration_samples: int = 3,
) -> MarketForecastService:
    return MarketForecastService(
        store,
        League.NBA,
        forecast_config=ForecastConfig(
            min_history_games=3,
            min_residual_games=3,
            team_prior_games=1.0,
            training_window=50,
            residual_window=30,
            result_available_after_hours=result_available_after_hours,
        ),
        calibration_config=CalibrationConfig(
            min_samples=min_calibration_samples,
            prior_strength=2,
            calibration_epoch_utc="2026-01-01T00:00:00Z",
        ),
    )


def test_all_old_markets_are_forecast_and_accas_use_calibrated_probabilities():
    report = service(populated_store()).build_report(
        as_of_utc=iso(AS_OF), acca_legs=2, acca_limit=4
    )
    current = [row for row in report["forecasts"] if row["game_id"] == "g12"]
    markets = {(row["period"], row["market"]) for row in current}
    assert ("full_game", "moneyline") in markets
    assert ("H1", "moneyline") in markets
    assert ("full_game", "total") in markets
    assert ("full_game", "spread") in markets
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
        assert acca["provisional_sample_gate_bypassed"] is False
    assert report["accas"]["strict_results"] == accas
    assert report["accas"]["paper_results"]
    assert all(
        row["provisional_sample_gate_bypassed"] is False
        for row in report["accas"]["paper_results"]
    )


def test_service_routes_sample_only_candidates_to_paper_results():
    report = service(
        populated_store(), min_calibration_samples=30
    ).build_report(as_of_utc=iso(AS_OF), acca_legs=2, acca_limit=4)
    accas = report["accas"]

    assert accas["mode"] == "paper_only"
    assert accas["results_policy"] == "strict"
    assert accas["paper_results_policy"] == "strict_plus_sample_only_provisional"
    assert accas["eligible_games"] == accas["strict_eligible_games"] == 0
    assert accas["results"] == accas["strict_results"] == []
    assert accas["paper_results"]
    assert accas["paper_result_count"] == len(accas["paper_results"])
    assert accas["provisional_sample_only_rows"] > 0
    assert accas["provisional_sample_only_games"] == 3
    assert accas["paper_candidate_games"] == 3
    assert report["execution"]["external_betting_writes"] is False

    for acca in accas["paper_results"]:
        assert acca["provisional_sample_gate_bypassed"] is True
        assert acca["provisional_leg_count"] == 2
        assert acca["ranking_status"] == "paper_only_provisional_sample_gate_bypassed"
        for leg in acca["legs"]:
            assert leg["eligibility"]["acca_eligible"] is False
            assert leg["eligibility"]["reasons"] == ["insufficient_walk_forward_samples"]
            assert leg["calibration"]["status"] == "provisional"
            assert 0 < leg["calibration"]["n"] < leg["calibration"]["minimum_samples"]
            assert leg["book"]["selected_decimal_price"] > 1.0


def test_service_stamps_model_config_and_runs_shadow_at_configured_delay():
    forecast_service = service(
        populated_store(finals=12, upcoming=3), result_available_after_hours=8
    )
    report = forecast_service.build_report(as_of_utc=iso(AS_OF), acca_legs=2, acca_limit=4)

    shadow = report["walk_forward"]["shadow_challenger"]
    assert shadow["status"] == "ok"
    assert shadow["result_delay_hours"] == 8
    assert shadow["rows_evaluated"] > 0
    assert shadow["challenger_version"].startswith("recency_v1-")

    current = report["forecasts"]
    model_row = next(row for row in current if row["model"] == "team_history_empirical_v1")
    assert model_row["calibration_version"] == forecast_service.model.calibration_version
    assert model_row["calibration"]["calibration_version"] == forecast_service.model.calibration_version
    shadow_explanation = model_row["explanation"]["shadow_calibration"]
    assert shadow_explanation["challenger_version"] == shadow["challenger_version"]
    assert shadow_explanation["n"] > 0
    assert 0.01 <= shadow_explanation["candidate_probability"] <= 0.99

    assert all(row["calibration_version"] for row in current)
    provider_row = next(row for row in current if row["model"] == "provider:espn")
    assert provider_row["calibration_version"].startswith("provider_prediction_v2-")
    assert (
        provider_row["explanation"]["segment"]["calibration_version"]
        == provider_row["calibration_version"]
    )
    assert provider_row["prediction"]["model"] == "espn-bpi"
    assert provider_row["prediction"]["model_version"] == "adapter:espn-bpi-v1"
    assert provider_row["prediction"]["produced_by"] == "ESPN BPI"
    hybrid_row = next(
        row for row in current if row["model"].startswith("provider:espn+")
    )
    assert hybrid_row["calibration_version"].startswith("history_provider_hybrid_v2-")

    forecasts_by_id = {row["forecast_id"]: row for row in current}
    for acca in report["accas"]["results"]:
        for leg in acca["legs"]:
            source = forecasts_by_id[leg["forecast_id"]]
            assert leg["calibrated_probability"] == source["calibrated_probability"]
            assert (
                leg["explanation"]["shadow_calibration"]["candidate_probability"]
                == source["explanation"]["shadow_calibration"]["candidate_probability"]
            )

@pytest.mark.parametrize("failure_stage", ["evaluation", "annotation"])
def test_shadow_failure_cannot_change_champion_or_acca_outputs(monkeypatch, failure_stage):
    forecast_service = service(populated_store(finals=12, upcoming=3))
    normal = forecast_service.build_report(as_of_utc=iso(AS_OF), acca_legs=2, acca_limit=4)
    partially_annotated_history = []

    if failure_stage == "evaluation":
        def fail_evaluation(rows, *args, **kwargs):
            if rows:
                rows[0].shadow_calibration = {"partial": True}
                partially_annotated_history.extend(rows)
            raise RuntimeError("injected challenger failure")

        monkeypatch.setattr(
            forecasting_service_module, "prequential_shadow_evaluation", fail_evaluation
        )
    else:
        def fail_annotation(forecasts, profiles, config):
            forecasts[0].shadow_calibration = {"corrupt": True}
            raise RuntimeError("injected annotation failure")

        monkeypatch.setattr(
            forecasting_service_module, "annotate_recency_profiles", fail_annotation
        )

    degraded = forecast_service.build_report(
        as_of_utc=iso(AS_OF), acca_legs=2, acca_limit=4
    )
    assert degraded["walk_forward"]["shadow_challenger"] == {
        "status": "unavailable",
        "error_type": "RuntimeError",
    }
    if failure_stage == "evaluation":
        assert partially_annotated_history
        assert all(not row.shadow_calibration for row in partially_annotated_history)

    def champion_surface(report):
        return {
            row["forecast_id"]: (
                row["raw_probability"],
                row["calibrated_probability"],
                row["calibration"],
                row["eligibility"],
            )
            for row in report["forecasts"]
        }

    def acca_surface(report):
        return [
            (
                row["acca_id"],
                row["estimated_joint_win_probability"],
                tuple(leg["forecast_id"] for leg in row["legs"]),
            )
            for row in report["accas"]["results"]
        ]

    assert champion_surface(degraded) == champion_surface(normal)
    assert acca_surface(degraded) == acca_surface(normal)
    assert all(
        "shadow_calibration" not in row["explanation"]
        for row in degraded["forecasts"]
    )


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
            if snapshot.market is Market.MONEYLINE and snapshot.period is Period.FULL_GAME
        }
        store.books[game_id] = kept
        store.frozen_books[game_id] = {
            key: value for key, value in store.frozen_books.get(game_id, {}).items() if key in kept
        }

    report = service(store).build_report(as_of_utc=iso(AS_OF), acca_legs=2, acca_limit=4)
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
        assert row["team_history"]["sufficient_for_accas"] is True
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


def test_league_prior_dominated_history_moneyline_cannot_supply_accas():
    store = populated_store(finals=12, upcoming=3)
    store.predictions = {}
    store.frozen_predictions = {}
    for game_id, snapshots in list(store.books.items()):
        kept = {
            key: snapshot
            for key, snapshot in snapshots.items()
            if snapshot.market is Market.MONEYLINE and snapshot.period is Period.FULL_GAME
        }
        store.books[game_id] = kept
        store.frozen_books[game_id] = {
            key: value for key, value in store.frozen_books.get(game_id, {}).items() if key in kept
        }

    report = MarketForecastService(
        store,
        League.NBA,
        forecast_config=ForecastConfig(
            min_history_games=3,
            min_residual_games=3,
            training_window=50,
            residual_window=30,
        ),
        calibration_config=CalibrationConfig(min_samples=3, prior_strength=2, calibration_epoch_utc="2026-01-01T00:00:00Z"),
    ).build_report(as_of_utc=iso(AS_OF), acca_legs=2, acca_limit=4)
    rows = [
        row
        for row in report["forecasts"]
        if row["market"] == "moneyline"
        and row["period"] == "full_game"
        and row["model"] == "team_history_empirical_v1"
    ]
    assert len(rows) == 3
    for row in rows:
        evidence = row["team_history"]
        assert evidence["required_effective_games_per_team"] == 8.0
        assert evidence["sufficient_for_accas"] is False
        assert row["eligibility"]["acca_eligible"] is False
        assert "insufficient_team_specific_history" in row["eligibility"]["reasons"]
    assert report["accas"]["results"] == []
    assert report["accas"]["eligible_games"] == 0
    # Fallback rows cannot manufacture eligible calibration, but diagnostics pool is visible.
    assert report["walk_forward"]["segments"]
    assert all(
        s["pool"] == "league_prior_in_diagnostics_only" for s in report["walk_forward"]["segments"]
    )
    assert all(not s["sufficient_for_accas"] for s in report["walk_forward"]["segments"])


def test_expected_total_never_becomes_a_book_line():
    store = populated_store(upcoming=1)
    store.books["g12"] = {}
    store.frozen_books["g12"] = {}
    report = service(store).build_report(as_of_utc=iso(AS_OF))
    current = [row for row in report["forecasts"] if row["game_id"] == "g12"]
    assert not any(row["market"] == "total" for row in current)
    moneyline = next(
        row for row in current if row["market"] == "moneyline" and row["period"] == "full_game"
    )
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
        assert datetime.fromisoformat(
            game.scheduled_utc.replace("Z", "+00:00")
        ) <= datetime.fromisoformat(target.scheduled_utc.replace("Z", "+00:00")) - timedelta(
            hours=4
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
    records = [_forecast(f"g{index}", 0.8, outcome="loss") for index in range(10)]
    config = CalibrationConfig(min_samples=5, prior_strength=2, calibration_epoch_utc="2026-01-01T00:00:00Z")
    profiles = build_profiles(records, config)
    profile = next(iter(profiles.values()))
    assert profile.observed_win_rate == 0
    assert profile.adjustment < 0
    calibrated = apply_profiles([_forecast("new", 0.8)], profiles, config)[0]
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
        assert result["estimated_joint_win_probability"] == pytest.approx(probability, abs=1e-7)


def test_paper_accas_admit_only_provisional_sample_only_legs():
    provisional = _forecast("g0", 0.74)
    provisional.calibrated_probability = 0.74
    provisional.calibration = {
        "status": "provisional",
        "n": 4,
        "minimum_samples": 30,
        "sufficient_for_accas": False,
    }
    provisional.eligibility = {
        "acca_eligible": False,
        "reasons": ["insufficient_walk_forward_samples"],
    }
    strict = _forecast("g1", 0.81)
    strict.calibrated_probability = 0.81
    strict.calibration = {
        "status": "calibrated",
        "n": 30,
        "minimum_samples": 30,
        "sufficient_for_accas": True,
    }
    strict.eligibility = {"acca_eligible": True, "reasons": []}

    assert build_accas([provisional, strict], legs=2, limit=1) == []
    paper = build_accas([provisional, strict], legs=2, limit=1, allow_provisional=True)
    assert len(paper) == 1
    assert paper[0]["provisional_sample_gate_bypassed"] is True
    assert paper[0]["provisional_leg_count"] == 1
    assert paper[0]["ranking_status"] == "paper_only_provisional_sample_gate_bypassed"
    assert {leg["game_id"] for leg in paper[0]["legs"]} == {"g0", "g1"}


@pytest.mark.parametrize(
    ("calibration", "reasons"),
    [
        (
            {
                "status": "provisional",
                "n": 4,
                "minimum_samples": 30,
                "sufficient_for_accas": False,
            },
            ["insufficient_walk_forward_samples", "no_exact_bookmaker_market"],
        ),
        (
            {
                "status": "provisional",
                "n": 4,
                "minimum_samples": 30,
                "sufficient_for_accas": False,
            },
            ["insufficient_walk_forward_samples", "insufficient_team_specific_history"],
        ),
        (
            {
                "status": "provisional",
                "n": 4,
                "minimum_samples": 30,
                "sufficient_for_accas": False,
            },
            ["insufficient_walk_forward_samples", "push_risk_requires_leg_void_rules"],
        ),
        (
            {
                "status": "provisional",
                "n": 4,
                "minimum_samples": 30,
                "sufficient_for_accas": False,
            },
            ["insufficient_walk_forward_samples", "below_calibrated_probability_floor"],
        ),
        (
            {
                "status": "provisional",
                "n": 0,
                "minimum_samples": 30,
                "sufficient_for_accas": False,
            },
            ["insufficient_walk_forward_samples"],
        ),
        (
            {
                "status": "no_history",
                "n": 0,
                "minimum_samples": 30,
                "sufficient_for_accas": False,
            },
            ["insufficient_walk_forward_samples"],
        ),
        (
            {
                "status": "provisional",
                "n": 30,
                "minimum_samples": 30,
                "sufficient_for_accas": False,
            },
            ["insufficient_walk_forward_samples"],
        ),
    ],
)
def test_paper_accas_do_not_bypass_other_gates_or_no_history(calibration, reasons):
    rejected = _forecast("g0", 0.74)
    rejected.calibrated_probability = 0.74
    rejected.calibration = calibration
    rejected.eligibility = {"acca_eligible": False, "reasons": reasons}
    strict = _forecast("g1", 0.81)
    strict.calibrated_probability = 0.81
    strict.eligibility = {"acca_eligible": True, "reasons": []}

    assert build_accas([rejected, strict], legs=2, limit=1, allow_provisional=True) == []


def test_bookmaker_price_never_breaks_acca_probability_ties():
    forecasts = []
    for game_id, price in (("g0", 2.0), ("g1", 3.0), ("g2", 15.0)):
        row = _forecast(game_id, 0.78)
        row.calibrated_probability = 0.78
        row.eligibility = {"acca_eligible": True}
        assert row.book is not None
        row.book.prices["home"] = price
        forecasts.append(row)

    result = build_accas(forecasts, legs=2, limit=1)
    assert len(result) == 1
    assert [leg["game_id"] for leg in result[0]["legs"]] == ["g0", "g1"]
    assert result[0]["combined_decimal_odds"] == 6.0


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
