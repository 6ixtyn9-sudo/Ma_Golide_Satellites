"""Tests for the availability-adjusted (context_availability) forecast surface.

Covers: bounded scale construction, missingness rejection, PIT inheritance,
base-calibration-version stability, model scaling behavior, surface scoping
(full-game only, empirical only), and service-level emission/dedup.
"""

from __future__ import annotations

from datetime import datetime, timedelta, timezone

import pytest

from forecasting.context_adjustment import (
    CONTEXT_SOURCE,
    ContextAdjustment,
    ContextAdjustmentConfig,
    build_context_adjustment,
)
from forecasting.model import ForecastConfig, HistoryModel
from forecasting.service import MarketForecastService
from forecasting.espn_context import has_confirmed_out_injury
from forecasting.types import MODEL_VERSION
from ingestion.models import (
    ActualResult,
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
AS_OF = datetime(2026, 1, 30, 18, tzinfo=UTC)


def iso(value: datetime) -> str:
    return value.isoformat().replace("+00:00", "Z")


def _players(team_id: str, ids: tuple[str, ...], points: str = "10") -> dict:
    return {
        "team": {"id": team_id},
        "statistics": [
            {
                "labels": ["MIN", "PTS"],
                "athletes": [
                    {
                        "athlete": {"id": athlete_id, "displayName": athlete_id},
                        "stats": ["30:00", points],
                    }
                    for athlete_id in ids
                ],
            }
        ],
    }


def _final(index: int, *, with_evidence: bool = True) -> Game:
    tip = datetime(2026, 1, 1, 22, tzinfo=UTC) + timedelta(days=2 * index)
    home_quarters = [28, 27, 29, 26]
    away_quarters = [22, 23, 21, 24]
    evidence = {}
    if with_evidence:
        evidence = {
            "espn": {
                "summary": {
                    "observed_utc": iso(tip - timedelta(hours=2)),
                    "player_boxscore": [
                        _players("h0", ("p1", "p2", "p3", "p4")),
                        _players("a0", ("q1", "q2", "q3", "q4")),
                    ],
                }
            }
        }
    return Game(
        game_id=f"final-{index}",
        league=League.NBA,
        season="2026",
        home_team_id="h0",
        away_team_id="a0",
        home_name="Home Zero",
        away_name="Away Zero",
        scheduled_utc=iso(tip),
        status=GameStatus.FINAL,
        actual=ActualResult(
            period_kind=PeriodKind.QUARTERS,
            home_quarters=home_quarters,
            away_quarters=away_quarters,
            home_final=sum(home_quarters),
            away_final=sum(away_quarters),
            source="official-test",
        ),
        evidence=evidence,
    )


def _upcoming(*, injuries=True, with_evidence=True) -> Game:
    tip = datetime(2026, 1, 30, 20, tzinfo=UTC)
    # Observed before the forecast cutoff (AS_OF), which is itself before tip:
    # evidence must be visible to the report, not merely pre-tip.
    summary: dict = {"observed_utc": iso(AS_OF - timedelta(minutes=30))}
    if with_evidence:
        summary["player_boxscore"] = [
            _players("h0", ("p1", "p2", "p3", "p4")),
            _players("a0", ("q1", "q2", "q3", "q4")),
        ]
    if injuries:
        summary["injuries"] = [
            {
                "team": {"id": "h0"},
                "injuries": [
                    {
                        "athlete": {"id": "p1", "displayName": "p1"},
                        "status": "Out",
                    }
                ],
            }
        ]
    return Game(
        game_id="upcoming-0",
        league=League.NBA,
        season="2026",
        home_team_id="h0",
        away_team_id="a0",
        home_name="Home Zero",
        away_name="Away Zero",
        scheduled_utc=iso(tip),
        status=GameStatus.SCHEDULED,
        evidence={"espn": {"summary": summary}} if with_evidence else {},
    )


def _candidate_record(
    *,
    injury_status="snapshot_present",
    home_window=5,
    away_window=5,
    home_out=("p1",),
    away_out=(),
    home_players=None,
    away_players=None,
):
    def team_rows(ids, window):
        return [
            {
                "athlete_id": athlete_id,
                "name": athlete_id,
                "appearances": window,
                "mean_minutes": 30.0,
                "mean_points": 10.0,
            }
            for athlete_id in ids
        ]

    return {
        "injury_snapshot_status": injury_status,
        "teams": {
            "home": {
                "historical_boxscore_games": home_window,
                "recent_players": (
                    team_rows(("p1", "p2", "p3", "p4"), home_window)
                    if home_players is None
                    else home_players
                ),
                "confirmed_out_player_ids": list(home_out),
            },
            "away": {
                "historical_boxscore_games": away_window,
                "recent_players": (
                    team_rows(("q1", "q2", "q3", "q4"), away_window)
                    if away_players is None
                    else away_players
                ),
                "confirmed_out_player_ids": list(away_out),
            },
        },
    }


LOOSE_CONFIG = ContextAdjustmentConfig(
    min_window_games=2, min_rotation_players=3, max_scale=0.25
)


def test_scale_is_exact_share_of_window_contribution():
    # 4 players at 10 points each; one confirmed out -> 10/40 = 0.25 lost.
    adjustment = build_context_adjustment(_candidate_record(), config=LOOSE_CONFIG)
    assert adjustment is not None
    assert adjustment.home_scale == pytest.approx(0.75)
    assert adjustment.away_scale == pytest.approx(1.0)
    assert adjustment.home_evidence["raw_lost_points_share"] == pytest.approx(0.25)


def test_share_is_capped_by_max_scale():
    adjustment = build_context_adjustment(
        _candidate_record(), config=ContextAdjustmentConfig(min_window_games=2, min_rotation_players=3, max_scale=0.05)
    )
    assert adjustment is not None
    assert adjustment.home_evidence["raw_lost_points_share"] > 0.05
    assert adjustment.home_evidence["applied_lost_points_share"] == pytest.approx(0.05)
    assert adjustment.home_scale == pytest.approx(0.95)


def test_appearance_rate_scales_the_loss():
    rows = [
        {
            "athlete_id": "p1",
            "name": "p1",
            "appearances": 2,
            "mean_minutes": 30.0,
            "mean_points": 10.0,
        },
        {
            "athlete_id": "p2",
            "name": "p2",
            "appearances": 5,
            "mean_minutes": 30.0,
            "mean_points": 10.0,
        },
        {
            "athlete_id": "p3",
            "name": "p3",
            "appearances": 5,
            "mean_minutes": 30.0,
            "mean_points": 10.0,
        },
        {
            "athlete_id": "p4",
            "name": "p4",
            "appearances": 5,
            "mean_minutes": 30.0,
            "mean_points": 10.0,
        },
    ]
    # p1 played only 2 of 5 window games: marginal loss = 0.4 * 10 = 4 of 34
    # (4 from p1 at appearance rate, 30 from the three full-rate teammates).
    adjustment = build_context_adjustment(
        _candidate_record(home_players=rows), config=LOOSE_CONFIG
    )
    assert adjustment is not None
    assert adjustment.home_evidence["raw_lost_points_share"] == pytest.approx(4.0 / 34.0)


def test_missingness_is_rejected_not_half_applied():
    gates = {
        "no_snapshot": _candidate_record(injury_status="no_point_in_time_injury_snapshot"),
        "short_home_window": _candidate_record(home_window=1),
        "short_away_window": _candidate_record(away_window=1),
        "thin_home_rotation": _candidate_record(
            home_players=[
                {
                    "athlete_id": "p1",
                    "name": "p1",
                    "appearances": 5,
                    "mean_minutes": 30.0,
                    "mean_points": 10.0,
                }
            ]
        ),
        "no_out_players": _candidate_record(home_out=(), away_out=()),
    }
    for name, record in gates.items():
        assert build_context_adjustment(record, config=LOOSE_CONFIG) is None, name


def test_non_dict_and_malformed_records_return_none():
    for record in (None, [], "record", {}, {"teams": {}}):
        assert build_context_adjustment(record, config=LOOSE_CONFIG) is None


def test_corrupt_numeric_production_fails_closed():
    nan_rows = [
        {
            "athlete_id": "p1",
            "name": "p1",
            "appearances": 5,
            "mean_minutes": 30.0,
            "mean_points": float("nan"),
        },
        {
            "athlete_id": "p2",
            "name": "p2",
            "appearances": 5,
            "mean_minutes": 30.0,
            "mean_points": 10.0,
        },
        {
            "athlete_id": "p3",
            "name": "p3",
            "appearances": 5,
            "mean_minutes": 30.0,
            "mean_points": 10.0,
        },
        {
            "athlete_id": "p4",
            "name": "p4",
            "appearances": 5,
            "mean_minutes": 30.0,
            "mean_points": 10.0,
        },
    ]
    # A NaN scoring value must be skipped without crashing the share math; the
    # remaining usable players still have to satisfy the rotation minimum.
    assert build_context_adjustment(
        _candidate_record(home_players=nan_rows), config=LOOSE_CONFIG
    ) is None  # only 3 usable players remain and p1's loss is unmeasurable


def test_metadata_is_deterministic_and_json_safe():
    import json

    first = build_context_adjustment(_candidate_record(), config=LOOSE_CONFIG)
    second = build_context_adjustment(_candidate_record(), config=LOOSE_CONFIG)
    assert first is not None and second is not None
    assert first.to_metadata() == second.to_metadata()
    assert json.loads(json.dumps(first.to_metadata())) == first.to_metadata()


def test_config_validation_rejects_unbounded_or_degenerate_knobs():
    invalid_kwargs = (
        {"points_coefficient": 0.0},
        {"points_coefficient": -1.0},
        {"max_scale": 0.0},
        {"max_scale": 0.5},
        {"min_window_games": 0},
        {"min_rotation_players": 0},
    )
    for kwargs in invalid_kwargs:
        with pytest.raises(ValueError):
            build_context_adjustment(
                _candidate_record(), config=ContextAdjustmentConfig(**kwargs)
            )


def test_adjustment_scales_are_bounded_and_ordered():
    with pytest.raises(ValueError):
        ContextAdjustment(home_scale=0.0, away_scale=1.0, home_evidence={}, away_evidence={})
    with pytest.raises(ValueError):
        ContextAdjustment(home_scale=1.5, away_scale=1.0, home_evidence={}, away_evidence={})


def test_context_knobs_never_reversion_base_calibration_segments():
    base = HistoryModel(ForecastConfig())
    tuned = HistoryModel(
        ForecastConfig(),
        ContextAdjustmentConfig(points_coefficient=1.5, max_scale=0.2),
    )
    assert base.calibration_version == tuned.calibration_version
    assert base.context_calibration_version != tuned.context_calibration_version
    assert base.context_calibration_version.startswith(f"{MODEL_VERSION}-context-")


def _model_and_history():
    model = HistoryModel(
        ForecastConfig(min_history_games=3, min_residual_games=3, team_prior_games=1.0)
    )
    finals = [_final(index) for index in range(6)]
    upcoming = _upcoming()
    return model, finals, upcoming


def test_distribution_scales_only_the_depleted_side():
    model, finals, upcoming = _model_and_history()
    base = model.distribution(upcoming, finals, Period.FULL_GAME)
    adjustment = ContextAdjustment(
        home_scale=0.75,
        away_scale=1.0,
        home_evidence={"applied_lost_points_share": 0.25},
        away_evidence={"applied_lost_points_share": 0.0},
    )
    adjusted = model.distribution(upcoming, finals, Period.FULL_GAME, context=adjustment)
    assert base is not None and adjusted is not None
    assert adjusted.baseline.home == pytest.approx(base.baseline.home * 0.75)
    assert adjusted.baseline.away == pytest.approx(base.baseline.away)
    assert adjusted.baseline.margin < base.baseline.margin
    assert adjusted.baseline.total < base.baseline.total


def test_adjusted_distribution_shifts_win_probability_toward_the_healthy_side():
    model, finals, upcoming = _model_and_history()
    base = model.distribution(upcoming, finals, Period.FULL_GAME)
    adjustment = ContextAdjustment(
        home_scale=0.75,
        away_scale=1.0,
        home_evidence={},
        away_evidence={},
    )
    adjusted = model.distribution(upcoming, finals, Period.FULL_GAME, context=adjustment)
    base_probabilities = model._winner_probabilities(base.margins)
    adjusted_probabilities = model._winner_probabilities(adjusted.margins)
    assert adjusted_probabilities["home"] < base_probabilities["home"]
    assert adjusted_probabilities["away"] > base_probabilities["away"]


def test_context_and_provider_prediction_are_mutually_exclusive():
    model, finals, upcoming = _model_and_history()
    prediction = Prediction(
        game_id=upcoming.game_id,
        provider="espn",
        observed_utc=iso(AS_OF - timedelta(hours=2)),
        model="espn-bpi",
        model_version="adapter:espn-bpi-v1",
        produced_by="ESPN BPI",
        prob_home=0.6,
        prob_away=0.4,
    )
    adjustment = ContextAdjustment(
        home_scale=0.9, away_scale=1.0, home_evidence={}, away_evidence={}
    )
    with pytest.raises(ValueError):
        model.distribution(
            upcoming, finals, Period.FULL_GAME, prediction, context=adjustment
        )
    with pytest.raises(ValueError):
        model.lines(
            upcoming,
            finals,
            [],
            {},
            prediction=prediction,
            context=adjustment,
        )


def test_period_winner_context_surface_identity_and_metadata():
    model, finals, upcoming = _model_and_history()
    adjustment = ContextAdjustment(
        home_scale=0.75, away_scale=1.0, home_evidence={}, away_evidence={}
    )
    forecasts = model.period_winner(
        upcoming, finals, [], {}, period=Period.FULL_GAME, context=adjustment
    )
    assert len(forecasts) == 1
    forecast = forecasts[0]
    assert forecast.source == CONTEXT_SOURCE
    assert forecast.source == "context_availability+team_history_empirical_v1"
    assert forecast.calibration_version == model.context_calibration_version
    assert forecast.context == adjustment.to_metadata()
    # Base empirical surface stays untouched.
    base_forecasts = model.period_winner(upcoming, finals, [], {}, period=Period.FULL_GAME)
    assert base_forecasts[0].source == MODEL_VERSION
    assert base_forecasts[0].context is None
    assert base_forecasts[0].forecast_id != forecast.forecast_id


def test_lines_context_applies_to_full_game_only():
    from datetime import datetime as dt

    model, finals, upcoming = _model_and_history()
    adjustment = ContextAdjustment(
        home_scale=0.75, away_scale=1.0, home_evidence={}, away_evidence={}
    )
    from ingestion.models import BookSnapshot

    line = 210.5
    observed = iso(dt(2026, 1, 30, 17, tzinfo=UTC))

    def snapshot_for(period):
        return BookSnapshot(
            game_id=upcoming.game_id,
            provider="theoddsapi",
            book="pinnacle",
            observed_utc=observed,
            market=Market.TOTAL,
            period=period,
            line=line,
            over_price_decimal=1.91,
            under_price_decimal=1.91,
            pre_match=True,
        )

    full_game = model.lines(
        upcoming,
        finals,
        [snapshot_for(Period.FULL_GAME)],
        {},
        context=adjustment,
    )
    assert len(full_game) == 1
    assert full_game[0].source == CONTEXT_SOURCE
    assert full_game[0].context is not None
    assert full_game[0].expected_total < line  # scaled mean shifts the total

    half = model.lines(
        upcoming,
        finals,
        [snapshot_for(Period.H1)],
        {},
        context=adjustment,
    )
    assert len(half) == 1
    assert half[0].source == MODEL_VERSION
    assert half[0].context is None


def _context_store(*, upcoming_evidence=True):
    store = InMemoryStore()
    for index in range(6):
        store.upsert_game(_final(index, with_evidence=upcoming_evidence))
    store.upsert_game(_upcoming(with_evidence=upcoming_evidence))
    return store


def _small_service(store, league=League.NBA):
    return MarketForecastService(
        store,
        league,
        forecast_config=ForecastConfig(
            min_history_games=3,
            min_residual_games=3,
            team_prior_games=1.0,
            residual_window=30,
        ),
        context_config=LOOSE_CONFIG,
    )


def test_service_emits_separate_context_surface_forecasts():
    service = _small_service(_context_store())
    report = service.build_report(as_of_utc=iso(AS_OF))
    forecasts = report["forecasts"]
    context_rows = [
        row
        for row in forecasts
        if row["model"].startswith("context_availability")
    ]
    assert context_rows, "availability surface must emit forecasts with injury evidence"
    assert {row["period"] for row in context_rows} == {"full_game"}
    assert all(row["context"] is not None for row in context_rows)
    assert all(row["context"]["home_scale"] == 0.75 for row in context_rows)
    assert all(
        row["calibration_version"] == service.model.context_calibration_version
        for row in context_rows
    )
    empirical_rows = [
        row
        for row in forecasts
        if row["model"] == MODEL_VERSION and row["period"] == "full_game"
    ]
    assert empirical_rows, "base empirical surface must remain present"
    by_market = {row["market"]: row for row in context_rows}
    base_by_market = {row["market"]: row for row in empirical_rows}
    assert by_market["moneyline"]["expected_margin"] != base_by_market["moneyline"]["expected_margin"]
    # Both surfaces survive training-row dedup: one book per (game, market,
    # period, SOURCE) — distinct sources are distinct rows.
    generated = service._generate(
        _upcoming(),
        [_final(index) for index in range(6)],
        as_of_utc=iso(AS_OF),
        historical=False,
    )
    training_keys = {
        (*row.market_key, row.source)
        for row in MarketForecastService._training_rows(generated)
    }
    assert any(key[3].startswith("context_availability") for key in training_keys)
    assert any(key[3] == MODEL_VERSION for key in training_keys)


def test_service_emits_no_context_surface_without_evidence():
    service = _small_service(_context_store(upcoming_evidence=False))
    report = service.build_report(as_of_utc=iso(AS_OF))
    assert not [
        row for row in report["forecasts"] if row["model"].startswith("context_availability")
    ]


def test_league_without_espn_path_never_builds_an_adjustment():
    service = _small_service(_context_store(), league=League.EUROLEAGUE)
    assert (
        service._context_adjustment(
            _upcoming(), [_final(index) for index in range(6)], as_of_utc=iso(AS_OF)
        )
        is None
    )


def test_audit_block_describes_the_promoted_surface():
    service = _small_service(_context_store())
    report = service.build_report(as_of_utc=iso(AS_OF))
    audit = report["feature_use_audit"]
    assert "context_availability_surface" in audit
    assert "context_availability" in audit["espn_player_context"]
    manifest = report["feature_use_manifest"]
    catalog = manifest["catalog"] if "catalog" in manifest else manifest
    assert "context_availability_surface" in catalog


def test_context_availability_is_never_applied_to_half_game_winner():
    model, finals, upcoming = _model_and_history()
    adjustment = ContextAdjustment(
        home_scale=0.75, away_scale=1.0, home_evidence={}, away_evidence={}
    )

    forecasts = model.period_winner(
        upcoming, finals, [], {}, period=Period.H1, context=adjustment
    )

    assert forecasts
    assert all(row.source == MODEL_VERSION and row.context is None for row in forecasts)
    with pytest.raises(ValueError, match="full-game"):
        model.distribution(upcoming, finals, Period.H1, context=adjustment)


def test_fast_injury_gate_uses_the_same_point_in_time_status_rules():
    assert has_confirmed_out_injury(_upcoming(), as_of_utc=iso(AS_OF)) is True
    assert has_confirmed_out_injury(
        _upcoming(injuries=False), as_of_utc=iso(AS_OF)
    ) is False


def test_walk_forward_skips_expensive_player_context_without_confirmed_out(monkeypatch):
    import forecasting.service as service_module

    service = _small_service(_context_store(upcoming_evidence=False))
    target = _upcoming(injuries=False, with_evidence=False)
    monkeypatch.setattr(
        service_module,
        "build_espn_player_context",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(
            AssertionError("player-history scan should be skipped without confirmed outs")
        ),
    )

    assert service._context_adjustment(
        target, [_final(index) for index in range(6)], as_of_utc=iso(AS_OF)
    ) is None
