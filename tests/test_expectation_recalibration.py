from __future__ import annotations

from dataclasses import asdict
from datetime import datetime, timedelta, timezone

import pytest

from forecasting.harness import config_fingerprint
from forecasting.model import (
    LEAGUE_FORECAST_OVERRIDES,
    RECALIBRATION_MAX_SLOPE,
    ForecastConfig,
    HistoryModel,
    league_forecast_config,
)
from forecasting.service import MarketForecastService
from forecasting.types import MODEL_VERSION, id_spelling
from ingestion.models import (
    ActualResult,
    BookSnapshot,
    Game,
    GameStatus,
    League,
    Market,
    Period,
    PeriodKind,
)
from ingestion.store import InMemoryStore

UTC = timezone.utc
START = datetime(2025, 11, 3, 23, tzinfo=UTC)
# True team strengths: the league mixes elite teams with very weak ones, as in
# early-season college basketball where 40-point mismatches are routine.
STRENGTHS = [-24.0, -16.0, -8.0, 0.0, 0.0, 8.0, 16.0, 24.0]
HOME_EDGE = 3.0


def iso(value: datetime) -> str:
    return value.isoformat().replace("+00:00", "Z")


def _game(index: int, home: int, away: int, *, final: bool = True) -> Game:
    # Deterministic pseudo-noise keeps the outcome distribution non-degenerate.
    noise = float((index * 37) % 21 - 10)
    margin = STRENGTHS[home] - STRENGTHS[away] + HOME_EDGE + noise
    actual = None
    if final:
        home_score = 70.0 + margin / 2
        away_score = 70.0 - margin / 2
        actual = ActualResult(
            period_kind=PeriodKind.QUARTERS,
            home_quarters=[home_score / 4] * 4,
            away_quarters=[away_score / 4] * 4,
            home_final=home_score,
            away_final=away_score,
        )
    return Game(
        game_id=f"g{index}",
        league=League.NCAA_W,
        season="2026",
        home_team_id=f"t{home}",
        away_team_id=f"t{away}",
        home_name=f"Team {home}",
        away_name=f"Team {away}",
        scheduled_utc=iso(START + timedelta(hours=8 * index)),
        status=GameStatus.FINAL if final else GameStatus.SCHEDULED,
        actual=actual,
    )


def _season(games: int = 240) -> list[Game]:
    teams = len(STRENGTHS)
    pairs = [(h, a) for h in range(teams) for a in range(teams) if h != a]
    return [_game(index, *pairs[(index * 11) % len(pairs)]) for index in range(games)]


def _config(**overrides) -> ForecastConfig:
    return ForecastConfig(
        min_history_games=20,
        min_residual_games=20,
        residual_window=100,
        **overrides,
    )


def _mismatch_distribution(config: ForecastConfig):
    finals = _season()
    target = _game(len(finals), home=7, away=0, final=False)
    model = HistoryModel(config)
    history = model.available_history(
        finals, scheduled_utc=target.scheduled_utc, as_of_utc=target.scheduled_utc
    )
    return model, target, history, model.distribution(target, history, Period.FULL_GAME)


def test_linear_recalibration_fits_and_guards_degenerate_windows():
    fit = HistoryModel._linear_recalibration
    assert fit([0.0, 1.0, 2.0, 3.0], [2.0, 5.0, 8.0, 11.0]) == pytest.approx((2.0, 3.0))
    # No predictor spread: keep the identity slope (only re-centre).
    assert fit([4.0, 4.0, 4.0], [1.0, 2.0, 3.0]) == pytest.approx((-2.0, 1.0))
    # Negative and explosive slopes are bounded.
    assert fit([0.0, 1.0, 2.0], [3.0, 2.0, 1.0])[1] == 0.0
    assert fit([0.0, 0.01, 0.02], [0.0, 10.0, 20.0])[1] == RECALIBRATION_MAX_SLOPE


def test_recalibration_restores_compressed_mismatch_margin():
    true_margin = STRENGTHS[7] - STRENGTHS[0] + HOME_EDGE  # 51 points
    _, _, _, raw = _mismatch_distribution(_config())
    _, _, _, recalibrated = _mismatch_distribution(_config(recalibrate_expectations=True))
    assert raw is not None and recalibrated is not None
    # The averaged for/against baseline halves rating gaps.
    assert raw.baseline.margin < 0.6 * true_margin
    assert abs(recalibrated.baseline.margin - true_margin) < 0.15 * true_margin
    assert abs(recalibrated.baseline.margin - true_margin) < abs(
        raw.baseline.margin - true_margin
    )
    # Residuals are re-fitted around the recalibrated mean, not reused raw.
    assert len(recalibrated.margins) == len(raw.margins)
    centre = sorted(recalibrated.margins)[len(recalibrated.margins) // 2]
    assert abs(centre - true_margin) < 0.2 * true_margin


def _spread_probability(config: ForecastConfig, line: float) -> float:
    model, target, history, _ = _mismatch_distribution(config)
    snapshot = BookSnapshot(
        game_id=target.game_id,
        provider="espn",
        book="DraftKings",
        market=Market.SPREAD,
        period=Period.FULL_GAME,
        observed_utc=target.scheduled_utc,
        line=line,
        home_price_decimal=1.91,
        away_price_decimal=1.91,
    )
    rows = model.lines(target, history, [snapshot], {})
    assert len(rows) == 1
    return rows[0].alternative_probabilities["home"]


def test_posted_line_weight_centres_spread_between_model_and_line():
    # The posted line disagrees sharply with the model (home -20.5 vs ~51).
    line = -20.5
    pure = _spread_probability(_config(recalibrate_expectations=True), line)
    full_weight = _spread_probability(
        _config(recalibrate_expectations=True, posted_line_model_weight=1.0), line
    )
    anchored = _spread_probability(
        _config(recalibrate_expectations=True, posted_line_model_weight=0.0), line
    )
    blended = _spread_probability(
        _config(recalibrate_expectations=True, posted_line_model_weight=0.2), line
    )
    assert pure > 0.9
    assert full_weight == pytest.approx(pure)
    assert 0.35 <= anchored <= 0.65
    assert anchored < blended < pure


def test_posted_line_model_weight_is_validated():
    with pytest.raises(ValueError, match="posted line model weight"):
        ForecastConfig(posted_line_model_weight=1.5)
    with pytest.raises(ValueError, match="posted line model weight"):
        ForecastConfig(posted_line_model_weight=-0.1)


def test_default_config_keeps_its_calibration_identity():
    legacy = asdict(ForecastConfig())
    legacy.pop("recalibrate_expectations")
    legacy.pop("posted_line_model_weight")
    expected = config_fingerprint({"model_version": id_spelling(MODEL_VERSION), **legacy})
    assert HistoryModel(ForecastConfig()).calibration_version == f"{MODEL_VERSION}-{expected}"
    assert (
        HistoryModel(league_forecast_config(League.NBA)).calibration_version
        == HistoryModel(ForecastConfig()).calibration_version
    )
    assert (
        HistoryModel(league_forecast_config(League.NCAA_W)).calibration_version
        != HistoryModel(ForecastConfig()).calibration_version
    )


def test_league_defaults_apply_only_to_backtested_college_leagues():
    assert set(LEAGUE_FORECAST_OVERRIDES) == {League.NCAA_W, League.NCAA_M}
    for league in (League.NCAA_W, League.NCAA_M):
        config = league_forecast_config(league)
        assert config.recalibrate_expectations is True
        assert config.posted_line_model_weight == 0.2
        service = MarketForecastService(InMemoryStore(), league)
        assert service.forecast_config == config
    for league in (League.NBA, League.WNBA, League.EUROLEAGUE, League.CBA):
        assert league_forecast_config(league) == ForecastConfig()
        assert MarketForecastService(InMemoryStore(), league).forecast_config == ForecastConfig()
    # Explicit settings still win over league defaults.
    assert league_forecast_config(
        League.NCAA_W, min_history_games=5, posted_line_model_weight=None
    ) == ForecastConfig(min_history_games=5, recalibrate_expectations=True)
