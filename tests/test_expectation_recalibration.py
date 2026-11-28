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
    legacy.pop("anchor_dispersion_rescale")
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
    assert set(LEAGUE_FORECAST_OVERRIDES) == {
        League.NCAA_W,
        League.NCAA_M,
        League.EUROLEAGUE_W,
    }
    for league in (League.NCAA_W, League.NCAA_M):
        config = league_forecast_config(league)
        assert config.recalibrate_expectations is True
        assert config.posted_line_model_weight == 0.2
        assert config.anchor_dispersion_rescale is True
        service = MarketForecastService(InMemoryStore(), league)
        assert service.forecast_config == config
    for league in (
        League.NBA,
        League.WNBA,
        League.EUROLEAGUE,
        League.EUROLEAGUE_W,
        League.CBA,
    ):
        assert league_forecast_config(league) == ForecastConfig()
        assert MarketForecastService(InMemoryStore(), league).forecast_config == ForecastConfig()
    # Explicit settings still win over league defaults.
    assert league_forecast_config(
        League.NCAA_W, min_history_games=5, posted_line_model_weight=None
    ) == ForecastConfig(
        min_history_games=5,
        recalibrate_expectations=True,
        anchor_dispersion_rescale=True,
    )


def test_posted_line_weight_anchors_once_to_main_line_across_alternate_lines():
    """Alternate lines must be priced against ONE main-line-anchored distribution.

    Previously, shifting toward each snapshot's own line stamped identical
    probabilities onto lines 1-2 points apart (2026-11-27 NCAA_M: Illinois
    176.5/177.5 both 0.592409; Hofstra 142.5/140.5 both 0.572409).
    """
    config = _config(recalibrate_expectations=True, posted_line_model_weight=0.2)
    model, target, history, _ = _mismatch_distribution(config)
    # Build spread main line (-20.5) plus alternate lines (-14.5, -26.5):
    # higher home handicap (-14.5) must have strictly higher home cover prob
    # than (-20.5), which must beat (-26.5).
    spread_snaps = [
        BookSnapshot(
            game_id=target.game_id,
            provider="espn",
            book="DraftKings",
            market=Market.SPREAD,
            period=Period.FULL_GAME,
            observed_utc=target.scheduled_utc,
            line=line,
            home_price_decimal=1.91,
            away_price_decimal=1.91,
            main_line=(line == -20.5),
        )
        for line in (-26.5, -20.5, -14.5)
    ]
    rows = model.lines(target, history, spread_snaps, {})
    by_line = {row.line: row.alternative_probabilities["home"] for row in rows}
    assert by_line[-26.5] < by_line[-20.5] < by_line[-14.5]
    # Single-line main snapshot must produce the exact same probability as the
    # main line when alt lines are present alongside it.
    solo = model.lines(target, history, [spread_snaps[1]], {})
    assert solo[0].alternative_probabilities["home"] == pytest.approx(by_line[-20.5])

    # Same monotonicity check for TOTAL on a season with non-degenerate total
    # variance: higher threshold -> strictly higher under prob and lower over prob.
    varied_finals = []
    for idx, g in enumerate(_season()):
        assert g.actual is not None
        total_shift = float((idx * 19) % 25 - 12)
        h_final = g.actual.home_final + total_shift / 2.0
        a_final = g.actual.away_final + total_shift / 2.0
        varied_finals.append(
            Game(
                game_id=g.game_id,
                league=g.league,
                season=g.season,
                home_team_id=g.home_team_id,
                away_team_id=g.away_team_id,
                home_name=g.home_name,
                away_name=g.away_name,
                scheduled_utc=g.scheduled_utc,
                status=g.status,
                actual=ActualResult(
                    period_kind=PeriodKind.QUARTERS,
                    home_quarters=[h_final / 4] * 4,
                    away_quarters=[a_final / 4] * 4,
                    home_final=h_final,
                    away_final=a_final,
                ),
            )
        )
    t_target = _game(len(varied_finals), home=7, away=0, final=False)
    t_history = model.available_history(
        varied_finals, scheduled_utc=t_target.scheduled_utc, as_of_utc=t_target.scheduled_utc
    )
    total_snaps = [
        BookSnapshot(
            game_id=t_target.game_id,
            provider="espn",
            book="DraftKings",
            market=Market.TOTAL,
            period=Period.FULL_GAME,
            observed_utc=t_target.scheduled_utc,
            line=line,
            over_price_decimal=1.91,
            under_price_decimal=1.91,
            main_line=(line == 140.5),
        )
        for line in (134.5, 140.5, 146.5)
    ]
    t_rows = model.lines(t_target, t_history, total_snaps, {})
    under_by_line = {row.line: row.alternative_probabilities["under"] for row in t_rows}
    over_by_line = {row.line: row.alternative_probabilities["over"] for row in t_rows}
    assert under_by_line[134.5] < under_by_line[140.5] < under_by_line[146.5]
    assert over_by_line[134.5] > over_by_line[140.5] > over_by_line[146.5]


def test_anchored_dispersion_scale_matches_orthogonal_variance_bounds():
    scale_fn = HistoryModel._anchored_dispersion_scale
    # Symmetric values around 0 with exact std = 14.5
    wide = [-14.5, 14.5] * 25
    # w=1.0 -> identity (no anchoring)
    assert scale_fn(wide, 0.0, 10.7, 1.0) == pytest.approx(1.0)
    # w=0.0 -> shrinks exactly to market sigma prior (10.7 / 14.5)
    assert scale_fn(wide, 0.0, 10.7, 0.0) == pytest.approx(10.7 / 14.5)
    # w=0.2 -> sqrt((1 - 0.04)*10.7^2 + 0.04*14.5^2) / 14.5; never drops below sigma_market!
    expected_sigma = ((1.0 - 0.04) * (10.7**2) + 0.04 * (14.5**2)) ** 0.5
    assert expected_sigma > 10.7
    assert scale_fn(wide, 0.0, 10.7, 0.2) == pytest.approx(expected_sigma / 14.5)
    # Tighter residual window (sigma_raw = 6.0 <= 10.7): never inflated above 1.0
    tight = [-6.0, 6.0] * 25
    assert scale_fn(tight, 0.0, 10.7, 0.2) == pytest.approx(1.0)


def test_anchor_dispersion_rescale_tightens_noisy_college_alt_line_tails():
    """High-noise college seasons must not price far alt lines with raw-rating width."""
    noisy_finals = []
    for idx, g in enumerate(_season()):
        assert g.actual is not None
        # Add wide margin noise (~17-point std) to mimic noisy early-season NCAA ratings.
        extra = float(((idx * 43) % 51) - 25)
        h_final = g.actual.home_final + extra / 2.0
        a_final = g.actual.away_final - extra / 2.0
        noisy_finals.append(
            Game(
                game_id=g.game_id,
                league=League.NCAA_M,
                season=g.season,
                home_team_id=g.home_team_id,
                away_team_id=g.away_team_id,
                home_name=g.home_name,
                away_name=g.away_name,
                scheduled_utc=g.scheduled_utc,
                status=g.status,
                actual=ActualResult(
                    period_kind=PeriodKind.QUARTERS,
                    home_quarters=[h_final / 4] * 4,
                    away_quarters=[a_final / 4] * 4,
                    home_final=h_final,
                    away_final=a_final,
                ),
            )
        )
    target = Game(
        game_id="g_target",
        league=League.NCAA_M,
        season="2026",
        home_team_id="t4",
        away_team_id="t3",
        home_name="Team 4",
        away_name="Team 3",
        scheduled_utc=iso(START + timedelta(hours=8 * len(noisy_finals))),
        status=GameStatus.SCHEDULED,
    )
    snaps = [
        BookSnapshot(
            game_id=target.game_id,
            provider="espn",
            book="DraftKings",
            market=Market.SPREAD,
            period=Period.FULL_GAME,
            observed_utc=target.scheduled_utc,
            line=line,
            home_price_decimal=1.91,
            away_price_decimal=1.91,
            main_line=(line == -4.5),
        )
        for line in (-14.5, -4.5, 5.5)
    ]
    unscaled_cfg = _config(
        recalibrate_expectations=True,
        posted_line_model_weight=0.2,
        anchor_dispersion_rescale=False,
    )
    scaled_cfg = _config(
        recalibrate_expectations=True,
        posted_line_model_weight=0.2,
        anchor_dispersion_rescale=True,
    )
    unscaled_model = HistoryModel(unscaled_cfg)
    scaled_model = HistoryModel(scaled_cfg)
    history = unscaled_model.available_history(
        noisy_finals, scheduled_utc=target.scheduled_utc, as_of_utc=target.scheduled_utc
    )
    unscaled_rows = {
        row.line: row.alternative_probabilities["home"]
        for row in unscaled_model.lines(target, history, snaps, {})
    }
    scaled_rows = {
        row.line: row.alternative_probabilities["home"]
        for row in scaled_model.lines(target, history, snaps, {})
    }
    # Alt-line monotonicity holds after rescaling.
    assert scaled_rows[-14.5] < scaled_rows[-4.5] < scaled_rows[5.5]
    # Far tail alt lines move away from 0.5 (less inflated tail mass).
    assert scaled_rows[-14.5] < unscaled_rows[-14.5]
    assert scaled_rows[5.5] > unscaled_rows[5.5]
    # Calibration version changes for scaled vs unscaled so cohorts stay separated.
    assert scaled_model.calibration_version != unscaled_model.calibration_version


