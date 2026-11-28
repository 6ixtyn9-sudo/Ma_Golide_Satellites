"""Tests for tri-state neutral-site home-court advantage handling (Item C)."""

from __future__ import annotations

from datetime import datetime, timedelta, timezone
from typing import Dict, List, Optional

import pytest

from forecasting.model import ForecastConfig, HistoryModel, _HistoryView, _IndexedHistory, league_forecast_config
from ingestion.models import ActualResult, Game, GameStatus, League, Period, PeriodKind


def _final_game(
    index: int,
    home: str,
    away: str,
    home_pts: float,
    away_pts: float,
    *,
    base_time: datetime,
    evidence: Optional[Dict] = None,
) -> Game:
    scheduled = (base_time + timedelta(days=index)).strftime("%Y-%m-%dT%H:%M:%SZ")
    hq = [home_pts / 4.0] * 4
    aq = [away_pts / 4.0] * 4
    return Game(
        game_id=f"hist-{index}",
        league=League.NCAA_M,
        season="2026-27",
        home_team_id=f"team:{home}",
        away_team_id=f"team:{away}",
        home_name=home,
        away_name=away,
        scheduled_utc=scheduled,
        status=GameStatus.FINAL,
        actual=ActualResult(
            period_kind=PeriodKind.QUARTERS,
            home_quarters=hq,
            away_quarters=aq,
            home_final=home_pts,
            away_final=away_pts,
            source="test",
        ),
        evidence=evidence or {},
    )


def _symmetric_history(
    *,
    neutral_extras: Optional[List[Game]] = None,
) -> tuple[List[Game], datetime]:
    """Strictly symmetric non-neutral history for Alpha & Beta with exact +8.0 HA.

    Alpha and Beta play identical per-team game sequences (same ages, same
    scores for/against) against sparring opponents Gamma and Delta, so their
    exponentially decayed team rates are bit-for-bit identical while league
    home_advantage is exactly +8.0 points.
    """
    base_time = datetime(2026, 11, 1, 18, 0, tzinfo=timezone.utc)
    games: List[Game] = []
    margins = [6.0, 10.0, 4.0, 12.0, 8.0, 8.0]
    totals = [150.0, 158.0, 146.0, 162.0, 154.0, 152.0]
    idx = 0
    for i, (margin, total) in enumerate(zip(margins, totals)):
        home_pts = (total + margin) / 2.0
        away_pts = (total - margin) / 2.0
        # Even rounds: Alpha and Beta both play at home against sparring teams.
        # Odd rounds: Alpha and Beta both play away against sparring teams.
        if i % 2 == 0:
            pairs = [("Alpha", "Gamma"), ("Beta", "Delta")]
        else:
            pairs = [("Gamma", "Alpha"), ("Delta", "Beta")]
        for home, away in pairs:
            games.append(
                _final_game(
                    idx,
                    home,
                    away,
                    home_pts,
                    away_pts,
                    base_time=base_time,
                    evidence={"espn": {"scoreboard": {"neutral_site": False}}},
                )
            )
            idx += 1
    if neutral_extras:
        games.extend(neutral_extras)
    return games, base_time + timedelta(days=len(games) + 2)


def _upcoming(
    game_id: str,
    scheduled: datetime,
    *,
    evidence: Optional[Dict] = None,
) -> Game:
    return Game(
        game_id=game_id,
        league=League.NCAA_M,
        season="2026-27",
        home_team_id="team:Alpha",
        away_team_id="team:Beta",
        home_name="Alpha",
        away_name="Beta",
        scheduled_utc=scheduled.strftime("%Y-%m-%dT%H:%M:%SZ"),
        status=GameStatus.SCHEDULED,
        evidence=evidence or {},
    )


def test_neutral_site_true_zeros_home_advantage_and_excludes_from_league_estimate():
    """Defect test (Item C): explicit neutral_site=True must apply 0.0 home advantage
    and must not distort the league home-advantage estimate from finals.
    """
    base_time = datetime(2026, 11, 1, 18, 0, tzinfo=timezone.utc)
    # Add 4 neutral-site finals with extreme -20.0 nominal home margin; they must
    # NOT drag down the league home_advantage computed from non-neutral finals.
    neutral_finals = [
        _final_game(
            12 + i,
            "Gamma",
            "Delta",
            65.0,
            85.0,
            base_time=base_time,
            evidence={"espn": {"scoreboard": {"neutral_site": True}}},
        )
        for i in range(4)
    ]
    history_clean, target_time = _symmetric_history()
    history_with_neutral, target_time_2 = _symmetric_history(neutral_extras=neutral_finals)
    target_time = max(target_time, target_time_2)

    config = ForecastConfig(min_history_games=4, min_residual_games=4, residual_window=50)
    model = HistoryModel(config)

    # 1) Neutral finals must not shift the non-neutral league home_advantage.
    _, ha_clean = model._league_rates(history_clean, Period.FULL_GAME)  # type: ignore[misc]
    _, ha_with_neutral = model._league_rates(history_with_neutral, Period.FULL_GAME)  # type: ignore[misc]
    assert ha_clean == pytest.approx(8.0)
    assert ha_with_neutral == pytest.approx(8.0)

    # 2) Explicit neutral_site=True (ESPN path) yields equal home/away baseline
    #    and equal moneyline probabilities for symmetric teams.
    neutral_espn = _upcoming(
        "target-espn-neutral",
        target_time,
        evidence={"espn": {"scoreboard": {"neutral_site": True}}},
    )
    baseline = model.baseline(neutral_espn, history_clean, Period.FULL_GAME)
    assert baseline is not None
    assert baseline.home == pytest.approx(baseline.away)
    assert baseline.margin == pytest.approx(0.0)

    # Also check BAL neutral=True path
    neutral_bal = _upcoming(
        "target-bal-neutral",
        target_time,
        evidence={"provider": "bal", "neutral": True},
    )
    baseline_bal = model.baseline(neutral_bal, history_clean, Period.FULL_GAME)
    assert baseline_bal is not None
    assert baseline_bal.home == pytest.approx(baseline_bal.away)

    # Moneyline probabilities on symmetric teams at neutral site are strictly
    # symmetric (50/50), whereas on a true home fixture home is favored (>50%).
    dist_neutral = model.distribution(neutral_espn, history_clean, Period.FULL_GAME)
    assert dist_neutral is not None
    assert dist_neutral.baseline.margin == pytest.approx(0.0)
    forecasts_neutral = model.period_winner(
        neutral_espn,
        history_clean,
        [],
        {},
        period=Period.FULL_GAME,
    )
    if forecasts_neutral:
        probs = forecasts_neutral[0].alternative_probabilities
        assert probs["home"] == pytest.approx(probs["away"], abs=1e-6)
    else:
        # When P(home) == P(away) == 0.5 (no strict winner), period_winner
        # legitimately declines to emit a biased side; confirm the underlying
        # residual distribution is balanced around 0.
        n = len(dist_neutral.margins)
        p_home = sum(1 for r in dist_neutral.margins if dist_neutral.baseline.margin + r > 0) / n
        p_away = sum(1 for r in dist_neutral.margins if dist_neutral.baseline.margin + r < 0) / n
        assert p_home == pytest.approx(p_away)

    # Control: same teams with neutral_site=False strongly favor home (> 0.5).
    home_espn = _upcoming(
        "target-espn-home",
        target_time,
        evidence={"espn": {"scoreboard": {"neutral_site": False}}},
    )
    forecasts_home = model.period_winner(
        home_espn,
        history_clean,
        [],
        {},
        period=Period.FULL_GAME,
    )
    assert forecasts_home
    assert forecasts_home[0].selection == "home"
    assert forecasts_home[0].raw_probability > 0.5


def test_neutral_site_false_keeps_home_advantage():
    history, target_time = _symmetric_history()
    config = ForecastConfig(min_history_games=4, min_residual_games=4, residual_window=50)
    model = HistoryModel(config)
    home_game = _upcoming(
        "target-home",
        target_time,
        evidence={"espn": {"scoreboard": {"neutral_site": False}}},
    )
    baseline = model.baseline(home_game, history, Period.FULL_GAME)
    assert baseline is not None
    assert baseline.margin == pytest.approx(8.0)


def test_neutral_site_absent_none_keeps_home_advantage():
    """Guard against the rejected None->no-edge proposal: lanes without a venue
    flag (EuroLeague, LBA, BBL, ACB, BetExplorer, sparse payloads) MUST keep
    normal home-court advantage unchanged.
    """
    history, target_time = _symmetric_history()
    config = ForecastConfig(min_history_games=4, min_residual_games=4, residual_window=50)
    model = HistoryModel(config)

    absent_game = _upcoming("target-absent", target_time, evidence={})
    explicit_none_game = _upcoming(
        "target-none",
        target_time,
        evidence={"espn": {"scoreboard": {"neutral_site": None}}},
    )
    home_game = _upcoming(
        "target-false",
        target_time,
        evidence={"espn": {"scoreboard": {"neutral_site": False}}},
    )
    b_absent = model.baseline(absent_game, history, Period.FULL_GAME)
    b_none = model.baseline(explicit_none_game, history, Period.FULL_GAME)
    b_false = model.baseline(home_game, history, Period.FULL_GAME)
    assert b_absent is not None and b_none is not None and b_false is not None
    assert b_absent.home == pytest.approx(b_false.home)
    assert b_absent.away == pytest.approx(b_false.away)
    assert b_absent.margin == pytest.approx(8.0)
    assert b_none.home == pytest.approx(b_false.home)
    assert b_none.away == pytest.approx(b_false.away)


def test_indexed_and_unindexed_baseline_paths_agree_for_all_tri_state_venues():
    """Pin exact parity between _indexed_baseline (_HistoryView) and baseline (list)."""
    base_time = datetime(2026, 11, 1, 18, 0, tzinfo=timezone.utc)
    neutral_extras = [
        _final_game(
            12 + i,
            "Gamma",
            "Delta",
            60.0 + i,
            82.0 - i,
            base_time=base_time,
            evidence={"espn": {"scoreboard": {"neutral_site": True}}},
        )
        for i in range(4)
    ]
    history, target_time = _symmetric_history(neutral_extras=neutral_extras)
    config = ForecastConfig(min_history_games=4, min_residual_games=4, residual_window=50)
    model = HistoryModel(config)
    indexed = _IndexedHistory(history, config.recent_half_life_games)
    view = _HistoryView(indexed, 0, len(history))

    # Also test sub-slice [2, len(history)) to verify prefix-sum range subtraction.
    sub_list = history[2:]
    sub_view = _HistoryView(indexed, 2, len(history))

    for flag_label, evidence in (
        ("true_espn", {"espn": {"scoreboard": {"neutral_site": True}}}),
        ("true_bal", {"provider": "bal", "neutral": True}),
        ("false_espn", {"espn": {"scoreboard": {"neutral_site": False}}}),
        ("none_absent", {}),
    ):
        g = _upcoming(f"target-{flag_label}", target_time, evidence=evidence)
        b_list = model.baseline(g, history, Period.FULL_GAME)
        b_view = model.baseline(g, view, Period.FULL_GAME)
        assert b_list is not None and b_view is not None
        assert b_list.home == pytest.approx(b_view.home), flag_label
        assert b_list.away == pytest.approx(b_view.away), flag_label
        assert b_list.margin == pytest.approx(b_view.margin), flag_label

        b_sub_list = model.baseline(g, sub_list, Period.FULL_GAME)
        b_sub_view = model.baseline(g, sub_view, Period.FULL_GAME)
        assert b_sub_list is not None and b_sub_view is not None
        assert b_sub_list.home == pytest.approx(b_sub_view.home), f"slice:{flag_label}"
        assert b_sub_list.away == pytest.approx(b_sub_view.away), f"slice:{flag_label}"


def test_all_neutral_history_yields_zero_home_advantage_without_error():
    base_time = datetime(2026, 11, 1, 18, 0, tzinfo=timezone.utc)
    history = [
        _final_game(
            i,
            "Alpha",
            "Beta",
            80.0 + (i % 3),
            70.0,
            base_time=base_time,
            evidence={"espn": {"scoreboard": {"neutral_site": True}}},
        )
        for i in range(6)
    ]
    config = ForecastConfig(min_history_games=4, min_residual_games=4, residual_window=50)
    model = HistoryModel(config)
    indexed = _IndexedHistory(history, config.recent_half_life_games)
    view = _HistoryView(indexed, 0, len(history))
    g = _upcoming(
        "target-home",
        base_time + timedelta(days=10),
        evidence={"espn": {"scoreboard": {"neutral_site": False}}},
    )
    b_list = model.baseline(g, history, Period.FULL_GAME)
    b_view = model.baseline(g, view, Period.FULL_GAME)
    assert b_list is not None and b_view is not None
    assert b_list.home == pytest.approx(b_view.home)
    assert b_list.away == pytest.approx(b_view.away)
    # No non-neutral games exist -> league home_advantage is 0.0 (no ZeroDivisionError)
    _, ha_list = model._league_rates(history, Period.FULL_GAME)  # type: ignore[misc]
    _, ha_view = indexed.period(Period.FULL_GAME).league_rates(0, len(history))  # type: ignore[misc]
    assert ha_list == pytest.approx(0.0)
    assert ha_view == pytest.approx(0.0)


def test_calibration_version_unchanged_by_neutral_site_fix():
    """Confirm calibration_version (SegmentKey[4]) is NOT bumped by Item C."""
    default_model = HistoryModel(ForecastConfig())
    legacy_ncaa_m_model = HistoryModel(
        ForecastConfig(recalibrate_expectations=True, posted_line_model_weight=0.2)
    )
    ncaa_m_model = HistoryModel(league_forecast_config(League.NCAA_M))
    assert default_model.calibration_version == "team_history_empirical-cb419eb48826"
    assert legacy_ncaa_m_model.calibration_version == "team_history_empirical-542f7e282358"
    assert ncaa_m_model.calibration_version == "team_history_empirical-0ae09a9196c1"
