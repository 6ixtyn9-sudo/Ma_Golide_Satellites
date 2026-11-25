"""Tests for the Ma Golide score cold-start shadow lane.

The scorer is a port of the retired ``calculateMaGolideScore`` from
``docs/Forecaster.gs``; these tests pin the behaviours that matter for the
shadow lane: neutral fallbacks, the RISKY threshold, logistic confidence
clamping, the variance penalty, and — above all — the lane restriction that
keeps the scorer off lanes that already have results machinery.
"""

from __future__ import annotations

import pytest

from forecasting.magolide_score import (
    MaGolideConfig,
    MaGolideFeatures,
    StandingsRow,
    build_standings_features,
    render_shadow_card,
    score_features,
    shadow_lane_leagues,
    shadow_rows,
)
from ingestion.models import Game, League


def _game(league: League = League.GREECE_BASKET_LEAGUE, game_id: str = "g-1") -> Game:
    return Game(
        game_id=game_id,
        league=league,
        season="2026",
        home_team_id="home",
        away_team_id="away",
        home_name="Panathinaikos BC",
        away_name="AEK Athens",
        scheduled_utc="2026-11-24T17:15:00Z",
    )


# ---------------------------------------------------------------------------
# scoring core
# ---------------------------------------------------------------------------


def test_all_zero_features_scores_risky_at_neutral_confidence():
    verdict = score_features(MaGolideFeatures(), MaGolideConfig(home_adv=0.0))
    assert verdict.prediction == "RISKY"
    assert verdict.score == 0.0
    assert verdict.uncalibrated_logistic_pct == 50.0
    assert verdict.meta["mode"] == "NEW"


def test_pct_difference_drives_home_and_away():
    config = MaGolideConfig(home_adv=0.0)
    features_home = build_standings_features(
        StandingsRow(team="H", pct=0.80), StandingsRow(team="A", pct=0.40)
    )
    verdict_home = score_features(features_home, config)
    assert verdict_home.prediction == "HOME"
    assert verdict_home.score > 0

    features_away = build_standings_features(
        StandingsRow(team="H", pct=0.40), StandingsRow(team="A", pct=0.80)
    )
    verdict_away = score_features(features_away, config)
    assert verdict_away.prediction == "AWAY"
    assert verdict_away.score < 0
    # The script reports the PREDICTED side's probability, so a perfect
    # mirror matchup scores the same probability for the other side.
    assert verdict_away.uncalibrated_logistic_pct == pytest.approx(verdict_home.uncalibrated_logistic_pct, abs=0.1)


def test_threshold_boundary_matches_retired_script():
    # pct weight 1 so score == pct_diff, mirroring the script's threshold-on-score.
    config = MaGolideConfig(pct=1.0, threshold=25.0, home_adv=0.0)
    small = score_features(MaGolideFeatures(pct_diff=20.0), config)
    boundary = score_features(MaGolideFeatures(pct_diff=25.0), config)
    assert small.prediction == "RISKY"
    # |score| == threshold is NOT below the threshold: HOME.
    assert boundary.prediction == "HOME"


def test_variance_is_always_a_penalty():
    config = MaGolideConfig(home_adv=0.0)
    without = score_features(MaGolideFeatures(pct_diff=10.0), config)
    with_var = score_features(
        MaGolideFeatures(pct_diff=10.0, variance_penalty=3.0), config
    )
    assert with_var.score < without.score
    assert with_var.factor_breakdown["variance"] == pytest.approx(-3.0)


def test_confidence_is_clamped_and_logistic():
    config = MaGolideConfig(home_adv=0.0)
    strong = score_features(MaGolideFeatures(pct_diff=60.0), config)
    assert strong.confidence <= config.conf_max
    assert strong.confidence >= config.conf_min
    assert strong.meta["confidence_mode"] == "PROBABILITY"

    legacy = score_features(MaGolideFeatures(pct_diff=60.0), MaGolideConfig(use_legacy_sigmoid=True, home_adv=0.0))
    assert legacy.meta["confidence_mode"] == "LEGACY_SIGMOID"


def test_risky_cap_limits_confidence_when_enabled():
    config = MaGolideConfig(
        home_adv=0.0, cap_risky_confidence=True, risky_confidence_cap=55.0
    )
    verdict = score_features(MaGolideFeatures(pct_diff=1.0), config)
    assert verdict.prediction == "RISKY"
    assert verdict.confidence <= 55.0


def test_legacy_mode_applies_when_no_new_feature_weights_are_set():
    legacy_config = MaGolideConfig(
        pct=0.0,
        net_rtg=0.0,
        home_court=0.0,
        momentum=0.0,
        streak=0.0,
        home_adv=0.0,
        form=2.5,
        h2h=1.5,
    )
    verdict = score_features(
        MaGolideFeatures(form_diff=3.0, h2h_diff=2.0), legacy_config
    )
    assert verdict.meta["mode"] == "LEGACY"
    # 2.5*3 + 1.5*2 = 10.5, above the 25 threshold -> RISKY still, but scored.
    assert verdict.score == pytest.approx(10.5)


def test_home_advantage_is_added_verbatim():
    config = MaGolideConfig(home_adv=5.0)
    verdict = score_features(MaGolideFeatures(), config)
    assert verdict.score == pytest.approx(5.0)
    # The retired script reports homeAdv inside the factor breakdown.
    assert verdict.factor_breakdown["homeAdv"] == pytest.approx(5.0)


# ---------------------------------------------------------------------------
# cold-start feature building
# ---------------------------------------------------------------------------


def test_missing_standings_rows_use_neutral_defaults():
    features = build_standings_features(None, None)
    assert features.rank_diff == 0.0
    assert features.pct_diff == 0.0
    assert features.home_court_effect == 0.0
    assert features.streak_diff == 0.0


def test_missing_team_falls_back_to_retired_script_defaults():
    features = build_standings_features(
        StandingsRow(team="H", pct=0.80, rank=1), None
    )
    # Away side defaults: rank 15, pct 0.5 — exactly calculateRankDifference_/…PCT…
    assert features.rank_diff == pytest.approx(14.0)
    assert features.pct_diff == pytest.approx((0.80 - 0.5) * 100)


def test_home_court_uses_home_pct_vs_away_pct_proxy():
    features = build_standings_features(
        StandingsRow(team="H", home_pct=0.90), StandingsRow(team="A", away_pct=0.60)
    )
    assert features.home_court_effect == pytest.approx(30.0)


def test_l10_form_beats_streak_proxy_when_present():
    features = build_standings_features(
        StandingsRow(team="H", streak=5, l10_wins=8, l10_losses=2),
        StandingsRow(team="A", streak=1, l10_wins=4, l10_losses=6),
    )
    assert features.form_diff == pytest.approx(8.0)  # (8-2) - (4-6)


# ---------------------------------------------------------------------------
# shadow-lane restriction
# ---------------------------------------------------------------------------


def test_shadow_lane_leagues_is_exactly_the_fixtures_only_class():
    leagues = shadow_lane_leagues()
    values = {spec.league.value for spec in leagues}
    # Leumit was retired on 2026-11-24 and a retired lane never qualifies,
    # whatever its sources say -- which is what this set asserts. NKL is in it
    # now: the admission rule stopped keying on "cannot be graded" (three lanes
    # gained a settlement route on 2026-11-25) and keyed on "no calibrated model
    # of its own", which is the reason this module exists.
    assert values == {
        "GREECE_BASKET_LEAGUE",
        "CHINA_WCBA_W",
        "AUSTRALIA_WNBL_W",
        "LITHUANIA_NKL",
    }
    for spec in leagues:
        assert spec.schedule_source == "betexplorer_fixtures"
        assert spec.status.value == "active"


def test_shadow_rows_refuse_lanes_with_results_machinery():
    with pytest.raises(ValueError, match="not a fixtures-only shadow lane"):
        shadow_rows(League.EUROLEAGUE, [_game(League.EUROLEAGUE)], {})
    with pytest.raises(ValueError, match="not a fixtures-only shadow lane"):
        shadow_rows(League.SPAIN_ACB, [_game(League.SPAIN_ACB)], {})


def test_shadow_rows_carry_shadow_eligibility_and_no_price():
    standings = {
        "Panathinaikos BC": StandingsRow(team="Panathinaikos BC", pct=0.82, rank=1),
    }
    rows = shadow_rows(
        League.GREECE_BASKET_LEAGUE, [_game()], standings, config=MaGolideConfig(home_adv=0.0)
    )
    assert len(rows) == 1
    row = rows[0]
    assert row.verdict.prediction == "HOME"  # 0.82 vs neutral 0.5 clears the threshold
    assert row.evidence["away_standings_present"] is False
    assert row.evidence["neutral_defaults_applied"] is True
    card_row = row.as_card_row()
    for marker in (
        "bucket=WATCHLIST_SHADOW",
        "price_evidence=NONE",
        "acca_eligible=False",
        "grade=NOT_SETTLEABLE",
        "no_settlement_route",
    ):
        assert marker in card_row


def test_rendered_card_never_claims_picks_or_settlement():
    card = render_shadow_card("GREECE_BASKET_LEAGUE", [], date_label="2026-11-24")
    assert "SHADOW ONLY" in card
    assert "not tickets" in card
    assert "(no fixtures supplied)" in card
    assert "settleable: 0" in card


# ---------------------------------------------------------------------------
# Guards added on review: the shadow lane's two dangerous tokens, and the one
# place it must never reach.
# ---------------------------------------------------------------------------


def test_the_card_never_prints_a_bare_probability_token():
    """``p=`` means a calibrated probability everywhere else in this repo.

    The logistic of an unmeasured hand-weighted score is not one, so the card
    must never render it with the token the pick cards use for a real fit.
    """

    game = _game()
    rows = shadow_rows(League.GREECE_BASKET_LEAGUE, [game], {})
    card = render_shadow_card("GREECE_BASKET_LEAGUE", rows, date_label="2026-11-24")

    assert "uncalibrated_logistic=" in card
    assert " p=" not in card
    # And no priced or eligible vocabulary leaks in from the pick vocabulary.
    for forbidden in ("@1.", "ELIGIBLE", "NAMED_BOOKMAKER_PRICE", "acca_eligible=True"):
        assert forbidden not in card


def test_an_explicit_zero_confidence_bound_is_kept():
    """safeNum_ keeps an explicit 0; truthiness would silently map it to 50."""

    verdict = score_features(
        MaGolideFeatures(pct_diff=40.0),
        MaGolideConfig(conf_min=0.0, conf_max=95.0),
    )
    assert verdict.meta["conf_bounds"]["min"] == 0.0


def test_the_shadow_scorer_is_not_reachable_from_the_forecast_pipeline():
    """A shadow row must never become a pipeline row by import accident.

    Nothing wires this module in today; this test is what keeps that true, so a
    future 'just call it from pregame' is a deliberate edit with this assertion
    to argue against, not a one-line slip.
    """

    from pathlib import Path

    root = Path(__file__).resolve().parent.parent
    watched = [
        root / "ingestion" / "pipeline.py",
        root / "forecasting" / "operator.py",
        root / "forecasting" / "service.py",
        root / "forecasting" / "accas.py",
        root / "scripts" / "daily.py",
    ]
    for path in watched:
        assert "magolide_score" not in path.read_text(), f"{path} must not import the shadow scorer"


def test_an_unversioned_default_is_never_presented_as_a_fit():
    """The retired system read a versioned config; a default must say so."""

    verdict = score_features(MaGolideFeatures(pct_diff=40.0), MaGolideConfig())

    assert verdict.meta["weights_source"] == "validator_defaults_not_a_fitted_config"
    assert verdict.meta["config_version"] == "unversioned-validator-defaults"

    supplied = MaGolideConfig(pct=4.0, config_version="v3")
    assert score_features(MaGolideFeatures(pct_diff=40.0), supplied).meta["config_version"] == "v3"
