"""Bounded ESPN history plans, including the NBA G League evidence bootstrap."""

from ingestion.models import League
from ingestion.pipeline import default_history_backfill_plan


def test_standard_espn_lanes_keep_the_conservative_recent_history_default():
    plan = default_history_backfill_plan(League.NBA)

    assert plan == {
        "history_days": 21,
        "history_request_budget": 21,
        "policy": "default_recent_scoreboard_backfill",
    }


def test_g_league_uses_a_larger_date_bounded_empirical_history_bootstrap():
    plan = default_history_backfill_plan(League.NBA_G_LEAGUE)

    # This is still the independently validated date-bounded scoreboard/result
    # route. It is deliberately not a core-index/deep-history claim and does
    # not alter any model, calibration, pricing, or ticket threshold.
    assert plan == {
        "history_days": 56,
        "history_request_budget": 56,
        "policy": "nba_g_league_date_bounded_empirical_history_bootstrap",
    }


def test_explicit_operator_values_are_not_overridden_by_the_g_league_plan():
    plan = default_history_backfill_plan(
        League.NBA_G_LEAGUE,
        history_days=28,
        history_request_budget=9,
    )

    assert plan == {
        "history_days": 28,
        "history_request_budget": 9,
        "policy": "operator_override",
    }
