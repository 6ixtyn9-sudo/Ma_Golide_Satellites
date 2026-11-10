"""Tests: prediction selection, coherence, freezing, and odds binding."""

import pytest

from ingestion.models import BookSnapshot, Game, League, Market, Period, Prediction
from ingestion.predictions import (
    freeze_game_predictions,
    latest_pre_tip,
    odds_for_prediction,
    resolve_frozen_bundle,
    select_prediction,
)
from ingestion.store import FrozenWriteError, InMemoryStore

NOW = "2026-11-07T20:00:00Z"


def _game():
    return Game(
        game_id="g1",
        league=League.NBA,
        season="2026",
        home_team_id="11",
        away_team_id="7",
        home_name="Indiana Pacers",
        away_name="Denver Nuggets",
        scheduled_utc="2026-11-07T22:00:00Z",
    )


def _pred(provider, observed, complete=True, tip="2026-11-07T22:00:00Z", **kw):
    base = dict(
        game_id="g1",
        provider=provider,
        observed_utc=observed,
        prob_home=0.55,
        prob_away=0.45,
    )
    if complete:
        base.update(
            expected_home_score=118.0,
            expected_away_score=112.0,
            expected_total=230.0,
            expected_margin=6.0,
        )
    base.update(kw)
    return Prediction(**base)


def _snap(provider="therundown", observed="2026-11-07T19:00:00Z", book="fanduel"):
    return BookSnapshot(
        game_id="g1",
        provider=provider,
        book=book,
        market=Market.MONEYLINE,
        period=Period.FULL_GAME,
        observed_utc=observed,
        home_price_decimal=1.67,
        away_price_decimal=2.30,
    )


# ---------------------------------------------------------------------------
# Selection: coherent bundles, priority, partials, freshness
# ---------------------------------------------------------------------------


def test_complete_bundle_selected_never_mixed():
    espn_p = _pred("espn", "2026-11-07T18:00:00Z")
    hw = _pred(
        "hardwood",
        "2026-11-07T17:50:00Z",
        expected_home_score=117.0,
        expected_away_score=110.0,
        expected_total=227.0,
        expected_margin=7.0,
    )
    result = select_prediction([espn_p, hw], _game(), NOW)
    assert result.provider == "espn" and result.reason == "complete_bundle"
    # No synthetic blend: the selection IS one provider's record.
    assert result.selected.expected_total == 230.0


def test_hardwood_fallback_when_no_espn():
    hw = _pred("hardwood", "2026-11-07T17:50:00Z")
    result = select_prediction([hw], _game(), NOW)
    assert result.provider == "hardwood"


def test_partial_bundle_only_probabilities():
    wnba_like = _pred("espn", NOW, complete=False)
    result = select_prediction([wnba_like], _game(), NOW)
    assert result.selected.prob_home == 0.55
    assert result.selected.expected_total is None
    assert result.reason == "partial_bundle_probabilities_only"
    assert set(result.partial_fields) == {
        "expected_home_score",
        "expected_away_score",
        "expected_total",
    }
    assert any(g.required_field == "prediction_scores" for g in result.gaps)


def test_stale_candidate_not_selected():
    stale = _pred("espn", "2026-11-06T06:00:00Z")  # > 12h before NOW
    result = select_prediction([stale], _game(), NOW, freshness_hours=12.0)
    assert result.selected is None
    assert result.reason == "no_fresh_candidate"
    assert result.gaps[0].required_field == "pregame_prediction"


def test_future_observation_rejected():
    future = _pred("espn", "2026-11-07T23:00:00Z")
    result = select_prediction([future], _game(), NOW)
    assert result.selected is None


def test_shadows_preserved_not_discarded():
    espn_p = _pred("espn", NOW)
    hw = _pred("hardwood", NOW)
    result = select_prediction([espn_p, hw], _game(), NOW)
    assert [s.provider for s in result.shadows] == ["hardwood"]


def test_rounded_scores_may_disagree_with_favorite_and_survive():
    """Rounded display can tie/disagree with the favorite; never repaired."""
    p = _pred(
        "espn",
        NOW,
        prob_home=0.55,
        prob_away=0.45,
        expected_home_score=110.4,
        expected_away_score=110.6,
        expected_total=221.0,
        expected_margin=-0.2,
    )
    result = select_prediction([p], _game(), NOW)
    assert result.selected.prob_home > result.selected.prob_away
    assert result.selected.expected_home_score == 110.4
    assert result.selected.expected_away_score == 110.6


def test_identity_mismatch_excluded():
    p = _pred("espn", NOW)
    p.game_id = "other-game"
    result = select_prediction([p], _game(), NOW)
    assert result.selected is None


# ---------------------------------------------------------------------------
# Freezing
# ---------------------------------------------------------------------------


def _seed(store, preds):
    for p in preds:
        store.upsert_prediction(p)


def test_freeze_uses_latest_pre_tip_observation():
    store = InMemoryStore()
    a = _pred("espn", "2026-11-07T10:00:00Z")  # pre-tip
    b = _pred("espn", "2026-11-07T23:00:00Z")  # post-tip arrival
    b.pre_tip = False  # pipeline stamps this at storage
    _seed(store, [a, b])
    game = _game()
    statuses = freeze_game_predictions(store, game, game.scheduled_utc)
    assert statuses["espn"] == "frozen"
    frozen_pred, was_frozen = resolve_frozen_bundle(store, "g1", game.scheduled_utc)
    assert was_frozen and frozen_pred.observed_utc == a.observed_utc


def test_frozen_prediction_never_replaced():
    store = InMemoryStore()
    a = _pred("espn", "2026-11-07T10:00:00Z")
    _seed(store, [a])
    game = _game()
    freeze_game_predictions(store, game, game.scheduled_utc)
    b = _pred("espn", "2026-11-07T23:00:00Z")  # post-tip update attempt
    with pytest.raises(FrozenWriteError):
        store.upsert_prediction(b, frozen=False)


def test_freeze_is_idempotent():
    store = InMemoryStore()
    a = _pred("espn", "2026-11-07T10:00:00Z")
    _seed(store, [a])
    game = _game()
    first = freeze_game_predictions(store, game, game.scheduled_utc)
    second = freeze_game_predictions(store, game, game.scheduled_utc)
    assert first["espn"] == "frozen"
    assert second["espn"] == "already_frozen"


def test_no_pregame_prediction_status():
    store = InMemoryStore()
    b = _pred("espn", "2026-11-07T23:00:00Z")  # only post-tip
    _seed(store, [b])
    game = _game()
    statuses = freeze_game_predictions(store, game, game.scheduled_utc)
    assert statuses["espn"] == "no_pregame_prediction"


def test_results_without_pregame_prediction_blank():
    store = InMemoryStore()
    pred, was_frozen = resolve_frozen_bundle(store, "missing-game", "2026-11-07T22:00:00Z")
    assert pred is None and not was_frozen


def test_results_use_pre_tip_not_post_tip_when_unfrozen():
    """Unattended run without a freeze step still renders pre-tip values."""
    store = InMemoryStore()
    a = _pred("espn", "2026-11-07T10:00:00Z", expected_total=230.0)
    b = _pred("espn", "2026-11-07T23:30:00Z", expected_total=999.0)
    b.pre_tip = False  # pipeline stamps this at storage
    _seed(store, [a, b])
    pred, was_frozen = resolve_frozen_bundle(store, "g1", "2026-11-07T22:00:00Z")
    assert not was_frozen
    assert pred.expected_total == 230.0


def test_latest_pre_tip_boundary_at_tip_is_excluded():
    before = _pred("espn", "2026-11-07T21:59:59Z")
    at_tip = _pred("espn", "2026-11-07T22:00:00Z")
    after = _pred("espn", "2026-11-07T22:00:01Z")
    assert latest_pre_tip([at_tip, before], "2026-11-07T22:00:00Z") is before
    assert latest_pre_tip([at_tip], "2026-11-07T22:00:00Z") is None
    assert latest_pre_tip([after], "2026-11-07T22:00:00Z") is None


# ---------------------------------------------------------------------------
# Odds binding
# ---------------------------------------------------------------------------


def test_odds_follow_selected_side_same_snapshot():
    snap = _snap()
    home_pred = _pred("espn", NOW, prob_home=0.55, prob_away=0.45)
    away_pred = _pred("espn", NOW, prob_home=0.45, prob_away=0.55)
    assert odds_for_prediction(snap, home_pred) == 1.67
    assert odds_for_prediction(snap, away_pred) == 2.30


def test_odds_none_without_prediction_or_market():
    assert odds_for_prediction(_snap(), None) is None
    total_snap = BookSnapshot(
        game_id="g1",
        provider="therundown",
        book="fanduel",
        market=Market.TOTAL,
        period=Period.FULL_GAME,
        observed_utc=NOW,
        line=232.0,
        over_price_decimal=1.9,
    )
    assert odds_for_prediction(total_snap, _pred("espn", NOW)) is None


def test_tied_probabilities_do_not_create_an_arbitrary_price_selection():
    tied = _pred("espn", NOW, prob_home=0.5, prob_away=0.5)
    assert odds_for_prediction(_snap(), tied) is None


def test_prediction_flip_recomputes_price_from_same_snapshot():
    """Selection change -> which side's price is used is recomputed."""
    snap = _snap()
    before = _pred("espn", NOW, prob_home=0.55, prob_away=0.45)
    after = _pred("espn", "2026-11-07T19:30:00Z", prob_home=0.40, prob_away=0.60)
    assert odds_for_prediction(snap, before) == 1.67
    assert odds_for_prediction(snap, after) == 2.30
