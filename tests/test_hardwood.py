"""Hardwood forecast artifact adapter tests.

Offline: the artifact comes from the committed fixture
``tests/fixtures/hardwood/game_forecasts.json`` (a slice captured LIVE from
the public artifact on 2026-11-06 — LIVE_VERIFIED).  Helpers carry the unique
``hw_`` prefix so they cannot collide with other agents' conftest helpers.

Probabilities remain canonical 0..1 fractions throughout acquisition.
"""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from ingestion.identity import normalize_team_name
from ingestion.models import CoverageGap, CoverageState, Game, League
from ingestion.providers import hardwood as hw

FIXTURES = Path(__file__).parent / "fixtures" / "hardwood"

# Deterministic clocks: the fixture's generated_at is 2026-11-06T17:50:17+00:00.
NOW_FRESH = "2026-11-06T18:00:00Z"  # ~10 min after generated_at -> OK
NOW_STALE = "2026-11-07T18:00:00Z"  # ~24.2h after generated_at -> STALE

ARTIFACT_SHA256 = (
    "9f61ad305c6cf38cb9acf494a6510530ee8f66e9c72f393be89f565804283ef0"
)


def hw_load_artifact() -> dict:
    with open(FIXTURES / "game_forecasts.json") as f:
        return json.load(f)


def hw_fetch(artifact=None):
    """FetchFn serving the artifact for the Hardwood URL."""
    data = hw_load_artifact() if artifact is None else artifact
    return lambda url: data


def hw_make_game(
    game_id: str = "401909287",
    home: str = "Atlanta Hawks",
    away: str = "Washington Wizards",
    scheduled: str = "2026-11-07T00:00:00Z",
    espn_id: str = "401909287",
) -> Game:
    """Canonical NBA game as ESPN discovery would produce it."""
    return Game(
        game_id=game_id,
        league=League.NBA,
        season="2026-27",
        home_team_id="ATL",
        away_team_id="WSH",
        home_name=home,
        away_name=away,
        scheduled_utc=scheduled,
        provider_ids={"espn": espn_id} if espn_id is not None else {},
    )


# ---------------------------------------------------------------------------
# load_forecasts / metadata validation
# ---------------------------------------------------------------------------


def test_artifact_metadata_validation_ok():
    artifact = hw.load_forecasts(hw_fetch())
    assert isinstance(artifact, hw.ForecastsArtifact)
    assert artifact.generated_at == "2026-11-06T17:50:17+00:00"
    assert artifact.ingestion_status == "ok"
    assert artifact.model_version == "2026.08.1+0d463d63"
    assert artifact.trained_through == "2026-11-06T03:00:00+00:00"
    assert artifact.artifact_sha256 == ARTIFACT_SHA256
    assert set(artifact.games_by_id) == {
        "401909287",
        "401909938",
        "401909288",
        "401909289",
        "401909290",
    }
    entry = artifact.games_by_id["401909287"]
    assert entry["home"]["name"] == "Atlanta Hawks"
    assert entry["p_home"] == pytest.approx(0.713845)

    # Structural breakage raises ValueError naming the missing field, so
    # callers treat the provider as unhealthy instead of parsing garbage.
    for dropped in ("generated_at", "trained_through", "artifact_sha256", "games"):
        broken = hw_load_artifact()
        broken.pop(dropped)
        with pytest.raises(ValueError, match=dropped):
            hw.load_forecasts(hw_fetch(broken))

    unparseable = hw_load_artifact()
    unparseable["generated_at"] = "not-a-timestamp"
    with pytest.raises(ValueError, match="generated_at"):
        hw.load_forecasts(hw_fetch(unparseable))


# ---------------------------------------------------------------------------
# fetch_prediction — happy path, join, provenance
# ---------------------------------------------------------------------------


def test_join_by_espn_game_id():
    prediction = hw.fetch_prediction(hw_make_game(), hw_fetch(), now_utc=NOW_FRESH)
    assert prediction is not None
    # Join key: the artifact's game_id == the ESPN event id.
    assert prediction.game_id == "401909287"
    # Probabilities stay 0..1 exactly as stored in the artifact.
    assert prediction.prob_home == pytest.approx(0.713845)
    assert prediction.prob_away == pytest.approx(0.286155)
    assert prediction.expected_home_score == pytest.approx(117.1)
    assert prediction.expected_away_score == pytest.approx(109.47)
    assert prediction.expected_total == pytest.approx(226.57)
    # exp_margin is home-minus-away and is stored as-is (never re-derived).
    assert prediction.expected_margin == pytest.approx(7.636)
    # Provenance.
    assert prediction.provider == "hardwood"
    assert prediction.model == "hardwood"
    assert prediction.model_version == "2026.08.1+0d463d63"
    assert prediction.produced_by == "roni-altshuler/nba_predictor"
    assert prediction.source_revision == ARTIFACT_SHA256
    assert prediction.observed_utc == "2026-11-06T17:50:17Z"
    assert prediction.source_modified_utc == "2026-11-06T17:50:17Z"
    assert prediction.pre_tip is True


def test_join_rejected_on_team_mismatch():
    """Same ESPN id but different teams -> None + gap; never a best-guess."""
    game = hw_make_game(home="X", away="Y")
    gaps: list = []
    prediction = hw.fetch_prediction(game, hw_fetch(), now_utc=NOW_FRESH, gaps=gaps)
    assert prediction is None
    assert len(gaps) == 1
    gap = gaps[0]
    assert isinstance(gap, CoverageGap)
    assert gap.reason is CoverageState.MISSING
    assert "team identity mismatch" in gap.detail


def test_join_rejected_on_time_mismatch():
    """Same matchup off by > 30 min (TIME_TOLERANCE) must not join across
    dates: team-only matching is forbidden."""
    game = hw_make_game(scheduled="2026-11-08T00:00:00Z")  # 24h later
    assert normalize_team_name(game.home_name) == "atlanta hawks"
    gaps: list = []
    prediction = hw.fetch_prediction(game, hw_fetch(), now_utc=NOW_FRESH, gaps=gaps)
    assert prediction is None
    assert len(gaps) == 1
    assert gaps[0].reason is CoverageState.MISSING
    assert "date mismatch" in gaps[0].detail


def test_stale_artifact_returns_none_or_state():
    """24h after generated_at (freshness 12h) the state is STALE and the
    join refuses to emit a prediction."""
    artifact = hw.load_forecasts(hw_fetch())
    assert hw.freshness_state(artifact, now_utc=NOW_STALE, freshness_hours=12.0) is (
        CoverageState.STALE
    )
    gaps: list = []
    prediction = hw.fetch_prediction(
        hw_make_game(), hw_fetch(), freshness_hours=12.0, now_utc=NOW_STALE, gaps=gaps
    )
    assert prediction is None
    assert len(gaps) == 1
    assert gaps[0].reason is CoverageState.STALE

    # Within the window the state is OK.
    assert hw.freshness_state(artifact, now_utc=NOW_FRESH) is CoverageState.OK


def test_never_averages_providers():
    """Hardwood returns ONE coherent single-provider bundle or nothing.

    No blending with ESPN BPI (whose golden total for 401909938 is 230.586 —
    deliberately different from Hardwood's 225.23): provider label, model and
    every number come from the one matched artifact entry.
    """
    prediction = hw.fetch_prediction(hw_make_game(), hw_fetch(), now_utc=NOW_FRESH)
    assert prediction is not None
    assert prediction.provider == "hardwood"  # exactly one provider label
    assert prediction.model == "hardwood"
    assert prediction.produced_by == "roni-altshuler/nba_predictor"
    # Every field belongs to the same single artifact entry.
    assert prediction.prob_home == pytest.approx(0.713845)
    assert prediction.prob_away == pytest.approx(0.286155)
    assert prediction.expected_home_score == pytest.approx(117.1)
    assert prediction.expected_away_score == pytest.approx(109.47)
    assert prediction.expected_total == pytest.approx(226.57)
    # The ESPN BPI golden total (230.586) must never leak into this bundle.
    assert prediction.expected_total != pytest.approx(230.586)


# ---------------------------------------------------------------------------
# freshness_state / rejection details
# ---------------------------------------------------------------------------


def test_failed_ingestion_status_is_provider_error_not_stale():
    broken = hw_load_artifact()
    broken["ingestion"]["status"] = "failed"
    artifact = hw.load_forecasts(hw_fetch(broken))
    assert hw.freshness_state(artifact, now_utc=NOW_FRESH) is (
        CoverageState.PROVIDER_ERROR
    )
    gaps: list = []
    assert hw.fetch_prediction(
        hw_make_game(), hw_fetch(broken), now_utc=NOW_FRESH, gaps=gaps
    ) is None
    assert gaps[0].reason is CoverageState.PROVIDER_ERROR


def test_structurally_broken_artifact_returns_none_not_raise():
    broken = hw_load_artifact()
    broken.pop("trained_through")
    gaps: list = []
    prediction = hw.fetch_prediction(
        hw_make_game(), hw_fetch(broken), now_utc=NOW_FRESH, gaps=gaps
    )
    assert prediction is None
    assert gaps[0].reason is CoverageState.PROVIDER_ERROR


def test_unknown_game_id_returns_missing_gap():
    gaps: list = []
    prediction = hw.fetch_prediction(
        hw_make_game(game_id="401999999", espn_id="401999999"),
        hw_fetch(),
        now_utc=NOW_FRESH,
        gaps=gaps,
    )
    assert prediction is None
    assert gaps[0].reason is CoverageState.MISSING
    assert "401999999" in gaps[0].detail


def test_swapped_home_away_is_identity_mismatch_gap_not_prediction():
    """Orientation matters: swapping the sides must not still join."""
    game = hw_make_game(home="Washington Wizards", away="Atlanta Hawks")
    prediction = hw.fetch_prediction(game, hw_fetch(), now_utc=NOW_FRESH)
    assert prediction is None


def test_self_contradictory_margin_quarantines_as_provider_error():
    """A margin that contradicts the entry's own expected scores (e.g. a
    flipped sign) is never rewritten: the prediction quarantines."""
    broken = hw_load_artifact()
    broken["games"][0]["exp_margin"] = -7.636  # contradicts 117.1 - 109.47
    gaps: list = []
    prediction = hw.fetch_prediction(
        hw_make_game(), hw_fetch(broken), now_utc=NOW_FRESH, gaps=gaps
    )
    assert prediction is None
    assert gaps[0].reason is CoverageState.PROVIDER_ERROR


def test_out_of_range_probabilities_quarantine():
    broken = hw_load_artifact()
    broken["games"][0]["p_home"] = 71.4  # percentage scale, not 0..1
    broken["games"][0]["p_away"] = 28.6
    prediction = hw.fetch_prediction(hw_make_game(), hw_fetch(broken), now_utc=NOW_FRESH)
    assert prediction is None
