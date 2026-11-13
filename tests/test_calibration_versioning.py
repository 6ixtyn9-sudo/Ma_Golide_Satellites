"""Version-aware calibration segmentation: deterministic fingerprints,
explicit legacy handling, and no silent merging across versions."""

import pytest

from forecasting.harness import (
    CalibrationConfig,
    apply_profiles,
    build_profiles,
    config_fingerprint,
)
from forecasting.types import LEGACY_VERSION_TAG, Forecast


def _record(
    game_id: str,
    scheduled: str,
    probability: float,
    outcome,
    *,
    source: str = "provider:test",
    version=None,
    team_history=None,
    league: str = "NBA",
) -> Forecast:
    return Forecast(
        game_id=game_id,
        league=league,
        scheduled_utc=scheduled,
        home_name="Home",
        away_name="Away",
        market="moneyline",
        period="full_game",
        selection="home",
        raw_probability=probability,
        alternative_probabilities={"home": probability, "away": 1 - probability},
        source=source,
        outcome=outcome,
        calibration_version=version,
        team_history=team_history,
    )


def test_config_fingerprint_is_deterministic_and_order_insensitive():
    first = config_fingerprint({"half_life": 14.0, "prior": 30})
    second = config_fingerprint({"prior": 30, "half_life": 14.0})
    assert first == second
    assert len(first) == 12
    # 30 and 30.0 are the same configuration
    assert config_fingerprint({"half_life": 14.0, "prior": 30.0}) == config_fingerprint(
        {"half_life": 14, "prior": 30}
    )


def test_config_fingerprint_changes_with_prediction_relevant_config():
    base = config_fingerprint({"shrinkage": 8.0, "half_life": 20.0})
    assert config_fingerprint({"shrinkage": 8.0, "half_life": 21.0}) != base
    assert config_fingerprint({"shrinkage": 9.0, "half_life": 20.0}) != base


def test_config_fingerprint_rejects_unserializable_components():
    with pytest.raises(ValueError):
        config_fingerprint({"bad": {1, 2}})


def test_records_without_version_carry_explicit_legacy_tag():
    row = _record("g1", "2026-11-12T20:00:00Z", 0.7, "win")
    assert row.version_tag == LEGACY_VERSION_TAG == "legacy"
    assert row.versioned_segment_key == (
        "NBA",
        "moneyline",
        "full_game",
        "provider:test",
        "legacy",
    )
    assert row.to_dict()["calibration_version"] is None


def test_versioned_records_carry_their_tag_through_serialization():
    row = _record("g1", "2026-11-12T20:00:00Z", 0.7, "win", version="a1b2c3d4e5f6")
    assert row.version_tag == "a1b2c3d4e5f6"
    assert row.to_dict()["calibration_version"] == "a1b2c3d4e5f6"


def test_legacy_grouping_matches_historical_four_tuple_semantics():
    rows = [
        _record("g1", "2026-11-12T20:00:00Z", 0.7, "win"),
        _record("g2", "2026-11-12T21:00:00Z", 0.6, "loss"),
        _record("g3", "2026-11-12T22:00:00Z", 0.8, "win", league="NCAA"),
        _record("g4", "2026-11-12T23:00:00Z", 0.9, "win", league="NCAA"),
        _record("g5", "2026-11-13T00:00:00Z", 0.5, "push"),
    ]
    profiles = build_profiles(rows)
    manual: dict = {}
    for row in rows:
        manual.setdefault(row.segment_key, []).append(row)
    assert {key[:4] for key in profiles} == set(manual)
    for key, profile in profiles.items():
        assert key[-1] == "legacy" and len(key) == 5
        expected_n = len(
            [r for r in manual[key[:4]] if r.outcome in {"win", "loss"}]
        )
        assert profile.n == expected_n
        serialized = profile.to_dict()
        assert serialized["calibration_version"] == "legacy"
        assert serialized["league"] == key[0]


def test_versions_never_merge_and_legacy_never_merges_with_versions():
    segment_rows = [
        _record("g1", "2026-11-12T20:00:00Z", 0.7, "win", version=None),
        _record("g2", "2026-11-12T21:00:00Z", 0.7, "win", version=None),
        *[
            _record(f"v{i}", "2026-11-12T22:00:00Z", 0.7, "win", version="ver_a")
            for i in range(5)
        ],
        *[
            _record(f"w{i}", "2026-11-12T23:00:00Z", 0.7, "loss", version="ver_b")
            for i in range(5)
        ],
    ]
    profiles = build_profiles(segment_rows)
    assert len(profiles) == 3
    tags = {key[4] for key in profiles}
    assert tags == {"legacy", "ver_a", "ver_b"}
    assert profiles[(*segment_rows[0].segment_key, "legacy")].adjustment > 0
    assert profiles[(*segment_rows[0].segment_key, "ver_b")].adjustment < 0


def test_apply_matches_only_same_version_profile():
    history = [
        _record(f"h{i}", "2026-11-12T20:00:00Z", 0.7, "win") for i in range(8)
    ]
    profiles = build_profiles(history)
    legacy_forecast = _record("f1", "2026-11-13T02:00:00Z", 0.6, None)
    versioned_forecast = _record("f2", "2026-11-13T02:00:00Z", 0.6, None, version="other")
    apply_profiles([legacy_forecast, versioned_forecast], profiles, CalibrationConfig())
    assert legacy_forecast.calibration["status"] == "provisional"
    assert legacy_forecast.calibrated_probability != legacy_forecast.raw_probability
    assert legacy_forecast.calibration["calibration_version"] == "legacy"
    # A different version must not borrow another version's evidence.
    assert versioned_forecast.calibration["status"] == "no_history"
    assert versioned_forecast.calibrated_probability == versioned_forecast.raw_probability
    assert versioned_forecast.calibration["calibration_version"] == "other"


def test_legacy_calibration_math_is_unchanged():
    rows = [
        _record(f"h{i}", "2026-11-12T20:00:00Z", 0.7, "win") for i in range(40)
    ]
    profiles = build_profiles(rows)
    forecast = _record("f1", "2026-11-13T02:00:00Z", 0.7, None)
    apply_profiles([forecast], profiles, CalibrationConfig())
    expected_adjustment = (40 / (40 + 30)) * (1.0 - 0.7)
    expected = min(0.99, max(0.01, 0.7 + expected_adjustment))
    assert forecast.calibrated_probability == pytest.approx(expected)
    assert forecast.eligibility["acca_eligible"] is False
    assert "no_exact_bookmaker_market" in forecast.eligibility["reasons"]


def test_diagnostic_pool_still_serialized_and_never_applied():
    insufficient = [
        _record(
            f"d{i}",
            "2026-11-12T20:00:00Z",
            0.7,
            "win",
            team_history={"sufficient_for_accas": False},
        )
        for i in range(6)
    ]
    profiles = build_profiles(insufficient)
    diag_keys = [key for key in profiles if key[-1] == "team_history_insufficient"]
    assert diag_keys and all(len(key) == 6 for key in diag_keys)
    serialized = profiles[diag_keys[0]].to_dict()
    assert serialized["pool"] == "league_prior_in_diagnostics_only"
    assert serialized["sufficient_for_accas"] is False
    assert serialized["calibration_version"] == "legacy"
    forecast = _record(
        "f1",
        "2026-11-13T02:00:00Z",
        0.7,
        None,
        team_history={"sufficient_for_accas": True},
    )
    apply_profiles([forecast], profiles, CalibrationConfig())
    assert forecast.calibration["status"] == "no_history"


def test_profile_serialization_is_deterministic():
    rows = [
        _record(f"h{i}", "2026-11-12T20:00:00Z", 0.6 + i * 0.01, "win" if i % 2 else "loss")
        for i in range(12)
    ]
    first = [p.to_dict() for _, p in sorted(build_profiles(rows).items())]
    second = [p.to_dict() for _, p in sorted(build_profiles(rows).items())]
    assert first == second
