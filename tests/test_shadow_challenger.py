"""Recency-decayed shadow challenger: leakage-free prequential evaluation,
ESS bounds, drift/sparsity behavior, and acca-input immutability."""

import copy
import random

import pytest

from forecasting.accas import build_accas
from forecasting.harness import CalibrationConfig, apply_profiles, build_profiles
from forecasting.shadow import (
    RecencyConfig,
    annotate_recency_profiles,
    build_recency_profiles,
    prequential_shadow_evaluation,
)
from forecasting.types import BookEvidence, Forecast


def _record(
    game_id: str,
    scheduled: str,
    probability: float,
    outcome,
    *,
    version=None,
    team_history=None,
    league: str = "NBA",
    source: str = "provider:test",
    market: str = "moneyline",
    selection: str = "home",
) -> Forecast:
    alternatives = (
        {"home": probability, "away": 1 - probability}
        if selection == "home"
        else {"home": 1 - probability, "away": probability}
    )
    return Forecast(
        game_id=game_id,
        league=league,
        scheduled_utc=scheduled,
        home_name="Home",
        away_name="Away",
        market=market,
        period="full_game",
        selection=selection,
        raw_probability=probability,
        alternative_probabilities=alternatives,
        source=source,
        outcome=outcome,
        calibration_version=version,
        team_history=team_history,
    )


def test_result_delay_respected_and_boundary_is_inclusive():
    cfg = RecencyConfig(result_delay_hours=4.0)
    history = _record("h1", "2026-11-12T10:00:00Z", 0.7, "win")
    before = _record("p1", "2026-11-12T13:59:00Z", 0.6, "loss")
    after = _record("p2", "2026-11-12T14:00:00Z", 0.6, "loss")
    prequential_shadow_evaluation([history, before, after], cfg)
    # The 10:00 result is available only at 14:00 under the delay rule.
    assert before.shadow_calibration["status"] == "no_history"
    assert before.shadow_calibration["probability"] == before.raw_probability
    assert after.shadow_calibration["n"] == 1
    assert after.shadow_calibration["adjustment"] > 0  # saw a win, raw was 0.6
    assert after.shadow_calibration["probability"] != after.raw_probability


def test_zero_delay_does_not_leak_outcomes_between_forecasts_for_same_game():
    tip = "2026-11-12T12:00:00Z"
    home = _record("same-game", tip, 0.7, "win", selection="home")
    away = _record("same-game", tip, 0.3, "loss", selection="away")
    prequential_shadow_evaluation(
        [home, away], RecencyConfig(result_delay_hours=0.0)
    )
    assert home.shadow_calibration["status"] == "no_history"
    assert away.shadow_calibration["status"] == "no_history"
    assert home.shadow_calibration["n"] == away.shadow_calibration["n"] == 0


def test_shadow_champion_mirror_uses_supplied_epoch_prior_and_probability_bounds():
    old = _record("old", "2026-11-11T10:00:00Z", 0.9, "win")
    history = [
        _record(f"h{i}", "2026-11-12T10:00:00Z", 0.6, "win") for i in range(3)
    ]
    probe = _record("probe", "2026-11-12T20:00:00Z", 0.7, "loss")
    champion_cfg = CalibrationConfig(
        prior_strength=2,
        min_samples=2,
        probability_floor=0.4,
        probability_ceiling=0.8,
        calibration_epoch_utc="2026-11-12T00:00:00Z",
    )
    challenger_cfg = RecencyConfig(half_life_days=1e9, result_delay_hours=1.0)
    profiles = build_profiles(history, champion_cfg)
    apply_profiles([probe], profiles, champion_cfg)

    summary = prequential_shadow_evaluation(
        [old, *history, probe], challenger_cfg, champion_config=champion_cfg
    )

    assert summary["rows_evaluated"] == 4
    assert summary["rows_ungraded_or_ineligible_skipped"] == 1
    assert probe.shadow_calibration["champion_n"] == 3
    assert probe.shadow_calibration["champion_status"] == "sufficient"
    assert probe.shadow_calibration["champion_probability"] == probe.calibrated_probability
    assert probe.calibrated_probability == 0.8  # champion-specific upper clamp


def test_input_order_does_not_change_results():
    cfg = RecencyConfig(result_delay_hours=2.0)
    rows = [
        _record("h1", "2026-11-12T08:00:00Z", 0.7, "win"),
        _record("h2", "2026-11-12T09:00:00Z", 0.65, "loss"),
        _record("p1", "2026-11-12T14:00:00Z", 0.6, "win"),
        _record("p2", "2026-11-12T15:00:00Z", 0.55, "loss"),
        _record("p3", "2026-11-12T16:00:00Z", 0.5, "win"),
    ]
    first = prequential_shadow_evaluation(list(rows), cfg)
    shuffled = rows[:]
    random.Random(7).shuffle(shuffled)
    second = prequential_shadow_evaluation(shuffled, cfg)
    assert first == second


def test_duplicate_forecast_ids_remain_deterministic_when_rows_are_reordered():
    cfg = RecencyConfig(result_delay_hours=0.0)
    duplicate_win = _record("dup", "2026-11-12T10:00:00Z", 0.65, "win")
    duplicate_loss = _record("dup", "2026-11-12T10:00:00Z", 0.75, "loss")
    later = _record("later", "2026-11-12T12:00:00Z", 0.6, "win")
    assert duplicate_win.forecast_id == duplicate_loss.forecast_id

    first = [duplicate_win, duplicate_loss, later]
    second = [duplicate_loss, later, duplicate_win]
    first_summary = prequential_shadow_evaluation(first, cfg)
    second_summary = prequential_shadow_evaluation(second, cfg)

    assert first_summary == second_summary
    assert duplicate_win.shadow_calibration["status"] == "no_history"
    assert duplicate_loss.shadow_calibration["status"] == "no_history"
    assert duplicate_win.shadow_calibration["probability"] == 0.65
    assert duplicate_loss.shadow_calibration["probability"] == 0.75
    assert later.shadow_calibration["n"] == 2


def test_only_earlier_available_outcomes_inform_each_prediction():
    cfg = RecencyConfig(result_delay_hours=1.0)
    a = _record("a", "2026-11-12T10:00:00Z", 0.7, "win")  # available 11:00
    b = _record("b", "2026-11-12T12:00:00Z", 0.65, "loss")  # available 13:00
    c = _record("c", "2026-11-12T11:10:00Z", 0.6, "loss")  # available 12:10
    d = _record("d", "2026-11-12T10:30:00Z", 0.55, "loss")  # available 11:30
    e = _record("e", "2026-11-12T13:00:00Z", 0.5, "loss")  # available 14:00
    prequential_shadow_evaluation([a, b, c, d, e], cfg)
    assert d.shadow_calibration["n"] == 0  # a not yet available at 10:30
    assert c.shadow_calibration["n"] == 1  # only a (11:00); d arrives 11:30
    assert b.shadow_calibration["n"] == 2  # a and d; c arrives 12:10
    # e at 13:00 sees a (11:00), d (11:30), c (12:10), b (13:00, boundary inclusive)
    assert e.shadow_calibration["n"] == 4


def test_kish_ess_bounds_and_decay():
    uniform = [
        _record(f"u{i}", "2026-11-12T19:00:00Z", 0.7, "win") for i in range(5)
    ]
    cfg = RecencyConfig(half_life_days=1e6, result_delay_hours=0.5)
    profile = build_recency_profiles(uniform, cfg, as_of_utc="2026-11-12T20:00:00Z")
    (value,) = profile.values()
    assert value.ess == pytest.approx(5.0)  # uniform weights: ess == n
    # One fresh and one heavily aged outcome: uneven weights shrink ess below n.
    mixed = [
        _record("fresh", "2026-11-12T19:00:00Z", 0.7, "win"),
        _record("aged", "2026-11-12T01:00:00Z", 0.7, "loss"),
    ]
    cfg_short = RecencyConfig(half_life_days=0.25, result_delay_hours=0.5)
    aged = build_recency_profiles(mixed, cfg_short, as_of_utc="2026-11-12T20:00:00Z")
    (value,) = aged.values()
    assert 0 < value.ess < value.n
    assert value.n == 2


def test_challenger_config_version_is_complete_and_part_of_profile_key():
    default = RecencyConfig()
    stricter = RecencyConfig(min_samples=31.0)
    assert default.version_tag != stricter.version_tag

    rows = [_record("versioned", "2026-11-12T19:00:00Z", 0.7, "win")]
    default_profiles = build_recency_profiles(
        rows, default, as_of_utc="2026-11-13T00:00:00Z"
    )
    stricter_profiles = build_recency_profiles(
        rows, stricter, as_of_utc="2026-11-13T00:00:00Z"
    )
    assert set(default_profiles).isdisjoint(stricter_profiles)
    (profile,) = default_profiles.values()
    assert profile.to_dict()["challenger"] == "recency_v1"
    assert profile.to_dict()["challenger_version"] == default.version_tag


@pytest.mark.parametrize(
    "kwargs",
    [
        {"half_life_days": float("nan")},
        {"prior_strength": float("inf")},
        {"min_samples": float("inf")},
        {"result_delay_hours": float("nan")},
        {"probability_floor": float("nan")},
        {"probability_ceiling": float("inf")},
    ],
)
def test_recency_config_rejects_non_finite_values(kwargs):
    with pytest.raises(ValueError, match="finite"):
        RecencyConfig(**kwargs)


def test_sparse_history_stays_shrunk_toward_zero():
    cfg = RecencyConfig(min_samples=30, result_delay_hours=0.5)
    rows = [
        _record("s1", "2026-11-12T19:00:00Z", 0.7, "win"),
        _record("s2", "2026-11-12T19:30:00Z", 0.7, "win"),
    ]
    profile = build_recency_profiles(rows, cfg, as_of_utc="2026-11-12T20:00:00Z")
    (value,) = profile.values()
    assert value.sufficient is False
    assert value.weighted_n < cfg.min_samples
    assert abs(value.adjustment) < 1.0 - 0.7  # shrinkage, never the raw gap


def test_recency_tracks_distribution_drift_champion_does_not():
    cfg = RecencyConfig(half_life_days=0.5, min_samples=30, result_delay_hours=0.5)
    old = [
        _record(f"o{i}", "2026-11-10T12:00:00Z", 0.7, "loss") for i in range(40)
    ]
    recent = [
        _record(f"r{i}", "2026-11-12T19:00:00Z", 0.7, "win") for i in range(40)
    ]
    challenger = build_recency_profiles(
        old + recent, cfg, as_of_utc="2026-11-12T20:00:00Z"
    )
    (value,) = challenger.values()
    assert value.adjustment > 0  # recency sees the recent wins
    champion = build_profiles(old + recent)
    (champion_profile,) = champion.values()
    assert champion_profile.adjustment < 0  # static pooling sees the old losses


def test_champion_mirror_matches_harness_math_without_decay():
    rows = [
        _record(
            f"h{i}",
            "2026-11-12T12:00:00Z",
            0.6 + (i % 5) * 0.05,
            "win" if i % 3 else "loss",
        )
        for i in range(37)
    ]
    probe = _record("probe", "2026-11-12T20:00:00Z", 0.65, "loss")
    profiles = build_profiles(rows)
    (profile,) = profiles.values()
    apply_profiles([probe], profiles, CalibrationConfig())
    # Huge half-life approximates the champion's undecayed pooling; the probe's
    # own result is not available before its tip (delay > 0), so no leakage.
    summary = prequential_shadow_evaluation(
        rows + [probe], RecencyConfig(half_life_days=1e9, result_delay_hours=1.0)
    )
    shadow = probe.shadow_calibration
    assert shadow["n"] == profile.n
    # shadow_calibration stores adjustment rounded to 6 decimals.
    assert shadow["adjustment"] == pytest.approx(profile.adjustment, abs=1e-6)
    assert shadow["probability"] == pytest.approx(probe.calibrated_probability, rel=1e-6)
    assert summary["rows_evaluated"] == 38


def _eligible_forecast(game_id: str) -> Forecast:
    return Forecast(
        game_id=game_id,
        league="NBA",
        scheduled_utc="2026-11-13T02:00:00Z",
        home_name="Home",
        away_name="Away",
        market="moneyline",
        period="full_game",
        selection="home",
        raw_probability=0.7,
        alternative_probabilities={"home": 0.7, "away": 0.3},
        source="provider:test",
        outcome="win",
        team_history={"sufficient_for_accas": True},
        book=BookEvidence(
            provider="test",
            book="book",
            observed_utc="2026-11-13T00:00:00Z",
            source_modified_utc=None,
            pre_match=True,
            main_line=True,
            frozen=True,
            prices={"home": 1.8, "away": 2.0},
        ),
    )


def test_shadow_evaluation_cannot_change_acca_inputs_or_selection():
    history = [
        _record(f"h{i}", "2026-11-12T12:00:00Z", 0.7, "win") for i in range(40)
    ]
    profiles = build_profiles(history)
    first = _eligible_forecast("f1")
    second = _eligible_forecast("f2")
    apply_profiles([first, second], profiles, CalibrationConfig())
    assert first.eligibility["acca_eligible"] and second.eligibility["acca_eligible"]
    accas_before = build_accas([first, second], legs=2)
    inputs_before = {
        "calibrated": [first.calibrated_probability, second.calibrated_probability],
        "calibration": [dict(first.calibration), dict(second.calibration)],
        "eligibility": [copy.deepcopy(first.eligibility), copy.deepcopy(second.eligibility)],
    }
    explanation_before = dict(first.to_dict()["explanation"])

    summary = prequential_shadow_evaluation(
        history + [first, second], RecencyConfig()
    )
    inputs_after = {
        "calibrated": [first.calibrated_probability, second.calibrated_probability],
        "calibration": [dict(first.calibration), dict(second.calibration)],
        "eligibility": [copy.deepcopy(first.eligibility), copy.deepcopy(second.eligibility)],
    }
    assert inputs_after == inputs_before
    # Selection identity is unchanged (acca_id hashes leg forecast_ids).
    def _identity(accas):
        return [
            (
                acca["acca_id"],
                acca["rank"],
                tuple(leg["forecast_id"] for leg in acca["legs"]),
                tuple(leg["calibrated_probability"] for leg in acca["legs"]),
            )
            for acca in accas
        ]

    accas_after = build_accas([first, second], legs=2)
    assert _identity(accas_after) == _identity(accas_before)
    # The serialized legs differ only inside the additive explanation block,
    # and only by the shadow ESS entry.
    legs_before = accas_before[0]["legs"][0]
    legs_after = accas_after[0]["legs"][0]
    assert {k for k in legs_after if legs_after[k] != legs_before[k]} == {"explanation"}
    assert set(legs_after["explanation"]) - set(legs_before["explanation"]) == {
        "effective_sample_size",
        "shadow_calibration",
    }
    # The explanation gains only explicit shadow diagnostics; nothing else moves.
    explanation_after = dict(first.to_dict()["explanation"])
    assert set(explanation_after) - set(explanation_before) == {
        "effective_sample_size",
        "shadow_calibration",
    }
    assert summary["rows_evaluated"] == 42
    assert first.shadow_calibration["status"] == "sufficient"


def test_current_shadow_annotation_is_diagnostic_only_for_accas():
    history = [
        _record(f"h{i}", "2026-11-12T12:00:00Z", 0.7, "win") for i in range(40)
    ]
    first, second = _eligible_forecast("f1"), _eligible_forecast("f2")
    apply_profiles([first, second], build_profiles(history), CalibrationConfig())
    before_inputs = [
        (forecast.calibrated_probability, copy.deepcopy(forecast.eligibility))
        for forecast in (first, second)
    ]
    before_accas = build_accas([first, second], legs=2)
    cfg = RecencyConfig()
    shadow_profiles = build_recency_profiles(
        history, cfg, as_of_utc="2026-11-13T01:00:00Z"
    )

    annotate_recency_profiles([first, second], shadow_profiles, cfg)

    after_inputs = [
        (forecast.calibrated_probability, copy.deepcopy(forecast.eligibility))
        for forecast in (first, second)
    ]
    after_accas = build_accas([first, second], legs=2)
    assert after_inputs == before_inputs
    assert [row["acca_id"] for row in after_accas] == [row["acca_id"] for row in before_accas]
    assert [row["estimated_joint_win_probability"] for row in after_accas] == [
        row["estimated_joint_win_probability"] for row in before_accas
    ]
    explanation = first.to_dict()["explanation"]["shadow_calibration"]
    assert explanation["challenger_version"] == cfg.version_tag
    assert explanation["n"] == 40
    assert explanation["champion_probability"] == first.calibrated_probability
    assert explanation["candidate_probability"] == first.shadow_calibration["probability"]


def test_incremental_matches_naive_rebuild_on_synthetic_stream():
    cfg = RecencyConfig(half_life_days=3.0, result_delay_hours=2.0, min_samples=5.0)
    rng = random.Random(20261113)
    rows = []
    hour = 0
    for i in range(300):
        hour += 1
        tip = f"2026-11-12T{hour % 24:02d}:00:00Z"
        probability = round(rng.uniform(0.55, 0.9), 3)
        outcome = "win" if rng.random() < probability - 0.05 else "loss"
        rows.append(_record(f"g{i}", tip, probability, outcome))
    prequential_shadow_evaluation(list(rows), cfg)
    checked = 0
    for row in rows:
        naive = build_recency_profiles(rows, cfg, as_of_utc=row.scheduled_utc)
        key = (*row.versioned_segment_key, cfg.version_tag)
        expected = naive.get(key)
        shadow = row.shadow_calibration
        if expected is None:
            assert shadow["status"] == "no_history"
        else:
            assert shadow["probability"] == pytest.approx(
                expected.apply(row.raw_probability, cfg), abs=1e-9
            )
            assert shadow["n"] == expected.n
            checked += 1
    assert checked > 250
