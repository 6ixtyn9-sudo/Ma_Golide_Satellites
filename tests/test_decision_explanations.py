"""Structured decision explanations: consistency with the calculations,
stable serialization, backward-compatible key surface, no invented fields."""

import json

from forecasting.harness import CalibrationConfig, apply_profiles, build_profiles
from forecasting.types import BookEvidence, Forecast

LEGACY_TO_DICT_KEYS = {
    "forecast_id",
    "game_id",
    "league",
    "scheduled_utc",
    "home",
    "away",
    "market",
    "period",
    "selection",
    "line",
    "role",
    "raw_probability",
    "calibrated_probability",
    "push_probability",
    "alternative_probabilities",
    "expected_margin",
    "expected_total",
    "model",
    "model_version",
    "training_games",
    "residual_games",
    "team_history",
    "prediction",
    "book",
    "outcome",
    "calibration",
    "eligibility",
}

KNOWN_ELIGIBILITY_REASONS = {
    "no_exact_bookmaker_market",
    "missing_selected_decimal_price",
    "insufficient_walk_forward_samples",
    "insufficient_team_specific_history",
    "push_risk_requires_leg_void_rules",
    "below_calibrated_probability_floor",
}


def _history() -> list:
    return [
        Forecast(
            game_id=f"h{i}",
            league="NBA",
            scheduled_utc="2026-11-12T12:00:00Z",
            home_name="Home",
            away_name="Away",
            market="moneyline",
            period="full_game",
            selection="home",
            raw_probability=0.7,
            alternative_probabilities={"home": 0.7, "away": 0.3},
            source="provider:test",
            outcome="win" if i % 3 else "loss",
            team_history={"sufficient_for_accas": True},
        )
        for i in range(40)
    ]


def _priced_forecast() -> Forecast:
    return Forecast(
        game_id="f1",
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
        team_history={"sufficient_for_accas": True},
        prediction_provider="test",
        prediction_observed_utc="2026-11-13T00:00:00Z",
        prediction_frozen=True,
        prediction_model="test-model",
        prediction_model_version="v1",
        prediction_produced_by="test producer",
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


def test_explanation_mirrors_the_calculations_that_produced_it():
    forecast = _priced_forecast()
    profiles = build_profiles(_history())
    apply_profiles([forecast], profiles, CalibrationConfig())
    explanation = forecast.to_dict()["explanation"]
    assert explanation["calibration"] == forecast.calibration
    assert explanation["eligibility"] == {
        "acca_eligible": forecast.eligibility["acca_eligible"],
        "reasons": forecast.eligibility["reasons"],
        "evidence_state": forecast.eligibility["evidence_state"],
    }
    assert explanation["segment"] == {
        "league": "NBA",
        "market": "moneyline",
        "period": "full_game",
        "source": "provider:test",
        "calibration_version": "legacy",
    }
    assert explanation["provenance"]["prediction"] == {
        "provider": "test",
        "model": "test-model",
        "model_version": "v1",
        "produced_by": "test producer",
        "observed_utc": "2026-11-13T00:00:00Z",
        "frozen": True,
    }
    assert explanation["provenance"]["book"] == {
        "book": "book",
        "provider": "test",
        "observed_utc": "2026-11-13T00:00:00Z",
        "pre_match": True,
        "frozen": True,
        "main_line": True,
    }
    assert explanation["prices"]["selected_decimal_price"] == forecast.decimal_price
    assert explanation["prices"]["opposite_decimal_price"] == forecast.opposite_decimal_price
    assert "effective_sample_size" not in explanation  # shadow never ran


def test_serialization_is_stable():
    forecast = _priced_forecast()
    apply_profiles([forecast], build_profiles(_history()), CalibrationConfig())
    first = forecast.to_dict()
    second = forecast.to_dict()
    assert first == second
    assert json.dumps(first, sort_keys=True) == json.dumps(second, sort_keys=True)


def test_to_dict_key_surface_stays_backward_compatible():
    forecast = _priced_forecast()
    keys = set(forecast.to_dict())
    assert keys == LEGACY_TO_DICT_KEYS | {"calibration_version", "explanation"}
    assert forecast.to_dict()["calibration_version"] is None


def test_no_invented_fields_or_labels():
    forecast = _priced_forecast()
    apply_profiles([forecast], build_profiles(_history()), CalibrationConfig())
    blob = json.dumps(forecast.to_dict()).lower()
    assert "confidence" not in blob
    assert "gold" not in blob
    explanation = forecast.to_dict()["explanation"]
    assert explanation["calibration"]["status"] in {"no_history", "calibrated", "provisional"}
    assert all(
        reason in KNOWN_ELIGIBILITY_REASONS
        for reason in explanation["eligibility"]["reasons"]
    )
    assert isinstance(explanation["eligibility"]["acca_eligible"], bool)


def test_bookless_and_unprocessed_forecasts_explain_their_actual_state():
    bare = Forecast(
        game_id="f2",
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
    )
    explanation = bare.to_dict()["explanation"]
    assert explanation["calibration"] is None  # calibration has not run
    assert explanation["eligibility"] is None
    assert explanation["provenance"]["book"] is None
    assert explanation["provenance"]["prediction"] is None
    assert explanation["prices"]["selected_decimal_price"] is None
    assert explanation["segment"]["calibration_version"] == "legacy"
    # After calibration without book evidence the gap is explained, not hidden.
    apply_profiles([bare], build_profiles(_history()), CalibrationConfig())
    explanation = bare.to_dict()["explanation"]
    assert "no_exact_bookmaker_market" in explanation["eligibility"]["reasons"]
    assert explanation["calibration"]["status"] == "calibrated"


def test_shadow_ess_appears_only_when_shadow_ran():
    from forecasting.shadow import RecencyConfig, prequential_shadow_evaluation

    forecast = _priced_forecast()
    forecast.outcome = "win"
    rows = _history() + [forecast]
    assert "effective_sample_size" not in forecast.to_dict()["explanation"]
    prequential_shadow_evaluation(rows, RecencyConfig())
    explanation = forecast.to_dict()["explanation"]
    assert isinstance(explanation["effective_sample_size"], (int, float))
    assert explanation["effective_sample_size"] == forecast.shadow_calibration["ess"]
    assert explanation["shadow_calibration"] == {
        **forecast.shadow_calibration,
        "candidate_probability": forecast.shadow_calibration["probability"],
    }
