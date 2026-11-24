"""Prior-heavy picks: visible as labelled paper rows, never promoted to eligible.

Live census (2026-11-24 latest dated pick files): 48 NCAA_M rows carried a model
probability, a matched named-book price and a calibrated segment (n=371), and
were withheld on exactly one reason -- ``insufficient_team_specific_history``,
because each team's own recency-weighted evidence was thinner than the 8-game
prior (``ForecastConfig.team_prior_games``).

Discarding them hid a real forecast; silently promoting them would break the
rule that a league-mean prior is not team evidence.  The compromise asserted
here: the row is rendered in its own labelled bucket with the shortfall stated
numerically, while eligibility (``acca_eligible``) is left exactly as the model
computed it -- no acca leg, no ticket, no ROI band promotion.

Superseded in part by operator decision 2026-11-22: ``apply_profiles`` no longer
emits this reason at all, so a *new* thin row is an eligible pick carrying
``advisories`` instead of a skipped one. This bucket therefore survives for the
dated ledgers already written under the old contract -- which is why these tests
build their rows by hand rather than from a live forecast.
"""

from forecasting.accas import is_provisional_sample_only
from forecasting.operator import (
    _bucket,
    _pick_row,
    _prior_heavy_disclosure,
    _render_picks,
)
from forecasting.types import Forecast

REASONS_TEAM_HISTORY = ["insufficient_team_specific_history"]


def _forecast_dict(**overrides) -> dict:
    row = {
        "forecast_id": "E2026_379:moneyline:full_game:home:team_history_empirical",
        "game_id": "E2026_379",
        "league": "EUROLEAGUE",
        "scheduled_utc": "2026-11-25T18:00:00Z",
        "home": "Real Madrid",
        "away": "Olympiacos",
        "market": "moneyline",
        "period": "full_game",
        "selection": "home",
        "line": None,
        "model": "team_history_empirical",
        "model_version": "team_history_empirical_v1",
        "raw_probability": 0.71,
        "calibrated_probability": 0.685,
        "push_probability": 0.0,
        "calibration": {
            "status": "calibrated",
            "n": 371,
            "sufficient_for_accas": True,
            "minimum_samples": 30,
        },
        "team_history": {
            "home_effective_games": 3.0,
            "away_effective_games": 6.5,
            "required_effective_games_per_team": 8.0,
            "sufficient_for_accas": False,
        },
        "book": {
            "provider": "betexplorer_direct",
            "book": "DraftKings",
            "selected_decimal_price": 1.62,
            "observed_utc": "2026-11-25T16:00:00Z",
            "pre_match": True,
            "frozen": False,
            "ticket_price_eligible": True,
            "source_permission": "unresolved",
        },
        "eligibility": {"acca_eligible": False, "reasons": list(REASONS_TEAM_HISTORY)},
        "lifecycle_status": "DRAFT",
    }
    row.update(overrides)
    return row


def test_thin_team_evidence_is_its_own_bucket():
    row = _pick_row(_forecast_dict(), "2026-11-25")

    assert row["bucket"] == "PRIOR_HEAVY"
    assert row["prior_heavy"] is True
    # Surfacing the row is not eligibility: the model's own verdict stands.
    assert row["eligibility"]["acca_eligible"] is False
    assert row["eligibility"]["reasons"] == REASONS_TEAM_HISTORY


def test_prior_heavy_does_not_widen_any_other_gate():
    # A second blocker keeps the row skipped; only the single-reason case is
    # promoted to a visible paper bucket.
    both = _forecast_dict(
        eligibility={
            "acca_eligible": False,
            "reasons": [
                "insufficient_team_specific_history",
                "below_calibrated_probability_floor",
            ],
        }
    )
    assert _bucket(both) == "SKIPPED_INELIGIBLE"

    unpriced = _forecast_dict(book=None)
    assert _bucket(unpriced) == "WATCHLIST_NO_ODDS"

    samples = _forecast_dict(
        eligibility={"acca_eligible": False, "reasons": ["insufficient_walk_forward_samples"]}
    )
    assert _bucket(samples) == "CAUTION"

    clean = _forecast_dict(eligibility={"acca_eligible": True, "reasons": []})
    assert _bucket(clean) == "ELIGIBLE_CLEAN"


def test_prior_heavy_rows_are_never_acca_legs():
    forecast = Forecast(
        game_id="E2026_379",
        league="EUROLEAGUE",
        scheduled_utc="2026-11-25T18:00:00Z",
        home_name="Real Madrid",
        away_name="Olympiacos",
        market="moneyline",
        period="full_game",
        selection="home",
        raw_probability=0.71,
        alternative_probabilities={"away": 0.29},
        source="team_history_empirical",
        team_history=_forecast_dict()["team_history"],
        calibrated_probability=0.685,
        calibration=_forecast_dict()["calibration"],
        eligibility={
            "acca_eligible": False,
            "reasons": list(REASONS_TEAM_HISTORY),
            "evidence_state": "current_pre_tip_unfrozen",
        },
    )
    assert is_provisional_sample_only(forecast) is False


def test_rendered_row_states_the_shortfall_numerically():
    row = _pick_row(_forecast_dict(), "2026-11-25")
    text = _render_picks(
        [row],
        league="EUROLEAGUE",
        target_date="2026-11-25",
        generated_utc="2026-11-25T15:00:00Z",
        timezone_name="Europe/Madrid",
    )

    assert "PRIOR-HEAVY — THIN TEAM EVIDENCE (PAPER ONLY)" in text
    assert "PRIOR_HEAVY=" in text
    disclosure = _prior_heavy_disclosure(row)
    assert "home=3.0/8.0" in disclosure
    assert "away=6.5/8.0" in disclosure
    assert "not acca-eligible, no ticket" in disclosure
    # Value at the observed quote is stated, never implied: 1.62 break-even is
    # 61.7%, the row's p is 68.5%, so the honest figure is +11.0%.
    assert "price break-even 61.7%" in disclosure
    assert "value +11.0%" in disclosure
    assert disclosure in text


def test_disclosure_survives_a_row_without_team_history():
    # Legacy frozen rows predate the team_history field; the line must still
    # print honestly instead of raising or inventing numbers.
    assert "home=n/a" in _prior_heavy_disclosure({"team_history": None})
