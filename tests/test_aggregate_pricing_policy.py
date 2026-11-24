"""The 2026-11-22 pricing and evidence policy: no silent zero-pick lanes.

Three separate refusals used to intersect so that a lane could publish a card
full of ``(none)`` for ever:

* an aggregate consensus price was never ticket-eligible, so every lane whose
  only route is the host's page average stopped at
  ``price_not_named_bookmaker_quote``;
* a short walk-forward sample and thin club history were *refusals*, so a lane
  that had settled nothing could never publish the picks that would have
  settled;
* the recency-weighted ROI band filter required a settled band, which made the
  same deadlock permanent at a second layer.

Each of those is now an explicit, labelled admission rather than a refusal, and
each admission is narrow: a provenance is only accepted if it has a published
number, evidence thinness is only an advisory, and cold start only applies while
the whole lane has settled nothing at all. These tests pin the edges of all
three, including the cases that must still be refused.
"""

from __future__ import annotations

from pathlib import Path


from forecasting.accas import filter_paper_candidate_picks
from forecasting.harness import CalibrationConfig, apply_profiles
from forecasting.types import TICKET_PRICE_PROVENANCE, BookEvidence, Forecast
from ingestion.models import (
    OperationalValidation,
    PriceProvenance,
    SourcePermission,
)

ROOT = Path(__file__).resolve().parent.parent


def _evidence(provenance: PriceProvenance, **overrides) -> BookEvidence:
    values = dict(
        provider="betexplorer",
        book="average",
        observed_utc="2026-11-22T12:00:00Z",
        source_modified_utc=None,
        pre_match=True,
        main_line=True,
        frozen=False,
        prices={"home": 1.8, "away": 2.1},
        price_provenance=provenance,
        source_permission=SourcePermission.UNRESOLVED,
        operational_validation=OperationalValidation.PROBATION,
        game_id="g",
        market="moneyline",
        period="full_game",
    )
    values.update(overrides)
    return BookEvidence(**values)


def _forecast(*, book=None, source="team_history_empirical", probability=0.6) -> Forecast:
    return Forecast(
        game_id="g",
        league="EUROLEAGUE",
        scheduled_utc="2026-11-22T18:00:00Z",
        home_name="Home",
        away_name="Away",
        market="moneyline",
        period="full_game",
        selection="home",
        raw_probability=probability,
        alternative_probabilities={"home": probability, "away": 1.0 - probability},
        source=source,
        book=book,
    )


def _pickable(
    game_id: str,
    *,
    price: float,
    probability: float,
    scheduled_utc: str = "2026-11-22T18:00:00Z",
    observed_utc: str = "2026-11-22T12:00:00Z",
    outcome: str | None = None,
) -> Forecast:
    """A priced, eligible forecast. ``decimal_price`` is the book's own quote.

    ``observed_utc`` must precede the tip, or the historical-scope filter is
    right to reject the row as post-tip evidence.
    """

    evidence = _evidence(
        PriceProvenance.AGGREGATE_CONSENSUS,
        game_id=game_id,
        observed_utc=observed_utc,
        prices={"home": price, "away": 1.0 + round(1.0 / max(price - 0.9, 0.05), 3)},
    )
    forecast = _forecast(book=evidence, probability=probability)
    forecast.game_id = game_id
    forecast.scheduled_utc = scheduled_utc
    forecast.calibrated_probability = probability
    forecast.outcome = outcome
    forecast.eligibility = {"acca_eligible": True, "reasons": []}
    return forecast


# -- the accepted provenances ------------------------------------------------


def test_the_accepted_provenances_are_named_in_one_place():
    assert TICKET_PRICE_PROVENANCE == {
        PriceProvenance.NAMED_BOOK,
        PriceProvenance.AGGREGATE_CONSENSUS,
    }


def test_an_aggregate_average_prices_a_paper_leg_but_is_never_called_a_book():
    evidence = _evidence(PriceProvenance.AGGREGATE_CONSENSUS)

    assert evidence.ticket_price_eligible is True
    assert evidence.price_provenance is PriceProvenance.AGGREGATE_CONSENSUS
    label = evidence.price_provenance_label
    assert "AGGREGATE" in label
    assert "NOT A NAMED-BOOK QUOTE" in label
    assert "named book" not in label.lower().replace("not a named-book quote", "")


def test_an_unpriced_record_and_an_exchange_match_still_never_price_a_leg():
    # A NO_PRICE record carries no number, and an exchange match is nobody's
    # published consensus -- so neither becomes a paper price by this decision.
    assert _evidence(PriceProvenance.NO_PRICE).ticket_price_eligible is False
    assert _evidence(PriceProvenance.EXCHANGE).ticket_price_eligible is False


def test_a_prohibited_source_is_refused_whatever_its_provenance():
    for provenance in (PriceProvenance.NAMED_BOOK, PriceProvenance.AGGREGATE_CONSENSUS):
        evidence = _evidence(provenance, source_permission=SourcePermission.PROHIBITED)
        assert evidence.ticket_price_eligible is False


# -- thin evidence is an advisory, not a refusal -----------------------------


def test_a_short_walk_forward_sample_publishes_and_names_itself():
    # A provider-only source needs no club history, so the short prequential
    # sample is the only shortfall left to name.
    calibrated = apply_profiles(
        [_forecast(book=_evidence(PriceProvenance.AGGREGATE_CONSENSUS), source="espn")],
        {},
        CalibrationConfig(min_samples=1),
    )[0]

    assert calibrated.eligibility["acca_eligible"] is True
    assert calibrated.eligibility["reasons"] == []
    assert calibrated.eligibility["advisories"] == ["insufficient_walk_forward_samples"]


def test_thin_club_evidence_publishes_and_names_itself():
    forecast = _forecast(book=_evidence(PriceProvenance.AGGREGATE_CONSENSUS))
    forecast.team_history = {
        "sufficient_for_accas": False,
        "home_effective_games": 1.0,
        "away_effective_games": 2.0,
        "required_effective_games_per_team": 8.0,
    }

    calibrated = apply_profiles([forecast], {}, CalibrationConfig(min_samples=1))[0]

    assert calibrated.eligibility["acca_eligible"] is True
    assert calibrated.eligibility["advisories"] == [
        "insufficient_walk_forward_samples",
        "insufficient_team_specific_history",
    ]


def test_a_push_risk_and_a_probability_floor_are_still_refusals():
    forecast = _forecast(book=_evidence(PriceProvenance.NAMED_BOOK), probability=0.2)
    forecast.push_probability = 0.4

    calibrated = apply_profiles(
        [forecast], {}, CalibrationConfig(min_samples=1, min_probability=0.5)
    )[0]

    assert calibrated.eligibility["acca_eligible"] is False
    assert "push_risk_requires_leg_void_rules" in calibrated.eligibility["reasons"]
    assert "below_calibrated_probability_floor" in calibrated.eligibility["reasons"]


def test_an_unpriced_forecast_is_still_refused():
    calibrated = apply_profiles(
        [_forecast()], {}, CalibrationConfig(min_samples=1)
    )[0]

    assert calibrated.eligibility["acca_eligible"] is False
    assert calibrated.eligibility["reasons"] == ["no_exact_bookmaker_market"]


# -- cold start admission ----------------------------------------------------


def _priced_pair():
    return [
        _pickable("g1", price=1.8, probability=0.6),
        _pickable("g2", price=2.0, probability=0.56),
    ]


def test_a_lane_that_has_settled_nothing_admits_its_bands_as_a_stamped_cold_start():
    qualifying, summary = filter_paper_candidate_picks(
        _priced_pair(), as_of_utc="2026-11-22T12:00:00Z"
    )

    assert {row.game_id for row in qualifying} == {"g1", "g2"}
    assert summary["cold_start_admission"]["applied_this_run"] is True
    assert summary["cold_start_admission"]["admitted_band_count"] >= 1
    assert summary["status"] == "applied_cold_start_no_settled_history"
    for band in summary["bands"]:
        assert band["qualification_basis"] == "cold_start_no_settled_history"
        assert band["historical_pick_count"] == 0


def test_the_cold_start_admission_can_be_turned_off():
    _, summary = filter_paper_candidate_picks(
        _priced_pair(),
        as_of_utc="2026-11-22T12:00:00Z",
        admit_cold_start_bands=False,
    )

    assert summary["qualifying_pick_count"] == 0
    assert summary["cold_start_admission"]["enabled"] is False
    assert summary["status"] == "no_settled_historical_picks"


def test_one_settled_pick_is_enough_to_end_the_admission_for_unseen_bands():
    settled = _pickable(
        "settled",
        price=1.8,
        probability=0.6,
        scheduled_utc="2026-11-20T18:00:00Z",
        observed_utc="2026-11-20T09:00:00Z",
        outcome="win",
    )

    _, summary = filter_paper_candidate_picks(
        _priced_pair(), [settled], as_of_utc="2026-11-22T12:00:00Z"
    )

    assert summary["cold_start_admission"]["applied_this_run"] is False
    # The measured band now decides, and a single win is a real measurement.
    assert all(
        band["qualification_basis"]
        != "cold_start_no_settled_history"
        for band in summary["bands"]
    )


def test_a_band_with_no_settled_picks_is_still_refused_once_the_lane_has_history():
    settled = _pickable(
        "settled",
        price=1.8,
        probability=0.6,
        scheduled_utc="2026-11-20T18:00:00Z",
        observed_utc="2026-11-20T09:00:00Z",
        outcome="loss",
    )

    # The candidates sit in a different probability and odds band from the
    # settled loser, so their band has no history of its own while the lane has
    # some -- which is exactly the case cold start must not cover.
    candidates = [
        _pickable("c1", price=1.05, probability=0.92),
        _pickable("c2", price=1.05, probability=0.93),
    ]

    _, summary = filter_paper_candidate_picks(
        candidates, [settled], as_of_utc="2026-11-22T12:00:00Z"
    )

    assert summary["cold_start_admission"]["applied_this_run"] is False
    assert summary["qualifying_pick_count"] == 0
    assert summary["status"] == "no_positive_historical_roi_band"




# -- what the card says ------------------------------------------------------


def _eligible_thin_row() -> dict:
    from forecasting.operator import _pick_row

    return _pick_row(
        {
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
                "status": "provisional",
                "n": 4,
                "sufficient_for_accas": False,
                "minimum_samples": 30,
            },
            "team_history": {
                "home_effective_games": 3.0,
                "away_effective_games": 6.5,
                "required_effective_games_per_team": 8.0,
                "sufficient_for_accas": False,
            },
            "book": {
                "provider": "betexplorer",
                "book": "average",
                "selected_decimal_price": 1.62,
                "observed_utc": "2026-11-25T16:00:00Z",
                "pre_match": True,
                "frozen": False,
                "ticket_price_eligible": True,
                "price_provenance": "aggregate_consensus",
                "source_permission": "unresolved",
            },
            "eligibility": {
                "acca_eligible": True,
                "reasons": [],
                "advisories": [
                    "insufficient_walk_forward_samples",
                    "insufficient_team_specific_history",
                ],
            },
            "lifecycle_status": "DRAFT",
        },
        "2026-11-25",
    )


def test_an_eligible_thin_row_is_a_pick_that_states_its_shortfall():
    from forecasting.operator import _render_picks

    row = _eligible_thin_row()
    assert row["bucket"] == "ELIGIBLE_CLEAN"

    text = _render_picks(
        [row],
        league="EUROLEAGUE",
        target_date="2026-11-25",
        generated_utc="2026-11-25T09:00:00Z",
        timezone_name="Europe/Paris",
    )

    assert "ELIGIBLE CLEAN PICKS" in text
    assert "acca_eligible=True" in text
    assert "price_evidence=AGGREGATE_CONSENSUS_PRICE" in text
    advisories = "evidence_advisories=insufficient_walk_forward_samples"
    assert advisories in text
    assert "insufficient_team_specific_history" in text
    assert "thin team evidence" in text
    assert "home=3.0/8.0" in text
    # It is a pick, and it must not claim the sample was sufficient.
    assert "not a sufficient-sample claim" in text
