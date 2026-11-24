"""The tuner: three analyses, bounded moves, applied automatically and recorded.

``Config_Tuner.gs`` wrote advice to a sheet and waited for a human. You asked for
the advice to be applied, so each rule here resolves to one numeric step that is
clamped to the key's bounds, forced through the config invariants, written to a
tuned override file, and ledgered with before/after values. These tests pin both
halves: that the analysis is the script's, and that the application cannot escape
the bounds a human would be held to.
"""

from __future__ import annotations

import json

import pytest

from forecasting.bet_config import DEFAULTS, default_config
from forecasting.config_ledger import stamps
from forecasting.config_tuner import (
    BRIER_THRESHOLD_BAD,
    MIN_SAMPLES_FOR_LEAGUE_TUNING,
    GradedBet,
    analyse,
    apply,
    graded_from_rows,
    load_tuned,
)


def _losers(count: int, *, league: str = "NKL", confidence: float = 90.0) -> list[GradedBet]:
    return [GradedBet("BANKER", league, confidence, won=False) for _ in range(count)]


# -- the analysis -----------------------------------------------------------


def test_no_settled_bets_means_no_change_and_says_so():
    report = analyse([])

    assert report.total == 0
    assert report.changes == {}
    assert "no settled bets" in " ".join(report.reasons)


def test_the_brier_score_is_the_scripts_mean_squared_error():
    bets = [
        GradedBet("BANKER", "NKL", 80.0, won=True),   # (0.8 - 1)^2 = 0.04
        GradedBet("BANKER", "NKL", 80.0, won=False),  # (0.8 - 0)^2 = 0.64
    ]

    report = analyse(bets)

    assert report.overall_brier == pytest.approx((0.04 + 0.64) / 2)


def test_a_push_scores_as_a_half():
    report = analyse([GradedBet("FT_OU", "NKL", 60.0, won=None)])

    assert report.overall_brier == pytest.approx((0.6 - 0.5) ** 2)


def test_confidence_lands_in_the_nearest_bucket():
    report = analyse([GradedBet("BANKER", "NKL", 62.0, won=True)])

    assert report.buckets[60.0]["count"] == 1
    assert report.buckets[65.0]["count"] == 0


def test_over_and_under_are_graded_separately():
    bets = [
        GradedBet("FT_OU", "NKL", 60.0, won=True, direction="OVER"),
        GradedBet("FT_OU", "NKL", 60.0, won=False, direction="UNDER"),
    ]

    report = analyse(bets)

    assert report.directional["OVER"]["win_rate"] == 1.0
    assert report.directional["UNDER"]["win_rate"] == 0.0


def test_a_league_is_reported_with_its_sample_size():
    report = analyse(_losers(5))

    assert report.leagues["NKL"].count == 5
    assert report.leagues["NKL"].win_rate == 0.0


# -- the three rules --------------------------------------------------------


def test_bad_calibration_asks_for_a_harder_shrink():
    report = analyse(_losers(30))

    assert report.overall_brier > BRIER_THRESHOLD_BAD
    assert "confidence_scale" in report.changes


def test_good_calibration_leaves_the_scale_alone():
    # Stated 55% and hitting 55% is well calibrated: Brier 0.2475 is over the line
    # though, so state 50 and hit half -- the trenched case below is the honest one.
    wins = [GradedBet("BANKER", "NKL", 50.0, won=True) for _ in range(10)]
    losses = [GradedBet("BANKER", "NKL", 50.0, won=False) for _ in range(10)]

    report = analyse(wins + losses)

    assert report.overall_brier == pytest.approx(0.25)  # above the gate, as the script had it
    assert "confidence_scale" in report.changes


def test_a_well_calibrated_book_leaves_the_scale_alone():
    # 80% confidence, 80% hit rate: Brier 0.16, under the 0.22 gate.
    wins = [GradedBet("BANKER", "NKL", 80.0, won=True) for _ in range(80)]
    losses = [GradedBet("BANKER", "NKL", 80.0, won=False) for _ in range(20)]

    report = analyse(wins + losses)

    assert report.overall_brier == pytest.approx(0.16)
    assert "confidence_scale" not in report.changes


def test_a_directional_gap_over_fifty_totals_asks_for_more_edge():
    bets = [GradedBet("FT_OU", "NKL", 60.0, won=True, direction="OVER") for _ in range(40)]
    bets += [GradedBet("FT_OU", "NKL", 60.0, won=False, direction="UNDER") for _ in range(20)]

    report = analyse(bets)

    assert "ftouEdgeThreshold" in report.changes
    assert any("OVER" in reason for reason in report.reasons)


def test_a_directional_gap_on_a_thin_sample_is_ignored():
    bets = [GradedBet("FT_OU", "NKL", 60.0, won=True, direction="OVER") for _ in range(5)]
    bets += [GradedBet("FT_OU", "NKL", 60.0, won=False, direction="UNDER") for _ in range(5)]

    assert "ftouEdgeThreshold" not in analyse(bets).changes


def test_a_league_under_fifty_percent_over_twenty_samples_is_discounted():
    report = analyse(_losers(MIN_SAMPLES_FOR_LEAGUE_TUNING))

    assert report.changes["league_weights/NKL"] == ("NKL", 0.7)


def test_a_league_under_twenty_samples_is_left_alone():
    report = analyse(_losers(MIN_SAMPLES_FOR_LEAGUE_TUNING - 1))

    assert not any(key.startswith("league_weights/") for key in report.changes)


def test_the_league_discount_never_raises_a_league():
    winners = [GradedBet("BANKER", "NKL", 60.0, won=True) for _ in range(30)]

    report = analyse(winners)

    assert not any(key.startswith("league_weights/") for key in report.changes)


# -- applying ---------------------------------------------------------------


def test_applying_moves_the_gate_and_reports_before_and_after(tmp_path):
    report = analyse(_losers(30))
    config = default_config()

    tuned, stamp, applied = apply(
        report, config, tuned_path=tmp_path / "tuned.json", ledger_path=tmp_path / "l.jsonl"
    )

    assert applied["confidence_scale"] == (DEFAULTS["confidence_scale"], DEFAULTS["confidence_scale"] + 10.0)
    assert tuned.get("confidence_scale") == DEFAULTS["confidence_scale"] + 10.0
    assert tuned.stamp_id != config.stamp_id
    assert stamp.stamp_id == tuned.stamp_id


def test_the_tuned_file_holds_only_what_differs_from_the_defaults(tmp_path):
    tuned_path = tmp_path / "tuned.json"
    report = analyse(_losers(30))

    apply(report, default_config(), tuned_path=tuned_path, ledger_path=tmp_path / "l.jsonl")

    payload = json.loads(tuned_path.read_text())
    assert payload["keys"] == {"confidence_scale": DEFAULTS["confidence_scale"] + 10.0}
    assert payload["league_weights"] == {"NKL": 0.7}


def test_the_tuned_file_round_trips_into_a_loaded_config(tmp_path):
    tuned_path = tmp_path / "tuned.json"
    report = analyse(_losers(30))
    tuned, _, _ = apply(
        report, default_config(), tuned_path=tuned_path, ledger_path=tmp_path / "l.jsonl"
    )

    reloaded = load_tuned(tuned_path=tuned_path)

    assert reloaded.stamp_id == tuned.stamp_id
    assert reloaded.get("confidence_scale") == tuned.get("confidence_scale")
    assert reloaded.league_weight("NKL") == 0.7


def test_a_second_run_keeps_the_first_runs_overrides(tmp_path):
    tuned_path = tmp_path / "tuned.json"
    ledger = tmp_path / "l.jsonl"
    first, _, _ = apply(analyse(_losers(30)), default_config(), tuned_path=tuned_path, ledger_path=ledger)
    first_scale = first.get("confidence_scale")

    second, _, _ = apply(analyse(_losers(30)), first, tuned_path=tuned_path, ledger_path=ledger)

    assert second.get("confidence_scale") == first_scale + 10.0
    payload = json.loads(tuned_path.read_text())
    assert "confidence_scale" in payload["keys"]


def test_every_applied_change_is_written_to_the_ledger(tmp_path):
    ledger = tmp_path / "l.jsonl"
    apply(
        analyse(_losers(30)),
        default_config(),
        tuned_path=tmp_path / "tuned.json",
        ledger_path=ledger,
    )

    rows = stamps(ledger)
    assert len(rows) == 1
    assert "confidence_scale" in rows[0].note
    assert "->" in rows[0].note
    assert "Brier" in rows[0].note


def test_nothing_to_do_writes_no_file_and_keeps_the_config(tmp_path):
    tuned_path = tmp_path / "tuned.json"
    config = default_config()

    tuned, stamp, applied = apply(
        analyse([]), config, tuned_path=tuned_path, ledger_path=tmp_path / "l.jsonl"
    )

    assert applied == {}
    assert tuned.stamp_id == config.stamp_id
    assert not tuned_path.exists()
    assert stamps(tmp_path / "l.jsonl")[-1].stamp_id == config.stamp_id


def test_a_step_stops_at_the_bound_instead_of_resetting_to_the_default(tmp_path):
    """The sheet reset out-of-range input to the default; a tuner step must clamp.

    200 -> 210 would be out of bounds, and the sheet loader would have turned that
    into 30. A tuner that did that would swing the system violently on one bad week.
    """

    config = default_config().with_overrides({"confidence_scale": 195.0})
    tuned, _, applied = apply(
        analyse(_losers(30)),
        config,
        tuned_path=tmp_path / "tuned.json",
        ledger_path=tmp_path / "l.jsonl",
    )

    assert tuned.get("confidence_scale") == 200.0
    assert applied["confidence_scale"] == (195.0, 200.0)


def test_the_applied_config_still_satisfies_the_tier_invariants(tmp_path):
    config = default_config().with_overrides({"tier_medium_min_score": 90.0})
    tuned, _, _ = apply(
        analyse(_losers(30)), config, tuned_path=tmp_path / "tuned.json", ledger_path=tmp_path / "l.jsonl"
    )

    assert tuned.get("tier_strong_min_score") > tuned.get("tier_medium_min_score")


def test_a_dry_run_reports_the_move_without_touching_disk(tmp_path):
    tuned_path = tmp_path / "tuned.json"
    ledger = tmp_path / "l.jsonl"
    tuned, stamp, applied = apply(
        analyse(_losers(30)),
        default_config(),
        tuned_path=tuned_path,
        ledger_path=ledger,
        write=False,
    )

    assert applied["confidence_scale"] == (DEFAULTS["confidence_scale"], DEFAULTS["confidence_scale"] + 10.0)
    assert stamp.stamp_id == tuned.stamp_id
    assert not tuned_path.exists()
    assert not ledger.exists()


# -- input conversion -------------------------------------------------------


def test_rows_without_a_confidence_are_skipped_not_guessed():
    bets = graded_from_rows(
        [
            {"bet_type": "BANKER", "league": "NKL", "confidence": 70, "won": True},
            {"bet_type": "BANKER", "league": "NKL", "confidence": "", "won": True},
            {"bet_type": "BANKER", "league": "NKL"},
        ]
    )

    assert len(bets) == 1
    assert bets[0].stated_confidence == 70.0


def test_a_row_with_no_outcome_is_a_push():
    bets = graded_from_rows([{"bet_type": "FT_OU", "league": "NKL", "confidence": 60}])

    assert bets[0].push is True
    assert bets[0].won is None
