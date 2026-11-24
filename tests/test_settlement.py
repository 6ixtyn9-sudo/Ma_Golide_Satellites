"""Settled picks out of the lanes' own ledgers, with every omission counted.

The tuner is only as honest as its input. A pending pick is not a loss, a moneyline
row was never recorded as a banker or a robber, and the same pick can sit in a dated
ledger and in the rolling audit at once. Each of those is a way a tuning run could
quietly lie, so each is a test.
"""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from forecasting.settlement import (
    SettlementSummary,
    bet_type_for,
    collect,
)


def _row(**overrides):
    row = {
        "game_id": "g1",
        "league": "LITHUANIA_NKL",
        "market": "moneyline",
        "period": "full_game",
        "selection": "home",
        "line": None,
        "raw_probability": 0.60,
        "calibrated_probability": 0.62,
        "outcome": "win",
        "grade_status": "win",
    }
    row.update(overrides)
    return row


def _dated(tmp_path: Path, rows, name: str = "picks_2026-11-20.json") -> Path:
    path = tmp_path / name
    path.write_text(json.dumps(rows))
    return path


def _rolling(tmp_path: Path, rows, name: str = "picks_audit_rolling.json") -> Path:
    path = tmp_path / name
    path.write_text(json.dumps({"settled_ledger": rows}))
    return path


# -- labels -----------------------------------------------------------------


@pytest.mark.parametrize(
    "market,period,expected",
    [
        ("moneyline", "full_game", "MONEYLINE"),
        ("moneyline", "H1", "FIRST_HALF_1X2"),
        ("moneyline", "Q3", "SNIPER_MARGIN"),
        ("total", "full_game", "FT_OU"),
        ("total", "Q1", "SNIPER_OU"),
        ("total", "H1", "FIRST_HALF_OU"),
        ("spread", "full_game", "SPREAD"),
        ("spread", "Q2", "SNIPER_MARGIN"),
        ("highest_quarter", "full_game", "HIGHEST_QUARTER"),
        (None, None, None),
        ("corners", "full_game", None),
    ],
)
def test_the_ledger_vocabulary_maps_to_the_bet_types(market, period, expected):
    assert bet_type_for(market, period) == expected


def test_a_moneyline_is_not_called_a_banker_by_guess():
    """The ledger never recorded banker vs robber; inventing one would skew counts."""

    assert bet_type_for("moneyline", "full_game") == "MONEYLINE"


# -- outcomes ---------------------------------------------------------------


def test_a_win_and_a_loss_come_through_as_outcomes(tmp_path):
    path = _dated(
        tmp_path,
        [_row(game_id="a", outcome="win"), _row(game_id="b", outcome="loss")],
    )

    summary = collect(path)

    assert [bet.won for bet in summary.bets] == [True, False]


def test_a_push_is_settled_but_not_a_win_or_a_loss(tmp_path):
    path = _dated(tmp_path, [_row(outcome="push")])

    summary = collect(path)

    assert summary.bets[0].won is None
    assert summary.bets[0].push is True


def test_a_pending_pick_is_counted_not_graded_as_a_loss(tmp_path):
    path = _dated(
        tmp_path,
        [_row(game_id="a", outcome=None, grade_status="pending"), _row(game_id="b")],
    )

    summary = collect(path)

    assert summary.total == 1
    assert summary.skipped["pending"] == 1


# -- confidence -------------------------------------------------------------


def test_the_calibrated_probability_is_preferred_over_the_raw_one(tmp_path):
    path = _dated(tmp_path, [_row(raw_probability=0.50, calibrated_probability=0.64)])

    assert collect(path).bets[0].stated_confidence == pytest.approx(64.0)


def test_the_raw_probability_is_used_when_no_calibration_exists(tmp_path):
    path = _dated(tmp_path, [_row(raw_probability=0.55, calibrated_probability=None)])

    assert collect(path).bets[0].stated_confidence == pytest.approx(55.0)


def test_a_pick_with_no_probability_is_counted_rather_than_given_one(tmp_path):
    path = _dated(
        tmp_path,
        [_row(game_id="a", raw_probability=None, calibrated_probability=None)],
    )

    summary = collect(path)

    assert summary.total == 0
    assert summary.skipped["no_probability"] == 1


def test_a_probability_outside_zero_to_one_is_refused(tmp_path):
    path = _dated(tmp_path, [_row(raw_probability=62.0, calibrated_probability=None)])

    assert collect(path).skipped["no_probability"] == 1


# -- direction --------------------------------------------------------------


def test_a_totals_pick_carries_its_direction(tmp_path):
    path = _dated(
        tmp_path,
        [
            _row(game_id="a", market="total", selection="over"),
            _row(game_id="b", market="total", selection="UNDER"),
        ],
    )

    assert [bet.direction for bet in collect(path).bets] == ["OVER", "UNDER"]


def test_a_moneyline_pick_has_no_direction(tmp_path):
    assert collect(_dated(tmp_path, [_row(selection="home")])).bets[0].direction is None


# -- sources ----------------------------------------------------------------


def test_the_rolling_audit_nests_its_rows_under_settled_ledger(tmp_path):
    path = _rolling(tmp_path, [_row()])

    assert collect(path).total == 1


def test_a_directory_contributes_both_dated_files_and_the_rolling_audit(tmp_path):
    _dated(tmp_path, [_row(game_id="a")], name="picks_2026-11-19.json")
    _dated(tmp_path, [_row(game_id="b")], name="picks_2026-11-20.json")
    _rolling(tmp_path, [_row(game_id="c")])

    summary = collect(tmp_path)

    assert summary.total == 3
    assert summary.files == 3


def test_the_same_pick_in_two_sources_counts_once(tmp_path):
    row = _row()
    _dated(tmp_path, [row])
    _rolling(tmp_path, [row])

    summary = collect(tmp_path)

    assert summary.total == 1
    assert summary.skipped["duplicate"] == 1


def test_two_different_lines_are_not_mistaken_for_a_duplicate(tmp_path):
    _dated(
        tmp_path,
        [
            _row(game_id="a", market="total", line=160.0),
            _row(game_id="a", market="total", line=160.5),
        ],
    )

    assert collect(tmp_path).total == 2


def test_an_unknown_market_is_counted_not_dropped_silently(tmp_path):
    summary = collect(_dated(tmp_path, [_row(market="corners")]))

    assert summary.total == 0
    assert summary.skipped["unmapped_market"] == 1


def test_a_missing_path_is_reported_and_not_fatal(tmp_path):
    summary = collect(tmp_path / "nowhere")

    assert summary.total == 0
    assert summary.skipped["path_missing"] == 1


def test_an_unreadable_file_is_counted_and_the_rest_still_loads(tmp_path):
    (tmp_path / "picks_2026-11-18.json").write_text("{broken")
    _dated(tmp_path, [_row()], name="picks_2026-11-19.json")

    summary = collect(tmp_path)

    assert summary.total == 1
    assert summary.skipped["unreadable_file"] == 1


def test_a_ledger_that_is_not_a_list_is_refused_as_unreadable(tmp_path):
    path = tmp_path / "picks_2026-11-18.json"
    path.write_text(json.dumps("not rows"))

    assert collect(path).skipped["unreadable_file"] == 1


# -- the summary ------------------------------------------------------------


def test_the_summary_counts_what_it_kept_and_what_it_left(tmp_path):
    summary = collect(_dated(tmp_path, [_row(), _row(game_id="b", outcome=None)]))

    assert isinstance(summary, SettlementSummary)
    assert summary.total == 1
    assert summary.skipped_total == 1
    assert summary.skip_line() == "pending 1"


def test_an_empty_sample_says_nothing_was_skipped(tmp_path):
    empty = collect(tmp_path)

    assert empty.total == 0
    assert empty.skip_line() == "nothing skipped"


def test_the_league_travels_with_the_pick(tmp_path):
    summary = collect(_dated(tmp_path, [_row(league="GREECE_ESAK")]))

    assert summary.bets[0].league == "GREECE_ESAK"
