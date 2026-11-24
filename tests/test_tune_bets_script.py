"""The tuner's CLI: analyse settled results, then apply or dry-run.

``Config_Tuner.gs`` wrote advice to a sheet. This script applies it -- and, more
importantly for a thing that edits its own gates, can be asked not to.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest

SCRIPTS = Path(__file__).resolve().parent.parent / "scripts"
sys.path.insert(0, str(SCRIPTS))

import tune_bets  # noqa: E402


def _results(tmp_path: Path, rows) -> Path:
    path = tmp_path / "settled.json"
    path.write_text(json.dumps(rows))
    return path


def _argv(tmp_path: Path, results: Path, *extra) -> list:
    return [
        "--results", str(results),
        "--tuned-path", str(tmp_path / "tuned.json"),
        "--ledger-path", str(tmp_path / "ledger.jsonl"),
        *extra,
    ]


def _losers(count: int = 30):
    return [
        {"bet_type": "BANKER", "league": "LITHUANIA_NKL", "confidence": 90, "won": False}
        for _ in range(count)
    ]


def test_a_badly_calibrated_book_is_tuned_and_written(tmp_path, capsys):
    results = _results(tmp_path, _losers())

    code = tune_bets.main(_argv(tmp_path, results))

    out = capsys.readouterr().out
    assert code == 0
    assert "settled bets: 30" in out
    assert "applied:" in out
    assert "confidence_scale: 30.0 -> 40.0" in out
    assert "config stamp: CFG_" in out
    assert (tmp_path / "tuned.json").exists()
    assert (tmp_path / "ledger.jsonl").exists()


def test_a_well_calibrated_book_writes_nothing(tmp_path, capsys):
    rows = [
        {"bet_type": "BANKER", "league": "LITHUANIA_NKL", "confidence": 80, "won": True}
        for _ in range(80)
    ] + [
        {"bet_type": "BANKER", "league": "LITHUANIA_NKL", "confidence": 80, "won": False}
        for _ in range(20)
    ]
    results = _results(tmp_path, rows)

    code = tune_bets.main(_argv(tmp_path, results))

    out = capsys.readouterr().out
    assert code == 0
    assert "no change:" in out
    assert not (tmp_path / "tuned.json").exists()


def test_a_dry_run_prints_the_move_and_writes_nothing(tmp_path, capsys):
    results = _results(tmp_path, _losers())

    code = tune_bets.main(_argv(tmp_path, results, "--dry-run"))

    out = capsys.readouterr().out
    assert code == 0
    assert "dry run: nothing written" in out
    assert not (tmp_path / "tuned.json").exists()
    assert not (tmp_path / "ledger.jsonl").exists()


def test_a_missing_results_file_is_refused(tmp_path, capsys):
    code = tune_bets.main(_argv(tmp_path, tmp_path / "absent.json"))

    assert code == 1
    assert "nothing to tune" in capsys.readouterr().err


def test_a_wrapped_payload_is_accepted(tmp_path):
    path = tmp_path / "settled.json"
    path.write_text(json.dumps({"bets": _losers(5)}))

    code = tune_bets.main(_argv(tmp_path, path, "--dry-run"))

    assert code == 0


def test_a_payload_that_is_not_an_array_is_refused(tmp_path):
    path = tmp_path / "settled.json"
    path.write_text(json.dumps("nope"))

    with pytest.raises(SystemExit):
        tune_bets.main(_argv(tmp_path, path))


def test_the_tuned_file_is_what_the_card_then_loads(tmp_path):
    from forecasting.config_tuner import load_tuned

    results = _results(tmp_path, _losers())
    tune_bets.main(_argv(tmp_path, results))

    loaded = load_tuned(tuned_path=tmp_path / "tuned.json")

    assert loaded.get("confidence_scale") == 40.0
    assert loaded.league_weight("LITHUANIA_NKL") == 0.7


# -- reading the lanes' own ledgers -----------------------------------------


def _lane(tmp_path: Path, rows, name: str = "picks_2026-11-20.json") -> Path:
    lane = tmp_path / "LITHUANIA_NKL"
    lane.mkdir(exist_ok=True)
    (lane / name).write_text(json.dumps(rows))
    return lane


def _ledger_row(game_id: str, *, won: bool, confidence: float = 0.90) -> dict:
    return {
        "game_id": game_id,
        "league": "LITHUANIA_NKL",
        "market": "moneyline",
        "period": "full_game",
        "selection": "home",
        "line": None,
        "calibrated_probability": confidence,
        "raw_probability": confidence,
        "outcome": "win" if won else "loss",
        "grade_status": "win" if won else "loss",
    }


def test_a_lane_directory_tunes_from_its_own_settled_ledger(tmp_path, capsys):
    lane = _lane(tmp_path, [_ledger_row(f"g{n}", won=False) for n in range(30)])

    code = tune_bets.main(
        [
            "--from-ledgers", str(lane),
            "--tuned-path", str(tmp_path / "tuned.json"),
            "--ledger-path", str(tmp_path / "ledger.jsonl"),
        ]
    )

    out = capsys.readouterr().out
    assert code == 0
    assert "settled bets: 30" in out
    assert "ledgers read: 1 file(s)" in out
    assert "confidence_scale: 30.0 -> 40.0" in out


def test_pending_rows_are_reported_as_not_counted(tmp_path, capsys):
    rows = [_ledger_row(f"g{n}", won=False) for n in range(30)]
    rows.append({**rows[0], "game_id": "pending-1", "outcome": None, "grade_status": "pending"})
    lane = _lane(tmp_path, [_ledger_row(f"x{n}", won=False) for n in range(5)] + rows[30:])

    code = tune_bets.main(
        [
            "--from-ledgers", str(lane),
            "--tuned-path", str(tmp_path / "tuned.json"),
            "--ledger-path", str(tmp_path / "ledger.jsonl"),
            "--dry-run",
        ]
    )

    out = capsys.readouterr().out
    assert code == 0
    assert "not counted: pending 1" in out


def test_an_empty_lane_changes_nothing(tmp_path, capsys):
    lane = tmp_path / "EMPTY"
    lane.mkdir()

    code = tune_bets.main(
        [
            "--from-ledgers", str(lane),
            "--tuned-path", str(tmp_path / "tuned.json"),
            "--ledger-path", str(tmp_path / "ledger.jsonl"),
        ]
    )

    out = capsys.readouterr().out
    assert code == 0
    assert "settled bets: 0" in out
    assert "no change:" in out
    assert not (tmp_path / "tuned.json").exists()


def test_the_two_sources_are_mutually_exclusive(tmp_path):
    with pytest.raises(SystemExit):
        tune_bets.main(
            [
                "--results", str(tmp_path / "r.json"),
                "--from-ledgers", str(tmp_path),
            ]
        )
