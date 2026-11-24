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
