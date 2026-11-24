"""Tuning happens in the tick, not in somebody's memory.

A self-tuning system that needs a human to invoke it is a manual system with extra
steps. ``scripts/daily.py`` therefore tunes from the lane's own settled ledgers on
every tick, before the run summary is written:

* a canonical tick (``--write``) applies the bounded steps and writes the lane-local
  ``magolide_bets.tuned.json`` plus ``config_ledger.jsonl``;
* a dry run reports exactly what it would have changed and writes nothing;
* ``--skip-tuning`` turns it off for one tick;
* a failure inside the tuner is reported as ``status: failed`` and never turns an
  otherwise complete lane red -- the canonical acquisition state is already
  committed by then.

These tests drive the helper directly: the full pipeline needs a warehouse and a
network, and neither is what this behaviour is about.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path


SCRIPTS = Path(__file__).resolve().parent.parent / "scripts"
sys.path.insert(0, str(SCRIPTS))

import daily  # noqa: E402


def _lane(tmp_path: Path, rows, name: str = "picks_2026-11-20.json") -> Path:
    lane = tmp_path / "lane"
    lane.mkdir(exist_ok=True)
    (lane / name).write_text(json.dumps(rows))
    return lane


def _settled(game_id: str, *, won: bool, confidence: float = 0.90) -> dict:
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


def _losers(count: int = 30) -> list:
    return [_settled(f"g{n}", won=False) for n in range(count)]


# -- the happy path ---------------------------------------------------------


def test_a_canonical_tick_tunes_from_the_lanes_own_ledger(tmp_path):
    lane = _lane(tmp_path, _losers())

    tuning = daily._tune_from_ledgers(lane, write=True, skip=False)

    assert tuning["enabled"] is True
    assert tuning["status"] == "applied"
    assert tuning["settled_picks"] == 30
    assert tuning["applied"]["confidence_scale"] == [30.0, 40.0]
    assert tuning["config_stamp"].startswith("CFG_")
    assert (lane / daily.TUNED_FILENAME).exists()
    assert (lane / daily.CONFIG_LEDGER_FILENAME).exists()


def test_the_tuning_record_is_json_serialisable_for_the_run_summary(tmp_path):
    lane = _lane(tmp_path, _losers())

    tuning = daily._tune_from_ledgers(lane, write=True, skip=False)

    assert json.loads(json.dumps(tuning)) == tuning


def test_a_well_calibrated_lane_reports_no_change_and_writes_no_config(tmp_path):
    rows = [_settled(f"g{n}", won=True, confidence=0.80) for n in range(80)]
    rows += [_settled(f"h{n}", won=False, confidence=0.80) for n in range(20)]
    lane = _lane(tmp_path, rows)

    tuning = daily._tune_from_ledgers(lane, write=True, skip=False)

    assert tuning["status"] == "no_change"
    assert tuning["settled_picks"] == 100
    assert not (lane / daily.TUNED_FILENAME).exists()


def test_a_second_tick_stacks_on_the_first_runs_config(tmp_path):
    lane = _lane(tmp_path, _losers())

    first = daily._tune_from_ledgers(lane, write=True, skip=False)
    second = daily._tune_from_ledgers(lane, write=True, skip=False)

    assert first["config_stamp"] != second["config_stamp"]
    assert second["applied"]["confidence_scale"] == [40.0, 50.0]


# -- safety -----------------------------------------------------------------


def test_a_dry_tick_reports_what_it_would_do_and_writes_nothing(tmp_path):
    lane = _lane(tmp_path, _losers())

    tuning = daily._tune_from_ledgers(lane, write=False, skip=False)

    assert tuning["status"] == "dry_run"
    assert tuning["applied"]["confidence_scale"] == [30.0, 40.0]
    assert not (lane / daily.TUNED_FILENAME).exists()
    assert not (lane / daily.CONFIG_LEDGER_FILENAME).exists()


def test_the_skip_flag_leaves_the_config_alone(tmp_path):
    lane = _lane(tmp_path, _losers())

    tuning = daily._tune_from_ledgers(lane, write=True, skip=True)

    assert tuning == {"enabled": False, "status": "skipped"}
    assert not (lane / daily.TUNED_FILENAME).exists()


def test_an_empty_lane_tunes_nothing_and_says_so(tmp_path):
    lane = tmp_path / "empty"
    lane.mkdir()

    tuning = daily._tune_from_ledgers(lane, write=True, skip=False)

    assert tuning["status"] == "no_change"
    assert tuning["settled_picks"] == 0
    assert tuning["reasons"]


def test_pending_rows_are_reported_not_counted(tmp_path):
    rows = _losers() + [
        {**_settled("pending-1", won=False), "outcome": None, "grade_status": "pending"}
    ]
    lane = _lane(tmp_path, rows)

    tuning = daily._tune_from_ledgers(lane, write=True, skip=False)

    assert tuning["settled_picks"] == 30
    assert tuning["not_counted"] == {"pending": 1}


def test_a_tuner_fault_cannot_fail_the_lane(tmp_path, monkeypatch):
    lane = _lane(tmp_path, _losers())

    def explode(*args, **kwargs):
        raise RuntimeError("ledger exploded")

    monkeypatch.setattr("forecasting.settlement.collect", explode)

    tuning = daily._tune_from_ledgers(lane, write=True, skip=False)

    assert tuning["status"] == "failed"
    assert "ledger exploded" in tuning["error"]


def test_the_lane_local_path_is_used_not_a_shared_one(tmp_path):
    """Each league is its own matrix job; a shared file would be a race."""

    lane = _lane(tmp_path, _losers())

    daily._tune_from_ledgers(lane, write=True, skip=False)

    assert (lane / "magolide_bets.tuned.json").exists()
    assert not (tmp_path / "magolide_bets.tuned.json").exists()


def test_the_card_loads_the_tuned_file_the_tick_wrote(tmp_path):
    from forecasting.config_tuner import load_tuned

    lane = _lane(tmp_path, _losers())
    tuning = daily._tune_from_ledgers(lane, write=True, skip=False)

    reloaded = load_tuned(tuned_path=lane / daily.TUNED_FILENAME)

    assert reloaded.stamp_id == tuning["config_stamp"]
    assert reloaded.league_weight("LITHUANIA_NKL") == 0.7
