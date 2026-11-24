"""Lane status: retired and off-season lanes leave the run, loudly.

The contract these tests hold:

* a dark lane is decided before any state read and before any schedule probe,
  so it costs nothing and never fabricates an authoritative empty slate;
* every dark lane carries a named reason, so its absence is explained rather
  than silent;
* retirement is reversible by one status flip and destroys no data; and
* forcing the gate open is a scheduling override, not a revival.
"""

from __future__ import annotations

from dataclasses import replace
from datetime import datetime, timezone
import json
from pathlib import Path
import subprocess
import sys

import pytest

from ingestion.adaptive_gate import (
    LEAGUES,
    LeagueGateConfig,
    ScheduleGame,
    evaluate_gate,
    evaluate_league,
    report_for,
)
from ingestion.league_registry import (
    LaneStatus,
    SPECS,
    league_spec,
    parse_league,
    retirement_guard,
    schedulable_specs,
)

NOW = datetime(2026, 11, 24, 12, 0, tzinfo=timezone.utc)

DARK_SPECS = tuple(spec for spec in SPECS if not spec.is_schedulable)


def _config(league_key: str) -> LeagueGateConfig:
    return next(config for config in LEAGUES if config.league_key == league_key)


def _must_not_fetch(_config, _date):
    raise AssertionError("a dark lane must not reach a schedule probe")


def test_dark_lane_population_is_exactly_the_measured_one():
    """The retired set is a decision, not an accident of the data."""

    retired = {spec.league.value for spec in SPECS if spec.status is LaneStatus.RETIRED}
    assert retired == {
        "ARGENTINA_LIGA_NACIONAL",
        "BASKETBALL_AFRICA_LEAGUE",
        "BASKETBALL_CHAMPIONS_LEAGUE",
        "BRAZIL_NBB",
        "FRANCE_LNB_PRO_A",
        "GERMANY_BBL",
        "ISRAEL_LIGA_LEUMIT",
        "ITALY_LBA",
        "SPAIN_ACB",
        "TURKEY_BSL",
    }
    # The operator's keep-list, stated explicitly so a future edit has to be
    # deliberate: the Chinese league is in season and the WNBA is dark only on
    # the calendar.
    assert league_spec(parse_league("CBA")).status is LaneStatus.ACTIVE
    assert league_spec(parse_league("WNBA")).status is LaneStatus.OFFSEASON
    # ABA League was considered and rejected as a retirement: its own state
    # holds priced, named-book picks, so retiring it would have been a mistake.
    assert league_spec(parse_league("ABA_LEAGUE")).status is LaneStatus.ACTIVE


def test_no_retired_lane_has_ever_priced_a_pick():
    """The rule that caught a wrong retirement, kept as a guard.

    A lane whose own frozen-pick state carries a matched quote has proved its
    price leg works.  Retiring it would hide a lane that is merely young behind
    a label that says it cannot work.
    """

    repo = Path(__file__).resolve().parents[1]
    offenders = []
    for spec in SPECS:
        if spec.status is not LaneStatus.RETIRED:
            continue
        state = repo / "localdata" / spec.state_key / "pick_lifecycle_state.json"
        if not state.exists():
            continue
        payload = json.loads(state.read_text())
        for games in (payload.get("frozen_by_date") or {}).values():
            for record in games.values():
                if any(row.get("book") for row in record.get("rows") or []):
                    offenders.append(spec.league.value)
                    break
    assert offenders == [], (
        "a retired lane has a matched-price pick in its own state; revisit the "
        f"retirement instead of hiding it: {sorted(set(offenders))}"
    )


@pytest.mark.parametrize("spec", DARK_SPECS, ids=lambda spec: spec.league.value)
def test_every_dark_lane_names_a_reason(spec):
    assert spec.status_detail.strip()
    assert spec.lane_reason.startswith(f"lane_{spec.status.value}: ")


@pytest.mark.parametrize("spec", DARK_SPECS, ids=lambda spec: spec.league.value)
def test_dark_lane_is_decided_before_state_or_fetch(tmp_path: Path, spec):
    config = _config(spec.league.value)

    decision = evaluate_league(
        config, state_root=tmp_path, now=NOW, schedule_fetcher=_must_not_fetch
    )

    assert decision.should_run is False
    assert decision.reasons == (config.lane_reason,)
    assert decision.scheduled_game_count is None
    # No state directory was created or read on behalf of the lane.
    assert not (tmp_path / spec.state_key).exists()


def test_offseason_lane_is_not_reported_as_a_failure(tmp_path: Path):
    wnba = _config("WNBA")

    decision = evaluate_league(
        wnba, state_root=tmp_path, now=NOW, schedule_fetcher=_must_not_fetch
    )

    assert decision.should_run is False
    assert decision.reasons[0].startswith("lane_offseason: ")
    # An off-season lane is stated as a calendar fact, never as a defect.
    assert decision.reasons[0] != "no_work"


def test_gate_matrix_contains_only_schedulable_lanes(tmp_path: Path):
    decisions = evaluate_gate(
        state_root=tmp_path, now=NOW, schedule_fetcher=_must_not_fetch
    )
    report = report_for(decisions, generated_utc=NOW)

    included = [row["league_key"] for row in report["matrix"]["include"]]
    assert included == [spec.league.value for spec in schedulable_specs()]
    assert all(row["should_run"] is False for row in report["decisions"] if row["league_key"] not in included)
    # Dark lanes stay visible, with a name for the reason.
    dark_rows = {
        row["league_key"]: row for row in report["decisions"] if not row["should_run"]
    }
    assert set(dark_rows) == {spec.league.value for spec in DARK_SPECS}
    assert all(row["reasons"][0].startswith("lane_") for row in dark_rows.values())


def test_forced_gate_override_does_not_resurrect_a_dark_lane(tmp_path: Path):
    decisions = evaluate_gate(
        state_root=tmp_path,
        now=NOW,
        force_all_reason="manual_read_only_rehearsal",
    )
    report = report_for(decisions, generated_utc=NOW)

    included = [row["league_key"] for row in report["matrix"]["include"]]
    assert set(included) == {spec.league.value for spec in schedulable_specs()}
    assert "WNBA" not in included
    retired_rows = [
        row
        for row in report["decisions"]
        if row["lane_status"] == LaneStatus.RETIRED.value
    ]
    assert retired_rows
    assert all(row["reasons"][0].startswith("lane_retired: ") for row in retired_rows)


def test_revival_is_a_one_line_flip(tmp_path: Path):
    retired = league_spec(parse_league("SPAIN_ACB"))
    assert retired.status is LaneStatus.RETIRED

    revived = replace(retired, status=LaneStatus.ACTIVE, status_detail="")
    assert revived.is_schedulable
    assert revived.lane_reason == ""
    # The rest of the contract survives the flip untouched.
    assert revived.price_sources == retired.price_sources
    assert revived.state_key == retired.state_key


def test_retirement_guard_names_the_reason_and_yields_to_explicit_opt_in():
    league = parse_league("SPAIN_ACB")

    refusal = retirement_guard(league)
    assert refusal is not None
    assert "retired" in refusal
    assert league_spec(league).status_detail[:20] in refusal
    assert retirement_guard(league, allow_retired=True) is None


def test_retirement_guard_does_not_block_an_active_or_offseason_lane():
    assert retirement_guard(parse_league("CBA")) is None
    assert retirement_guard(parse_league("WNBA")) is None


def test_daily_cli_refuses_a_retired_lane_without_the_opt_in(tmp_path: Path):
    repo = Path(__file__).resolve().parents[1]
    result = subprocess.run(
        [
            sys.executable,
            str(repo / "scripts" / "daily.py"),
            "--league",
            "SPAIN_ACB",
            "--date",
            "2026-11-24",
            "--warehouse",
            str(tmp_path / "w.duckdb"),
            "--localdata",
            str(tmp_path / "localdata"),
        ],
        cwd=repo,
        capture_output=True,
        text=True,
        check=False,
    )

    assert result.returncode == 2, result.stderr
    payload = json.loads(result.stderr)
    assert payload["status"] == "refused"
    assert payload["lane_status"] == "retired"
    assert "retired" in payload["reason"]


def _idle_operator_state(root: Path, config: LeagueGateConfig) -> None:
    """Minimal valid operator state so the gate can reach its schedule probe."""

    state_dir = root / config.state_key
    state_dir.mkdir(parents=True, exist_ok=True)
    (state_dir / "auto_tickets_state.json").write_text(json.dumps({"frozen_by_date": {}}))
    (state_dir / "auto_tickets_performance.json").write_text(
        json.dumps(
            {
                "league": config.league_key,
                "pending_positions": 0,
                "pending_accas": 0,
                "pending_singles": 0,
                "awaiting_game_legs": 0,
                "unresolved_result_legs": 0,
                "result_unresolvable_legs": 0,
            }
        )
    )


def test_active_lane_still_reaches_its_schedule_probe(tmp_path: Path):
    """Anti-regression: the guard must not swallow an active lane."""

    nba = _config("NBA")
    _idle_operator_state(tmp_path, nba)
    calls: list[str] = []

    def fetch(config, date):
        calls.append(date)
        return [ScheduleGame(NOW, "scheduled")]

    decision = evaluate_league(
        nba, state_root=tmp_path, now=NOW, schedule_fetcher=fetch
    )

    assert calls, "an active lane must still be probed"
    assert decision.scheduled_game_count == 1
