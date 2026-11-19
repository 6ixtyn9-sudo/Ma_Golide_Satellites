from __future__ import annotations

from datetime import datetime, timedelta, timezone
import json
from pathlib import Path

from ingestion.adaptive_gate import (
    LEAGUES,
    ScheduleGame,
    evaluate_gate,
    evaluate_league,
    fetch_schedule,
    report_for,
)

NOW = datetime(2026, 11, 18, 12, 0, tzinfo=timezone.utc)


def _write_json(path: Path, payload: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload))


def _operator_state(root: Path, config, *, date: str, frozen: str | None = None) -> None:
    state_dir = root / config.state_key
    locks: dict[str, dict] = {}
    if frozen == "ticket":
        locks[date] = {"final": True, "receipt_path": f"auto_tickets_{date}_decision.json"}
        _write_json(
            state_dir / f"auto_tickets_{date}_decision.json",
            {"date": date, "league": config.league, "status": "frozen"},
        )
    elif frozen == "no_bet":
        locks[date] = {"final": True, "no_bet": True, "no_bet_path": f"auto_tickets_{date}_no_bet.json"}
        _write_json(
            state_dir / f"auto_tickets_{date}_no_bet.json",
            {"target_date": date, "league": config.league, "status": "frozen_no_bet"},
        )
    _write_json(state_dir / "auto_tickets_state.json", {"frozen_by_date": locks})
    _write_json(
        state_dir / "auto_tickets_performance.json",
        {
            "league": config.league,
            "pending_positions": 0,
            "pending_accas": 0,
            "pending_singles": 0,
            "awaiting_game_legs": 0,
            "unresolved_result_legs": 0,
            "result_unresolvable_legs": 0,
        },
    )


def _schedule(*games: ScheduleGame):
    def fetch(_config, _date):
        return games

    return fetch


def test_gate_skips_a_valid_idle_league(tmp_path: Path):
    nba = LEAGUES[0]
    _operator_state(tmp_path, nba, date="2026-11-18")

    decision = evaluate_league(
        nba,
        state_root=tmp_path,
        now=NOW,
        schedule_fetcher=_schedule(
            ScheduleGame(NOW + timedelta(hours=4), "scheduled"),
        ),
    )

    assert not decision.should_run
    assert decision.reasons == ("no_work",)
    assert decision.scheduled_game_count == 1


def test_gate_publishes_an_identified_fixture_draft_before_its_freeze_window(tmp_path: Path):
    nba = LEAGUES[0]
    _operator_state(tmp_path, nba, date="2026-11-18")

    decision = evaluate_league(
        nba,
        state_root=tmp_path,
        now=NOW,
        schedule_fetcher=_schedule(ScheduleGame(NOW + timedelta(hours=4), "scheduled", "game-1")),
    )

    assert decision.should_run
    assert decision.reasons == ("pick_draft_publication_due",)


def test_gate_runs_for_an_unfrozen_fixture_freeze_window(tmp_path: Path):
    nba = LEAGUES[0]
    _operator_state(tmp_path, nba, date="2026-11-18")
    state_dir = tmp_path / nba.state_key
    _write_json(state_dir / "picks_2026-11-18.json", [])

    decision = evaluate_league(
        nba,
        state_root=tmp_path,
        now=NOW,
        schedule_fetcher=_schedule(ScheduleGame(NOW + timedelta(minutes=60), "scheduled", "game-1")),
    )

    assert decision.should_run
    assert decision.reasons == (
        "fixture_pick_freeze_due",
        "ticket_creation_or_acquisition_due",
    )


def test_gate_runs_for_an_unfrozen_slate_inside_the_ticket_window(tmp_path: Path):
    nba = LEAGUES[0]
    _operator_state(tmp_path, nba, date="2026-11-18")

    decision = evaluate_league(
        nba,
        state_root=tmp_path,
        now=NOW,
        schedule_fetcher=_schedule(
            ScheduleGame(NOW + timedelta(minutes=60), "scheduled"),
        ),
    )

    assert decision.should_run
    assert decision.reasons == ("ticket_creation_or_acquisition_due",)


def test_gate_accepts_a_valid_frozen_no_bet_for_a_due_game(tmp_path: Path):
    nba = LEAGUES[0]
    _operator_state(tmp_path, nba, date="2026-11-18", frozen="no_bet")

    decision = evaluate_league(
        nba,
        state_root=tmp_path,
        now=NOW,
        schedule_fetcher=_schedule(ScheduleGame(NOW, "live")),
    )

    assert not decision.should_run
    assert decision.reasons == ("no_work",)


def test_gate_fails_open_when_pick_lifecycle_state_is_invalid(tmp_path: Path):
    nba = LEAGUES[0]
    _operator_state(tmp_path, nba, date="2026-11-18")
    state_dir = tmp_path / nba.state_key
    _write_json(state_dir / "picks_2026-11-18.json", [])
    _write_json(
        state_dir / "pick_lifecycle_state.json",
        {"schema_version": "unknown", "league": "NBA", "pending_frozen_games": []},
    )

    decision = evaluate_league(
        nba,
        state_root=tmp_path,
        now=NOW,
        schedule_fetcher=_schedule(ScheduleGame(NOW + timedelta(hours=8), "scheduled", "game-later")),
    )

    assert decision.should_run
    assert decision.reasons == ("state_uncertain: pick lifecycle state has an unknown schema",)


def test_gate_rechecks_a_tipped_frozen_pick_until_it_settles(tmp_path: Path):
    nba = LEAGUES[0]
    _operator_state(tmp_path, nba, date="2026-11-18", frozen="ticket")
    state_dir = tmp_path / nba.state_key
    _write_json(state_dir / "picks_2026-11-18.json", [])
    _write_json(
        state_dir / "pick_lifecycle_state.json",
        {
            "schema_version": "magolide.pick_lifecycle_state",
            "league": "NBA",
            "pending_frozen_games": [["game-yesterday", "2026-11-18T11:00:00Z"]],
        },
    )

    decision = evaluate_league(
        nba,
        state_root=tmp_path,
        now=NOW,
        schedule_fetcher=_schedule(ScheduleGame(NOW + timedelta(hours=8), "scheduled", "game-later")),
    )

    assert decision.should_run
    assert decision.reasons == ("frozen_pick_settlement_due",)


def test_gate_runs_pending_positions_and_awaiting_results_even_when_the_slate_is_later(
    tmp_path: Path,
):
    nba = LEAGUES[0]
    _operator_state(tmp_path, nba, date="2026-11-18", frozen="ticket")
    performance_path = tmp_path / nba.state_key / "auto_tickets_performance.json"
    performance = json.loads(performance_path.read_text())
    performance.update({"pending_positions": 1, "awaiting_game_legs": 3})
    _write_json(performance_path, performance)

    decision = evaluate_league(
        nba,
        state_root=tmp_path,
        now=NOW,
        schedule_fetcher=_schedule(ScheduleGame(NOW + timedelta(hours=8), "scheduled")),
    )

    assert decision.should_run
    assert decision.reasons == ("pending_settlement", "result_reconciliation")


def test_gate_fails_open_when_a_schedule_cannot_be_read(tmp_path: Path):
    nba = LEAGUES[0]
    _operator_state(tmp_path, nba, date="2026-11-18")

    def unavailable(_config, _date):
        raise RuntimeError("temporary provider outage")

    decision = evaluate_league(nba, state_root=tmp_path, now=NOW, schedule_fetcher=unavailable)

    assert decision.should_run
    assert decision.reasons == ("schedule_uncertain: temporary provider outage",)


def test_gate_fails_open_when_performance_is_from_the_wrong_league(tmp_path: Path):
    nba = LEAGUES[0]
    _operator_state(tmp_path, nba, date="2026-11-18")
    performance_path = tmp_path / nba.state_key / "auto_tickets_performance.json"
    performance = json.loads(performance_path.read_text())
    performance["league"] = "CBA"
    _write_json(performance_path, performance)

    decision = evaluate_league(
        nba,
        state_root=tmp_path,
        now=NOW,
        schedule_fetcher=_schedule(),
    )

    assert decision.should_run
    assert decision.reasons == (
        "state_uncertain: operator performance belongs to a different or unknown league",
    )


def test_gate_fails_open_when_a_state_lock_lacks_its_immutable_receipt(tmp_path: Path):
    nba = LEAGUES[0]
    _operator_state(tmp_path, nba, date="2026-11-18")
    state_path = tmp_path / nba.state_key / "auto_tickets_state.json"
    _write_json(state_path, {"frozen_by_date": {"2026-11-18": {"final": True}}})

    decision = evaluate_league(
        nba,
        state_root=tmp_path,
        now=NOW,
        schedule_fetcher=_schedule(),
    )

    assert decision.should_run
    assert decision.reasons[0].startswith("state_uncertain: missing frozen decision")


def test_fiba_elw_gate_uses_official_competition_schedule_and_canonical_utc(monkeypatch, tmp_path: Path):
    elw = next(config for config in LEAGUES if config.league == "EUROLEAGUE_W")
    _operator_state(tmp_path, elw, date="2026-10-14")
    fixture = json.loads(
        (
            Path(__file__).parent
            / "fixtures"
            / "fiba_euroleague_w"
            / "schedule_209127_20261118.json"
        ).read_text()
    )
    urls: list[str] = []

    def read(url, **_kwargs):
        urls.append(url)
        return fixture

    monkeypatch.setattr("ingestion.adaptive_gate._read_json_response", read)
    games = list(fetch_schedule(elw, "2026-10-14"))

    assert "gdapCompetitionId=209127" in urls[0]
    assert len(games) == 1
    assert games[0].scheduled_utc == datetime(2026, 10, 14, 17, 0, tzinfo=timezone.utc)
    assert games[0].status == "final"


def test_euroleague_unbounded_schedule_path_is_explicitly_fail_open(tmp_path: Path):
    euroleague = next(config for config in LEAGUES if config.league == "EuroLeague")
    _operator_state(tmp_path, euroleague, date="2026-11-18")

    def no_date_bounded_source(_config, _date):
        raise RuntimeError("EuroLeague has no verified date-bounded schedule source")

    decision = evaluate_league(
        euroleague,
        state_root=tmp_path,
        now=NOW,
        schedule_fetcher=no_date_bounded_source,
    )

    assert decision.should_run
    assert decision.reasons == (
        "schedule_uncertain: EuroLeague has no verified date-bounded schedule source",
    )


def test_forced_gate_matrix_keeps_all_leagues_and_reports_only_work_rows(tmp_path: Path):
    decisions = evaluate_gate(
        state_root=tmp_path,
        now=NOW,
        force_all_reason="manual_read_only_rehearsal",
    )
    report = report_for(decisions, generated_utc=NOW)

    assert report["run_any"] is True
    assert [row["league"] for row in report["matrix"]["include"]] == [
        config.league for config in LEAGUES
    ]
    assert all(row["gate_reason"] == "manual_read_only_rehearsal" for row in report["matrix"]["include"])
