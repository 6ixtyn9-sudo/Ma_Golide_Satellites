from __future__ import annotations

import argparse
import json
from pathlib import Path

import pytest

from forecasting.service import MarketForecastService
from ingestion.models import BookSnapshot, Game, League, Market, Period
from ingestion.store import InMemoryStore
from scripts.daily import _market_cutoff, _write_aba_history_receipts, _write_coverage_gap_marker
from scripts.daily import main as daily_main
from scripts.daily import run_daily


def _args(tmp_path, **changes):
    values = dict(
        auto_once=True,
        date="2026-11-07",
        league="NBA",
        tz=None,
        warehouse=str(tmp_path / "localdata" / "warehouse.duckdb"),
        localdata=str(tmp_path / "localdata"),
        keep_days=30,
        run_id="test-run",
        write=False,
        offline_fixture=True,
        skip_results=False,
        freshness_hours=12.0,
        book=None,
        book_provider="none",
        skip_markets=False,
        acca_legs=3,
        acca_limit=5,
        notify=False,
        send_notifications=False,
        started_utc="2026-11-07T18:00:00Z",
        # Every test here writes a one-game warehouse into tmp_path, so comparing
        # it against the repository's published checkpoint is meaningless -- and
        # the comparison is not inert: scripts.daily falls back to
        # localdata/<state_key>/archive_manifest.json relative to the working
        # directory, so the moment the scheduled tick publishes a real checkpoint
        # onto main, this file's tests fail on a checkout that has done nothing
        # wrong. That is exactly what turned every scheduled run red after the
        # 2026-11-23 10:34Z publish. The guard itself is covered by
        # tests/test_daily_checkpoint_guard.py, which exercises both outcomes
        # against a manifest it controls.
        allow_checkpoint_regression=True,
        skip_backfill=False,
    )
    values.update(changes)
    return argparse.Namespace(**values)


def test_source_unavailable_coverage_marker_is_stable_and_write_gated(tmp_path):
    assert _write_coverage_gap_marker(
        tmp_path, league="FRANCE_LNB_PRO_A", target_date="2026-11-18", write_enabled=False
    ) is None
    path = _write_coverage_gap_marker(
        tmp_path, league="FRANCE_LNB_PRO_A", target_date="2026-11-18", write_enabled=True
    )
    assert path is not None
    first = Path(path).read_bytes()
    assert b'"status": "source_unavailable"' in first
    assert _write_coverage_gap_marker(
        tmp_path, league="FRANCE_LNB_PRO_A", target_date="2026-11-18", write_enabled=True
    ) == path
    assert Path(path).read_bytes() == first
    assert _write_coverage_gap_marker(
        tmp_path, league="NBA", target_date="2026-11-18", write_enabled=True
    ) is None


def test_a_failed_scan_writes_a_marker_and_an_empty_slate_does_not(tmp_path):
    """Zero rows must say which zero it was.

    2026-11-25: nine lanes scanned zero rows and published no marker, after
    writing markers for days beforehand. The writer only fired for
    `schedule_source == "source_unavailable"`, so the day a real provider started
    failing -- exactly when the operator needed the signal -- the markers stopped.
    """

    provider_gap = {
        "game_id": None,
        "required_field": "schedule",
        "reason": "provider_error",
        "provider": "betexplorer",
        "detail": "U2001: eurocup schedule for U2001 is empty; a wrong season code answers 200 with no rows",
        "observed_utc": "2026-11-25T04:00:00Z",
    }

    path = _write_coverage_gap_marker(
        tmp_path,
        league="EUROCUP",
        target_date="2026-11-25",
        write_enabled=True,
        scan_rows=0,
        fixtures_found=0,
        gaps=(provider_gap,),
    )
    assert path is not None
    marker = json.loads(Path(path).read_text())
    assert marker["status"] == "scan_failed"
    assert marker["gap_reasons"] == ["provider_error"]
    assert "U2001" in marker["detail"]
    assert marker["attempted_utc"]
    assert marker["scan_rows"] == 0

    # An honest empty slate is not a gap: no fixtures and no failure says nothing
    # about whether the lane is broken.
    assert (
        _write_coverage_gap_marker(
            tmp_path,
            league="NBA",
            target_date="2026-11-25",
            write_enabled=True,
            scan_rows=0,
            fixtures_found=0,
            gaps=(),
        )
        is None
    )
    # A lane that scanned rows is not a gap either, whatever its gaps list says.
    assert (
        _write_coverage_gap_marker(
            tmp_path,
            league="EUROCUP",
            target_date="2026-11-26",
            write_enabled=True,
            scan_rows=26,
            fixtures_found=4,
            gaps=(provider_gap,),
        )
        is None
    )
    # A transport failure with no recorded gap is still a failed scan.
    transport = _write_coverage_gap_marker(
        tmp_path,
        league="GREECE_BASKET_LEAGUE",
        target_date="2026-11-25",
        write_enabled=True,
        scan_rows=0,
        fixtures_found=0,
        gaps=(),
        transport_errors=12,
    )
    assert transport is not None
    assert json.loads(Path(transport).read_text())["transport_errors"] == 12


def test_aba_history_receipts_publish_only_completed_historical_seasons(tmp_path):
    historical = {
        "season_end_year": 2026,
        "calendar_final_count": 185,
        "reconciled_final_count": 183,
        "administrative_disposition_count": 2,
        "not_applicable_rows": 0,
        "blocking_gap_count": 0,
        "status": "calendar_terminal_count_reconciled",
        "source": "aba_liga",
        "evidence": "ROUTE_CONTRACT_REFERENCED",
        "completed_at_utc": "2026-11-20T13:40:00Z",
    }
    current = {**historical, "season_end_year": 2027, "completed_at_utc": "2026-11-20T14:00:00Z"}

    path = _write_aba_history_receipts(
        tmp_path, [historical, current], current_season_end_year=2027
    )

    assert path is not None
    summary = json.loads(Path(path).read_text())
    assert summary == {
        "schema_version": "magolide.aba_history_receipts",
        "league": "ABA_LEAGUE",
        "provider": "aba_liga",
        "receipts": {"2026": historical},
    }
    # An equivalent re-attestation preserves the original durable receipt and
    # does not cause a timestamp-only publication change.
    assert _write_aba_history_receipts(
        tmp_path,
        [{**historical, "completed_at_utc": "2026-11-20T15:00:00Z"}],
        current_season_end_year=2027,
    ) == path
    assert json.loads(Path(path).read_text()) == summary
    with pytest.raises(ValueError, match="conflicts with durable evidence"):
        _write_aba_history_receipts(
            tmp_path,
            [{**historical, "calendar_final_count": 186, "reconciled_final_count": 184}],
            current_season_end_year=2027,
        )


def test_aba_legacy_receipt_cannot_hide_an_administrative_count_conflict(tmp_path):
    legacy = {
        "season_end_year": 2026,
        "calendar_final_count": 185,
        "reconciled_final_count": 185,
        "not_applicable_rows": 0,
        "blocking_gap_count": 0,
        "status": "calendar_final_count_reconciled",
        "source": "aba_liga",
        "evidence": "ROUTE_CONTRACT_REFERENCED",
        "completed_at_utc": "2026-11-20T13:40:00Z",
    }
    (tmp_path / "history_receipts.json").write_text(
        json.dumps(
            {
                "schema_version": "magolide.aba_history_receipts",
                "league": "ABA_LEAGUE",
                "provider": "aba_liga",
                "receipts": {"2026": legacy},
            }
        )
    )
    incoming = {
        **legacy,
        "reconciled_final_count": 183,
        "administrative_disposition_count": 2,
        "status": "calendar_terminal_count_reconciled",
    }

    with pytest.raises(ValueError, match="conflicts with durable evidence"):
        _write_aba_history_receipts(
            tmp_path, [incoming], current_season_end_year=2027
        )


def test_daily_dry_run_creates_no_warehouse_and_builds_market_report(tmp_path):
    summary = run_daily(_args(tmp_path))
    assert summary["mode"] == "dry_run"
    assert summary["system_of_record"] == "duckdb_csv"
    assert summary["commands"]["pregame"]["fixtures_found"] == 1
    assert summary["commands"]["pregame"]["generated_utc"] == "2026-11-07T00:00:00Z"
    assert summary["commands"]["pregame"]["persistence"]["kind"] == "duckdb"
    assert not Path(summary["warehouse"]).exists()
    assert summary["ledgers"] == {}
    assert summary["markets"]["status"] == "decision_support_only"
    assert summary["markets"]["accas"]["results"] == []
    assert summary["markets"]["forecasts"][0]["prediction"]["frozen"] is False
    assert Path(summary["market_report_path"]).exists()
    assert Path(summary["report_path"]).exists()
    assert summary["operator"]["paper_only"] is True
    assert summary["operator"]["counts"]["current_picks"] == 1
    assert Path(summary["operator"]["picks_text"]).exists()
    assert Path(summary["operator"]["performance_text"]).exists()


def test_dry_run_prunes_daily_and_market_report_retention(tmp_path):
    import json

    reports = tmp_path / "localdata" / "reports"
    reports.mkdir(parents=True)
    (reports / "daily_old.json").write_text(json.dumps({"finished_utc": "2000-01-01T00:00:00Z"}))
    (reports / "markets_old.json").write_text(json.dumps({"generated_utc": "2000-01-01T00:00:00Z"}))
    malformed = reports / "malformed_old.json"
    malformed.write_text("not json")
    import os

    os.utime(malformed, (946684800, 946684800))  # 2000-01-01 UTC
    summary = run_daily(_args(tmp_path))
    assert summary["retention"]["report_files_pruned"] == 3
    assert not (reports / "daily_old.json").exists()
    assert not (reports / "markets_old.json").exists()
    assert not malformed.exists()


def test_fixture_persistence_and_notification_have_separate_guards(tmp_path, monkeypatch):
    monkeypatch.setenv("ALLOW_MAGOLIDE_WRITES", "true")
    monkeypatch.delenv("ALLOW_FIXTURE_WRITES", raising=False)
    with pytest.raises(RuntimeError, match="ALLOW_FIXTURE_WRITES"):
        run_daily(_args(tmp_path, write=True))

    monkeypatch.setenv("ALLOW_FIXTURE_WRITES", "true")
    with pytest.raises(RuntimeError, match="cannot trigger external"):
        run_daily(_args(tmp_path, write=True, notify=True, send_notifications=True))


def test_fixture_write_refuses_canonical_state_paths(tmp_path, monkeypatch):
    monkeypatch.setenv("ALLOW_MAGOLIDE_WRITES", "true")
    monkeypatch.setenv("ALLOW_FIXTURE_WRITES", "true")
    canonical = tmp_path / "live" / "warehouse.duckdb"
    monkeypatch.setenv("MAGOLIDE_WAREHOUSE", str(canonical))

    with pytest.raises(RuntimeError, match="isolated from canonical state"):
        run_daily(
            _args(
                tmp_path,
                write=True,
                warehouse=str(canonical),
                localdata=str(canonical.parent),
            )
        )
    assert not canonical.exists()
    assert not list(canonical.parent.glob("reports/*.json"))


def test_offline_cli_pins_its_forecast_cutoff(tmp_path, monkeypatch, capsys):
    import json
    import sys

    monkeypatch.setattr(
        sys,
        "argv",
        [
            "daily.py",
            "--auto-once",
            "--date",
            "20261107",
            "--offline-fixture",
            "--warehouse",
            str(tmp_path / "warehouse.duckdb"),
            "--localdata",
            str(tmp_path),
            "--run-id",
            "offline-cli",
        ],
    )
    assert daily_main() == 0
    summary = json.loads(capsys.readouterr().out)
    assert summary["target_date"] == "2026-11-07"
    assert summary["started_utc"] == "2026-11-07T18:00:00Z"
    assert summary["markets"]["forecast_count"] == 1
    assert not (tmp_path / "warehouse.duckdb").exists()


def test_live_cutoff_includes_books_acquired_after_process_start(tmp_path, monkeypatch):
    started = "2026-11-07T20:02:25Z"
    acquired = "2026-11-07T20:02:32Z"
    after_acquisition = "2026-11-07T20:02:40Z"
    monkeypatch.setattr("scripts.daily._utc_now", lambda: after_acquisition)
    cutoff = _market_cutoff(_args(tmp_path, offline_fixture=False, started_utc=started))
    assert cutoff == after_acquisition

    store = InMemoryStore()
    game = Game(
        game_id="401909938",
        league=League.NBA,
        season="2027",
        home_team_id="11",
        away_team_id="7",
        home_name="Indiana Pacers",
        away_name="Denver Nuggets",
        scheduled_utc="2026-11-07T22:00:00Z",
    )
    store.upsert_game(game)
    snapshot = BookSnapshot(
        game_id=game.game_id,
        provider="espn",
        book="DraftKings",
        market=Market.MONEYLINE,
        period=Period.FULL_GAME,
        observed_utc=acquired,
        home_price_decimal=3.61,
        away_price_decimal=1.30,
        pre_match=True,
    )
    store.upsert_book_snapshot(snapshot)
    service = MarketForecastService(store, League.NBA)
    assert service._books(game, as_of_utc=started, historical=False) == []
    assert service._books(game, as_of_utc=cutoff, historical=False) == [snapshot]


def test_daily_write_persists_duckdb_csv_and_idempotent_ledgers(tmp_path, monkeypatch):
    monkeypatch.setenv("ALLOW_MAGOLIDE_WRITES", "true")
    monkeypatch.setenv("ALLOW_FIXTURE_WRITES", "true")
    first = run_daily(_args(tmp_path, write=True))
    assert first["counts"]["games"] == 1
    assert first["counts"]["predictions"] == 1
    assert first["counts"]["pipeline_runs"] == 3
    assert set(first["ledgers"]) == {"pregame", "freeze", "results"}
    assert all(item["created"] for item in first["ledgers"].values())
    assert len(first["csv_exports"]) == 6
    assert Path(first["warehouse"]).exists()

    ledger_bytes = {
        kind: Path(item["path"]).read_bytes() for kind, item in first["ledgers"].items()
    }
    # A later wall clock must not rewrite deterministic fixture evidence.
    from ingestion.providers import espn

    monkeypatch.setattr(espn, "now_utc_iso", lambda: "2099-01-01T00:00:00Z")
    second = run_daily(_args(tmp_path, write=True))
    assert second["counts"] == first["counts"]
    for kind, item in second["ledgers"].items():
        assert item["created"] is False, (kind, item, first["ledgers"][kind])
        assert Path(item["path"]).read_bytes() == ledger_bytes[kind]

    from ingestion.warehouse import connect

    con = connect(first["warehouse"])
    try:
        assert con.execute("select count(*) from games").fetchone()[0] == 1
        assert con.execute("select count(*) from prediction_facts").fetchone()[0] == 1
    finally:
        con.close()


def test_the_daily_tick_records_what_the_tuner_did(tmp_path):
    """Tuning is part of the tick, not a command somebody has to remember."""

    summary = run_daily(_args(tmp_path))

    tuning = summary["tuning"]
    assert tuning["enabled"] is True
    # A dry tick reports rather than writes; either way the stamp is reported so a
    # card rendered later can be matched to the config this tick ran under.
    assert tuning["status"] in {"dry_run", "no_change"}
    assert tuning["config_stamp"].startswith("CFG_")
    assert isinstance(tuning["settled_picks"], int)


def test_the_daily_tick_can_skip_tuning(tmp_path):
    summary = run_daily(_args(tmp_path, skip_tuning=True))

    assert summary["tuning"] == {"enabled": False, "status": "skipped"}


EXPECTED_STAGE_KEYS = {
    "store_open",
    "backfill",
    "acquisition",
    "predictions",
    "odds",
    "pregame_rows",
    "pregame_persist",
    "forecast",
    "freeze",
    "settle",
    "retention",
    "publish",
    "tuning",
    "exports",
    "notification",
    "summary_write",
    "store_close",
}


def test_stage_seconds_covers_the_full_tick_and_reconciles_with_wall_time(tmp_path):
    """Every stage of run_daily is timed once; persisted receipt matches wall time."""

    import time

    started = time.perf_counter()
    summary = run_daily(_args(tmp_path))
    wall_seconds = time.perf_counter() - started

    pregame_stages = summary["commands"]["pregame"]["stage_seconds"]
    assert set(pregame_stages) == {"acquisition", "predictions", "odds", "pregame_rows"}
    assert all(isinstance(v, float) and v >= 0.0 for v in pregame_stages.values())

    stages = summary["stage_seconds"]
    assert set(stages) == EXPECTED_STAGE_KEYS
    assert "pregame" not in stages  # must not double-count pregame sub-stages
    assert all(isinstance(v, float) and v >= 0.0 for v in stages.values())

    persisted = json.loads(Path(summary["report_path"]).read_text())
    assert persisted["stage_seconds"] == stages

    total = sum(stages.values())
    assert abs(total - wall_seconds) <= 5.0
    # Guard against double-counting: sum of disjoint stages cannot exceed wall
    # time by more than rounding granularity across 17 keys.
    assert total <= wall_seconds + 0.2

    # Skipping optional stages preserves the schema with explicit 0.0 entries.
    skipped = run_daily(
        _args(tmp_path / "skip", skip_backfill=True, skip_results=True, skip_markets=True)
    )
    assert set(skipped["stage_seconds"]) == EXPECTED_STAGE_KEYS
    assert skipped["stage_seconds"]["backfill"] == 0.0
    assert skipped["stage_seconds"]["settle"] == 0.0
    assert skipped["stage_seconds"]["forecast"] == 0.0


def test_stage_seconds_finally_handles_mid_tick_failure_cleanly(tmp_path, monkeypatch):
    """An exception before summary creation must propagate unmodified after store.close()."""

    from ingestion.pipeline import Pipeline

    def _boom(self):
        raise RuntimeError("injected_pregame_failure")

    monkeypatch.setattr(Pipeline, "pregame", _boom)
    with pytest.raises(RuntimeError, match="injected_pregame_failure"):
        run_daily(_args(tmp_path))


def test_an_offline_fixture_tick_touches_no_network(tmp_path):
    """Offline fixture replay must never make live HTTP requests or block on network."""

    from ingestion.pipeline import PipelineConfig

    # Default stays False so non-offline callers and injected-fetch tests are untouched.
    assert PipelineConfig(league="NBA", date_yyyymmdd="20261107").offline_fixture is False

    summary = run_daily(_args(tmp_path))
    pregame = summary["commands"]["pregame"]
    acquisition = pregame["acquisition"]

    assert acquisition["transport_attempts"] == 0
    assert acquisition["transport_errors"] == 0
    assert summary["stage_seconds"]["odds"] < 5.0

    book_status = pregame["book_provider_status"]
    for provider in ("theoddsapi", "betexplorer_direct"):
        assert provider in book_status
        assert book_status[provider]["enabled"] is False
        assert book_status[provider]["status"] == "disabled"
        assert "offline fixture replay" in book_status[provider]["detail"]


