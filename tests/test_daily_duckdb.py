from __future__ import annotations

import argparse
from pathlib import Path

import pytest

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
    )
    values.update(changes)
    return argparse.Namespace(**values)


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
    (reports / "daily_old.json").write_text(
        json.dumps({"finished_utc": "2000-01-01T00:00:00Z"})
    )
    (reports / "markets_old.json").write_text(
        json.dumps({"generated_utc": "2000-01-01T00:00:00Z"})
    )
    summary = run_daily(_args(tmp_path))
    assert summary["retention"]["report_files_pruned"] == 2
    assert not (reports / "daily_old.json").exists()
    assert not (reports / "markets_old.json").exists()


def test_fixture_persistence_and_notification_have_separate_guards(
    tmp_path, monkeypatch
):
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
