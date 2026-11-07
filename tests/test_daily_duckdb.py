from __future__ import annotations

import argparse
from pathlib import Path

import pytest

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
        quarter_csv=None,
        notify=False,
        send_notifications=False,
        started_utc="2026-11-07T18:00:00Z",
    )
    values.update(changes)
    return argparse.Namespace(**values)


def test_daily_dry_run_uses_no_google_and_creates_no_warehouse(tmp_path, monkeypatch):
    import ingestion.cli as cli

    monkeypatch.setattr(
        cli,
        "open_persistent_store",
        lambda *a, **k: (_ for _ in ()).throw(
            AssertionError("legacy Google store must not be selected")
        ),
    )
    summary = run_daily(_args(tmp_path))
    assert summary["mode"] == "dry_run"
    assert summary["system_of_record"] == "duckdb_csv"
    assert summary["commands"]["pregame"]["fixtures_found"] == 1
    assert summary["commands"]["pregame"]["persistence"]["kind"] == "duckdb"
    assert not Path(summary["warehouse"]).exists()
    assert summary["ledgers"] == {}
    assert Path(summary["report_path"]).exists()


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
