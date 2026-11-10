from __future__ import annotations

import csv
from dataclasses import replace
import gzip
import json

import pytest

from ingestion.duckdb_store import DuckDBStore
from ingestion.ledger import (
    LedgerConflictError,
    canonical_snapshot,
    write_immutable_json,
)
from ingestion.models import (
    ActualResult,
    BookSnapshot,
    CoverageState,
    Game,
    GameStatus,
    League,
    Market,
    Period,
    PeriodKind,
    Prediction,
    ProviderHealth,
)
from ingestion.store import FrozenWriteError
from ingestion.warehouse import connect


def _game() -> Game:
    return Game(
        game_id="espn-1",
        league=League.NBA,
        season="2027",
        home_team_id="11",
        away_team_id="7",
        home_name="Indiana Pacers",
        away_name="Denver Nuggets",
        scheduled_utc="2026-11-07T22:00:00Z",
        provider_ids={"espn": "401909938"},
        actual=ActualResult(
            period_kind=PeriodKind.QUARTERS,
            home_quarters=[25, 25, 25, 25],
            away_quarters=[20, 20, 20, 20],
            home_final=100,
            away_final=80,
            source="espn",
        ),
        status=GameStatus.FINAL,
    )


def _prediction(**changes) -> Prediction:
    values = dict(
        game_id="espn-1",
        provider="espn",
        observed_utc="2026-11-07T18:00:00Z",
        model="nba_predictor",
        model_version="1",
        prob_home=0.6,
        prob_away=0.4,
        expected_home_score=110.0,
        expected_away_score=105.0,
        expected_total=215.0,
        expected_margin=5.0,
        pre_tip=True,
    )
    values.update(changes)
    return Prediction(**values)


def _snapshot(**changes) -> BookSnapshot:
    values = dict(
        game_id="espn-1",
        provider="theoddsapi",
        book="draftkings",
        market=Market.TOTAL,
        period=Period.FULL_GAME,
        observed_utc="2026-11-07T18:00:00Z",
        line=220.5,
        over_price_decimal=1.91,
        under_price_decimal=1.91,
        pre_match=True,
        main_line=True,
    )
    values.update(changes)
    return BookSnapshot(**values)


def test_duckdb_roundtrip_preserves_all_canonical_records(tmp_path):
    path = tmp_path / "warehouse.duckdb"
    store = DuckDBStore(str(path))
    store.upsert_game(_game())
    store.upsert_prediction(_prediction(), frozen=True)
    store.upsert_prediction(_prediction(pre_tip=False, observed_utc="2026-11-07T23:00:00Z"))
    store.upsert_book_snapshot(_snapshot(), frozen=True)
    store.upsert_book_snapshot(
        _snapshot(pre_match=False, observed_utc="2026-11-07T23:00:00Z", line=221.5)
    )
    store.upsert_team_alias("Denver Nuggets", ["Denver", "Nuggets"])
    store.upsert_provider_health(
        ProviderHealth(
            provider="espn",
            league=League.NBA,
            last_attempt_utc="2026-11-07T18:00:00Z",
            last_success_utc="2026-11-07T18:00:00Z",
            coverage={"complete": 1, "total": 1},
            state=CoverageState.OK,
        )
    )
    store.record_run("run-1", "pregame", "2026-11-07", {"generated_utc": "2026-11-07T18:00:00Z"})
    store.record_run("run-1", "freeze", "2026-11-07", {"generated_utc": "2026-11-07T18:01:00Z"})
    store.close()

    loaded = DuckDBStore(str(path), persist=False)
    try:
        assert loaded.get_game("espn-1").actual.home_final == 100
        assert len(loaded.get_predictions("espn-1")) == 2
        assert len(loaded.get_book_snapshots("espn-1")) == 2
        assert loaded.is_prediction_frozen("espn-1", "espn")
        assert loaded.is_book_frozen("espn-1", "theoddsapi")
        assert loaded.get_team_aliases()["Denver Nuggets"] == ["Denver", "Nuggets"]
        assert loaded.get_provider_health("espn", League.NBA).coverage["complete"] == 1
        assert loaded.counts()["pipeline_runs"] == 2
        post = [row for row in loaded.table_rows("predictions") if not row["pre_tip"]]
        assert len(post) == 1 and not post[0]["frozen"]

        with pytest.raises(FrozenWriteError):
            loaded.upsert_prediction(_prediction(prob_home=0.7), frozen=True)
        # Idempotent replay of the exact frozen value is allowed.
        loaded.upsert_prediction(_prediction(), frozen=True)
    finally:
        loaded.close()


def test_first_half_draw_price_roundtrip_and_schema_migration(tmp_path):
    import duckdb

    path = tmp_path / "warehouse.duckdb"
    store = DuckDBStore(str(path))
    snapshot = _snapshot(
        market=Market.MONEYLINE,
        period=Period.H1,
        line=None,
        over_price_decimal=None,
        under_price_decimal=None,
        home_price_decimal=1.85,
        away_price_decimal=2.1,
        draw_price_decimal=11.0,
    )
    store.upsert_book_snapshot(snapshot, frozen=True)
    store.close()

    # Simulate a version-1 warehouse: reopening must add and backfill the new
    # typed draw column without touching the complete JSON source payload.
    con = duckdb.connect(str(path))
    con.execute("DROP INDEX idx_books_game_market_period")
    con.execute("ALTER TABLE book_snapshots DROP COLUMN draw_price_decimal")
    con.close()
    migrated = DuckDBStore(str(path))
    try:
        loaded = migrated.get_book_snapshots("espn-1")[0]
        assert loaded.draw_price_decimal == 11.0
        row = migrated.table_rows("book_snapshots")[0]
        assert row["draw_price_decimal"] == 11.0
        assert (
            migrated._con.execute(
                "select value from schema_meta where key='schema_version'"
            ).fetchone()[0]
            == "2"
        )
    finally:
        migrated.close()


def test_operational_run_retention_does_not_touch_canonical_facts(tmp_path):
    store = DuckDBStore(str(tmp_path / "warehouse.duckdb"))
    store.upsert_game(_game())
    store.record_run("old", "pregame", "2020-01-01", {"generated_utc": "2020-01-01T00:00:00Z"})
    store.record_run("new", "pregame", "2026-11-07", {"generated_utc": "2026-11-07T00:00:00Z"})
    assert store.prune_pipeline_runs("2021-01-01T00:00:00Z") == 1
    assert store.counts()["pipeline_runs"] == 1
    assert store.get_game("espn-1") is not None
    assert store.table_rows("pipeline_runs")[0]["run_id"] == "new"
    store.close()


def test_dry_run_never_creates_or_changes_warehouse(tmp_path):
    path = tmp_path / "warehouse.duckdb"
    dry = DuckDBStore(str(path), persist=False)
    dry.upsert_game(_game())
    assert dry.get_game("espn-1") is not None
    dry.close()
    assert not path.exists()

    writable = DuckDBStore(str(path))
    writable.upsert_game(_game())
    writable.close()
    before = path.read_bytes()

    readonly = DuckDBStore(str(path), persist=False)
    readonly.upsert_team_alias("New Alias", ["Not persisted"])
    readonly.close()
    assert path.read_bytes() == before
    check = DuckDBStore(str(path), persist=False)
    assert "New Alias" not in check.get_team_aliases()
    check.close()


def test_csv_export_and_sql_views_expose_observed_typed_facts(tmp_path):
    path = tmp_path / "warehouse.duckdb"
    store = DuckDBStore(str(path))
    store.upsert_game(_game())
    store.upsert_prediction(_prediction(), frozen=True)
    store.upsert_book_snapshot(_snapshot(), frozen=True)
    exported = store.export_csv(str(tmp_path / "csv"))
    store.close()

    with gzip.open(exported["book_snapshots"], "rt", newline="") as handle:
        rows = list(csv.DictReader(handle))
    assert rows[0]["line"] == "220.5"
    assert json.loads(rows[0]["payload_json"])["over_price_decimal"] == 1.91

    con = connect(str(path))
    try:
        prediction = con.execute(
            "select prob_home, expected_total, frozen from prediction_facts"
        ).fetchone()
        book = con.execute("select line, period, frozen from book_facts").fetchone()
        result = con.execute("select home_final, result_source from game_results").fetchone()
        assert prediction == (0.6, 215.0, True)
        assert book == (220.5, Period.FULL_GAME.value, True)
        assert result == ("100", "espn")
    finally:
        con.close()

    recovered = DuckDBStore(str(tmp_path / "recovered.duckdb"))
    counts = recovered.restore_csv(str(tmp_path / "csv"))
    assert counts["games"] == 1
    assert counts["predictions"] == 1
    assert counts["book_snapshots"] == 1
    assert recovered.is_prediction_frozen("espn-1", "espn")
    assert recovered.is_book_frozen("espn-1", "theoddsapi")
    assert recovered.get_game("espn-1").actual.home_final == 100
    with pytest.raises(RuntimeError, match="empty warehouse"):
        recovered.restore_csv(str(tmp_path / "csv"))
    recovered.close()


def test_ledger_snapshot_is_scoped_to_league_local_slate(tmp_path):
    store = DuckDBStore(str(tmp_path / "warehouse.duckdb"))
    store.upsert_game(_game())
    # 00:30 UTC on November 8 is still November 7 in New York.
    store.upsert_game(replace(_game(), game_id="espn-2", scheduled_utc="2026-11-08T00:30:00Z"))
    store.upsert_game(replace(_game(), game_id="espn-3", scheduled_utc="2026-11-08T22:00:00Z"))
    snapshot = canonical_snapshot(store, League.NBA, "2026-11-07")
    assert {game["game_id"] for game in snapshot["games"]} == {"espn-1", "espn-2"}
    next_slate = canonical_snapshot(store, League.NBA, "20261108")
    assert {game["game_id"] for game in next_slate["games"]} == {"espn-3"}
    store.close()


def test_immutable_ledger_noops_on_identical_bytes_and_rejects_conflict(tmp_path):
    path = tmp_path / "ledger.json"
    payload = {"games": [{"game_id": "g1"}], "frozen": True}
    assert write_immutable_json(str(path), payload) == (str(path), True)
    assert write_immutable_json(str(path), payload) == (str(path), False)
    with pytest.raises(LedgerConflictError):
        write_immutable_json(str(path), {"games": []})
    assert json.loads(path.read_text()) == payload
