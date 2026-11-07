"""Sheets data-store round-trip and CLI guard tests."""

import os

import pytest

from ingestion.models import (
    BookSnapshot,
    CoverageState,
    Game,
    GameStatus,
    League,
    Market,
    Period,
    Prediction,
    ProviderHealth,
)
from ingestion.sheets_store import (
    BOOKS_TAB,
    GAMES_TAB,
    PREDICTIONS_TAB,
    SheetsDataStore,
)
from ingestion.sheets_writer import FakeWorksheet
from ingestion.store import FrozenWriteError

GAME = Game(
    game_id="401909938", league=League.NBA, season="2026",
    home_team_id="11", away_team_id="7",
    home_name="Indiana Pacers", away_name="Denver Nuggets",
    scheduled_utc="2026-11-07T22:00:00Z", status=GameStatus.SCHEDULED,
)


def _store():
    return SheetsDataStore({
        GAMES_TAB: FakeWorksheet(["game_id", "league", "payload_json"]),
        PREDICTIONS_TAB: FakeWorksheet(["game_id", "provider", "frozen", "payload_json"]),
        BOOKS_TAB: FakeWorksheet(["snapshot_key", "game_id", "provider", "frozen", "payload_json"]),
    })


def test_store_roundtrip_game_prediction_book():
    store = _store()
    store.upsert_game(GAME)
    pred = Prediction(
        game_id="401909938", provider="espn", observed_utc="2026-11-07T18:00:00Z",
        prob_home=0.23062, prob_away=0.76938,
        expected_home_score=110.54, expected_away_score=120.046,
        expected_total=230.586, expected_margin=9.506,
    )
    store.upsert_prediction(pred)
    snap = BookSnapshot(
        game_id="401909938", provider="sportsgameodds", book="draftkings",
        market=Market.TOTAL, period=Period.Q1, observed_utc="2026-11-07T18:00:00Z",
        line=58.5, pre_match=True, main_line=True,
    )
    store.upsert_book_snapshot(snap)

    fresh = SheetsDataStore({
        GAMES_TAB: store._ws[GAMES_TAB],
        PREDICTIONS_TAB: store._ws[PREDICTIONS_TAB],
        BOOKS_TAB: store._ws[BOOKS_TAB],
    })
    got_game = fresh.get_game("401909938")
    assert got_game is not None and got_game.home_name == "Indiana Pacers"
    preds = fresh.get_predictions("401909938")
    assert len(preds) == 1 and preds[0].expected_total == 230.586
    snaps = fresh.get_book_snapshots("401909938")
    assert len(snaps) == 1 and snaps[0].line == 58.5


def test_store_frozen_enforcement_across_reload():
    store = _store()
    pred = Prediction(
        game_id="401909938", provider="espn", observed_utc="2026-11-07T18:00:00Z",
        prob_home=0.23, prob_away=0.77, expected_total=230.5,
    )
    store.upsert_prediction(pred, frozen=True)

    # Same value re-upsert is idempotent; a changed value is rejected.
    same = Prediction(**{**pred.__dict__})
    store.upsert_prediction(same, frozen=False)
    changed = Prediction(**{**pred.__dict__, "expected_total": 999.0})
    with pytest.raises(FrozenWriteError):
        store.upsert_prediction(changed, frozen=False)

    # Enforcement survives a reload from the sheets-backed tables.
    reloaded = SheetsDataStore({
        GAMES_TAB: store._ws[GAMES_TAB],
        PREDICTIONS_TAB: store._ws[PREDICTIONS_TAB],
        BOOKS_TAB: store._ws[BOOKS_TAB],
    })
    with pytest.raises(FrozenWriteError):
        reloaded.upsert_prediction(changed, frozen=False)


def test_store_pre_tip_slot_protected_from_post_tip_displacement():
    store = _store()
    pre = Prediction(
        game_id="401909938", provider="espn", observed_utc="2026-11-07T18:00:00Z",
        prob_home=0.23, prob_away=0.77, expected_total=230.5,
    )
    post = Prediction(
        game_id="401909938", provider="espn", observed_utc="2026-11-07T23:00:00Z",
        prob_home=0.5, prob_away=0.5, expected_total=200.0,
    )
    post.pre_tip = False
    store.upsert_prediction(pre)
    store.upsert_prediction(post)  # post-tip lands in its own slot
    preds = {p.observed_utc: p for p in store.get_predictions("401909938")}
    assert preds["2026-11-07T18:00:00Z"].expected_total == 230.5  # preserved


def test_store_rewrites_trim_stale_rows():
    """A shrinking table must not leave stale rows behind after rewrite."""
    store = _store()
    store.upsert_team_alias("Real Madrid", ["RMB"])
    store.upsert_team_alias("FC Barcelona", ["BAR"])
    store.upsert_team_alias("Zalgiris Kaunas", ["ZAL"])
    assert len(store.get_team_aliases()) == 3
    store.upsert_team_alias("Zalgiris Kaunas", ["ZAL"])  # unchanged content
    store._aliases.pop("Zalgiris Kaunas")
    store._persist_aliases()  # 2 aliases now; old third row must be blanked
    values = store._ws["TeamAliases"].get_all_values()
    body = [r for r in values[1:] if any(str(v or "").strip() for v in r)]
    assert len(body) == 2
    fresh = SheetsDataStore({**store._ws})
    assert len(fresh.get_team_aliases()) == 2


def test_provider_health_roundtrip():
    store = _store()
    health = ProviderHealth(
        provider="espn", league=League.NBA,
        last_attempt_utc="2026-11-07T20:00:00Z",
        last_success_utc="2026-11-07T20:00:01Z",
        coverage={"bpi_complete": {"complete": 7, "total": 7}},
        state=CoverageState.OK,
    )
    store.upsert_provider_health(health)
    got = store.get_provider_health("espn", League.NBA)
    assert got is not None
    assert got.coverage["bpi_complete"]["complete"] == 7
    assert got.state is CoverageState.OK


# ---------------------------------------------------------------------------
# CLI guards (offline)
# ---------------------------------------------------------------------------


def test_cli_write_requires_env_flag(monkeypatch, tmp_path):
    from ingestion.cli import main

    monkeypatch.setenv("ALLOW_MAGOLIDE_WRITES", "true")
    monkeypatch.delenv("ALLOW_SHEET_WRITES", raising=False)
    with pytest.raises(SystemExit) as e:
        main(["pregame", "--league", "NBA", "--date", "20261107", "--write",
              "--sheet-id", "TEST", "--store-json", str(tmp_path / "store.json")])
    assert "ALLOW_SHEET_WRITES" in str(e.value)


def test_cli_write_uses_duckdb_without_sheet_id(monkeypatch, tmp_path, capsys):
    import json
    from ingestion.cli import main

    monkeypatch.setenv("ALLOW_MAGOLIDE_WRITES", "true")
    monkeypatch.setenv("ALLOW_FIXTURE_WRITES", "true")
    warehouse = tmp_path / "warehouse.duckdb"
    assert main([
        "pregame", "--league", "NBA", "--date", "20261107", "--write",
        "--offline-fixture", "--warehouse", str(warehouse),
    ]) == 0
    report = json.loads(capsys.readouterr().out)
    assert report["persistence"]["kind"] == "duckdb"
    assert report["persistence"]["write_enabled"] is True
    assert report["mode"] == "write"
    assert warehouse.exists()


def test_cli_fixture_write_refuses_configured_canonical(monkeypatch, tmp_path):
    from ingestion.cli import main

    monkeypatch.setenv("ALLOW_MAGOLIDE_WRITES", "true")
    monkeypatch.setenv("ALLOW_FIXTURE_WRITES", "true")
    canonical = tmp_path / "live" / "warehouse.duckdb"
    monkeypatch.setenv("MAGOLIDE_WAREHOUSE", str(canonical))
    with pytest.raises(SystemExit, match="isolated from canonical state"):
        main([
            "pregame", "--league", "NBA", "--date", "20261107", "--write",
            "--offline-fixture", "--warehouse", str(canonical),
        ])
    assert not canonical.exists()


def test_cli_dry_run_is_default(monkeypatch, tmp_path):
    """Without --write nothing authenticates or writes; offline fixtures only."""
    from ingestion.cli import main

    monkeypatch.delenv("ALLOW_SHEET_WRITES", raising=False)
    rc = main(["pregame", "--league", "NBA", "--date", "20261107",
               "--store-json", str(tmp_path / "store.json"),
               "--offline-fixture"])
    assert rc == 0
