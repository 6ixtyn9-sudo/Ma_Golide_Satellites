"""Regression tests for production issues found during artifact review."""

from __future__ import annotations

import datetime as dt
import json
from types import SimpleNamespace

import pytest

from ingestion.identity import parse_legacy_display_date
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
from ingestion.pipeline import Pipeline, PipelineConfig
from ingestion.sheets_store import GAMES_HEADERS, GAMES_TAB, SheetsDataStore
from ingestion.sheets_writer import FakeWorksheet, GspreadWorksheetAdapter
from ingestion.store import FrozenWriteError, InMemoryStore, JsonFileStore


def _game(game_id="espn-1"):
    return Game(
        game_id=game_id,
        league=League.NBA,
        season="2027",
        home_team_id="11",
        away_team_id="7",
        home_name="Indiana Pacers",
        away_name="Denver Nuggets",
        scheduled_utc="2026-11-07T22:00:00Z",
        status=GameStatus.SCHEDULED,
        provider_ids={"espn": game_id},
    )


def _snapshot(provider="one", event="event-1", period=Period.FULL_GAME, **kw):
    values = dict(
        game_id=event,
        provider=provider,
        provider_event_id=event,
        book="draftkings",
        market=Market.TOTAL,
        period=period,
        observed_utc="2026-11-07T18:00:00Z",
        line=230.5 if period is Period.FULL_GAME else 57.5,
        pre_match=True,
        main_line=True,
        home_name="Indiana Pacers",
        away_name="Denver Nuggets",
        scheduled_utc="2026-11-07T22:00:00Z",
    )
    values.update(kw)
    return BookSnapshot(**values)


def _pipeline(store=None):
    return Pipeline(
        PipelineConfig(league="NBA", date_yyyymmdd="20261107"),
        store or InMemoryStore(),
        fetch=lambda url: (_ for _ in ()).throw(AssertionError("no network")),
    )


def test_ambiguous_legacy_date_is_never_locale_guessed():
    assert parse_legacy_display_date("11/01/2026") is None
    assert parse_legacy_display_date("2026-11-01").date() == dt.date(2026, 11, 1)
    assert parse_legacy_display_date("20/01/2026").date() == dt.date(2026, 1, 20)
    assert parse_legacy_display_date("11/20/2026").date() == dt.date(2026, 11, 20)


def test_book_slate_date_uses_league_timezone_not_utc_only():
    from ingestion.providers.theoddsapi import _on_utc_date

    commence = "2026-11-08T00:30:00Z"  # 7:30pm ET on November 7
    assert not _on_utc_date(commence, "20261107")
    assert _on_utc_date(commence, "20261107", "America/New_York")


def test_book_lines_from_multiple_provider_ids_are_combined(monkeypatch):
    import ingestion.pipeline as module

    game = _game()
    store = InMemoryStore()
    pipeline = _pipeline(store)
    snaps = {
        "one": [_snapshot("one", "event-one")],
        "two": [_snapshot("two", "event-two", Period.Q1)],
    }
    modules = {
        provider: SimpleNamespace(
            is_enabled=lambda: True,
            fetch_odds=lambda *args, p=provider: snaps[p],
        )
        for provider in snaps
    }
    monkeypatch.setattr(module, "BOOK_PROVIDERS", tuple(modules))
    monkeypatch.setattr(module, "_book_module", modules.__getitem__)

    mapped = pipeline._acquire_book_lines([game])

    assert {s.provider for s in mapped[game.game_id]} == {"one", "two"}
    assert {s.provider for s in store.get_book_snapshots(game.game_id)} == {"one", "two"}
    assert game.provider_ids["one"] == "event-one"
    assert game.provider_ids["two"] == "event-two"


def test_contradictory_known_provider_id_is_quarantined(monkeypatch):
    import ingestion.pipeline as module

    game = _game()
    game.provider_ids["one"] = "event-one"
    bad = _snapshot(event="event-one", home_name="Boston Celtics")
    mod = SimpleNamespace(is_enabled=lambda: True, fetch_odds=lambda *args: [bad])
    monkeypatch.setattr(module, "BOOK_PROVIDERS", ("one",))
    monkeypatch.setattr(module, "_book_module", lambda provider: mod)
    store = InMemoryStore()
    pipeline = _pipeline(store)

    assert pipeline._acquire_book_lines([game]) == {}
    assert store.get_book_snapshots(game.game_id) == []
    assert pipeline.gaps[-1].required_field == "book_identity"
    assert "contradicts" in pipeline.gaps[-1].detail


def test_unmapped_provider_id_without_identity_is_not_trusted(monkeypatch):
    import ingestion.pipeline as module

    game = _game("coincident-id")
    raw = _snapshot(event="coincident-id", home_name=None, away_name=None, scheduled_utc=None)
    mod = SimpleNamespace(is_enabled=lambda: True, fetch_odds=lambda *args: [raw])
    monkeypatch.setattr(module, "BOOK_PROVIDERS", ("one",))
    monkeypatch.setattr(module, "_book_module", lambda provider: mod)
    store = InMemoryStore()

    assert _pipeline(store)._acquire_book_lines([game]) == {}
    assert store.get_book_snapshots(game.game_id) == []


def test_pre_and_post_book_slots_survive_and_every_pre_market_freezes():
    store = InMemoryStore()
    game = _game()
    pre_total = _snapshot(game_id=game.game_id)
    pre_quarter = _snapshot(game_id=game.game_id, period=Period.Q1)
    post_total = _snapshot(
        game_id=game.game_id, observed_utc="2026-11-07T23:00:00Z",
        pre_match=False, line=999.0
    )
    for snap in (pre_total, pre_quarter, post_total):
        store.upsert_book_snapshot(snap)

    pipeline = _pipeline(store)
    assert pipeline._freeze_books(game) == 2
    assert store.is_book_snapshot_frozen(game.game_id, pre_total)
    assert store.is_book_snapshot_frozen(game.game_id, pre_quarter)
    assert not store.is_book_snapshot_frozen(game.game_id, post_total)
    assert len(store.get_book_snapshots(game.game_id)) == 3

    changed_pre = _snapshot(game_id=game.game_id, line=231.0)
    with pytest.raises(FrozenWriteError):
        store.upsert_book_snapshot(changed_pre)
    # The post slot remains independently updateable and cannot erase pre-tip.
    store.upsert_book_snapshot(_snapshot(
        game_id=game.game_id, observed_utc="2026-11-07T23:05:00Z",
        pre_match=False, line=888.0
    ))
    pre = [s for s in store.get_book_snapshots(game.game_id)
           if s.period is Period.FULL_GAME and s.pre_match]
    assert len(pre) == 1 and pre[0].line == 230.5


def test_json_store_book_slots_and_health_roundtrip(tmp_path):
    path = tmp_path / "state.json"
    store = JsonFileStore(str(path))
    store.upsert_book_snapshot(_snapshot(game_id=_game().game_id))
    store.upsert_book_snapshot(_snapshot(
        game_id=_game().game_id, pre_match=False,
        observed_utc="2026-11-07T23:00:00Z",
    ))
    store.upsert_provider_health(ProviderHealth(
        provider="espn", league=League.NBA, state=CoverageState.OK,
        coverage={"complete": 1, "total": 1},
    ))

    fresh = JsonFileStore(str(path))
    assert len(fresh.get_book_snapshots(_game().game_id)) == 2
    health = fresh.get_provider_health("espn", League.NBA)
    assert health is not None
    assert health.league is League.NBA and health.state is CoverageState.OK


def test_bound_http_client_quota_reaches_adapter():
    from ingestion.providers import theoddsapi

    class Client:
        last_quota = {"x-requests-remaining": "41"}

        def get_json(self, url):
            return []

    theoddsapi._reset_gaps()
    payload, headers = theoddsapi._call(Client().get_json, "https://example.invalid")
    assert payload == []
    assert headers == {"x-requests-remaining": "41"}
    assert theoddsapi.last_quota() == {"x-requests-remaining": 41}


def test_sheets_store_dry_run_buffers_without_rewriting():
    ws = FakeWorksheet(GAMES_HEADERS)
    store = SheetsDataStore({GAMES_TAB: ws}, persist=False)
    store.upsert_game(_game())
    assert store.get_game(_game().game_id) is not None
    assert ws.rows() == [] and ws.update_calls == []


def test_sheets_store_rate_limits_real_rewrites_with_injected_sleep():
    calls = []
    store = SheetsDataStore(
        {GAMES_TAB: FakeWorksheet(GAMES_HEADERS)}, sleep=calls.append
    )
    store.upsert_game(_game())
    assert calls == [1.1]


def test_sheets_store_pre_post_slots_survive_reload():
    ws = FakeWorksheet(["snapshot_key", "game_id", "provider", "frozen", "payload_json"])
    store = SheetsDataStore({"BookSnapshots": ws})
    store.upsert_book_snapshot(_snapshot(game_id=_game().game_id))
    store.upsert_book_snapshot(_snapshot(
        game_id=_game().game_id, pre_match=False,
        observed_utc="2026-11-07T23:00:00Z",
    ))
    fresh = SheetsDataStore({"BookSnapshots": ws})
    assert len(fresh.get_book_snapshots(_game().game_id)) == 2


def test_gspread_adapter_uses_formula_reads_and_json_safe_values():
    class Spreadsheet:
        def __init__(self):
            self.body = None

        def values_batch_update(self, body):
            json.dumps(body)  # must not raise on date/time objects
            self.body = body

    class Worksheet:
        def __init__(self):
            self.spreadsheet = Spreadsheet()
            self.read_options = None
            self.appended = None

        def get_all_values(self, **kwargs):
            self.read_options = kwargs
            return []

        def append_rows(self, rows, value_input_option):
            json.dumps(rows)
            self.appended = (rows, value_input_option)

    raw = Worksheet()
    ws = GspreadWorksheetAdapter(raw)
    ws.get_all_values()
    assert raw.read_options["value_render_option"].value == "FORMULA"
    assert raw.read_options["date_time_render_option"].value == "FORMATTED_STRING"

    values = [[dt.date(2026, 11, 7), dt.time(17, 0)]]
    ws.batch_update_ranges([("E2:F2", values)])
    assert raw.spreadsheet.body["data"][0]["values"] == [["2026-11-07", "17:00:00"]]
    ws.append_rows(values)
    assert raw.appended == ([["2026-11-07", "17:00:00"]], "USER_ENTERED")


def test_cli_uses_only_explicit_legacy_store_for_coverage(monkeypatch, capsys):
    from ingestion import cli

    calls = []

    def open_store(sheet_id, write=False):
        calls.append((sheet_id, write))
        return InMemoryStore()

    # The old environment variable must not silently restore Sheets as the
    # system of record. Legacy access requires the explicit CLI boundary.
    monkeypatch.setenv("DATA_STORE_SPREADSHEET_ID", "store-from-env")
    monkeypatch.setattr(cli, "open_persistent_store", open_store)
    assert cli.main([
        "coverage", "--league", "NBA", "--data-store-id", "store-from-flag"
    ]) == 0
    assert calls == [("store-from-flag", False)]
    assert json.loads(capsys.readouterr().out)["fixtures_found"] == 0


def test_offline_cli_dry_run_does_not_create_json_store(tmp_path, capsys):
    from ingestion import cli

    path = tmp_path / "dry-run-state.json"
    rc = cli.main([
        "pregame", "--league", "NBA", "--date", "20261107",
        "--offline-fixture", "--store-json", str(path),
    ])
    assert rc == 0 and not path.exists()
    report = json.loads(capsys.readouterr().out)
    assert report["persistence"]["kind"] == "json_file_local"
    assert report["persistence"]["write_enabled"] is False
    assert not any(
        "invalid metadata" in gap["detail"] for gap in report["gaps"]
    )


def test_offline_smoke_labels_hardwood_as_fixture_verified(capsys):
    from ingestion import cli

    assert cli.main([
        "smoke", "--league", "NBA", "--date", "20261107", "--offline-fixture"
    ]) == 0
    report = json.loads(capsys.readouterr().out)
    assert report["verification"] == "FIXTURE_VERIFIED"
    assert report["hardwood"]["verification"] == "FIXTURE_VERIFIED"


def test_newer_partial_prediction_does_not_displace_fresh_complete_bundle():
    from ingestion.predictions import select_prediction

    game = _game()
    complete = Prediction(
        game_id=game.game_id, provider="espn", observed_utc="2026-11-07T10:00:00Z",
        prob_home=0.23, prob_away=0.77, expected_home_score=110.5,
        expected_away_score=120.0, expected_total=230.5,
    )
    newer_partial = Prediction(
        game_id=game.game_id, provider="espn", observed_utc="2026-11-07T11:00:00Z",
        prob_home=0.25, prob_away=0.75,
    )
    result = select_prediction(
        [complete, newer_partial], game, "2026-11-07T12:00:00Z"
    )
    assert result.selected is complete
    assert result.reason == "complete_bundle"


def test_stored_prediction_survives_transient_provider_outage(monkeypatch):
    from ingestion.identity import now_utc_iso
    from ingestion.providers import espn, hardwood

    store = InMemoryStore()
    game = _game()
    stored = Prediction(
        game_id=game.game_id, provider="espn", observed_utc=now_utc_iso(),
        prob_home=0.23, prob_away=0.77, expected_home_score=110.5,
        expected_away_score=120.0, expected_total=230.5,
    )
    store.upsert_prediction(stored)
    monkeypatch.setattr(espn, "fetch_bpi", lambda *args: (_ for _ in ()).throw(RuntimeError("outage")))
    monkeypatch.setattr(hardwood, "load_forecasts", lambda *args: (_ for _ in ()).throw(RuntimeError("outage")))
    selections = _pipeline(store)._acquire_predictions([game])
    assert selections[game.game_id].selected is stored


def test_stored_odds_fallback_uses_predicted_side_only():
    store = InMemoryStore()
    game = _game()
    prediction = Prediction(
        game_id=game.game_id, provider="espn", observed_utc="2026-11-07T18:00:00Z",
        prob_home=0.2, prob_away=0.8,
    )
    moneyline = BookSnapshot(
        game_id=game.game_id, provider="one", book="draftkings",
        market=Market.MONEYLINE, period=Period.FULL_GAME,
        observed_utc="2026-11-07T18:00:00Z", pre_match=True, main_line=True,
        home_price_decimal=1.5, away_price_decimal=2.5,
    )
    store.upsert_book_snapshot(moneyline)
    pipeline = _pipeline(store)
    import ingestion.pipeline as module
    module.now_utc_iso = lambda: "2026-11-07T19:00:00Z"
    try:
        assert pipeline._stored_odds_for(game, prediction) == 2.5
        store.upsert_book_snapshot(BookSnapshot(
            game_id=game.game_id, provider="two", book="fanduel",
            market=Market.MONEYLINE, period=Period.FULL_GAME,
            observed_utc="2026-11-07T18:30:00Z", pre_match=True, main_line=True,
            home_price_decimal=1.4, away_price_decimal=None,
        ))
        assert pipeline._stored_odds_for(game, prediction) is None
    finally:
        module.now_utc_iso = __import__("ingestion.identity", fromlist=["now_utc_iso"]).now_utc_iso


def test_coverage_sums_quarter_lines_across_games():
    store = InMemoryStore()
    one, two = _game("one"), _game("two")
    for game in (one, two):
        store.upsert_game(game)
        store.upsert_book_snapshot(_snapshot(event=game.game_id, period=Period.Q1))
    report = _pipeline(store).coverage()
    assert report["games_in_store"] == 2
    assert report["quarter_lines"]["Q1"] == 2
