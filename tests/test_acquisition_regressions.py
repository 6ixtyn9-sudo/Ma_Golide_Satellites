"""Regression tests for canonical acquisition, identity, freezing, and CLI."""

from __future__ import annotations

import json
from types import SimpleNamespace

import pytest

from ingestion.models import (
    ActualResult,
    BookSnapshot,
    Game,
    GameStatus,
    League,
    Market,
    Period,
    PeriodKind,
    Prediction,
)
from ingestion.pipeline import Pipeline, PipelineConfig
from ingestion.store import FrozenWriteError, InMemoryStore


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


def _snapshot(provider="one", event="event-1", period=Period.FULL_GAME, **changes):
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
    values.update(changes)
    return BookSnapshot(**values)


def _pipeline(store=None):
    return Pipeline(
        PipelineConfig(
            league="NBA",
            date_yyyymmdd="20261107",
            fixture_now_utc="2026-11-07T19:00:00Z",
        ),
        store or InMemoryStore(),
        fetch=lambda url: (_ for _ in ()).throw(AssertionError("no network")),
    )


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
    snapshots = {
        "one": [_snapshot("one", "event-one")],
        "two": [_snapshot("two", "event-two", Period.Q1)],
    }
    modules = {
        provider: SimpleNamespace(
            is_enabled=lambda: True,
            fetch_odds=lambda *args, p=provider: snapshots[p],
        )
        for provider in snapshots
    }
    monkeypatch.setattr(module, "BOOK_PROVIDERS", tuple(modules))
    monkeypatch.setattr(module, "_book_module", modules.__getitem__)

    mapped = pipeline._acquire_book_lines([game])
    assert {snap.provider for snap in mapped[game.game_id]} == {"one", "two"}
    assert {snap.provider for snap in store.get_book_snapshots(game.game_id)} == {
        "one",
        "two",
    }
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


def test_new_provider_event_cannot_replace_existing_provider_mapping(monkeypatch):
    import ingestion.pipeline as module

    game = _game()
    game.provider_ids["one"] = "event-old"
    raw = _snapshot(event="event-new")
    mod = SimpleNamespace(is_enabled=lambda: True, fetch_odds=lambda *args: [raw])
    monkeypatch.setattr(module, "BOOK_PROVIDERS", ("one",))
    monkeypatch.setattr(module, "_book_module", lambda provider: mod)
    store = InMemoryStore()
    pipeline = _pipeline(store)
    assert pipeline._acquire_book_lines([game]) == {}
    assert game.provider_ids["one"] == "event-old"
    assert "existing mapping" in pipeline.gaps[-1].detail


def test_unmapped_provider_id_without_identity_is_not_trusted(monkeypatch):
    import ingestion.pipeline as module

    game = _game("coincident-id")
    raw = _snapshot(
        event="coincident-id", home_name=None, away_name=None, scheduled_utc=None
    )
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
        game_id=game.game_id,
        observed_utc="2026-11-07T23:00:00Z",
        pre_match=False,
        line=999.0,
    )
    for snapshot in (pre_total, pre_quarter, post_total):
        store.upsert_book_snapshot(snapshot)

    pipeline = _pipeline(store)
    assert pipeline._freeze_books(game) == 2
    assert store.is_book_snapshot_frozen(game.game_id, pre_total)
    assert store.is_book_snapshot_frozen(game.game_id, pre_quarter)
    assert not store.is_book_snapshot_frozen(game.game_id, post_total)
    assert len(store.get_book_snapshots(game.game_id)) == 3

    with pytest.raises(FrozenWriteError):
        store.upsert_book_snapshot(_snapshot(game_id=game.game_id, line=231.0))
    store.upsert_book_snapshot(
        _snapshot(
            game_id=game.game_id,
            observed_utc="2026-11-07T23:05:00Z",
            pre_match=False,
            line=888.0,
        )
    )
    pre = [
        snap
        for snap in store.get_book_snapshots(game.game_id)
        if snap.period is Period.FULL_GAME and snap.pre_match
    ]
    assert len(pre) == 1 and pre[0].line == 230.5


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


def test_offline_cli_dry_run_does_not_create_warehouse(tmp_path, capsys):
    from ingestion import cli

    warehouse = tmp_path / "dry-run.duckdb"
    rc = cli.main(
        [
            "pregame",
            "--league",
            "NBA",
            "--date",
            "20261107",
            "--offline-fixture",
            "--warehouse",
            str(warehouse),
        ]
    )
    assert rc == 0 and not warehouse.exists()
    report = json.loads(capsys.readouterr().out)
    assert report["persistence"]["kind"] == "duckdb"
    assert report["persistence"]["write_enabled"] is False
    assert report["persistence"]["system_of_record"] == "duckdb_csv"


def test_offline_smoke_labels_hardwood_as_fixture_verified(tmp_path, capsys):
    from ingestion import cli

    warehouse = tmp_path / "smoke.duckdb"
    assert cli.main(
        [
            "smoke",
            "--league",
            "NBA",
            "--date",
            "20261107",
            "--offline-fixture",
            "--warehouse",
            str(warehouse),
        ]
    ) == 0
    report = json.loads(capsys.readouterr().out)
    assert report["verification"] == "FIXTURE_VERIFIED"
    assert report["hardwood"]["verification"] == "FIXTURE_VERIFIED"
    assert not warehouse.exists()


def test_newer_partial_prediction_does_not_displace_fresh_complete_bundle():
    from ingestion.predictions import select_prediction

    game = _game()
    complete = Prediction(
        game_id=game.game_id,
        provider="espn",
        observed_utc="2026-11-07T10:00:00Z",
        prob_home=0.23,
        prob_away=0.77,
        expected_home_score=110.5,
        expected_away_score=120.0,
        expected_total=230.5,
    )
    newer_partial = Prediction(
        game_id=game.game_id,
        provider="espn",
        observed_utc="2026-11-07T11:00:00Z",
        prob_home=0.25,
        prob_away=0.75,
    )
    result = select_prediction(
        [complete, newer_partial], game, "2026-11-07T12:00:00Z"
    )
    assert result.selected is complete
    assert result.reason == "complete_bundle"


def test_stored_prediction_survives_transient_provider_outage(monkeypatch):
    from ingestion.providers import espn, hardwood

    store = InMemoryStore()
    game = _game()
    stored = Prediction(
        game_id=game.game_id,
        provider="espn",
        observed_utc="2026-11-07T18:30:00Z",
        prob_home=0.23,
        prob_away=0.77,
        expected_home_score=110.5,
        expected_away_score=120.0,
        expected_total=230.5,
    )
    store.upsert_prediction(stored)
    monkeypatch.setattr(
        espn,
        "fetch_bpi",
        lambda *args, **kwargs: (_ for _ in ()).throw(RuntimeError("outage")),
    )
    monkeypatch.setattr(
        hardwood,
        "load_forecasts",
        lambda *args: (_ for _ in ()).throw(RuntimeError("outage")),
    )
    selections = _pipeline(store)._acquire_predictions([game])
    assert selections[game.game_id].selected is stored


def test_odds_fallback_uses_latest_complete_price_for_predicted_side():
    store = InMemoryStore()
    game = _game()
    prediction = Prediction(
        game_id=game.game_id,
        provider="espn",
        observed_utc="2026-11-07T18:00:00Z",
        prob_home=0.2,
        prob_away=0.8,
    )
    older = BookSnapshot(
        game_id=game.game_id,
        provider="one",
        book="draftkings",
        market=Market.MONEYLINE,
        period=Period.FULL_GAME,
        observed_utc="2026-11-07T18:00:00Z",
        home_price_decimal=1.5,
        away_price_decimal=2.5,
    )
    newer_partial = BookSnapshot(
        game_id=game.game_id,
        provider="two",
        book="fanduel",
        market=Market.MONEYLINE,
        period=Period.FULL_GAME,
        observed_utc="2026-11-07T18:30:00Z",
        home_price_decimal=1.4,
        away_price_decimal=None,
    )
    store.upsert_book_snapshot(older)
    store.upsert_book_snapshot(newer_partial)
    price, evidence = _pipeline(store)._odds_evidence_for(
        store.get_book_snapshots(game.game_id), prediction
    )
    assert price == 2.5
    assert evidence["book"] == "draftkings"


def test_prediction_selection_orders_timestamps_as_instants():
    from ingestion.predictions import select_prediction

    game = _game()
    earlier = Prediction(
        game_id=game.game_id,
        provider="espn",
        observed_utc="2026-11-07T10:30:00-02:00",  # 12:30Z
        prob_home=0.6,
        prob_away=0.4,
        expected_home_score=110,
        expected_away_score=100,
        expected_total=210,
    )
    later = Prediction(
        game_id=game.game_id,
        provider="espn",
        observed_utc="2026-11-07T13:00:00Z",
        prob_home=0.7,
        prob_away=0.3,
        expected_home_score=112,
        expected_away_score=100,
        expected_total=212,
    )
    selected = select_prediction([later, earlier], game, "2026-11-07T14:00:00Z")
    assert selected.selected is later


def _actual(**changes):
    values = dict(
        period_kind=PeriodKind.QUARTERS,
        home_quarters=[25, 25, 25, 25],
        away_quarters=[20, 20, 20, 20],
        home_final=100,
        away_final=80,
        source="test",
    )
    values.update(changes)
    return ActualResult(**values)


@pytest.mark.parametrize(
    "actual",
    [
        _actual(home_quarters=[25, 25, 25]),
        _actual(home_overtime=5, away_overtime=None),
        _actual(home_overtime=5, away_overtime=4, overtime_periods=0),
        _actual(home_overtime=None, away_overtime=None, overtime_periods=1),
        _actual(home_final=float("nan")),
        _actual(away_quarters=[20, 20, 20, -1], away_final=79),
    ],
)
def test_actual_reconciliation_rejects_incomplete_or_impossible_scores(actual):
    assert actual.reconciles()[0] is False


def test_actual_reconciliation_accepts_balanced_overtime():
    actual = _actual(
        home_quarters=[25, 25, 25, 20],
        away_quarters=[20, 20, 20, 20],
        home_overtime=10,
        away_overtime=5,
        home_final=105,
        away_final=85,
        overtime_periods=1,
    )
    assert actual.reconciles() == (True, "")


def test_schedule_rediscovery_does_not_erase_official_final():
    store = InMemoryStore()
    final = _game()
    final.status = GameStatus.FINAL
    final.actual = _actual()
    store.upsert_game(final)
    store.upsert_game(_game())
    loaded = store.get_game(final.game_id)
    assert loaded.status is GameStatus.FINAL
    assert loaded.actual.home_final == 100


def test_nonfinite_or_negative_prediction_expectations_are_rejected():
    for field in ("expected_home_score", "expected_away_score", "expected_total"):
        for value in (float("nan"), float("inf"), -1.0):
            with pytest.raises(ValueError):
                Prediction(
                    game_id="g1",
                    provider="espn",
                    observed_utc="2026-11-07T18:00:00Z",
                    **{field: value},
                )
    with pytest.raises(ValueError):
        Prediction(
            game_id="g1",
            provider="espn",
            observed_utc="2026-11-07T18:00:00Z",
            expected_margin=float("nan"),
        )


def test_nonfinite_lines_and_invalid_decimal_prices_are_rejected():
    for line in (float("nan"), float("inf")):
        with pytest.raises(ValueError):
            _snapshot(line=line)
    for field in (
        "over_price_decimal",
        "under_price_decimal",
        "home_price_decimal",
        "away_price_decimal",
        "draw_price_decimal",
    ):
        for price in (float("nan"), float("inf"), 0.99):
            with pytest.raises(ValueError):
                _snapshot(**{field: price})


def test_invalid_timestamp_is_rejected_not_merely_suffix_checked():
    with pytest.raises(ValueError, match="valid ISO"):
        Prediction(
            game_id="g",
            provider="test",
            observed_utc="2026-99-99T00:00:00Z",
        )
