"""WNBA is an isolated, fully routed ESPN league, not an enum-only scaffold."""

from __future__ import annotations

import json

import pytest

from ingestion.books import BookConfig
from ingestion.duckdb_store import DuckDBStore
from ingestion.models import League, Market, Period
from ingestion.pipeline import Pipeline, PipelineConfig
from ingestion.providers import theoddsapi


def test_wnba_pipeline_persists_native_periods_context_and_espn_books(tmp_path, monkeypatch):
    monkeypatch.setattr(theoddsapi, "is_enabled", lambda: False)
    calls = []

    def fetch(url):
        calls.append(url)
        if "scoreboard" in url:
            assert "/wnba/" in url
            assert "groups=50" not in url
            path = "tests/fixtures/espn/scoreboard_wnba_409900001.json"
        elif "competitions/409900001/predictor" in url:
            assert "/wnba/" in url
            path = "tests/fixtures/espn/predictor_409900001_partial.json"
        elif "summary?event=409900001" in url:
            assert "/wnba/" in url
            path = "tests/fixtures/espn/summary_wnba_409900001.json"
        else:  # pragma: no cover - defensive route assertion
            raise AssertionError(url)
        with open(path) as handle:
            return json.load(handle)

    warehouse = tmp_path / "localdata" / "wnba" / "warehouse.duckdb"
    store = DuckDBStore(str(warehouse), persist=True)
    report = Pipeline(
        PipelineConfig(
            league="WNBA",
            date_yyyymmdd="20261107",
            book=BookConfig(provider="espn"),
            fixture_now_utc="2026-11-07T18:00:00Z",
        ),
        store,
        fetch=fetch,
        allow_writes=True,
    ).pregame()

    assert report["fixtures_found"] == 1
    assert report["predictions_selected"] == {"espn": 1}
    assert report["espn_context"]["enriched"] == 1
    assert not [gap for gap in report["gaps"] if gap["required_field"] == "espn_bpi"]
    row = report["rows"][0]
    assert row["prediction_provider"] == "espn"
    assert row["prediction_reason"] == "partial_bundle_probabilities_only"
    assert row["prediction"]["expected_total"] is None
    assert row["full_game_line"] == pytest.approx(168.5)
    assert row["odds"] == pytest.approx(1 + 100 / 170)
    assert row["odds_evidence"] == {
        "provider": "espn",
        "book": "DraftKings",
        "observed_utc": "2026-11-07T18:00:00Z",
        "source_modified_utc": None,
        "payload_hash": row["odds_evidence"]["payload_hash"],
    }
    assert row["halves"] == {}
    assert set(row["quarters"]) == {"Q1", "Q2", "Q3", "Q4"}
    assert set(row["native_periods"]) == {"Q1", "Q2", "Q3", "Q4"}

    game = store.get_game("409900001")
    assert game.league is League.WNBA
    assert game.home_name == "Las Vegas Aces"
    assert game.away_name == "New York Liberty"
    assert game.evidence["espn"]["scoreboard"]["teams"]["home"]["records"][0]["summary"] == "20-10"
    assert game.evidence["espn"]["summary"]["article"]["headline"] == "Liberty at Aces"
    books = store.get_book_snapshots(game.game_id)
    assert {snapshot.market for snapshot in books} == {
        Market.MONEYLINE,
        Market.SPREAD,
        Market.TOTAL,
    }
    assert all(snapshot.period is Period.FULL_GAME for snapshot in books)
    assert not store.list_games(league=League.NCAA_W)
    store.close()

    reopened = DuckDBStore(str(warehouse), persist=False)
    try:
        recovered = reopened.get_game("409900001")
        assert recovered.league is League.WNBA
        assert recovered.evidence == game.evidence
        assert len(reopened.get_book_snapshots(game.game_id)) == 3
    finally:
        reopened.close()
