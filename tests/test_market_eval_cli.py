"""Read-only warehouse integration checks for the market-evaluation CLI."""

from __future__ import annotations

import json
import sys

from ingestion.duckdb_store import DuckDBStore
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
from scripts.market_eval import main


def _game(game_id: str, *, reconciled: bool = True) -> Game:
    actual = ActualResult(
        period_kind=PeriodKind.QUARTERS,
        home_quarters=[27, 27, 27, 27],
        away_quarters=[25, 25, 25, 25],
        home_final=108 if reconciled else 109,
        away_final=100,
        source="test",
    )
    return Game(
        game_id=game_id,
        league=League.NBA,
        season="2026",
        home_team_id="home-id",
        away_team_id="away-id",
        home_name="Home",
        away_name="Away",
        scheduled_utc="2026-11-16T18:00:00Z",
        status=GameStatus.FINAL,
        actual=actual,
    )


def test_cli_uses_frozen_forecasts_separate_models_books_and_reconciled_finals(
    tmp_path, monkeypatch, capsys
):
    warehouse = tmp_path / "warehouse.duckdb"
    store = DuckDBStore(str(warehouse))
    try:
        store.upsert_game(_game("valid"))
        store.upsert_game(_game("bad-result", reconciled=False))
        for provider, version, frozen in (
            ("feed-a", "model-a", True),
            ("feed-b", "model-b", True),
            ("feed-unfrozen", "model-c", False),
        ):
            store.upsert_prediction(
                Prediction(
                    game_id="valid",
                    provider=provider,
                    observed_utc="2026-11-16T10:00:00Z",
                    model=f"forecast-{provider}",
                    model_version=version,
                    prob_home=0.65,
                    prob_away=0.35,
                ),
                frozen=frozen,
            )
        for book, observed, home_price, away_price in (
            ("Pinnacle", "2026-11-16T09:30:00Z", 2.10, 1.80),
            ("OtherBook", "2026-11-16T09:45:00Z", 2.00, 1.90),
        ):
            store.upsert_book_snapshot(
                BookSnapshot(
                    game_id="valid",
                    provider="odds-feed",
                    book=book,
                    market=Market.MONEYLINE,
                    period=Period.FULL_GAME,
                    observed_utc=observed,
                    home_price_decimal=home_price,
                    away_price_decimal=away_price,
                    pre_match=True,
                ),
                frozen=True,
            )
    finally:
        store.close()

    monkeypatch.setattr(
        sys,
        "argv",
        [
            "market_eval",
            "--warehouse",
            str(warehouse),
            "--as-of",
            "2026-11-17T00:00:00Z",
            "--bootstrap-iterations",
            "0",
        ],
    )
    assert main() == 0
    output = capsys.readouterr().out
    assert "NaN" not in output
    report = json.loads(output)

    assert report["data_audit"]["settled_games"] == 1
    assert report["data_audit"]["rejected_nonfinal_or_unreconciled_result_rows"] == 1
    assert report["data_audit"]["unfrozen_pre_tip_prediction_rows"] == 1
    assert report["game_model_pairs_considered"] == 2
    assert set(report["model_results"]) == {
        "feed-a::forecast-feed-a::model-a",
        "feed-b::forecast-feed-b::model-b",
    }
    for model_result in report["model_results"].values():
        assert set(model_result["by_book"]) == {"Pinnacle", "OtherBook"}
    assert report["join_exclusions"]["forecast_not_frozen"] == 1
