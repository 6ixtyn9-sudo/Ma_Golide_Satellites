from __future__ import annotations

import hashlib
import json
from pathlib import Path
import sys

from ingestion.duckdb_store import DuckDBStore
from ingestion.models import Game, League, Prediction
from scripts import build_markets


def _digest(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def test_standalone_market_cli_reads_duckdb_without_modifying_it(tmp_path, monkeypatch, capsys):
    warehouse = tmp_path / "warehouse.duckdb"
    output = tmp_path / "markets.json"
    store = DuckDBStore(str(warehouse))
    store.upsert_game(
        Game(
            game_id="upcoming",
            league=League.NBA,
            season="2027",
            home_team_id="home",
            away_team_id="away",
            home_name="Home",
            away_name="Away",
            scheduled_utc="2026-11-08T00:00:00Z",
        )
    )
    store.upsert_prediction(
        Prediction(
            game_id="upcoming",
            provider="espn",
            observed_utc="2026-11-07T18:00:00Z",
            prob_home=0.6,
            prob_away=0.4,
            expected_home_score=110,
            expected_away_score=100,
            expected_total=210,
            expected_margin=10,
        )
    )
    store.close()
    before = _digest(warehouse)

    monkeypatch.setattr(
        sys,
        "argv",
        [
            "build_markets.py",
            "--warehouse",
            str(warehouse),
            "--league",
            "NBA",
            "--date",
            "20261108",
            "--tz",
            "Africa/Johannesburg",
            "--as-of",
            "2026-11-07T18:00:00Z",
            "--legs",
            "2",
            "--limit",
            "3",
            "--output",
            str(output),
        ],
    )
    assert build_markets.main() == 0
    summary = json.loads(capsys.readouterr().out)
    report = json.loads(output.read_text())
    assert summary["forecast_count"] == 1
    assert summary["external_writes"] is False
    assert report["forecasts"][0]["market"] == "moneyline"
    assert report["forecasts"][0]["eligibility"]["reasons"] == [
        "no_exact_bookmaker_market",
        "insufficient_walk_forward_samples",
        "insufficient_team_specific_history",
    ]
    assert report["accas"]["results"] == []
    assert _digest(warehouse) == before
