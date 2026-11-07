from __future__ import annotations

import json

from ingestion.duckdb_store import DuckDBStore
from ingestion.mining import mine_prediction_assay, prediction_outcome_rows, write_assay
from ingestion.models import (
    ActualResult,
    Game,
    GameStatus,
    League,
    PeriodKind,
    Prediction,
)


def _final_game(index: int, tied: bool = False) -> Game:
    away_final = 100 if tied else 90
    return Game(
        game_id=f"g{index}",
        league=League.NBA,
        season="2027",
        home_team_id="1",
        away_team_id="2",
        home_name="Home",
        away_name="Away",
        scheduled_utc=f"2026-11-{index + 1:02d}T22:00:00Z",
        status=GameStatus.FINAL,
        actual=ActualResult(
            period_kind=PeriodKind.QUARTERS,
            home_quarters=[25, 25, 25, 25],
            away_quarters=[25, 25, 25, away_final - 75],
            home_final=100,
            away_final=away_final,
            source="espn",
        ),
    )


def test_assay_uses_only_frozen_pretip_bundles_and_never_certifies(tmp_path):
    store = DuckDBStore(str(tmp_path / "warehouse.duckdb"))
    for index in range(10):
        game = _final_game(index)
        store.upsert_game(game)
        store.upsert_prediction(
            Prediction(
                game_id=game.game_id,
                provider="espn",
                observed_utc=game.scheduled_utc.replace("22:00:00Z", "21:00:00Z"),
                prob_home=0.6,
                prob_away=0.4,
            ),
            frozen=True,
        )
        # A contradictory post-tip shadow must not contaminate the assay.
        store.upsert_prediction(
            Prediction(
                game_id=game.game_id,
                provider="espn",
                observed_utc=game.scheduled_utc.replace("22:00:00Z", "23:00:00Z"),
                prob_home=0.1,
                prob_away=0.9,
                pre_tip=False,
            )
        )
    store.upsert_game(_final_game(10, tied=True))
    store.upsert_prediction(
        Prediction(
            game_id="g10",
            provider="espn",
            observed_utc="2026-11-11T21:00:00Z",
            prob_home=0.6,
            prob_away=0.4,
        ),
        frozen=True,
    )

    rows = prediction_outcome_rows(store, League.NBA)
    report = mine_prediction_assay(store, League.NBA)
    store.close()

    assert len(rows) == 10
    assert {row["outcome"] for row in rows} == {"win"}
    assert report["sample"] == {"n": 10, "wins": 10, "losses": 0}
    assert report["status"] == "research_only"
    assert report["certified_picks"] == 0
    assert len(report["edges"]) == 1
    assert report["edges"][0]["certified"] is False
    assert report["edges"][0]["segment"][-1] == "espn"
    assert report["edges"][0]["wilson_lower_95"] > 0
    assert "expected_total_to_ou_conversion" in report["exclusions"]


def test_assay_artifact_is_stable_on_identical_rerun(tmp_path):
    path = tmp_path / "assay.json"
    report = {"status": "research_only", "certified_picks": 0, "edges": []}
    assert write_assay(str(path), report) == str(path)
    first = path.read_bytes()
    write_assay(str(path), report)
    assert path.read_bytes() == first
    assert json.loads(path.read_text())["certified_picks"] == 0
