"""Warehouse-native research assay for frozen prediction bundles.

Only an explicit home/away probability pick and an official reconciled final
produce an outcome row. Expected totals are not converted into O/U bets, and
post-tip shadow observations never enter the historical pregame assay.
"""

from __future__ import annotations

from collections import Counter
import json
import os
from pathlib import Path
from typing import Any, Dict, List

from assayer.assayer_engine import assay_prediction_outcomes
from .duckdb_store import DuckDBStore
from .identity import parse_utc
from .models import GameStatus, League


def prediction_outcome_rows(store: DuckDBStore, league: League) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    for game in store.list_games(league=league):
        if game.status is not GameStatus.FINAL:
            continue
        if game.actual is None or not game.actual.reconciles()[0]:
            continue
        if game.actual.home_final is None or game.actual.away_final is None:
            continue
        if game.actual.home_final == game.actual.away_final:
            continue
        actual_side = "home" if game.actual.home_final > game.actual.away_final else "away"
        for prediction in store.get_predictions(game.game_id):
            if not prediction.pre_tip or not store.is_prediction_frozen(
                game.game_id, prediction.provider
            ):
                continue
            if prediction.prob_home is None or prediction.prob_away is None:
                continue
            if abs(prediction.prob_home + prediction.prob_away - 1.0) > 0.01:
                continue
            if parse_utc(prediction.observed_utc) >= parse_utc(game.scheduled_utc):
                continue
            if prediction.prob_home == prediction.prob_away:
                continue
            selected_side = "home" if prediction.prob_home > prediction.prob_away else "away"
            rows.append(
                {
                    "game_id": game.game_id,
                    "league": league.value,
                    "period": "full_game",
                    "market": "moneyline",
                    "side": selected_side,
                    "confidence": max(prediction.prob_home, prediction.prob_away),
                    "outcome": "win" if selected_side == actual_side else "loss",
                    "provider": prediction.provider,
                    "observed_utc": prediction.observed_utc,
                    "result_source": game.actual.source,
                }
            )
    return sorted(rows, key=lambda row: (row["game_id"], row["provider"], row["observed_utc"]))


def mine_prediction_assay(store: DuckDBStore, league: League) -> Dict[str, Any]:
    outcome_rows = prediction_outcome_rows(store, league)
    edges = assay_prediction_outcomes(outcome_rows)
    for edge in edges:
        edge["status"] = "research_only"
        edge["certified"] = False
        edge["certification_reason"] = "walk_forward_validation_required"
    outcomes = Counter(row["outcome"] for row in outcome_rows)
    return {
        "schema_version": "magolide.research_assay.v2",
        "league": league.value,
        "status": "research_only",
        "certified_picks": 0,
        "method": "frozen_pre_tip_moneyline_vs_reconciled_official_final",
        "exclusions": [
            "post_tip_predictions",
            "unfrozen_predictions",
            "missing_tied_or_incoherent_probabilities",
            "predictions_not_observed_strictly_before_tip",
            "unreconciled_or_tied_finals",
            "expected_total_to_ou_conversion",
        ],
        "sample": {
            "n": len(outcome_rows),
            "wins": outcomes.get("win", 0),
            "losses": outcomes.get("loss", 0),
        },
        "edges": edges,
    }


def write_assay(path: str, report: Dict[str, Any]) -> str:
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    raw = (json.dumps(report, indent=2, sort_keys=True, default=str) + "\n").encode()
    if target.exists() and target.read_bytes() == raw:
        return str(target)
    temporary = target.with_suffix(".json.tmp")
    temporary.write_bytes(raw)
    os.replace(temporary, target)
    return str(target)
