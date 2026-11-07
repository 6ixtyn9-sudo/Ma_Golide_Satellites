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

from assayer.assayer_engine import assay_side_data
from .duckdb_store import DuckDBStore
from .models import League


def prediction_outcome_rows(store: DuckDBStore, league: League) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    for game in store.list_games(league=league):
        if game.actual is None or not game.actual.reconciles()[0]:
            continue
        if game.actual.home_final is None or game.actual.away_final is None:
            continue
        if game.actual.home_final == game.actual.away_final:
            continue
        actual_side = (
            "home" if game.actual.home_final > game.actual.away_final else "away"
        )
        for prediction in store.get_predictions(game.game_id):
            if not prediction.pre_tip or not store.is_prediction_frozen(
                game.game_id, prediction.provider
            ):
                continue
            if prediction.prob_home is None or prediction.prob_away is None:
                continue
            if prediction.prob_home == prediction.prob_away:
                continue
            selected_side = (
                "home" if prediction.prob_home > prediction.prob_away else "away"
            )
            rows.append(
                {
                    "game_id": game.game_id,
                    "league": league.value,
                    "quarter": "full_game",
                    "tier": "provider_moneyline",
                    "side": selected_side,
                    "confidence": max(prediction.prob_home, prediction.prob_away),
                    "outcome": "win" if selected_side == actual_side else "loss",
                    "provider": prediction.provider,
                    "observed_utc": prediction.observed_utc,
                    "result_source": game.actual.source,
                }
            )
    return sorted(
        rows, key=lambda row: (row["game_id"], row["provider"], row["observed_utc"])
    )


def mine_prediction_assay(store: DuckDBStore, league: League) -> Dict[str, Any]:
    outcome_rows = prediction_outcome_rows(store, league)
    edges = []
    for provider in sorted({row["provider"] for row in outcome_rows}):
        provider_rows = [row for row in outcome_rows if row["provider"] == provider]
        for edge in assay_side_data(provider_rows, source_label=provider):
            edge["status"] = "research_only"
            edge["certified"] = False
            edge["certification_reason"] = "walk_forward_validation_required"
            edges.append(edge)
    edges.sort(key=lambda edge: (edge["segment"][-1], -edge["win_rate"], -edge["n"]))
    outcomes = Counter(row["outcome"] for row in outcome_rows)
    return {
        "schema_version": "magolide.research_assay.v1",
        "league": league.value,
        "status": "research_only",
        "certified_picks": 0,
        "method": "frozen_pre_tip_moneyline_vs_reconciled_official_final",
        "exclusions": [
            "post_tip_predictions",
            "unfrozen_predictions",
            "missing_or_tied_probabilities",
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
