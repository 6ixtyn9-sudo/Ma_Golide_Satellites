#!/usr/bin/env python3
"""Evaluate frozen pre-tip forecasts against decision-time named-book prices.

Research-only market-comparison lane. Reads the canonical warehouse through
the read-only research views, evaluates each frozen provider/model-version
forecast independently for each named book against settled results under the
pre-registered entry policy, and reports:

- probability quality: paired Brier/log-loss, model vs devigged market,
  with a date-block bootstrap interval;
- paper stake diagnostics: flat-stake yield on stake (primary), return on
  the stated bank (secondary), drawdown on the stated basis, and a capped
  fractional-Kelly diagnostic (full Kelly is refused);
- an explicit data audit: quote counts, missing quote timestamps, and every
  join exclusion reason.

Profit is never claimed. When timestamped pre-tip named-book prices are
missing, the report says profit cannot be evaluated instead of substituting
prices that were not available at decision time.
"""

from __future__ import annotations

import argparse
import json
import math
import sys
from pathlib import Path
from typing import List

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from forecasting.market_eval import (  # noqa: E402
    GameOutcome,
    MarketEvalConfig,
    MarketQuote,
    ModelForecast,
    betting_diagnostics,
    evaluate_pairs,
    join_for_evaluation,
    parse_utc,
)
from ingestion.models import ActualResult, PeriodKind  # noqa: E402
from ingestion.warehouse import DEFAULT_WAREHOUSE, connect  # noqa: E402

LEAGUES = ["NBA", "EUROLEAGUE", "EuroLeague", "EUROLEAGUE_W", "CBA", "WNBA", "NCAA_M", "NCAA_W"]


def _json_safe(value):
    """Convert non-finite diagnostic values to standards-compliant JSON nulls."""
    if isinstance(value, float) and not math.isfinite(value):
        return None
    if isinstance(value, dict):
        return {key: _json_safe(item) for key, item in value.items()}
    if isinstance(value, list):
        return [_json_safe(item) for item in value]
    return value


def _model_identity(provider, model_name, model_version) -> str:
    return "::".join(
        (provider or "unknown", model_name or "unknown", model_version or "unknown")
    )


def _not_evaluable(reason: str, audit: dict) -> dict:
    return {
        "cannot_evaluate": {
            "profit": reason,
            "reason_detail": (
                "observed quotes without decision-time timestamps cannot be "
                "turned into an entry policy without inventing evidence"
            ),
        },
        "data_audit": audit,
    }


def _settled_outcomes(rows):
    """Keep only final rows whose stored period scores reconcile to finals."""
    outcomes = []
    rejected = 0
    for game_id, scheduled_utc, status, actual_json in rows:
        if status != "FINAL" or not actual_json:
            rejected += 1
            continue
        try:
            payload = json.loads(actual_json)
            actual = ActualResult(
                period_kind=PeriodKind(payload["period_kind"]),
                home_quarters=payload["home_quarters"],
                away_quarters=payload["away_quarters"],
                home_overtime=payload.get("home_overtime"),
                away_overtime=payload.get("away_overtime"),
                home_final=payload.get("home_final"),
                away_final=payload.get("away_final"),
                overtime_periods=payload.get("overtime_periods", 0),
                source=payload.get("source", ""),
            )
            reconciled, _ = actual.reconciles()
            if not reconciled or actual.home_final is None or actual.away_final is None:
                rejected += 1
                continue
            if not float(actual.home_final).is_integer() or not float(actual.away_final).is_integer():
                rejected += 1
                continue
            scheduled = parse_utc(scheduled_utc)
            if scheduled is None:
                rejected += 1
                continue
            outcomes.append(
                GameOutcome(
                    game_id=game_id,
                    home_final=int(actual.home_final),
                    away_final=int(actual.away_final),
                    scheduled_utc=scheduled,
                )
            )
        except (KeyError, TypeError, ValueError, json.JSONDecodeError):
            rejected += 1
    return outcomes, rejected


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--warehouse", default=None)
    parser.add_argument("--league", default="NBA", choices=LEAGUES)
    parser.add_argument("--book", default=None, help="restrict quotes to one named book")
    parser.add_argument(
        "--as-of", default=None, help="only games scheduled at or before this UTC time (ISO-8601)"
    )
    parser.add_argument("--output", default=None, help="optional JSON artifact path")
    parser.add_argument("--bootstrap-iterations", type=int, default=2000)
    parser.add_argument("--seed", type=int, default=20261117)
    parser.add_argument("--max-quote-age-hours", type=float, default=72.0)
    args = parser.parse_args()
    if args.warehouse is None:
        args.warehouse = (
            "localdata/euroleague_w/warehouse.duckdb"
            if args.league == "EUROLEAGUE_W"
            else DEFAULT_WAREHOUSE
        )
    # The canonical enum is EUROLEAGUE; keep the familiar display spelling as
    # a CLI alias without querying an impossible warehouse league value.
    if args.league == "EuroLeague":
        args.league = "EUROLEAGUE"

    if not Path(args.warehouse).exists():
        parser.error(f"warehouse does not exist: {args.warehouse}")

    con = connect(args.warehouse, read_only=True)
    try:
        league_params: List = [args.league]
        as_of_clause = ""
        if args.as_of:
            as_of_clause = (
                " AND TRY_CAST(g.scheduled_utc AS TIMESTAMPTZ) "
                "<= TRY_CAST(? AS TIMESTAMPTZ)"
            )
            league_params = [args.league, args.as_of]

        result_rows = con.execute(
            "SELECT game_id, scheduled_utc, status, actual_json "
            "FROM game_results g WHERE g.league = ?" + as_of_clause,
            league_params,
        ).fetchall()
        prediction_rows = con.execute(
            "SELECT p.game_id, p.provider, p.captured_at, p.model_version, "
            "       json_extract_string(p.payload_json, '$.model') AS model_name, "
            "       p.prob_home, p.prob_away, p.frozen "
            "FROM prediction_facts p JOIN games g ON g.game_id = p.game_id "
            "WHERE p.pre_tip AND g.league = ?" + as_of_clause,
            league_params,
        ).fetchall()
        quote_rows = con.execute(
            "SELECT b.game_id, b.book, b.observed_utc, b.home_price_decimal, "
            "       b.away_price_decimal, b.pre_match, b.frozen "
            "FROM book_facts b JOIN games g ON g.game_id = b.game_id "
            "WHERE b.market = 'moneyline' AND b.period = 'full_game' "
            "AND g.league = ?" + as_of_clause,
            league_params,
        ).fetchall()
    finally:
        con.close()

    outcomes, rejected_result_rows = _settled_outcomes(result_rows)
    missing_quote_timestamps = sum(1 for row in quote_rows if row[2] is None)
    audit = {
        "league": args.league,
        "result_rows_with_actual": len(result_rows),
        "settled_games": len(outcomes),
        "rejected_nonfinal_or_unreconciled_result_rows": rejected_result_rows,
        "pre_tip_prediction_rows": len(prediction_rows),
        "frozen_pre_tip_prediction_rows": sum(1 for row in prediction_rows if row[7]),
        "unfrozen_pre_tip_prediction_rows": sum(1 for row in prediction_rows if not row[7]),
        "forecast_rows_by_provider_model": {},
        "moneyline_full_game_quote_rows": len(quote_rows),
        "quote_rows_missing_observed_utc": missing_quote_timestamps,
        "as_of": args.as_of,
        "book_filter": args.book,
    }

    for row in prediction_rows:
        model_key = _model_identity(row[1], row[4], row[3])
        counts = audit["forecast_rows_by_provider_model"]
        counts[model_key] = counts.get(model_key, 0) + 1

    if not quote_rows or missing_quote_timestamps == len(quote_rows):
        report = _not_evaluable(
            "no timestamped moneyline/full-game quotes in the warehouse for this league", audit
        )
    else:
        skipped_probabilities = sum(
            1 for row in prediction_rows if row[5] is None or row[6] is None
        )
        forecasts = [
            ModelForecast(
                game_id=row[0],
                model=_model_identity(row[1], row[4], row[3]),
                prob_home=row[5],
                prob_away=row[6],
                captured_at=parse_utc(row[2]),
                frozen=bool(row[7]),
            )
            for row in prediction_rows
            if row[5] is not None and row[6] is not None
        ]
        quotes = [
            MarketQuote(
                game_id=row[0],
                book=row[1],
                home_price_decimal=row[3],
                away_price_decimal=row[4],
                observed_utc=parse_utc(row[2]),
                pre_match=bool(row[5]),
                frozen=bool(row[6]),
            )
            for row in quote_rows
        ]
        audit["prediction_rows_missing_probabilities"] = skipped_probabilities
        config = MarketEvalConfig(
            book=args.book,
            quote_max_age_hours=args.max_quote_age_hours,
            bootstrap_iterations=args.bootstrap_iterations,
            bootstrap_seed=args.seed,
        )
        joined = join_for_evaluation(forecasts, quotes, outcomes, config)
        pairs_by_model_book = {}
        for pair in joined.pairs:
            pairs_by_model_book.setdefault((pair.model, pair.book), []).append(pair)
        model_results = {}
        for (model, book), rows in sorted(pairs_by_model_book.items()):
            model_result = model_results.setdefault(model, {"by_book": {}})
            model_result["by_book"][book] = {
                "probability_comparison": evaluate_pairs(rows, config),
                "paper_stake_diagnostics": betting_diagnostics(rows, config),
            }
        report = {
            "entry_policy": {
                "description": (
                    "latest frozen pre-tip forecast revision per game and provider/model "
                    "version; newest named-book moneyline/full-game quote observed at or "
                    "before that forecast's decision cutoff, within the maximum quote age"
                ),
                "frozen_forecast_required": config.require_frozen_forecast,
                "frozen_quote_required": config.require_frozen,
                "book": args.book,
                "max_quote_age_hours": args.max_quote_age_hours,
                "bootstrap_iterations": args.bootstrap_iterations,
                "bootstrap_seed": args.seed,
            },
            "game_model_pairs_considered": joined.games_considered,
            "join_exclusions": joined.exclusions,
            "model_results": model_results,
            "data_audit": audit,
            "interpretation": (
                "research-only comparisons are reported separately for each acquired "
                "provider/model version, against market-implied probabilities on identical "
                "games; observed quotes are not guaranteed execution and positive paper "
                "stakes are not evidence of profitability"
            ),
        }

    rendered = json.dumps(_json_safe(report), indent=2, default=str, allow_nan=False)
    if args.output:
        Path(args.output).write_text(rendered + "\n", encoding="utf-8")
    print(rendered)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
