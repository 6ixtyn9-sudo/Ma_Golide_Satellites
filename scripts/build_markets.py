#!/usr/bin/env python3
"""Build calibrated multi-market forecasts and top estimated-probability accas.

The warehouse is always opened read-only. The JSON report is a derived run
artifact; no bookmaker, notification, or other external write is performed.
"""

from __future__ import annotations

import argparse
import json
import os
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from forecasting.harness import CalibrationConfig  # noqa: E402
from forecasting.model import ForecastConfig  # noqa: E402
from forecasting.operator import load_settled_pick_history  # noqa: E402
from forecasting.service import MarketForecastService, write_market_report  # noqa: E402
from ingestion.books import BookConfig  # noqa: E402
from ingestion.duckdb_store import DuckDBStore  # noqa: E402
from ingestion.identity import now_utc_iso  # noqa: E402
from ingestion.models import League  # noqa: E402
from ingestion.warehouse import DEFAULT_WAREHOUSE  # noqa: E402


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--warehouse", default=DEFAULT_WAREHOUSE)
    parser.add_argument(
        "--localdata",
        default=os.environ.get("LOCALDATA"),
        help="operator history directory (defaults to localdata/<league-lowercase>)",
    )
    parser.add_argument(
        "--league",
        default="NBA",
        choices=["NBA", "EUROLEAGUE", "CBA", "WNBA", "NCAA_M", "NCAA_W"],
    )
    parser.add_argument("--date", default=None, help="optional local YYYYMMDD slate")
    parser.add_argument("--tz", default="UTC", help="IANA timezone for --date")
    parser.add_argument("--as-of", default=None, help="forecast cutoff ISO timestamp")
    parser.add_argument("--book-provider", default=None)
    parser.add_argument("--book", default=None)
    parser.add_argument("--legs", type=int, default=3)
    parser.add_argument("--limit", type=int, default=5)
    parser.add_argument("--min-history-games", type=int, default=20)
    parser.add_argument("--min-residual-games", type=int, default=20)
    parser.add_argument("--min-calibration-samples", type=int, default=30)
    parser.add_argument("--calibration-prior", type=float, default=30.0)
    parser.add_argument("--output", default=None)
    args = parser.parse_args()
    warehouse = Path(args.warehouse)
    if not warehouse.exists():
        parser.error(f"warehouse does not exist: {warehouse}")
    output = Path(args.output or f"localdata/markets_{args.league}_{args.date or 'upcoming'}.json")
    if output.suffix.lower() != ".json":
        parser.error("--output must be a .json report")
    if output.resolve() == warehouse.resolve():
        parser.error("--output cannot overwrite the warehouse")

    provider = None if args.book_provider in (None, "", "none") else args.book_provider
    forecast_config = ForecastConfig(
        min_history_games=args.min_history_games,
        min_residual_games=args.min_residual_games,
    )
    calibration_config = CalibrationConfig(
        prior_strength=args.calibration_prior,
        min_samples=args.min_calibration_samples,
    )
    store = DuckDBStore(str(warehouse), persist=False)
    localdata = Path(args.localdata) if args.localdata else Path("localdata") / args.league.lower()
    try:
        report = MarketForecastService(
            store,
            League[args.league],
            forecast_config=forecast_config,
            calibration_config=calibration_config,
            book_config=BookConfig(provider=provider, book=args.book),
            paper_pick_history=load_settled_pick_history(localdata),
        ).build_report(
            as_of_utc=args.as_of or now_utc_iso(),
            date_yyyymmdd=args.date,
            timezone_name=args.tz,
            acca_legs=args.legs,
            acca_limit=args.limit,
        )
        write_market_report(str(output), report)
    finally:
        store.close()
    print(
        json.dumps(
            {
                "output": str(output),
                "status": report["status"],
                "forecast_count": report["forecast_count"],
                "graded_walk_forward_forecasts": report["walk_forward"]["graded_forecasts"],
                "calibration_segments": len(report["walk_forward"]["segments"]),
                "eligible_acca_games": report["accas"]["eligible_games"],
                "accas": len(report["accas"]["results"]),
                "paper_accas": len(report["accas"].get("paper_results") or []),
                "external_writes": False,
            },
            indent=2,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
