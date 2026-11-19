#!/usr/bin/env python3
"""Mine research-only moneyline assay rows from the canonical warehouse."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from ingestion.duckdb_store import DuckDBStore  # noqa: E402
from ingestion.mining import mine_prediction_assay, write_assay  # noqa: E402
from ingestion.league_registry import cli_league_choices, league_spec, parse_league  # noqa: E402
from ingestion.models import League  # noqa: E402
from ingestion.warehouse import DEFAULT_WAREHOUSE  # noqa: E402


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--warehouse", default=None)
    parser.add_argument(
        "--league",
        default="NBA",
        choices=cli_league_choices(),
    )
    parser.add_argument(
        "--output",
        default=None,
        help="optional derived JSON artifact; stdout is always available",
    )
    args = parser.parse_args()
    args.league = parse_league(args.league).value
    if args.warehouse is None:
        args.warehouse = (
            str(Path("localdata") / league_spec(parse_league(args.league)).state_key / "warehouse.duckdb")
            if args.league not in {"NBA", "EUROLEAGUE", "CBA", "WNBA", "NCAA_M", "NCAA_W"}
            else DEFAULT_WAREHOUSE
        )
    if not Path(args.warehouse).exists():
        parser.error(f"warehouse does not exist: {args.warehouse}")
    league = League[args.league.upper()]
    store = DuckDBStore(args.warehouse, persist=False)
    try:
        report = mine_prediction_assay(store, league)
        if args.output:
            write_assay(args.output, report)
        print(json.dumps(report, indent=2, default=str))
    finally:
        store.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
