"""CLI for the DuckDB-first acquisition pipeline. Dry run is the default.

    python -m ingestion.cli smoke     --league NBA --date 20261107
    python -m ingestion.cli pregame   --league NBA --date 20261107
    ALLOW_MAGOLIDE_WRITES=true python -m ingestion.cli pregame \
        --league NBA --date 20261107 --write
    python -m ingestion.cli freeze    --league NBA --date 20261107
    python -m ingestion.cli results   --league NBA --date 20261105
    python -m ingestion.cli coverage  --league NBA

Canonical state defaults to ``localdata/warehouse.duckdb``.  ``--write``
requires ``ALLOW_MAGOLIDE_WRITES=true``.  Google Sheets stores/worksheets are
available only through explicit legacy flags and additionally require
``ALLOW_SHEET_WRITES=true``; they are never selected implicitly.

Exit codes: 0 success, 1 operational failure, 2 usage error.
"""

from __future__ import annotations

import argparse
import json
import os
import sys
from typing import List, Optional

from .duckdb_store import DuckDBStore
from .pipeline import Pipeline, PipelineConfig, open_persistent_store, open_satellite_worksheets
from .store import InMemoryStore, JsonFileStore
from .warehouse import DEFAULT_WAREHOUSE
from .write_guards import assert_fixture_write_isolation


def _build_store(args, write: bool = False) -> object:
    if getattr(args, "store_json", None):
        return JsonFileStore(args.store_json, persist=write)
    data_store_id = getattr(args, "data_store_id", None)
    if data_store_id:
        return open_persistent_store(data_store_id, write=write)
    warehouse = (
        getattr(args, "warehouse", None)
        or os.environ.get("MAGOLIDE_WAREHOUSE")
        or DEFAULT_WAREHOUSE
    )
    return DuckDBStore(warehouse, persist=write)


def _resolve_sheet(args, write: bool):
    if not args.sheet_id:
        return None, None
    return open_satellite_worksheets(args.sheet_id, write=write)


def _guard_write(args) -> None:
    if not args.write:
        return
    if str(os.environ.get("ALLOW_MAGOLIDE_WRITES", "")).strip().lower() != "true":
        raise SystemExit(
            "--write requires ALLOW_MAGOLIDE_WRITES=true; refusing canonical "
            "warehouse writes (default is dry run)"
        )
    if args.offline_fixture and str(
        os.environ.get("ALLOW_FIXTURE_WRITES", "")
    ).strip().lower() != "true":
        raise SystemExit(
            "fixture persistence requires ALLOW_FIXTURE_WRITES=true and an isolated warehouse"
        )
    if args.offline_fixture and not (args.store_json or args.data_store_id):
        warehouse = args.warehouse or os.environ.get(
            "MAGOLIDE_WAREHOUSE", DEFAULT_WAREHOUSE
        )
        try:
            assert_fixture_write_isolation(warehouse)
        except RuntimeError as exc:
            raise SystemExit(str(exc)) from exc
    if (args.sheet_id or args.data_store_id) and str(
        os.environ.get("ALLOW_SHEET_WRITES", "")
    ).strip().lower() != "true":
        raise SystemExit(
            "legacy Google Sheets writes additionally require ALLOW_SHEET_WRITES=true"
        )


def _print(report: dict) -> None:
    print(json.dumps(report, indent=1, default=str))


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(prog="ingestion", description=__doc__)
    sub = parser.add_subparsers(dest="command", required=True)

    def add_common(p):
        p.add_argument("--league", default="NBA", help="NBA or EuroLeague")
        p.add_argument("--date", default=None, help="slate date YYYYMMDD (league-local)")
        p.add_argument("--tz", default=None, help="display timezone (IANA)")
        p.add_argument("--warehouse", default=None,
                       help="canonical DuckDB path (default: MAGOLIDE_WAREHOUSE or localdata/warehouse.duckdb)")
        p.add_argument("--store-json", default=None, help="local JSON store path (development only)")
        p.add_argument("--sheet-id", default=None,
                       help="LEGACY satellite worksheet export target; not the system of record")
        p.add_argument("--data-store-id", default=None,
                       help="LEGACY Google Sheets store id; explicit migration/compatibility only")
        p.add_argument("--write", action="store_true",
                       help="opt-in persistent writes (needs ALLOW_MAGOLIDE_WRITES=true)")
        p.add_argument("--dry-run", action="store_true", help="explicit dry run (default anyway)")
        p.add_argument("--freshness-hours", type=float, default=12.0)
        p.add_argument("--book", default=None, help="configured bookmaker name (e.g. draftkings)")
        p.add_argument("--book-provider", default="none",
                       help="bookmaker provider for merge policy (therundown/sportsgameodds/theoddsapi)")
        p.add_argument("--quarter-csv", default=None, help="manual quarter-line CSV fallback")
        p.add_argument("--offline-fixture", action="store_true",
                       help="serve committed fixtures instead of live endpoints "
                            "(hermetic demo; coverage labelled FIXTURE_VERIFIED)")

    add_common(sub.add_parser("smoke"))
    add_common(sub.add_parser("pregame"))
    add_common(sub.add_parser("freeze"))
    add_common(sub.add_parser("results"))
    coverage_parser = sub.add_parser("coverage")
    coverage_parser.add_argument("--league", default="NBA")
    coverage_parser.add_argument("--warehouse", default=None)
    coverage_parser.add_argument("--data-store-id", default=None,
                                 help="LEGACY Google Sheets store id")
    coverage_parser.add_argument("--store-json", default=None)

    args = parser.parse_args(argv)

    if args.command == "coverage":
        config = PipelineConfig(league=args.league, date_yyyymmdd="00000000")
        store = _build_store(args)
        try:
            report = Pipeline(config, store).coverage()
            _print(report)
        finally:
            close = getattr(store, "close", None)
            if close:
                close()
        return 0

    if not args.date:
        parser.error(f"--date is required for {args.command}")

    _guard_write(args)
    config = PipelineConfig(
        league=args.league,
        date_yyyymmdd=args.date.replace("-", ""),
        tz_name=args.tz,
        sheet_id=args.sheet_id,
        freshness_hours=args.freshness_hours,
        quarter_csv=args.quarter_csv,
        book=_book_config(args),
        fixture_now_utc=(
            f"{args.date.replace('-', '')}T00:00:00Z"
            if args.offline_fixture
            else None
        ),
    )
    allow_writes = bool(args.write)

    store = InMemoryStore() if args.command == "smoke" else _build_store(args, write=allow_writes)
    if type(store) is InMemoryStore and args.write and args.command != "smoke":
        raise SystemExit("--write requires a persistent warehouse/store")

    try:
        upcoming_ws, results_ws = _resolve_sheet(args, write=allow_writes)
        fetch, fetch_factory = (None, None)
        if getattr(args, "offline_fixture", False):
            fetch, fetch_factory = _offline_fixture_fetch()
        pipeline = Pipeline(
            config, store, fetch=fetch, fetch_factory=fetch_factory,
            allow_writes=allow_writes,
            upcoming_ws=upcoming_ws, results_ws=results_ws,
        )
        if args.command == "smoke":
            _print(pipeline.smoke())
        elif args.command == "pregame":
            _print(pipeline.pregame())
        elif args.command == "freeze":
            _print(pipeline.freeze())
        elif args.command == "results":
            _print(pipeline.results())
    finally:
        close = getattr(store, "close", None)
        if close:
            close()
    return 0


def _offline_fixture_fetch():
    """Hermetic demo fetch serving the committed golden fixtures.

    Output is labelled FIXTURE_VERIFIED by construction; never use it to
    claim live coverage.
    """
    import os

    fixtures = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                            "tests", "fixtures")

    def fetch(url: str):
        import json

        mapping = {
            "scoreboard": "espn/scoreboard_synth_golden.json",
            # The Hardwood URL contains "nba_predictor"; match its filename
            # before the ESPN predictor substring or it returns the wrong JSON.
            "game_forecasts": "hardwood/game_forecasts.json",
            "competitions/401909938/predictor": "espn/predictor_401909938.json",
            "seasons/E2026/games": "euroleague/schedule.json",
            "Boxscore": "euroleague/boxscore_97.json",
            "standings": "espn/standings.json",
        }
        for needle, rel in mapping.items():
            if needle in url:
                with open(os.path.join(fixtures, rel)) as f:
                    return json.load(f)
        from .http import HttpError

        raise HttpError(404, url, "no offline fixture mapped")

    return fetch, (lambda provider: fetch)


def _book_config(args):
    from .books import BookConfig

    return BookConfig(
        provider=args.book_provider,
        book=args.book,
    )


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        sys.exit(130)
