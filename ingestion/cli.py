"""Headless DuckDB acquisition CLI. Dry run is always the default.

Examples:
    python -m ingestion.cli smoke --league NBA --date 20261107
    python -m ingestion.cli pregame --league NBA --date 20261107
    python -m ingestion.cli freeze --league NBA --date 20261107
    python -m ingestion.cli results --league NBA --date 20261107
    python -m ingestion.cli coverage --league NBA

Canonical writes require both ``--write`` and ``ALLOW_MAGOLIDE_WRITES=true``.
"""

from __future__ import annotations

import argparse
import json
import os
from pathlib import Path
import sys
from typing import List, Optional

from .books import BookConfig
from .duckdb_store import DuckDBStore
from .pipeline import Pipeline, PipelineConfig
from .store import InMemoryStore
from .warehouse import DEFAULT_WAREHOUSE
from .write_guards import assert_fixture_write_isolation


def _warehouse(args) -> str:
    return (
        getattr(args, "warehouse", None)
        or os.environ.get("MAGOLIDE_WAREHOUSE")
        or DEFAULT_WAREHOUSE
    )


def _book_config(args) -> BookConfig:
    provider = getattr(args, "book_provider", None)
    return BookConfig(
        provider=None if provider in (None, "", "none") else provider,
        book=getattr(args, "book", None),
    )


def _guard_write(args) -> None:
    if not args.write:
        return
    if os.environ.get("ALLOW_MAGOLIDE_WRITES", "").strip().lower() != "true":
        raise SystemExit(
            "--write requires ALLOW_MAGOLIDE_WRITES=true; refusing canonical "
            "warehouse writes (default is dry run)"
        )
    if args.offline_fixture:
        if os.environ.get("ALLOW_FIXTURE_WRITES", "").strip().lower() != "true":
            raise SystemExit(
                "fixture persistence requires ALLOW_FIXTURE_WRITES=true and "
                "an isolated warehouse"
            )
        try:
            assert_fixture_write_isolation(_warehouse(args))
        except RuntimeError as exc:
            raise SystemExit(str(exc)) from exc


def _print(report: dict) -> None:
    print(json.dumps(report, indent=2, default=str))


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(prog="ingestion", description=__doc__)
    subcommands = parser.add_subparsers(dest="command", required=True)

    def add_common(command_parser) -> None:
        command_parser.add_argument(
            "--league",
            default="NBA",
            choices=["NBA", "EuroLeague", "CBA", "NCAA_M", "NCAA_W"],
        )
        command_parser.add_argument(
            "--date", default=None, help="league-local slate date YYYYMMDD"
        )
        command_parser.add_argument(
            "--tz", default=None, help="league-local timezone (IANA)"
        )
        command_parser.add_argument(
            "--warehouse",
            default=None,
            help="DuckDB path (default: MAGOLIDE_WAREHOUSE or localdata/warehouse.duckdb)",
        )
        command_parser.add_argument(
            "--write",
            action="store_true",
            help="persist canonical writes (also requires ALLOW_MAGOLIDE_WRITES=true)",
        )
        command_parser.add_argument(
            "--dry-run", action="store_true", help="explicit no-write mode (the default)"
        )
        command_parser.add_argument("--freshness-hours", type=float, default=12.0)
        command_parser.add_argument("--book", default=None, help="bookmaker name")
        command_parser.add_argument(
            "--book-provider",
            default="none",
            choices=["none", "therundown", "sportsgameodds", "theoddsapi", "isportsapi"],
            help="restrict book evidence to one provider",
        )
        command_parser.add_argument(
            "--offline-fixture",
            action="store_true",
            help="use committed fixtures; output is FIXTURE_VERIFIED, never live coverage",
        )

    for command in ("smoke", "pregame", "freeze", "results"):
        add_common(subcommands.add_parser(command))
    coverage_parser = subcommands.add_parser("coverage")
    coverage_parser.add_argument(
        "--league",
        default="NBA",
        choices=["NBA", "EuroLeague", "CBA", "NCAA_M", "NCAA_W"],
    )
    coverage_parser.add_argument("--warehouse", default=None)

    args = parser.parse_args(argv)
    if args.command == "smoke" and args.write:
        parser.error("smoke is read-only; remove --write")
    if args.command == "coverage":
        warehouse = _warehouse(args)
        if warehouse != ":memory:" and not Path(warehouse).exists():
            parser.error(f"warehouse does not exist: {warehouse}")
        store = DuckDBStore(warehouse, persist=False)
        try:
            config = PipelineConfig(league=args.league, date_yyyymmdd="00000000")
            _print(Pipeline(config, store).coverage())
        finally:
            store.close()
        return 0

    if not args.date:
        parser.error(f"--date is required for {args.command}")
    if args.freshness_hours <= 0:
        parser.error("--freshness-hours must be positive")
    _guard_write(args)

    config = PipelineConfig(
        league=args.league,
        date_yyyymmdd=args.date.replace("-", ""),
        tz_name=args.tz,
        freshness_hours=args.freshness_hours,
        book=_book_config(args),
        fixture_now_utc=(
            f"{args.date.replace('-', '')}T00:00:00Z"
            if args.offline_fixture
            else None
        ),
    )
    store = (
        InMemoryStore()
        if args.command == "smoke"
        else DuckDBStore(_warehouse(args), persist=bool(args.write))
    )
    try:
        fetch, fetch_factory = (
            _offline_fixture_fetch() if args.offline_fixture else (None, None)
        )
        pipeline = Pipeline(
            config,
            store,
            fetch=fetch,
            fetch_factory=fetch_factory,
            allow_writes=bool(args.write),
        )
        report = getattr(pipeline, args.command)()
        _print(report)
    finally:
        close = getattr(store, "close", None)
        if close:
            close()
    return 0


def _offline_fixture_fetch():
    """Serve committed golden fixtures without network or credentials."""

    fixtures = Path(__file__).resolve().parent.parent / "tests" / "fixtures"
    mapping = {
        "scoreboard": "espn/scoreboard_synth_golden.json",
        # Match the Hardwood artifact before the ESPN predictor substring.
        "game_forecasts": "hardwood/game_forecasts.json",
        "competitions/401909938/predictor": "espn/predictor_401909938.json",
        "seasons/E2026/games": "euroleague/schedule.json",
        "Boxscore": "euroleague/boxscore_97.json",
        "standings": "espn/standings.json",
    }

    def fetch(url: str):
        for needle, relative in mapping.items():
            if needle in url:
                return json.loads((fixtures / relative).read_text())
        from .http import HttpError

        raise HttpError(404, url, "no offline fixture mapped")

    return fetch, (lambda _provider: fetch)


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        sys.exit(130)
