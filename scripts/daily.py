#!/usr/bin/env python3
"""One resumable Ma Golide orchestrator for unattended three-hour runs.

Capture -> canonical DuckDB -> calibrated markets/accas -> immutable ledgers ->
research assay -> CSV exports -> optional operational notification.
``--auto-once`` is intentionally an alias for one invocation so GitHub Actions,
cron and local operators use the
same code path. Only DuckDB, CSV.GZ, local files, and optional notifications
are used.
"""

from __future__ import annotations

import argparse
from datetime import datetime, timedelta, timezone
import json
import os
from pathlib import Path
import re
import sys
from typing import Any, Dict
from zoneinfo import ZoneInfo

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from forecasting.operator import publish_operator_reports  # noqa: E402
from forecasting.service import MarketForecastService, write_market_report  # noqa: E402
from ingestion.books import BookConfig  # noqa: E402
from ingestion.cli import _offline_fixture_fetch  # noqa: E402
from ingestion.duckdb_store import DuckDBStore  # noqa: E402
from ingestion.ledger import write_canonical_ledger  # noqa: E402
from ingestion.mining import mine_prediction_assay, write_assay  # noqa: E402
from ingestion.pipeline import ACQUISITION_TZ, Pipeline, PipelineConfig  # noqa: E402
from ingestion.warehouse import DEFAULT_WAREHOUSE  # noqa: E402
from ingestion.write_guards import assert_fixture_write_isolation  # noqa: E402


def _utc_now() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def _market_cutoff(args: argparse.Namespace) -> str:
    """Cut off live markets after acquisition, not before it.

    Live predictions/books are observed after ``args.started_utc``. Using the
    process start here makes every fresh observation invisible to the forecast
    service. Fixture replay retains its deliberately pinned historical cutoff.
    """

    return args.started_utc if args.offline_fixture else _utc_now()


def _safe_run_id(value: str) -> str:
    return re.sub(r"[^A-Za-z0-9_.-]", "_", value)


def _write_run_report(localdata: Path, run_id: str, report: dict) -> str:
    target = localdata / "reports" / f"daily_{_safe_run_id(run_id)}.json"
    target.parent.mkdir(parents=True, exist_ok=True)
    temporary = target.with_suffix(".json.tmp")
    temporary.write_text(json.dumps(report, indent=2, sort_keys=True, default=str) + "\n")
    os.replace(temporary, target)
    return str(target)


def _prune_report_files(localdata: Path, cutoff_utc: str) -> int:
    """Prune parseable derived telemetry, never canonical immutable ledgers."""
    pruned = 0
    for path in (localdata / "reports").glob("*.json"):
        try:
            report = json.loads(path.read_text())
            generated = str(report.get("finished_utc") or report.get("generated_utc") or "")
        except (OSError, ValueError):
            continue
        if generated and generated < cutoff_utc:
            path.unlink()
            pruned += 1
    return pruned


def run_daily(args: argparse.Namespace) -> Dict[str, Any]:
    if args.write and os.environ.get("ALLOW_MAGOLIDE_WRITES", "").lower() != "true":
        raise RuntimeError("--write requires ALLOW_MAGOLIDE_WRITES=true")
    if (
        args.offline_fixture
        and args.write
        and os.environ.get("ALLOW_FIXTURE_WRITES", "").lower() != "true"
    ):
        raise RuntimeError(
            "fixture persistence requires ALLOW_FIXTURE_WRITES=true and an isolated warehouse"
        )
    if args.offline_fixture and args.write:
        assert_fixture_write_isolation(args.warehouse, localdata=args.localdata)
    if args.send_notifications and args.offline_fixture:
        raise RuntimeError("fixture reports cannot trigger external notifications")
    if args.keep_days < 1:
        raise RuntimeError("--keep-days must be positive")
    if args.freshness_hours <= 0:
        raise RuntimeError("--freshness-hours must be positive")
    if args.send_notifications and not args.notify:
        raise RuntimeError("--send-notifications requires --notify")
    if args.send_notifications and not args.write:
        raise RuntimeError("--send-notifications requires canonical --write")
    if (
        args.send_notifications
        and os.environ.get("ALLOW_NOTIFICATION_WRITES", "").lower() != "true"
    ):
        raise RuntimeError("--send-notifications requires ALLOW_NOTIFICATION_WRITES=true")

    localdata = Path(args.localdata)
    localdata.mkdir(parents=True, exist_ok=True)
    run_id = (
        args.run_id
        or os.environ.get("GITHUB_RUN_ID")
        or datetime.now(timezone.utc).strftime("local_%Y%m%dT%H%M%S%fZ")
    )
    if os.environ.get("GITHUB_RUN_ATTEMPT"):
        run_id = f"{run_id}_attempt_{os.environ['GITHUB_RUN_ATTEMPT']}"

    store = DuckDBStore(args.warehouse, persist=args.write)
    reports: Dict[str, Any] = {}
    ledgers: Dict[str, Any] = {}
    try:
        fetch, fetch_factory = _offline_fixture_fetch() if args.offline_fixture else (None, None)
        config = PipelineConfig(
            league=args.league,
            date_yyyymmdd=args.date.replace("-", ""),
            tz_name=args.tz,
            freshness_hours=args.freshness_hours,
            history_days=(0 if args.offline_fixture else getattr(args, "history_days", 7)),
            history_request_budget=getattr(args, "history_requests", 7),
            deep_history_event_budget=(
                0 if args.offline_fixture else getattr(args, "deep_history_events", 25)
            ),
            book=BookConfig(
                provider=(None if args.book_provider in (None, "", "none") else args.book_provider),
                book=args.book,
            ),
            fixture_now_utc=(
                f"{args.date.replace('-', '')}T00:00:00Z" if args.offline_fixture else None
            ),
        )
        pipeline = Pipeline(
            config,
            store,
            fetch=fetch,
            fetch_factory=fetch_factory,
            allow_writes=args.write,
        )

        backfill_report = pipeline.backfill()
        backfill_ledgers: Dict[str, Any] = {}
        if backfill_report["backfill"]["status"] not in {"disabled", "unsupported_league"}:
            store.record_run(run_id, "backfill", args.date, backfill_report)
            if args.write:
                for date in backfill_report["backfill"].get("ledger_dates", []):
                    backfill_ledgers[date] = write_canonical_ledger(
                        store,
                        config.league_enum,
                        date,
                        "backfill",
                        str(localdata / "ledgers"),
                        config.timezone,
                    )

        commands = ["pregame", "freeze"] + ([] if args.skip_results else ["results"])
        market_report = None
        market_report_path = None
        for command in commands:
            report = getattr(pipeline, command)()
            reports[command] = report
            store.record_run(run_id, command, args.date, report)
            if args.write:
                ledgers[command] = write_canonical_ledger(
                    store,
                    config.league_enum,
                    args.date,
                    command,
                    str(localdata / "ledgers"),
                    config.timezone,
                )
            # Current markets use visible pre-tip evidence. The following
            # freeze stage preserves tip-passed rows for later grading.
            if command == "pregame" and not getattr(args, "skip_markets", False):
                market_report = MarketForecastService(
                    store,
                    config.league_enum,
                    book_config=config.book,
                ).build_report(
                    as_of_utc=_market_cutoff(args),
                    date_yyyymmdd=config.date_yyyymmdd,
                    timezone_name=config.timezone,
                    acca_legs=getattr(args, "acca_legs", 3),
                    acca_limit=getattr(args, "acca_limit", 5),
                )
                market_report_path = write_market_report(
                    str(
                        localdata / "reports" / f"markets_{args.league}_{config.date_yyyymmdd}.json"
                    ),
                    market_report,
                )

        cutoff_utc = (datetime.now(timezone.utc) - timedelta(days=args.keep_days)).strftime(
            "%Y-%m-%dT%H:%M:%SZ"
        )
        pruned_runs = store.prune_pipeline_runs(cutoff_utc) if args.write else 0
        # Dry-run schedules also write reports, so retention cannot depend on
        # canonical writes being enabled.
        pruned_reports = _prune_report_files(localdata, cutoff_utc)

        operator = publish_operator_reports(
            store,
            config.league_enum,
            market_report,
            localdata,
            target_date=args.date,
            generated_utc=args.started_utc,
            run_id=run_id,
        )

        assay = mine_prediction_assay(store, config.league_enum)
        assay_path = (
            write_assay(str(localdata / f"assay_{args.league}.json"), assay) if args.write else None
        )
        exports = store.export_csv(str(localdata / "csv")) if args.write else {}

        summary = {
            "schema_version": "magolide.daily.v1",
            "run_id": run_id,
            "target_date": args.date,
            "league": args.league,
            "mode": "write" if args.write else "dry_run",
            "verification": "FIXTURE_VERIFIED" if args.offline_fixture else "LIVE_ATTEMPT",
            "system_of_record": "duckdb_csv",
            "warehouse": args.warehouse,
            "started_utc": args.started_utc,
            "finished_utc": _utc_now(),
            "backfill": backfill_report,
            "backfill_ledgers": backfill_ledgers,
            "commands": reports,
            "markets": market_report,
            "market_report_path": market_report_path,
            "operator": operator,
            "ledgers": ledgers,
            "assay": assay,
            "assay_path": assay_path,
            "csv_exports": exports,
            "counts": store.counts(),
            "retention": {
                "keep_days": args.keep_days,
                "pipeline_runs_pruned": pruned_runs,
                "report_files_pruned": pruned_reports,
            },
        }
        notification: Dict[str, Any] = {"enabled": args.notify, "status": "disabled"}
        if args.notify:
            from scripts.notify import notify_summary

            try:
                notification = notify_summary(
                    summary,
                    send=args.send_notifications,
                    ledger_path=str(localdata / "notification_ledger.json"),
                )
                notification["enabled"] = True
            except Exception as exc:
                # Notification failure is visible but cannot roll back or block
                # the already-committed canonical acquisition state.
                notification = {
                    "enabled": True,
                    "sent": False,
                    "status": "failed",
                    "error": f"{type(exc).__name__}: {str(exc)[:200]}",
                }
        summary["notification"] = notification
        summary["report_path"] = _write_run_report(localdata, run_id, summary)
        return summary
    finally:
        store.close()


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--auto-once",
        action="store_true",
        help="run exactly one autonomous iteration (compatibility/explicitness flag)",
    )
    parser.add_argument(
        "--date",
        default=None,
        help="slate date (default: today in the league's acquisition timezone)",
    )
    parser.add_argument(
        "--league",
        default="NBA",
        choices=["NBA", "EuroLeague", "CBA", "WNBA", "NCAA_M", "NCAA_W"],
    )
    parser.add_argument("--tz", default=None)
    parser.add_argument(
        "--warehouse", default=os.environ.get("MAGOLIDE_WAREHOUSE", DEFAULT_WAREHOUSE)
    )
    parser.add_argument("--localdata", default="localdata")
    parser.add_argument(
        "--keep-days",
        type=int,
        default=30,
        help="retain bounded run telemetry; immutable ledgers are retained",
    )
    parser.add_argument("--run-id", default=None)
    parser.add_argument("--write", action="store_true", help="persist canonical DuckDB/CSV/ledgers")
    parser.add_argument("--offline-fixture", action="store_true")
    parser.add_argument(
        "--history-days",
        type=int,
        default=21,
        help="rolling ESPN result lookback (offline fixtures disable backfill)",
    )
    parser.add_argument(
        "--history-requests",
        type=int,
        default=21,
        help="maximum historical scoreboard requests per league run",
    )
    parser.add_argument(
        "--deep-history-events",
        type=int,
        default=50,
        help="maximum indexed historical events per write-enabled league run",
    )
    parser.add_argument("--skip-results", action="store_true")
    parser.add_argument("--skip-markets", action="store_true")
    parser.add_argument("--acca-legs", type=int, default=3)
    parser.add_argument("--acca-limit", type=int, default=5)
    parser.add_argument("--freshness-hours", type=float, default=12.0)
    parser.add_argument("--book", default=None)
    parser.add_argument(
        "--book-provider",
        default="none",
        choices=["none", "therundown", "sportsgameodds", "theoddsapi", "isportsapi", "espn"],
    )
    parser.add_argument("--notify", action="store_true", help="build an operations heartbeat")
    parser.add_argument(
        "--send-notifications",
        action="store_true",
        help="send the heartbeat through configured WhatsApp credentials",
    )
    args = parser.parse_args()
    if args.date is None:
        args.date = datetime.now(ZoneInfo(ACQUISITION_TZ[args.league.upper()])).strftime("%Y-%m-%d")
    try:
        args.date = datetime.strptime(args.date.replace("-", ""), "%Y%m%d").strftime("%Y-%m-%d")
    except ValueError:
        parser.error("--date must be YYYY-MM-DD or YYYYMMDD")
    args.started_utc = f"{args.date}T18:00:00Z" if args.offline_fixture else _utc_now()

    try:
        summary = run_daily(args)
    except Exception as exc:
        print(
            json.dumps({"status": "failed", "error": str(exc)}, indent=2),
            file=sys.stderr,
        )
        return 1
    print(json.dumps(summary, indent=2, default=str))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
