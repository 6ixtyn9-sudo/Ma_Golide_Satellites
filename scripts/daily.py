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

from forecasting.operator import (  # noqa: E402
    load_settled_pick_history,
    publish_operator_reports,
)
from forecasting.service import MarketForecastService, write_market_report  # noqa: E402
from ingestion.books import BookConfig  # noqa: E402
from ingestion.cli import _offline_fixture_fetch  # noqa: E402
from ingestion.duckdb_store import DuckDBStore  # noqa: E402
from ingestion.ledger import write_canonical_ledger  # noqa: E402
from ingestion.league_registry import cli_league_choices, league_spec, parse_league  # noqa: E402
from ingestion.mining import mine_prediction_assay, write_assay  # noqa: E402
from ingestion.pipeline import (  # noqa: E402
    ACQUISITION_TZ,
    Pipeline,
    PipelineConfig,
    default_history_backfill_plan,
)
from ingestion.providers import aba_liga  # noqa: E402
from ingestion.warehouse import DEFAULT_WAREHOUSE  # noqa: E402
from scripts.release_archive import _checkpoint_regressions  # noqa: E402
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


ABA_HISTORY_RECEIPTS_FILE = "history_receipts.json"
ABA_HISTORY_RECEIPTS_SCHEMA = "magolide.aba_history_receipts"
_ABA_RECEIPT_FIELDS = frozenset(
    {
        "season_end_year",
        "calendar_final_count",
        "reconciled_final_count",
        "administrative_disposition_count",
        "not_applicable_rows",
        "blocking_gap_count",
        "status",
        "source",
        "evidence",
        "completed_at_utc",
    }
)
_ABA_RECEIPT_IDENTITY_FIELDS = _ABA_RECEIPT_FIELDS - {"completed_at_utc"}
_ABA_LEGACY_RECEIPT_FIELDS = _ABA_RECEIPT_FIELDS - {"administrative_disposition_count"}


def _validated_aba_history_receipt(receipt: object, *, description: str) -> dict[str, Any]:
    # Receipts written before generic administrative dispositions existed had
    # only reconciled played finals. Preserve those immutable attestations as a
    # zero-administrative-count form rather than stranding prior Git evidence.
    if isinstance(receipt, dict) and set(receipt) == _ABA_LEGACY_RECEIPT_FIELDS:
        if receipt.get("status") != "calendar_final_count_reconciled":
            raise ValueError(f"invalid ABA history receipt: {description}")
        receipt = {
            **receipt,
            "administrative_disposition_count": 0,
            "status": "calendar_terminal_count_reconciled",
        }
    if not isinstance(receipt, dict) or set(receipt) != _ABA_RECEIPT_FIELDS:
        raise ValueError(f"invalid ABA history receipt: {description}")
    season = receipt["season_end_year"]
    counts = (
        receipt["calendar_final_count"],
        receipt["reconciled_final_count"],
        receipt["administrative_disposition_count"],
        receipt["not_applicable_rows"],
        receipt["blocking_gap_count"],
    )
    if (
        isinstance(season, bool)
        or not isinstance(season, int)
        or any(isinstance(value, bool) or not isinstance(value, int) or value < 0 for value in counts)
        or receipt["calendar_final_count"]
        != receipt["reconciled_final_count"] + receipt["administrative_disposition_count"]
        or receipt["blocking_gap_count"] != 0
        or receipt["status"] != "calendar_terminal_count_reconciled"
        or receipt["source"] != aba_liga.PROVIDER
        or not isinstance(receipt["evidence"], str)
        or not receipt["evidence"]
        or not isinstance(receipt["completed_at_utc"], str)
        or not receipt["completed_at_utc"]
    ):
        raise ValueError(f"invalid ABA history receipt: {description}")
    return dict(receipt)


def _load_aba_history_receipts(path: Path) -> dict[str, dict[str, Any]]:
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise ValueError(f"invalid ABA history receipt summary: {path}") from exc
    if (
        not isinstance(data, dict)
        or data.get("schema_version") != ABA_HISTORY_RECEIPTS_SCHEMA
        or data.get("league") != "ABA_LEAGUE"
        or data.get("provider") != aba_liga.PROVIDER
        or not isinstance(data.get("receipts"), dict)
    ):
        raise ValueError(f"invalid ABA history receipt summary: {path}")
    receipts: dict[str, dict[str, Any]] = {}
    for season, receipt in data["receipts"].items():
        if not isinstance(season, str) or not season.isdigit():
            raise ValueError(f"invalid ABA history receipt season: {path}")
        validated = _validated_aba_history_receipt(receipt, description=f"{path}:{season}")
        if validated["season_end_year"] != int(season):
            raise ValueError(f"ABA history receipt season mismatch: {path}:{season}")
        receipts[season] = validated
    return receipts


def _write_aba_history_receipts(
    localdata: Path,
    receipts: object,
    *,
    current_season_end_year: int,
) -> str | None:
    """Persist only immutable completed ABA season-count attestations.

    The current season's count can legitimately change as new finals post, so
    it remains in the warehouse/provider-health cursor. Completed historical
    seasons are merged into one Git-durable, count-only summary; raw source
    pages, game rows, prices, and ticket evidence never enter this file.
    """

    target = localdata / ABA_HISTORY_RECEIPTS_FILE
    existing = _load_aba_history_receipts(target) if target.exists() else {}
    if not isinstance(receipts, list):
        raise ValueError("ABA history receipt batch must be a list")
    merged = dict(existing)
    changed = False
    for index, raw_receipt in enumerate(receipts):
        receipt = _validated_aba_history_receipt(raw_receipt, description=f"batch[{index}]")
        season = receipt["season_end_year"]
        if season >= current_season_end_year:
            continue
        key = str(season)
        previous = merged.get(key)
        if previous is not None:
            if any(previous[field] != receipt[field] for field in _ABA_RECEIPT_IDENTITY_FIELDS):
                raise ValueError(
                    f"ABA history receipt conflicts with durable evidence for season {season}"
                )
            continue
        merged[key] = receipt
        changed = True
    if not merged:
        return None
    if changed or not target.exists():
        payload = {
            "schema_version": ABA_HISTORY_RECEIPTS_SCHEMA,
            "league": "ABA_LEAGUE",
            "provider": aba_liga.PROVIDER,
            "receipts": {key: merged[key] for key in sorted(merged, key=int)},
        }
        raw = json.dumps(payload, indent=2, sort_keys=True) + "\n"
        temporary = target.with_name(target.name + ".tmp")
        temporary.write_text(raw, encoding="utf-8")
        os.replace(temporary, target)
    return str(target)


def _write_coverage_gap_marker(
    localdata: Path, *, league: str, target_date: str, write_enabled: bool
) -> str | None:
    """Publish one stable daily marker for a known-unavailable source lane.

    The marker is not sports evidence and does not claim an empty slate. It is
    the gate's durable proof that the explicit CoverageGap already ran today,
    allowing idle unavailable lanes to avoid repeating a full pipeline every
    scheduler tick while still fail-opening on missing/corrupt state.
    """

    spec = league_spec(parse_league(league))
    if not write_enabled or spec.schedule_source != "source_unavailable":
        return None
    payload = {
        "schema_version": "magolide.coverage_gap_marker",
        "league": spec.league.value,
        "target_date": target_date,
        "schedule_source": spec.schedule_source,
        "results_source": spec.results_source,
        "status": "source_unavailable",
        "detail": spec.source_detail,
    }
    target = localdata / f"coverage_gap_{target_date}.json"
    raw = json.dumps(payload, indent=2, sort_keys=True) + "\n"
    if target.is_file() and target.read_text() == raw:
        return str(target)
    temporary = target.with_name(target.name + ".tmp")
    temporary.write_text(raw)
    os.replace(temporary, target)
    return str(target)


def _prune_report_files(localdata: Path, cutoff_utc: str) -> int:
    """Prune derived telemetry, never canonical immutable ledgers.

    Parseable reports use their generated clock. A malformed report cannot keep
    consuming cache space forever: its filesystem mtime is a bounded fallback.
    """

    pruned = 0
    for path in (localdata / "reports").glob("*.json"):
        try:
            report = json.loads(path.read_text())
            generated = str(report.get("finished_utc") or report.get("generated_utc") or "")
        except (OSError, ValueError):
            try:
                generated = datetime.fromtimestamp(
                    path.stat().st_mtime, timezone.utc
                ).strftime("%Y-%m-%dT%H:%M:%SZ")
            except OSError:
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
        if getattr(args, "write", False):
            assert_checkpoint_not_regressed(args, store)
        fetch, fetch_factory = _offline_fixture_fetch() if args.offline_fixture else (None, None)
        history_plan = default_history_backfill_plan(
            parse_league(args.league),
            history_days=getattr(args, "history_days", None),
            history_request_budget=getattr(args, "history_requests", None),
        )
        config = PipelineConfig(
            league=args.league,
            date_yyyymmdd=args.date.replace("-", ""),
            tz_name=args.tz,
            freshness_hours=args.freshness_hours,
            history_days=(0 if args.offline_fixture else int(history_plan["history_days"])),
            history_request_budget=(
                0 if args.offline_fixture else int(history_plan["history_request_budget"])
            ),
            deep_history_event_budget=(
                0 if args.offline_fixture else getattr(args, "deep_history_events", 25)
            ),
            book=BookConfig(
                provider=(None if args.book_provider in (None, "", "none") else args.book_provider),
                book=args.book,
            ),
            acquire_book_lines=not getattr(args, "skip_book_acquisition", False),
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
        deep_history = backfill_report["backfill"].get("deep_history") or {}
        history_receipts_path = None
        if config.league_enum.value == "ABA_LEAGUE":
            current_aba_season = aba_liga.season_end_year(
                datetime.strptime(config.date_yyyymmdd, "%Y%m%d").date()
            )
            history_receipts_path = _write_aba_history_receipts(
                localdata,
                deep_history.get("season_receipts") or [],
                current_season_end_year=current_aba_season,
            )
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
                # Late finals for earlier days belong in those days' ledgers.
                for date in report.get("rechecked_dates") or []:
                    ledgers[f"{command}:{date}"] = write_canonical_ledger(
                        store,
                        config.league_enum,
                        date,
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
                    paper_pick_history=load_settled_pick_history(localdata),
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

        # Live acquisition and forecasting can take minutes. Use the actual
        # publication time for the pre-tip ticket decision rather than process
        # start, while fixtures retain their fixed replay timestamp.
        operator_generated_utc = args.started_utc if args.offline_fixture else _utc_now()
        operator = publish_operator_reports(
            store,
            config.league_enum,
            market_report,
            localdata,
            target_date=args.date,
            generated_utc=operator_generated_utc,
            run_id=run_id,
            slate_timezone_name=config.timezone,
        )

        assay = mine_prediction_assay(store, config.league_enum)
        assay_path = (
            write_assay(str(localdata / f"assay_{args.league}.json"), assay) if args.write else None
        )
        exports = store.export_csv(str(localdata / "csv")) if args.write else {}
        coverage_gap_marker = _write_coverage_gap_marker(
            localdata,
            league=config.league_enum.value,
            target_date=args.date,
            write_enabled=args.write,
        )

        summary = {
            "schema_version": "magolide.daily",
            "run_id": run_id,
            "target_date": args.date,
            "league": args.league,
            "mode": "write" if args.write else "dry_run",
            "verification": "FIXTURE_VERIFIED" if args.offline_fixture else "LIVE_ATTEMPT",
            "system_of_record": "duckdb_csv",
            "warehouse": args.warehouse,
            "started_utc": args.started_utc,
            "finished_utc": _utc_now(),
            "acquisition_plan": {
                **history_plan,
                "enabled": not args.offline_fixture,
                "effective_history_days": config.history_days,
                "effective_history_request_budget": config.history_request_budget,
            },
            "backfill": backfill_report,
            "backfill_ledgers": backfill_ledgers,
            "history_receipts_path": history_receipts_path,
            "commands": reports,
            "markets": market_report,
            "market_report_path": market_report_path,
            "operator": operator,
            "ledgers": ledgers,
            "assay": assay,
            "assay_path": assay_path,
            "csv_exports": exports,
            "coverage_gap_marker": coverage_gap_marker,
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


def _checkpoint_manifest(args) -> Path | None:
    """Locate the committed checkpoint manifest for this lane, if any."""

    override = os.environ.get("MAGOLIDE_CHECKPOINT_MANIFEST")
    if override:
        return Path(override)
    state_key = league_spec(parse_league(args.league)).state_key
    for candidate in (
        Path(args.localdata) / "archive_manifest.json",
        Path("localdata") / state_key / "archive_manifest.json",
    ):
        if candidate.is_file():
            return candidate
    return None


def assert_checkpoint_not_regressed(args, store) -> None:
    """Refuse a write run whose warehouse is thinner than its checkpoint.

    A cold cache can leave a nearly empty DuckDB behind. That warehouse then
    runs a full day, exports thin CSVs, and re-packs itself as the canonical
    checkpoint, so the loss becomes permanent. Failing closed here keeps the
    published checkpoint authoritative: the job goes red with
    ``checkpoint_regression`` instead of publishing a degraded corpus.
    """

    manifest_path = _checkpoint_manifest(args)
    if manifest_path is None:
        return
    if getattr(args, "allow_checkpoint_regression", False):
        return
    try:
        manifest = json.loads(manifest_path.read_text())
    except (OSError, json.JSONDecodeError):
        return
    regressions = _checkpoint_regressions(manifest, store.counts())
    if regressions:
        raise ValueError(
            "checkpoint_regression: warehouse holds fewer canonical rows than "
            f"the committed checkpoint {manifest_path}: "
            f"{json.dumps(regressions, sort_keys=True)}; recover from the verified "
            "release asset or pass --allow-checkpoint-regression once the loss is "
            "understood"
        )


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
        choices=cli_league_choices(),
    )
    parser.add_argument("--tz", default=None)
    parser.add_argument("--warehouse", default=None)
    parser.add_argument(
        "--localdata",
        default=None,
        help="operator state directory (ELW defaults to localdata/euroleague_w)",
    )
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
        default=None,
        help=(
            "rolling ESPN result lookback (default: 21 days; NBA G League: "
            "56-day date-bounded empirical bootstrap; offline fixtures disable backfill)"
        ),
    )
    parser.add_argument(
        "--history-requests",
        type=int,
        default=None,
        help=(
            "maximum historical scoreboard requests per league run (default "
            "matches --history-days; an explicit lower value caps the lookback)"
        ),
    )
    parser.add_argument(
        "--deep-history-events",
        type=int,
        default=50,
        help="maximum indexed historical events per write-enabled league run",
    )
    parser.add_argument("--skip-results", action="store_true")
    parser.add_argument("--skip-markets", action="store_true")
    parser.add_argument(
        "--skip-book-acquisition",
        action="store_true",
        help=(
            "defer paid named-price providers during pregame; the bounded direct "
            "BetExplorer fallback still fills missing full-game moneylines"
        ),
    )
    parser.add_argument("--acca-legs", type=int, default=3)
    parser.add_argument("--acca-limit", type=int, default=5)
    parser.add_argument("--freshness-hours", type=float, default=12.0)
    parser.add_argument("--book", default=None)
    parser.add_argument(
        "--book-provider",
        default="none",
        choices=[
            "none",
            "theoddsapi",
            "betexplorer",
            "betexplorer_direct",
            "espn",
        ],
    )
    parser.add_argument(
        "--allow-checkpoint-regression",
        action="store_true",
        help="run even though the warehouse holds fewer rows than its checkpoint manifest",
    )
    parser.add_argument("--notify", action="store_true", help="build an operations heartbeat")
    parser.add_argument(
        "--send-notifications",
        action="store_true",
        help="send the heartbeat through configured WhatsApp credentials",
    )
    args = parser.parse_args()
    args.league = parse_league(args.league).value
    spec = league_spec(parse_league(args.league))
    if args.localdata is None:
        # Preserve the established default state locations while new leagues
        # begin isolated from legacy NBA-localdata state.
        args.localdata = (
            "localdata/euroleague_w"
            if args.league == "EUROLEAGUE_W"
            else "localdata"
            if args.league in {"NBA", "EUROLEAGUE", "CBA", "WNBA", "NCAA_M", "NCAA_W"}
            else str(Path("localdata") / spec.state_key)
        )
    if args.warehouse is None:
        args.warehouse = os.environ.get("MAGOLIDE_WAREHOUSE") or (
            str(Path(args.localdata) / "warehouse.duckdb")
            if args.league not in {"NBA", "EUROLEAGUE", "CBA", "WNBA", "NCAA_M", "NCAA_W"}
            else DEFAULT_WAREHOUSE
        )
    if args.date is None:
        args.date = datetime.now(ZoneInfo(ACQUISITION_TZ[args.league])).strftime("%Y-%m-%d")
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
