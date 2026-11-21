#!/usr/bin/env python3
"""Reserve and reconcile the bounded The Odds API free-tier budget.

This script contains no provider credential and never calls The Odds API.  The
``reserve`` command creates Git-durable maximum-cost reservations before any
matrix runner receives a price-provider key.  The ``reconcile`` command records
only redacted numeric quota headers returned by completed runners.
"""

from __future__ import annotations

import argparse
from datetime import datetime, timezone
import json
from pathlib import Path
import sys
from typing import Any

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from ingestion.odds_budget import (  # noqa: E402
    BudgetStateError,
    load_state,
    reconcile_provider_quota,
    reserve,
    write_state,
)


def _now(value: str | None) -> datetime:
    if value is None:
        return datetime.now(timezone.utc).replace(microsecond=0)
    try:
        parsed = datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError as exc:
        raise BudgetStateError("--now must be ISO-8601") from exc
    if parsed.tzinfo is None:
        raise BudgetStateError("--now must include a timezone")
    return parsed.astimezone(timezone.utc)


def _matrix(value: str) -> dict[str, Any]:
    try:
        parsed = json.loads(value)
    except json.JSONDecodeError as exc:
        raise BudgetStateError("--matrix is not JSON") from exc
    if not isinstance(parsed, dict):
        raise BudgetStateError("--matrix is not an object")
    return parsed


def _quota_from_artifacts(root: Path) -> list[dict[str, Any]]:
    """Extract only provider quota headers from uploaded daily summaries."""

    observations: list[dict[str, Any]] = []
    for path in sorted(root.rglob("daily-output.json")):
        try:
            report = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            continue
        if not isinstance(report, dict):
            continue
        quota = (((report.get("commands") or {}).get("pregame") or {}).get("quota") or {}).get(
            "theoddsapi"
        )
        if isinstance(quota, dict):
            observations.append(quota)
    return observations


def reserve_command(args: argparse.Namespace) -> int:
    current = _now(args.now)
    state = load_state(args.state_file, current)
    next_state, permits = reserve(
        state,
        _matrix(args.matrix),
        run_id=args.run_id,
        now=current,
        automated_tick=args.automated_tick,
    )
    write_state(args.state_file, next_state)
    payload = json.dumps(permits, sort_keys=True, separators=(",", ":"))
    if args.output_file:
        Path(args.output_file).write_text(payload + "\n", encoding="utf-8")
    if args.github_output:
        with Path(args.github_output).open("a", encoding="utf-8") as output:
            output.write(f"permits={payload}\n")
    print(
        json.dumps(
            {
                "automated_tick": bool(args.automated_tick),
                "permits_issued": len(permits),
                "state_file": str(args.state_file),
            },
            sort_keys=True,
        )
    )
    return 0


def reconcile_command(args: argparse.Namespace) -> int:
    current = _now(args.now)
    state = load_state(args.state_file, current)
    observations = _quota_from_artifacts(Path(args.artifacts))
    next_state = reconcile_provider_quota(state, observations, now=current)
    write_state(args.state_file, next_state)
    print(
        json.dumps(
            {
                "observations_used": bool(next_state["last_provider_quota"]),
                "state_file": str(args.state_file),
            },
            sort_keys=True,
        )
    )
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest="command", required=True)

    reserve_parser = commands.add_parser("reserve", help="reserve maximum-cost routine permits")
    reserve_parser.add_argument("--state-file", required=True)
    reserve_parser.add_argument("--matrix", required=True, help="gate matrix JSON")
    reserve_parser.add_argument("--run-id", required=True)
    reserve_parser.add_argument("--automated-tick", action="store_true")
    reserve_parser.add_argument("--github-output")
    reserve_parser.add_argument("--output-file")
    reserve_parser.add_argument("--now", help="ISO-8601 test clock")
    reserve_parser.set_defaults(handler=reserve_command)

    reconcile_parser = commands.add_parser("reconcile", help="record completed-job quota headers")
    reconcile_parser.add_argument("--state-file", required=True)
    reconcile_parser.add_argument("--artifacts", required=True)
    reconcile_parser.add_argument("--now", help="ISO-8601 test clock")
    reconcile_parser.set_defaults(handler=reconcile_command)

    args = parser.parse_args()
    try:
        return args.handler(args)
    except BudgetStateError as exc:
        print(json.dumps({"status": "failed", "error": str(exc)}), file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
