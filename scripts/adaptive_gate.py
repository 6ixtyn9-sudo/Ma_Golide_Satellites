#!/usr/bin/env python3
"""Emit the fail-open adaptive Actions matrix for Ma Golide scheduler ticks.

This is intentionally a read-only schedule/state check.  It does not restore
or mutate caches, canonical warehouses, operator state, ticket cards, or Git.
The workflow restores the latest Git-authoritative operator state before
calling it and consumes its JSON matrix only when a league actually needs the
full pipeline.
"""

from __future__ import annotations

import argparse
from datetime import datetime
import json
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from ingestion.adaptive_gate import evaluate_gate, report_for, utc_now  # noqa: E402


def _parse_now(value: str) -> datetime:
    try:
        parsed = datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError as exc:
        raise argparse.ArgumentTypeError("--now-utc must be ISO 8601") from exc
    if parsed.tzinfo is None:
        raise argparse.ArgumentTypeError("--now-utc must include a timezone")
    return parsed


def _write_github_output(path: Path, report: dict) -> None:
    """Write compact one-line Action outputs without exposing any credentials."""

    matrix = json.dumps(report["matrix"], sort_keys=True, separators=(",", ":"))
    with path.open("a", encoding="utf-8") as handle:
        handle.write(f"run_any={'true' if report['run_any'] else 'false'}\n")
        handle.write(f"matrix={matrix}\n")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--state-root", type=Path, default=Path("localdata"))
    parser.add_argument(
        "--force-all-reason",
        default=None,
        help="fail open every league with this auditable reason",
    )
    parser.add_argument(
        "--now-utc",
        type=_parse_now,
        default=None,
        help="testing/replay clock; defaults to the current UTC time",
    )
    parser.add_argument(
        "--github-output",
        type=Path,
        default=None,
        help="optional GITHUB_OUTPUT-compatible destination",
    )
    args = parser.parse_args()

    now = args.now_utc or utc_now()
    try:
        decisions = evaluate_gate(
            state_root=args.state_root,
            now=now,
            force_all_reason=args.force_all_reason,
        )
    except Exception as exc:  # defensive final boundary: gate errors always fail open
        decisions = evaluate_gate(
            state_root=args.state_root,
            now=now,
            force_all_reason=f"gate_internal_uncertainty: {type(exc).__name__}",
        )
    report = report_for(decisions, generated_utc=now)
    if args.github_output is not None:
        _write_github_output(args.github_output, report)
    print(json.dumps(report, indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
