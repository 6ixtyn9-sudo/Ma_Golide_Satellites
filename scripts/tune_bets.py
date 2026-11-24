#!/usr/bin/env python3
"""Run the config tuner over settled results and apply what it finds.

The retired system wrote its tuning advice to ``Config_Tuner_Suggestions`` and
waited for a human. You asked for the advice to be applied, so this script runs the
same three analyses and writes the resulting overrides to
``localdata/magolide_bets.tuned.json``, which ``scripts/research_standings.py`` loads
on every card. Every applied change is recorded in the config ledger with the
before and after value, so the card can still say which configuration it ran under.

Input is a JSON array of settled bets::

    [
      {"bet_type": "BANKER", "league": "LITHUANIA_NKL",
       "confidence": 72.5, "won": true},
      {"bet_type": "FT_OU", "league": "GREECE_ESAK", "confidence": 61.0,
       "won": false, "direction": "OVER"}
    ]

``won: null`` is a push and is scored as half a win, exactly as the sheet did.

Or read the ledgers the lanes already keep -- dated pick files and the rolling
audit -- and let this script grade nothing, only count what is already settled::

    python scripts/tune_bets.py --from-ledgers localdata/<lane>
    python scripts/tune_bets.py --from-ledgers localdata --dry-run
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from forecasting import settlement  # noqa: E402
from forecasting.config_ledger import DEFAULT_LEDGER_PATH  # noqa: E402
from forecasting.config_tuner import (  # noqa: E402
    TUNED_PATH,
    analyse,
    apply,
    graded_from_rows,
    load_tuned,
)


def _load_results(path: Path) -> list:
    payload = json.loads(path.read_text())
    if isinstance(payload, dict):
        payload = payload.get("bets", [])
    if not isinstance(payload, list):
        raise SystemExit(f"{path}: expected a JSON array of settled bets")
    return payload


def report_lines(report, applied: dict, config, summary=None) -> list:
    """The human-readable half of the run: what was seen, what was changed."""

    lines = [
        f"settled bets: {report.total}",
        f"overall Brier: {'n/a' if report.overall_brier is None else round(report.overall_brier, 4)}",
    ]
    if report.leagues:
        worst = sorted(report.leagues.values(), key=lambda item: item.win_rate)[:3]
        lines.append(
            "leagues: "
            + ", ".join(f"{entry.league} {entry.win_rate:.1%} ({entry.count})" for entry in worst)
        )
    if not applied:
        lines.append("no change: " + ("; ".join(report.reasons) or "nothing to tune"))
    else:
        lines.append("applied:")
        for key, (before, after) in sorted(applied.items()):
            lines.append(f"  {key}: {before} -> {after}")
        lines.append(f"config stamp: {config.stamp_id}")
    if report.reasons and applied:
        lines.append("why: " + "; ".join(report.reasons))
    if summary is not None:
        lines.append(f"ledgers read: {summary.files} file(s)")
        lines.append(f"not counted: {summary.skip_line()}")
    return lines


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    source = parser.add_mutually_exclusive_group(required=True)
    source.add_argument("--results", help="JSON array of settled bets")
    source.add_argument(
        "--from-ledgers",
        nargs="+",
        metavar="PATH",
        help="pick ledger files or lane directories to read settled picks from",
    )
    parser.add_argument("--tuned-path", default=str(TUNED_PATH))
    parser.add_argument("--ledger-path", default=str(DEFAULT_LEDGER_PATH))
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="analyse and print, write neither the tuned file nor the ledger",
    )
    args = parser.parse_args(argv)

    summary = None
    if args.from_ledgers:
        summary = settlement.collect(*args.from_ledgers)
        bets = list(summary.bets)
    else:
        results = Path(args.results)
        if not results.exists():
            print(f"no results file at {results}: nothing to tune", file=sys.stderr)
            return 1
        bets = graded_from_rows(_load_results(results))

    report = analyse(bets)
    config = load_tuned(tuned_path=args.tuned_path)
    tuned, stamp, applied = apply(
        report,
        config,
        tuned_path=args.tuned_path,
        ledger_path=args.ledger_path,
        write=not args.dry_run,
    )

    for line in report_lines(report, applied, tuned, summary):
        print(line)
    if args.dry_run:
        print("dry run: nothing written")
    else:
        print(f"tuned config: {args.tuned_path} (stamp {stamp.stamp_id})")
        print(f"ledger: {args.ledger_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
