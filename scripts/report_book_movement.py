#!/usr/bin/env python3
"""Report observed book movement from immutable canonical ledgers only."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from ingestion.book_movement import observed_book_movement  # noqa: E402


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--ledger-dir", default="localdata/ledgers")
    parser.add_argument("--league", default=None)
    parser.add_argument("--date", default=None, help="target local slate date YYYYMMDD")
    parser.add_argument("--output", default=None, help="optional derived JSON output")
    args = parser.parse_args()
    directory = Path(args.ledger_dir)
    paths = sorted(directory.glob("*.json")) if directory.exists() else []
    report = observed_book_movement(paths, league=args.league, target_date=args.date)
    encoded = json.dumps(report, indent=2, sort_keys=True)
    if args.output:
        output = Path(args.output)
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(encoded + "\n", encoding="utf-8")
    print(encoded)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
