#!/usr/bin/env python3
"""Materialize canonical DuckDB tables as CSV.GZ research evidence."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from ingestion.warehouse import DEFAULT_CSV_DIR, DEFAULT_WAREHOUSE, export_csv  # noqa: E402


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--warehouse", default=DEFAULT_WAREHOUSE)
    parser.add_argument("--csv-dir", default=DEFAULT_CSV_DIR)
    args = parser.parse_args()
    exported = export_csv(args.warehouse, args.csv_dir)
    print(json.dumps({"warehouse": args.warehouse, "csv": exported}, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
