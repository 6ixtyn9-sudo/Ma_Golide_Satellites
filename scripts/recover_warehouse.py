#!/usr/bin/env python3
"""Recover a new canonical DuckDB file from the six CSV.GZ tables.

The destination must not exist. Recovery happens at a temporary path and is
atomically moved into place only after every table restores successfully.
"""

from __future__ import annotations

import argparse
import json
import os
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from ingestion.duckdb_store import DuckDBStore  # noqa: E402
from ingestion.warehouse import DEFAULT_CSV_DIR, DEFAULT_WAREHOUSE  # noqa: E402


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--warehouse",
        default=DEFAULT_WAREHOUSE,
        help="new destination; an existing file is never overwritten",
    )
    parser.add_argument("--csv-dir", default=DEFAULT_CSV_DIR)
    args = parser.parse_args()
    if os.environ.get("ALLOW_MAGOLIDE_WRITES", "").lower() != "true":
        parser.error("recovery requires ALLOW_MAGOLIDE_WRITES=true")
    destination = Path(args.warehouse)
    if destination.exists():
        parser.error(f"refusing to overwrite existing warehouse: {destination}")
    destination.parent.mkdir(parents=True, exist_ok=True)
    temporary = destination.with_name(f".{destination.name}.recovering")
    if temporary.exists():
        parser.error(f"temporary recovery path already exists: {temporary}")

    store = None
    try:
        store = DuckDBStore(str(temporary), persist=True)
        counts = store.restore_csv(args.csv_dir)
        store.close()
        store = None
        os.replace(temporary, destination)
    except Exception:
        if store is not None:
            store.close()
        if temporary.exists():
            temporary.unlink()
        raise
    print(json.dumps({"warehouse": str(destination), "restored": counts}, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
