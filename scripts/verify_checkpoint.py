#!/usr/bin/env python3
"""Compare a warehouse's canonical row counts with its committed checkpoint manifest.

A warm cache is trusted only while it still holds at least the rows the committed
checkpoint promises. Exiting 1 with ``checkpoint_regression`` means the warehouse
is thinner than the checkpoint it is supposed to represent, so the caller must
discard it and recover from the verified release asset instead of running on a
corpus that can no longer see its own history.
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from ingestion.duckdb_store import DuckDBStore  # noqa: E402
from scripts.release_archive import (  # noqa: E402
    REGRESSION_GUARDED_TABLES,
    _checkpoint_regressions,
)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--warehouse", required=True)
    parser.add_argument("--manifest", required=True)
    parser.add_argument("--state-key", default=None)
    args = parser.parse_args()

    warehouse = Path(args.warehouse)
    manifest_path = Path(args.manifest)
    report: dict[str, object] = {
        "warehouse": str(warehouse),
        "manifest": str(manifest_path),
        "state_key": args.state_key,
    }
    if not manifest_path.is_file():
        report["status"] = "no_manifest"
        print(json.dumps(report, indent=2, sort_keys=True))
        return 0
    try:
        manifest = json.loads(manifest_path.read_text())
    except (OSError, json.JSONDecodeError) as exc:
        raise SystemExit(f"invalid checkpoint manifest: {manifest_path}") from exc
    committed = manifest.get("canonical_row_counts")
    if not isinstance(committed, dict) or not committed:
        # Older manifests carry no counts; the next build records them.
        report["status"] = "no_manifest_counts"
        print(json.dumps(report, indent=2, sort_keys=True))
        return 0
    if not warehouse.is_file():
        report["status"] = "checkpoint_regression"
        report["reason"] = "warehouse_missing"
        report["regressions"] = {
            table: {"committed": int(committed[table]), "new": 0}
            for table in REGRESSION_GUARDED_TABLES
            if isinstance(committed.get(table), int) and int(committed[table]) > 0
        }
        print(json.dumps(report, indent=2, sort_keys=True))
        return 1

    store = DuckDBStore(str(warehouse), persist=False)
    try:
        counts = store.counts()
    finally:
        store.close()
    regressions = _checkpoint_regressions(manifest, counts)
    report["status"] = "checkpoint_regression" if regressions else "ok"
    report["counts"] = counts
    report["committed"] = committed
    if regressions:
        report["regressions"] = regressions
    print(json.dumps(report, indent=2, sort_keys=True))
    return 1 if regressions else 0


if __name__ == "__main__":
    raise SystemExit(main())
