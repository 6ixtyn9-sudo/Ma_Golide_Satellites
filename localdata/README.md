# Local runtime state

This directory is restored and saved by the scheduled workflow. Runtime files
are ignored by Git; this README is the exception.

Expected files after a write-enabled daily run:

- `warehouse.duckdb` — canonical facts and bounded operational telemetry.
- `ledgers/<kind>_<league>_<date>_<sha256>.json` — immutable acquisition evidence.
- `assay_<LEAGUE>.json` — derived, research-only mining reports.
- `csv/*.csv.gz` — six canonical CSV.GZ tables.
- `notification_ledger.json` — content hashes for heartbeat dedupe.

Use an isolated directory for fixtures or local checks. `warehouse.duckdb.tmp`
is an atomic-build staging file and is never a second canonical store. Canonical
writes require `ALLOW_MAGOLIDE_WRITES=true`; fixture persistence additionally
requires `ALLOW_FIXTURE_WRITES=true` and an isolated path.

Never put provider credentials, service-account JSON, notification tokens, or
other secrets in this directory. Use the Actions repository/environment secret
store. Do not commit the DuckDB database, exports, ledgers, or delivery hashes.
