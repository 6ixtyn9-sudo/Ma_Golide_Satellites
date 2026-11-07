# Contributor guide

## Required checks

```bash
python -m pip install -e '.[dev]'
python -m pytest -q
ruff check .
```

For a reproducible environment with `uv`:

```bash
uv lock
uv sync --extra dev
uv run pytest -q
uv run ruff check .
```

Do not weaken a test to preserve a removed interface. Update callers and tests
to canonical behavior, and add a regression when a review finds a leakage,
identity, reconciliation, or write-guard edge.

## Offline-first tests

Tests must not require credentials, network access, a web application, or
external writes. Committed fixtures should document whether they are captured
live evidence or synthetic examples. Live smoke tests stay deselected by
default and must use an explicit opt-in.

Use fixed UTC clocks for acquisition, freezing, walk-forward, retention, and
notification tests. Assert parsed-time behavior at boundaries, especially
strictly before tip and the four-hour result-availability delay.

Run a complete fixture rehearsal after changing orchestration:

```bash
python scripts/daily.py \
  --auto-once --date 2026-11-07 --offline-fixture \
  --warehouse /tmp/magolide-review/warehouse.duckdb \
  --localdata /tmp/magolide-review
```

Confirm no DuckDB/CSV canonical files appear in a dry run. Keep generated
state out of Git.

## Adding or changing a provider

1. Normalize into `Game`, `Prediction`, `BookSnapshot`, `ActualResult`, or
   `ProviderHealth`; do not return presentation dictionaries.
2. Reconcile provider IDs and team orientation explicitly. Never guess a side
   from array position when IDs contradict it.
3. Parse timestamps as real timezone-aware values and retain source revision
   and observation times.
4. Preserve exact market and period. Q1, H1, full game, and overtime are not
   interchangeable.
5. Preserve missing values. Do not synthesize a line, draw price, score,
   quarter, or bookmaker attribution.
6. Redact credentials and record structured coverage gaps/quota evidence.
7. Add fixture tests for valid, malformed, stale, wrong-team, wrong-period,
   and post-tip payloads.

Documentation-only endpoint behavior must remain labelled as such until a
real authenticated response verifies it.

## Changing forecasting behavior

A valid change must preserve all of these invariants:

- historical rows use frozen pre-tip evidence and reconciled finals only;
- each simulated round excludes future and unavailable results;
- every residual baseline is itself history-only;
- books are deduplicated for calibration so multi-book games do not dominate;
- expected totals never become bookmaker lines;
- exact ties do not produce arbitrary winners;
- current acca legs require exact visible pre-tip prices and expose freeze
  state, while every historical training row requires frozen evidence;
- same-game legs remain prohibited unless a real dependence model replaces
  the current explicit policy;
- rankings remain probability-ranked estimates, never guarantees.

When changing probability adjustment, add a synthetic walk-forward example
where the correct past-only result can be calculated independently. For acca
search changes, retain the brute-force small-case equivalence test.

## Storage and schema changes

Update `ingestion/models.py`, DuckDB typed columns/payload round trips, CSV
export/recovery, SQL views, and schema migration together. Add a migration
test using an older shape. Never drop a canonical table merely to avoid
writing a migration.

Frozen writes must remain immutable. Exact replay may be idempotent; changed
content must fail. Dry runs must not create or alter the configured warehouse.
New external side effects require a separate explicit guard and fixture-send
protection.

## Project boundaries

Do not reintroduce spreadsheet stores, worksheet registries, Apps Script,
Flask, Supabase, or JSON-file persistence. Derived JSON reports are allowed;
they are not canonical state. Do not add a second scheduler alongside
`scripts/daily.py`.

Avoid claims of live coverage, accuracy, profitability, or certification
unless the corresponding evidence actually exists. Record verification
commands and remaining limitations in `docs/VERIFICATION.md`.
