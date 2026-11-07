# Documentation index

Start with the repository [README](../README.md) for setup and the canonical
market list.

| Document | Purpose |
| --- | --- |
| [ARCHITECTURE.md](ARCHITECTURE.md) | Headless data flow, package boundaries, DuckDB schema, freezes, exports, and orchestration |
| [CONTRACT.md](CONTRACT.md) | Exact evidence, timing, market, walk-forward, calibration, and acca rules |
| [OPERATIONS.md](OPERATIONS.md) | Dry runs, deliberate writes, scheduling, SQL, retention, notification, export, and recovery commands |
| [PROVIDERS.md](PROVIDERS.md) | Implemented provider coverage, authentication, identity, and evidence status |
| [RESEARCH.md](RESEARCH.md) | Provider assay versus market harness, metrics, dependence, and interpretation |
| [VERIFICATION.md](VERIFICATION.md) | Executed checks, fixture rehearsal results, and remaining unverified limitations |
| [CONTRIBUTING.md](CONTRIBUTING.md) | Required checks and invariants for provider, model, and schema changes |

The documents describe the current headless implementation only. Historical
spreadsheet applications, worksheet contracts, and external web-service
deployment notes were intentionally removed rather than retained as alternate
runtime paths.
