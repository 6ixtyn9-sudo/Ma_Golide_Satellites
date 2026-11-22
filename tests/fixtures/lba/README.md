# LBA fixtures

Provenance, byte counts and hashes for every file here live in
[`../CAPTURE-PROVENANCE.md`](../CAPTURE-PROVENANCE.md). In short:

- `championships_2026.json` (867 B) and `calendar_602.json` (11636 B) are
  **byte-exact captures** of the legabasket.it JSON, re-hashed by
  `tests/test_capture_fidelity.py`. The calendar is the league's current focus
  day at capture time: eight rows across 2026-11-21 and 2026-11-22, with
  `filters.days` marking round 9 as focused rather than any single date, and a
  status mix of `game_status` `"2"` (final, 3 rows), `"1"` (live, 4 rows) and
  `"0"` (not started, 1 row). Two of the live rows already publish real scores,
  which is why status is read from the code and never inferred from numbers.
- `pbp_25426.json` (regulation, 22-20/26-30/24-22/28-14 → 100-86) and
  `pbp_25422_ot.json` (23-29/41-50/62-69/87-87/92-106 → overtime 5-19) hold the
  **verbatim captured period-end actions**; only the `{"pbp": {"actions": []}}`
  envelope and the re-serialisation are this repository's. The full bodies are
  314 KB each and were never transmissible.
- `pbp_synthetic_noncontiguous_ot.json` is **synthetic**, not a capture: four
  regulation anchors plus a lone period 6, used to prove non-contiguous
  overtime is refused rather than filled in.
