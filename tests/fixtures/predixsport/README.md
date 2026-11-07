# Predixsport fixtures — DOCUMENTATION_ONLY

These payloads are **SYNTHETIC**, authored on 2026-11-06 from Predixsport's
public documentation / OpenAPI description (`https://api.predixsport.com/docs`,
`https://api.predixsport.com/openapi.json`).

Unauthenticated baseline calls returned `Missing Authorization header`; no
authenticated live response has ever been observed, so the real response
schema is unconfirmed. These fixtures exist only to test the adapter's
parsing, derivation and defensive-skip logic against the *documented* schema.

Files:
- `upcoming_nba.json` — upcoming-list rows (probabilities + margin, no
  totals; one deliberately malformed entry to exercise defensive skipping).
- `detail_match.json` — match detail with `predicted_total_points` 228.4 and
  `predicted_point_diff` **-6.2** (HOME minus AWAY; away favoured):
  home = (228.4 - 6.2) / 2 = 111.1, away = (228.4 + 6.2) / 2 = 117.3.
- `detail_disagreement.json` — independent-model disagreement fixture: the
  probability model favours away (0.58) while the margin model favours home
  (+6.2). `Pred` must be 2 (from probability) while `expected_margin` stays
  positive; both preserved, never reconciled.

Evidence label: **DOCUMENTATION_ONLY** (never `LIVE_VERIFIED` or
`FIXTURE_VERIFIED` from real data). Upgrade the label only after a real-key
authenticated live smoke test is run and its response recorded.
