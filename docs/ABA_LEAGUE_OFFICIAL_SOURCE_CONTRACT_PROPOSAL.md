# ABA League official-source contract

**Status: active canonical schedule/result/history adapter; source-use terms remain unreviewed.**

This records the narrow contract implemented in
`ingestion/providers/aba_liga.py` and wired through ABA League's registry,
pipeline discovery/results/pending-final rechecks, and resumable history
cursor. It does not change the independent pricing contract: BetExplorer
observations remain aggregate research-only and ABA still has no ticket, ROI,
bank, receipt, or execution price lane.

## Why this exists

An independently maintained MIT-licensed adapter repository documented a
candidate official ABA League contract using its own captures from 2026-10-12.
That work is a useful interface-design reference, but it is not Ma Golide's
capture, source-use clearance, or live validation. We did not copy its source
page fixtures or parser implementation. The Ma Golide tests use small,
synthetic HTML examples written only to exercise the proposed invariants.

The canonical source is the official domain `www.aba-liga.com`, with these
contract routes:

| Role | Proposed route | Required scope |
| --- | --- | --- |
| Complete season calendar | `/calendar/{sez}/1/` | `sez = season_end_year - 2001`; ABA League `league_id = 1` |
| Completed-game detail | `/match/{game_no}/{sez}/1/Overview` | `(game_no, sez, league_id)` must echo on page |

For example, canonical 2025-26 maps to `sez=25`, and the stable event namespace
is `aba_liga:25:1:{game_no}`. Game number alone is never globally unique.

## Canonical fields and invariants

The active adapter is deliberately strict and accepts only the following
minimal contract:

1. **Calendar scope:** a selected calendar season must exactly equal
   `{sez}/1/`. Every match link must repeat the same `sez` and `league_id`.
2. **Fixture identity:** the calendar provides the event number, home/away
   display names, final-present indicator, and local wall-clock date/time.
   Calendar names are not used as immutable team IDs.
3. **Clock semantics:** the site display is reported as `CET` year-round. The
   parser reads its wall-clock time in `Europe/Belgrade`, including DST, before
   converting it to UTC. A fixed `+01:00` conversion is rejected by tests.
4. **Detail scope and identity:** the match page must echo `match_data_id`,
   `match_data_sez`, and `match_data_lea`; expose two scoped team links with
   distinct official team IDs; repeat the calendar home/away order; and agree
   with the calendar tip time. Any drift produces a coverage gap rather than a
   name-only canonical game.
5. **Finality:** a calendar score only identifies a candidate final. A canonical
   `ActualResult` requires the individual match page's numeric final and a
   Q1--Q4 table. Each OT period, if present, is retained as combined overtime
   only after both sides' values are present.
6. **Reconciliation:** regulation plus overtime must exactly add to both final
   scores. There is no final-score-only fallback, derived quarter score, or
   array-order inference.
7. **History:** season results are enumerated from the complete calendar in
   ascending `(game_no)` order. The helper returns attempted game numbers and
   stops on a malformed detail page so the pipeline cursor cannot falsely
   advance past an unwritten result.

`tests/test_aba_liga.py` covers normal discovery/result parsing, DST, scope and
team/time contradiction refusal, period-total refusal, malformed calendar
refusal, and bounded history behavior.

## Operational guardrails currently in code

The adapter is registered for normal pipeline discovery, results, pending-final
rechecks, and bounded history. It still fails closed at every page boundary:

- `is_enabled()` is `True`; live retrieval uses the official text pages.
- Hermetic tests inject response text without changing live behavior.
- The evidence string is `CONTRACT_CAPTURE_REFERENCED`, not `LIVE_VERIFIED`.
- The module acquires no price data and cannot affect the existing BetExplorer
  aggregate-consensus research lane.
- The adaptive gate fails open rather than duplicating a whole-season HTML
  retrieval in a frequent scheduler probe.

## Source-use and capture follow-up

A reviewer must still record a source-use decision and a compact first-party
capture dossier. At minimum it should contain:

1. the official site's relevant access, automation, storage, and redistribution
   terms as reviewed on the capture date;
2. one current season calendar response and two detail responses (a regulation
   final and an overtime final if available), retained only as allowed;
3. an upcoming fixture example, postponed/cancelled or correction behavior if
   available, and proof that no final is fabricated when periods are absent;
4. a direct check that the selected season, game ID, team IDs, home/away order,
   timestamp, final, and periods agree across the documented pages;
5. a documented historical floor, page/rate limits, and acceptable retention
   plan for raw source bodies versus compact durable coverage receipts; and
6. confirmation that the provider is ABA League (`league_id=1`), rather than
   ABA 2 or another competition.

The current adapter remains limited to canonical schedule/result/history
acquisition while this follow-up is completed. It does not establish ABA
calibration, predictive edge, priced ROI, ticket eligibility, or permission to
reuse raw page bodies beyond the approved retention basis.
