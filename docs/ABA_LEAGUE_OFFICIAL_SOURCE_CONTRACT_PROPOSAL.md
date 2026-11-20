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
candidate official ABA League contract using its own reported 2026-10-12
captures. That work is a useful interface-design reference, but it is not a Ma
Golide capture, source-use clearance, or live validation. We did not copy its
source page fixtures or parser implementation. The Ma Golide tests use small,
synthetic HTML examples written only to exercise the contract invariants.

The active adapter targets the official domain `www.aba-liga.com`, with these
externally referenced contract routes:

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
   Calendar names are not used as immutable team IDs. A future date range,
   date-only row, or one-cell unscored/undated pairing without a single tip is
   a non-canonical placeholder: it is skipped with explicit coverage telemetry,
   never assigned an invented timestamp. This check occurs before team-name
   parsing because future bracket placeholders can also have no assigned teams.
3. **Clock semantics:** `CET`, `CEST`, and the source's `Local time` suffix
   are wall-clock values in `Europe/Belgrade`, including DST, before conversion
   to UTC; an optional trailing parenthetical local-time annotation is
   informational only. `Local time` is a displayed source label, not permission
   to infer a missing timestamp. A fixed `+01:00` conversion is rejected by
   tests.
4. **Detail scope and identity:** the match page must echo `match_data_id`,
   `match_data_sez`, and `match_data_lea`; expose two scoped team links with
   distinct official team IDs; repeat the calendar home/away display names; and
   agree with the calendar tip time. All of that remains mandatory for a
   canonical played game. An administrative disposition is not a game record:
   it anchors identity to the scoped event ID and distinct official team-ID pair
   because historical city/sponsor display suffixes can drift while the event
   remains the same. Any unresolved scope/ID/tip contradiction produces a
   coverage gap.
5. **Finality:** a calendar score only identifies a candidate terminal row. A
   canonical played-game `ActualResult` requires the individual match page's
   numeric final and a Q1--Q4 table. Each OT period, if present, is retained as
   combined overtime only after both sides' values are present.
6. **Reconciliation:** regulation plus overtime must exactly add to both final
   scores. There is no final-score-only fallback, derived quarter score, or
   array-order inference.
7. **Administrative dispositions:** a `20:0` or `0:20` is a FIBA default-score
   *candidate*, never a basketball final by itself. The adapter can classify it
   as an event-scoped administrative disposition only after the official detail
   verifies the scoped event and distinct team IDs and either (a) has no period
   cells (an empty retained table shell is equivalent), or (b) expressly states
   in match content that a played match was later registered with an official
   result. A default-score row with valid reconciling periods stays a played
   final; a malformed/contradictory non-empty period table with no explicit
   official-registration statement remains a blocking gap. An unpinned default
   award is never assigned a tip. Administrative classification is history-only:
   discovery and current-result paths continue to fail closed rather than emit a
   tipped default score as a live canonical game. This is generic logic — no
   club, season, or hand-maintained event allowlist is used.
8. **History:** season results are enumerated from the complete calendar in
   ascending `(game_no)` order. The helper returns attempted game numbers and
   the full calendar-final count. A malformed detail stops the cursor at the
   last written game; after a completed current-season scan, a changed calendar
   final count triggers a bounded replay rather than a silent late-result hole.
   A clean exhausted scan also retains a compact per-season receipt in the
   history provider-health cursor: season, calendar terminal count,
   period-reconciled played-final count, verified administrative-disposition
   count, non-canonical schedule-only `not_applicable` row count, zero blocking
   gaps, evidence label, and completion time. Administrative dispositions have
   their own count and do not inflate the schedule-placeholder telemetry.
   Unpinned rows are deliberately excluded from both terminal populations. A receipt is
   withheld and the cursor rewinds to game zero unless
   `calendar_terminal_count == reconciled_final_count +
   administrative_disposition_count`. Administrative records are never written
   as games or used by results/model/ROI paths. For seasons older than the active
   season, the daily publisher additionally merges those count-only receipts
   into `localdata/aba_league/history_receipts.json`. The Git-durable summary
   contains no raw page body, game rows, prices, or ticket evidence. A changed
   historical count conflicts visibly instead of allowing a stale artifact to
   overwrite the prior durable attestation; the still-active season remains in
   the warehouse/provider-health cursor because its calendar total can change.

### Historical-completeness policy

The historical floor remains 2001-02 (`season_end_year=2002`), with two explicit
operator exclusions:

- **2008-09** (`season_end_year=2009`): legacy calendar layout without a
  verified canonical parser contract.
- **2019-20** (`season_end_year=2020`): the season was cancelled by the ABA
  League due to COVID-19, and its official calendar retains an unregistered
  awarded-score row without reconcilable periods. The season-level COVID
  cancellation does not itself validate that individual score as canonical.

The pipeline does not fetch, fabricate, or label either excluded season as
canonical history. Instead it emits an `aba_liga_history_policy / not_applicable`
gap, records `excluded_by_operator_policy` in the bounded history report, and
advances only to the next older season. A terminal run therefore reports
`complete_to_policy_floor_with_exclusions`, not an unqualified complete-to-floor
claim. These exclusions do not cover any additional season. Other awarded or
forfeit rows are handled only by the generic administrative-disposition contract
above; a score-shaped row that does not meet that first-party detail contract
remains a hard failure. Ordinary scored unpinned rows still fail closed.

`tests/test_aba_liga.py` uses synthetic contract examples to cover normal
discovery/result parsing, DST interpretation, scope and team/time contradiction
refusal, regulation and multi-OT reconciliation, generic administrative
outcomes, malformed calendar refusal, and bounded history/completeness-replay
behavior. It is not live-source verification.

### Administrative-outcome evidence boundary

The generic contract was added after direct first-party review of the league's
2020/21 archive. An ABA notice for an unplayed Round 10 fixture says the game
would not take place and records the `0:20` award under the league bylaws
([official notice](https://www.aba-liga.com/news/44251/round-10-game-borac-koper-primorska-will-not-take-place/)).
The current official match archive can also state that a match took place but
was later registered with an official result, with player statistics annulled
([example event](https://www.aba-liga.com/match/63/20/1/Boxscore/q1/1/home/krka-koper-primorska/)).
Direct review on 2026-11-20 also observed that the latter archive page displays
its historical tip as `Local time`, while an ordinary 2020/21 played final
continues to display `CET`; the former's score is administrative despite retained
unofficial statistics. These are page-contract observations only: no raw source
body is retained in this repository.

Those sources establish two *event dispositions*, not a team-specific rule,
financial-status inference, or approval to import a club's history. The adapter
retains neither page body nor player statistics; Git-durable receipts retain
counts only.

## Operational guardrails currently in code

The adapter is registered for normal pipeline discovery, results, pending-final
rechecks, and bounded history. It still fails closed at every page boundary:

- `is_enabled()` is `True`; live retrieval uses the official text pages.
- Hermetic tests inject response text without changing live behavior.
- The evidence string is `ROUTE_CONTRACT_REFERENCED`, not `LIVE_VERIFIED`;
  it explicitly states that source captures are external references and tests
  are synthetic.
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
   plan for raw source bodies versus compact durable coverage receipts. The
   current season receipt survives with provider health in the canonical
   warehouse/cache until its count is stable; completed historical seasons are
   also retained as count-only Git summaries, but that does not grant a basis
   to retain raw page bodies or replace an approved long-term source-retention
   policy; and
6. confirmation that the provider is ABA League (`league_id=1`), rather than
   ABA 2 or another competition.

The current adapter remains limited to canonical schedule/result/history
acquisition while this follow-up is completed. It does not establish ABA
calibration, predictive edge, priced ROI, ticket eligibility, or permission to
reuse raw page bodies beyond the approved retention basis.
