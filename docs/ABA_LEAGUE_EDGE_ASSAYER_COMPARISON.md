# ABA League research comparison and implementation decision

**Scope:** `ABA_LEAGUE` only. This records the research decision and its
subsequent narrow implementation. It compares Ma Golide's local Edge Factory
review with the independently supplied Z.ai `aba_league_design_review.md` and
`aba_league_implementation_plan.md`.

**Decision:** adopt the design review's evidence-first direction, with the
corrections below. Do **not** implement the five-PR Z.ai implementation plan as
written. In particular, do not register OddsPortal as an ABA schedule/result
source until its contract is demonstrated and reviewed. The existing
BetExplorer adapter may be exercised only as a probationary aggregate research
observation; it is not a named-book, ROI, ticket, or canonical-source lane.

This decision and its constrained implementation make no claim about ABA
predictive edge, calibration, ticket eligibility, return on investment, price
availability, live source validation, or historical coverage.

**Implementation update (2026-11-18):** the adapter targeting the documented
official-site route is active for canonical ABA schedule/result/history
acquisition; its route contract is an external reference, not a retained Ma
Golide first-party response. The implementation and remaining
source-use/capture follow-up are in
[`ABA_LEAGUE_OFFICIAL_SOURCE_CONTRACT_PROPOSAL.md`](ABA_LEAGUE_OFFICIAL_SOURCE_CONTRACT_PROPOSAL.md).
This change does not relax price, ticket, ROI, bank, receipt, or execution
restrictions: BetExplorer remains an aggregate research-only lane.

## 1. Executive comparison

| Topic | Local Edge/Ma Golide review | Z.ai design review | Z.ai implementation plan | Decision |
| --- | --- | --- | --- | --- |
| Scope | ABA only; no broad multi-league work | ABA only | ABA only | Retain ABA-only scope. |
| Evidence ordering | Source contract -> canonical history -> historical replay/research -> price/ROI later | Same ordering | Reverses it by making aggregate prices and an unvalidated source registry change PR1 | Use the evidence-first ordering. |
| Canonical results | A compliant, date-bounded result source is a prerequisite | Same | Calls an unverified OddsPortal archive canonical | Do not approve that source designation without proof and a source-permission decision. |
| Empirical forecasts | Reconstruct them chronologically from prior reconciled finals; do not manufacture provider observations | Correctly identifies existing replay capability | Correct in prose, but later treats the provider-prediction assay as if it can assess reconstructed rows | Keep reconstruction separate from `predictions` and build any research projection from the market-service replay. |
| Prices and ROI | Exact historical named-book prices are a later prerequisite for ROI | Same | Adds aggregate BetExplorer prices first and describes later historical price evidence without implementing it | Aggregate quotes, if ever validated, stay research-only. They are not a prerequisite for history/calibration and cannot support ROI. |
| Assayer | Existing Wilson/shrinkage safeguards are useful but not an edge certificate | Same | Adds a write guard to the wrong assay contract and promises metrics that its proposed code does not produce | Preserve existing provider assay; design a separate, explicit empirical-replay research report after history exists. |
| Storage | Use the existing per-league warehouse and compact lifecycle state; avoid sidecars | Same intent | Calls CSV.GZ Git-durable | Correct this: CSV.GZ and DuckDB are ignored by Git and need an explicit cache/artifact retention plan. |

## 2. Findings that are confirmed and transferable

The following parts of the two research outputs agree with the checked code and
are adopted.

1. **ABA is a typed, isolated canonical lane.** `League.ABA_LEAGUE` has state
   key `aba_league`, timezone `Europe/Belgrade`, and the official
   `aba_liga` schedule/result/history adapter in `ingestion/league_registry.py`.
   The adapter refuses page scope, identity, clock, finality, and reconciliation
   drift rather than treating an empty or malformed page as a successful slate.

2. **The system already has the right canonical storage boundary.** An ABA
   warehouse path resolves to `localdata/aba_league/warehouse.duckdb`; canonical
   games, results, aliases, predictions, books, provider health, and pipeline
   reports already have a common store interface. A future ABA importer should
   extend those components, not introduce a per-game JSON family.

3. **The existing forecast service is the correct historical-research base.**
   `MarketForecastService.build_report` replays each reconciled final at its own
   scheduled tip using earlier available finals. Its empirical team-history
   surface is deliberately reconstructible without a stored provider prediction.
   This is appropriate for calibration research and must not be presented as a
   captured external pre-tip prediction.

4. **The existing `ingestion.mining` assay has a narrower job.** It assesses
   frozen provider predictions only: a prediction must be pre-tip, frozen,
   coherent, and paired with a reconciled non-tied final. It is useful for an
   external ABA prediction provider if one is later approved, but it cannot
   directly assay reconstructed team-history forecasts today.

5. **The useful Edge Factory patterns are methodological rather than
   structural.** Retain loss-aware history recovery, immutable identity checks,
   conservative sample gates, Wilson lower bounds, chronological/decay
   diagnostics, and a regression-to-zero protection for durable research
   outputs. Do not copy soccer-provider taxonomies, weak date/team-name joins,
   source-per-view SQL, raw capture families, or price/ROI conclusions without
   exact prices.

6. **The existing sample safeguards remain unchanged.** The descriptive
   prediction assay uses one-sided Wilson (`z=1.645`), Beta(2,2) shrinkage,
   segment suppression below 10 rows, and a 30-row minimum-sample marker. The
   forecast service separately uses chronological calibration profiles and
   probability metrics. None of these certify an edge.

## 3. Evidence corrections required before any implementation

### 3.1 The new source claims are leads, not verified contracts

The Z.ai implementation plan asserts that BetExplorer exposes the desired ABA
fixtures table and that OddsPortal has an 18-season ABA archive with final and
period data. No captured response, endpoint contract, source terms decision,
stable event-ID demonstration, or official cross-check is in this repository.
Those assertions therefore remain **unvalidated leads** here.

This environment cannot independently fetch either external site under the
current outbound-network allowlist. That is not a reason to relax the
requirement: a candidate must remain disabled until captured evidence is
reviewed in a normal permitted research environment.

A non-official odds/archive site must not be labelled the canonical result
source merely because it has historical scores. At most it may later serve as a
cross-check or a research-price source, subject to permission and identity
validation. The canonical schedule/result source should be official or an
explicitly approved, demonstrably reliable replacement.

### 3.2 BetExplorer evidence does not transfer from one competition to ABA

`ingestion/providers/betexplorer.py` is live-verified only for the existing
EuroLeague and CBA URLs. An ABA route inherits a parser and an unreviewed reuse
position, not an ABA verification. It can therefore be exercised only as a
probationary aggregate research observation; it does not supply canonical
fixtures/results or named-book evidence. Capture and test an ABA page, verify
its competition identity and clock semantics against the selected canonical
schedule, and make a source-use policy decision before promoting its evidence
label or depending on it operationally.

The adapter provides only an observed, aggregate average moneyline pair before
tip. It has no historical archive and no named-book attribution. It cannot
justify an ABA ROI figure, strict ticket, receipt, bank, or execution claim.

### 3.3 CSV.GZ is recoverable materialization, not Git-durable history

The Z.ai review and plan describe CSV.GZ exports as Git-durable. That is
incorrect for this repository: `.gitignore` ignores `localdata/*` except for a
narrow compact-card/lifecycle allowlist, and it also ignores `*.duckdb`.
`docs/ARCHITECTURE.md` describes CSV.GZ as an export/recovery materialization
that must be retained through the cache/artifact policy.

Before a meaningful ABA backfill, decide how full CSV.GZ exports, source
captures, and completeness receipts will be retained outside Git. The compact
Git-durable record should be a source-contract/completeness summary, not raw
multi-season CSV data.

### 3.4 Aggregate price acquisition is not the first dependency

Canonical schedule/results and sufficient reconciled history are needed before
empirical walk-forward research. Price acquisition is neither a substitute for
those facts nor a shortcut to a priced assay. It should remain a separate,
later source contract.

## 4. Why the proposed implementation sequence is not executable as written

The plan is valuable as a list of hypotheses, but its code blocks are not ready
to apply. These are material issues, not style preferences.

1. **PR1 would falsely promote an unvalidated source.** It changes ABA's
   registry schedule, results, and gate sources from `source_unavailable` to
   `oddsportal` before an adapter, fixtures, source approval, or verified
   contract exists. Registry source values are operational contracts, not
   candidate names.

2. **The registry flip would break the current gate's special unavailable-lane
   semantics.** `adaptive_gate.fetch_schedule` supports `espn`, `cba`,
   `fiba_elw`, and the deliberately fail-open EuroLeague condition. It has no
   `oddsportal` route. The proposed registry state would therefore produce a
   generic uncertain gate, while the existing explicit coverage-gap marker
   contract specifically expects `source_unavailable`.

3. **The proposed OddsPortal parser cannot run its own fixture tests.** Its
   `_fetch_text` calls `_require_enabled()` before accepting injected fixture
   text, while `is_enabled()` is hard-coded to `False`. The second and third
   displayed tests call `fetch_results`, so they would raise before parsing.
   A fixtures-first parser needs a separately callable parse function or an
   explicit, test-only contract mode; it must never make a disabled live
   provider indistinguishable from a disabled parser test.

4. **The candidate adapter invents an HTML contract.** Its regexes, date field,
   fixture/results URL usage, event-id format, team attributes, and period-score
   field are assumptions. No fixture supports them. `discover_games` also reads
   a results URL for future fixtures. That is precisely what Phase 0 is meant to
   discover rather than hard-code.

5. **The proposed day-walk is not an 18-season recovery design.** Re-fetching
   a date-filtered result page for `history_request_budget` recent days neither
   locates prior seasons nor persists a season/page cursor. A real history
   importer must use the candidate source's demonstrated archive pagination or
   season identifiers and provide a bounded, resumable cursor.

6. **The plan omits result-path integration.** A supported source needs coherent
   handling in `_discover_games`, `_fetch_finals`, pending-final rechecks, and
   bounded historical backfill. The proposed snippets touch discovery and
   backfill but leave `_fetch_finals` and `_recheck_pending_finals` unsupported
   for ABA.

7. **The proposed pricing tests do not test the registry boundary correctly.**
   `betexplorer.fetch_odds` returns adapter-created snapshots; registry evidence
   stamping occurs in the pipeline. In addition, the displayed import
   `from ingestion import betexplorer` is not supplied by `ingestion.__init__`.
   The stated test must use `ingestion.providers` and exercise a pipeline path
   to prove registry evidence axes.

8. **The unpriced-forecast PR is more than one membership line.** Adding ABA to
   `_UNPRICED_EMPIRICAL_LEAGUES` would set part of the acca-policy behavior, but
   `MarketForecastService` also has explicit capability labels for the two
   current leagues. The proposed `Forecast.to_dict()` assertion for
   `pricing_status` cannot pass: that field does not exist. Tests should assert
   the report-level capability and `accas.pricing_evidence_status`, along with
   absent/ineligible price evidence.

9. **The assay PR targets the wrong data model and wrong function contract.**
   `mine_prediction_assay` returns `sample` and `edges`, not `outcome_rows`, and
   does not receive reconstructed empirical forecasts. Current `write_assay`
   has signature `write_assay(path, report) -> str`; the proposed code reverses
   the arguments and changes its return type without reconciling existing
   callers/tests. Its zero-row guard therefore would not protect the present
   report shape and would not create Brier, log-loss, folds, or provenance rows.

10. **Historical aggregate prices are promised but not parsed or stored.** The
    proposed OddsPortal adapter returns `Game` rows only. No `BookSnapshot`
    parsing, source timestamps, exact market identity, or historical price
    admission logic exists in the displayed code. It cannot support the later
    price claims it describes.

## 5. Approved ABA path

### Phase 0 — source evidence and retention follow-up

The `aba_liga` adapter targeting the externally referenced official route is
active for canonical discovery, reconciled results, pending-final rechecks, and
resumable season/game-number history. The remaining operational follow-up is
compact: retain allowed current calendar and detail captures; record source
automation/storage/reuse terms; verify postponement/correction behavior and
historical limits; and retain a durable completeness receipt outside Git. Do
not commit raw archives, browser exports, or price history.

### Phase 1 — canonical-history operation

The active path covers scheduled-game discovery, detail-page team identity,
period reconciliation, coverage gaps, provider-health cursors, and existing
per-league DuckDB/CSV materialization. A malformed result stops the cursor at
the last successful game, so it is retried instead of skipped; a changed
current-season calendar-final count triggers bounded replay from game zero.
Synthetic hermetic tests cover DST, scope/identity contradictions,
missing/reconciled periods, metadata-only tip corrections, and cursor/completeness
behavior. These facts establish implementation behavior only; they do not
establish live-source validation, model calibration, prediction quality, price
permission, ROI, or ticket eligibility.

### Phase 2 — transparent unpriced empirical forecasts

Once sufficient ABA results are actually in the warehouse, use the existing
team-history reconstruction. If ABA is added to the explicit unpriced empirical
policy set, update every relevant report-level policy label and test it using
realistic fixture-store history; do not introduce a fake `pricing_status` field
or add reconstructed rows to the `predictions` table.

This phase remains ticket-ineligible in the absence of an exact named-book
contract. Aggregate consensus observations, if separately approved later, stay
out of strict tickets, receipts, bank, ROI, and execution.

### Phase 3 — empirical-replay research report

Before adding a new assay module, define a compact projection of the existing
market-service replay. It must clearly state `reconstructed_from_prior_finals`
and include model/config/calibration identity, chronological folds, Brier score,
log loss, Wilson/shrinkage summaries, sample counts, exclusions, and the
reconciled-result source.

This is distinct from the provider-prediction `mine_prediction_assay` and must
not make reconstructed forecasts appear to have been observed or frozen before
tip. A regression-to-zero protection can be added only after the report schema,
retention location, and existing-report comparison key are specified correctly.

### Phase 4 — optional price/ROI research

This phase requires a separate approved source that provides exact, timestamped,
pre-tip, named-book prices for the selected ABA market. It must preserve book,
market, period, side, observation time, identity resolution, and all three
registry evidence axes. ROI can then use only the priced, settled denominator.
Aggregate average/closing prices never meet this bar.

## 6. Assayer admission policy

An ABA empirical research report may emit descriptive calibration metrics only
for reconciled canonical finals replayed with information available at each
historical tip. It may emit a segment-level research verdict only after the
existing sample gates and chronological-fold requirements are met. Its output
must remain `research_only` and `certified: false`.

Any future price/ROI analysis additionally requires an exact historical
named-book price for the exact selected leg. Without it, ROI is `null` with an
explicit missing-price reason, never zero and never estimated from a consensus
quote.

No phase authorizes automatic ticket activation, a bet slip, edge certification,
or modification of the existing named-book evidence policy.

## 7. Remaining operator decisions

1. What source-use and raw-capture retention basis applies to the active
   official ABA League adapter?
2. What archive recovery policy retains the ignored DuckDB/CSV source data and
   compact completeness receipts outside Git?
3. What minimum season depth and competition phases are in scope before any
   ABA empirical report is useful for research?
4. Is `moneyline` full-game plus H1 the initial research scope, or is
   full-game moneyline alone preferred until result/period coverage is proven?
5. Should ABA remain watchlist/research-only until a named-book price contract
   exists, rather than publishing its unpriced forecasts into the ticket-card
   lifecycle?

## 8. Recommended next action

Operate and observe the active canonical ABA path, record approved current
captures/retention terms, and let reconciled history accumulate. Only then
consider the separately gated empirical-replay report. Exact historical
named-book prices remain a prerequisite for any ROI or ticket work.
