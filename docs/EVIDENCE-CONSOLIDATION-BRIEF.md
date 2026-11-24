# Hand-off brief: consolidate every piece of project evidence into one file

Paste this whole document to a fresh agent. It assumes **no** prior conversation.

---

## Your job

You are taking over a bootstrapped basketball forecasting/settlement project
(`6ixtyn9-sudo/Ma_Golide_Satellites`, branch off `main`, all work via PRs).
Nothing has been read into your context, and the evidence for a week of work is
scattered across GitHub PRs, run logs pasted by the operator, three reports from
an external prober, and half a dozen repo documents.

**Your single deliverable is one file: `docs/EVIDENCE.md`.**

It must let a careful stranger answer, without opening anything else: *what do we
actually know, how do we know it, when was it observed, and what does it not
prove?* If a claim cannot be sourced, it goes in the file marked `UNSOURCED` —
never smoothed over.

Do not open a PR that changes behaviour. This is a documentation task. One file.

---

## Non-negotiable rules

1. **No invented facts.** Every number, digest, byte count, status code, SHA and
   URL in `docs/EVIDENCE.md` must be traceable to a named source. If you cannot
   open the source, write `MISSING: <what you looked for>` and say which claim
   it blocks.
2. **Separate first-hand from second-hand.** Logs and bodies are first-hand;
   an external prober's report is second-hand and must be labelled as such in
   every row that relies on it. Never present second-hand evidence as your own
   observation.
3. **State what a capture does *not* prove.** Each evidence group gets a
   `Does not prove:` line. Example: a 200 on a fixtures page proves the route
   exists, not that a price is published.
4. **Three different capabilities, never merged:** *can ingest* (fixtures
   land), *can grade* (a market price exists to score against), *can bet*
   (a ticket is emitted). A lane may be the first and none of the others.
5. **Do not scrape, and do not defeat bot protection.** Cloudflare/CloudFront
   challenges are out of bounds. If a question needs a live probe, write it as a
   probe request for a human, not as something you try.
6. **Do not edit workflows, do not loosen tests, do not add a price source.**
   If you believe something in `main` is wrong, say so in the file and leave the
   code alone.
7. **Never print or commit a credential.** If a report quotes a key, redact it.

---

## Where the evidence is

### A. This repository (`git`/`gh` are authenticated)

- `localdata/*/auto_tickets_performance.json` — the paper record, one file per
  league. This is the primary source for performance.
- `docs/PERFORMANCE-2026-11-23.md` — a snapshot of that record taken on
  2026-11-23; re-derive the numbers, do not trust them.
- `docs/FREE-API-RESEARCH.md` — price-avenue reconnaissance (GitLab, PyPI, npm,
  bot-protection reasoning) added 2026-11-23.
- `docs/OPERATIONS.md`, `docs/ARCHITECTURE.md` — cadence, lane cap, concurrency.
- `ingestion/pipeline.py` — the dispatch chain, the fail-closed guards, the
  prediction acquisition path. Line-level claims need line-level quotes.
- `ingestion/providers/*.py` — the adapters, each with a probe (`last_probe()`).
- `tests/` — every fail-closed behaviour the project claims is pinned by a test.
  A claim with a test is stronger than a claim without one; say which applies.
- `workflows/ingestion.yml` and `.github/workflows/ingestion.yml` — must be
  identical; cadence lives outside the repo (no `cron:`).

Merged PRs to read for intent and dated decisions:
`#173 eb89fd3` · `#174 72e5847` · `#175 1a025e2` · `#176 dc455c4` ·
`#177 9f576d6` · `#178 1bd7931` · `#179 e217938` · `#180 e03fe6e` ·
`#181 2094beb` · `#182 c928508` · `#183 1eeb6d7` · `#184 23d0cf4` ·
`#185 1250b9a`.

Runs:
`45218527700` 2026-11-23T16:37Z head `dc455c4` (lane logs showed four lanes at
`fixtures_found: 0` with a void guard firing) · `45240293582` 18:37Z head
`e03fe6e`, 54 min, success · `45250804751` 19:37Z head `eb45dd2`, 44 min,
success — **this is the run whose lane logs exist and matter most** · then eight
consecutive failures `45261222884` 20:37Z → `45320753471` 03:37Z at head
`b6619a5`, all in preflight, all within ~2 minutes.

### B. Operator-pasted artifacts

Files the operator uploaded during the previous session, usually at
`/tmp/arena-coding-uploads/`. **The path may not exist in your sandbox** — if it
does not, say so in `docs/EVIDENCE.md` under `MISSING` and list what each
missing file was needed for.

Lane logs (stdout of one job each): `2_Japan`, `3_France`, `3_Basketball`,
`4_ABA`, `4_Italy`, `6_Italy`, `6_NCAA women`, `8_Japan`, `8_Spain` (×2),
`10_Basketball Africa League`, `10_Israel`, `11_Spain ACB`, `12_Basketball`,
`13_Italy LBA`, `14_Japan`, `15_Israel Premier League`, `16_Italy`,
`17_Israel`, `17_Japan B.League` (names approximate; sizes 117 KB–4 MB).

Reports (second-hand, from an external prober the operator calls "Zai"):
`lane-audit-2026-11-23.md` · `runner-nojoy-diagnosis-2026-11-23.md` ·
`price-source-probes-2026-11-23.md` · `lane-probe-2026-11-23.md` ·
`betexplorer-probe-2026-11-23.md` · `bleague_six_leagues_probe_report.md` ·
`period_probe_report.md` · `espn_deep_history_backfill_fix.md` ·
`lnb_robots_curl.txt`, `lnb_temporeal.json`, `liga_fixture_laliga.json`.

`evidence-2026-11-22.md` (123 KB) is an **earlier consolidation attempt**. Read
it, but verify: treat it as a source of claims to check, not as truth. Your file
supersedes it; say so explicitly and note anything you could not re-verify.

---

## Claims you must verify rather than copy

Each of these is a number or a code fact to be re-derived from the source. A row
in your file that merely repeats one of these without the check is a failure.

| # | Claim | How to check |
|---|---|---|
| 1 | Japan lane, run `45250804751`: `transport_attempts: 22`, `transport_responses: 22`, `transport_errors: 0` beside `network_requests: 0` | grep the `17_Japan B.League` log for both keys |
| 2 | Japan health row populated: `state: missing`, `slate_games: 10`, `off_date_dropped: 10`, `requests: 22`, `bytes: 5733851` | same log, inside `provider_health` |
| 3 | Japan `day=24` page digest `d093393c…` matches the digest an external probe recorded independently | log vs `price-source-probes-2026-11-23.md` |
| 4 | The string `treat the slate as unknown, not empty` appears **0** times in all seven lane logs of run `45250804751` | grep every log; report the exact per-file count |
| 5 | Italy LBA: `fixtures_found: 6`, `off_date_dropped: 8`, `gaps=0` | grep `13_Italy LBA` |
| 6 | Italy LBA and Israel both: `games_with_any_pre_tip_prediction: 0`, `predictions_selected: {}`, *and* a "fresh named full-game moneyline already exists for every slate game" detail | grep both logs |
| 7 | BAL: `rows: 42`, `games: 0`, `off_date_dropped: 0` — unexplained drop, nothing records why | grep `10_Basketball Africa League`; confirm in `ingestion/providers/bal.py` that no counter covers that path |
| 8 | ABA: `fixtures_found: 4252` — a season-sized number, not a slate | grep `4_ABA`; check what the adapter returns |
| 9 | **Structural blocker:** `_predixsport_candidates_for_games` returns early unless the league is `NBA` or `EUROLEAGUE`, so non-ESPN lanes have no probability source at all | read `ingestion/pipeline.py`, quote the line and its number |
| 10 | Paper record: 33 settled positions, 11W/22L, 33.0u staked, +9.659u, ROI 29.27%; 23 of 25 leagues have zero settled positions | re-derive from `localdata/*/auto_tickets_performance.json` with your own script and paste the totals |
| 11 | ROI is not significant: interval spans zero (±32.4% at a single-bet-like spread; wider on the 29 acca positions) | recompute; state the assumption you used |
| 12 | NBA leg Brier 0.2766 (worse than a coin's 0.25), NCAA_M 0.2160 | the per-league JSON |
| 13 | Day-param settlement: four `day=22/23/24/25` fetches, 143,266 B each, same ten keys 506602–506611, bodies differing by exactly the 2-byte echoed `day`; digests `1fca1127…`, `bb0f2aa7…`, `d093393c…`, `6ed7646c…` | `price-source-probes-2026-11-23.md` + the probe fields it produced (`day_param`, `pages[].new_keys`) in the code and logs |
| 14 | B.League publishes no price: S3 snapshot 129 (543,231 B) has zero matches for `odd|bet|quota|cuota|moneyline|handicap|spread`; `/game/detail/{key}/` returns 404 | second-hand; label as such, keep the "JS-rendered site not exhaustively swept" caveat |
| 15 | ACB match-header has no odds: future `matchId 105456` → 200 with the site `X-APIKEY` (2,070 B), 401 with (0 B) without; 18 fields, identical key set to played `105375` (2,355 B) | second-hand; also confirm `acb.py` always sends the header |
| 16 | BetExplorer: ACB fixtures route 203,990 B with empty odds cells (`data-oid=""`, 0 `data-odd`) for the unopened 12-05/06 round; played games carry 162 `data-odd` values spanning 1.04–11.45, none `1.00` | second-hand; cite `betexplorer-probe-2026-11-23.md` |
| 17 | Workflow incident: lane timeout raised 60→360 at 2026-11-23T19:59Z (commits `7400bc3`, `50f82b2`); a test pinned `<= 60`; every tick from 20:37Z failed in preflight until `#184` replaced the constant with the concurrency property | `gh run list`, `gh api .../commits?path=.github/workflows/ingestion.yml`, `tests/test_workflow_sync.py` |
| 18 | Concurrency is serialising: `group: magolide-automatic-available-leagues`, `cancel-in-progress: false`; no `cron:` anywhere in the workflows | grep both workflow files |
| 19 | Exchange prices were previously inexpressible: `PriceProvenance` has `named_book`, `aggregate_consensus`, `no_price`, `exchange` (added `#181`); `betfair`/`betfair_exchange` map to an `acquire=False` exchange policy even when unregistered | `ingestion/models.py`, `ingestion/league_registry.py`, `tests/test_price_provenance_exchange.py` |
| 20 | Package reconnaissance: `betfairlightweight` 2.24.0 (2026-09-04, MIT) is the only live price-avenue client; every npm client checked last released 2014–2019; `oddspapi` 404s on PyPI; GitLab public search returned no usable odds project | re-check with `curl https://pypi.org/pypi/<name>/json` and `https://registry.npmjs.org/<name>` if you have network; otherwise cite `docs/FREE-API-RESEARCH.md` and label it repo-recorded |

Anything you cannot check becomes an explicit `UNVERIFIED` row. That is a
successful outcome, not a gap to hide.

---

## Required structure of `docs/EVIDENCE.md`

1. **Scope and method** — what this file covers, the cut-off timestamp, the
   source hierarchy (first-hand capture > operator-pasted log > external prober
   report > repo code > inference), and how contradictions are resolved.
2. **Executive summary** — ≤ 15 sentences. Must include: what the system can do
   today; what it has actually earned; and the single biggest structural blocker.
   No adjectives that the evidence cannot carry.
3. **Timeline** — dated table: commit → what changed → what it was meant to fix →
   whether the next run confirmed it.
4. **Lane table** — one row per league: fixtures on the target day, whether the
   source was read (with the receipt), whether a prediction source exists,
   whether a price exists, and the resulting label (`can ingest` / `can grade` /
   `can bet`, only the true ones).
5. **Evidence ledger** — the core of the file. Columns:
   `id | claim | exact value/quote | source (path or URL) | observed at (UTC) | first/second-hand | strength | does not prove`.
6. **Contradictions and corrections** — every claim that was made and later
   withdrawn, by whom, and on what evidence. Include the ones where the previous
   agent was wrong. This section is mandatory and must not be softened.
7. **Open questions** — each with the exact probe or code change that would
   settle it, and whether it needs a human.
8. **Reproducibility** — the commands you ran to re-derive the headline numbers,
   with their output. A reader must be able to repeat them.
9. **Index of sources** — every file and URL cited, with size or digest when the
   source is a body.

House style: tables over prose, exact quotes over paraphrase, no hedging verbs
that hide an absence of evidence. If something is unknown, write `UNKNOWN` and
say what would resolve it.

---

## Sanity checks before you finish

- Re-derive the paper record yourself; your totals must match 33 settled
  positions / 11W / 22L / +9.659u / 29.27% ROI, or you must explain the
  difference.
- `grep -c` the void-guard string across every lane log and report the numbers.
- Confirm the prediction-scope line exists in `ingestion/pipeline.py` and quote it
  with its line number.
- Confirm `cron:` appears nowhere in either workflow file.
- Count your ledger rows by strength and state the split (first-hand vs
  second-hand vs repo code). If more than half of the load-bearing rows are
  second-hand, say so in the executive summary.

## What success looks like

A person who has never seen this project reads `docs/EVIDENCE.md` and can state,
without asking a follow-up question: what the system ingests today, what it has
earned on paper and whether that is distinguishable from luck, where the price
layer actually stands, why the revived lanes have never emitted a pick, and
exactly which evidence is missing to change any of those three answers.
