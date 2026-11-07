# ESPN acquisition, history, and coverage limits

Verification date: **2026-11-07**. This document separates live observations,
hermetic regression coverage, and facts that still require scheduled-run
verification. Passing fixtures do not prove a provider's current availability.

## Scope

ESPN now supplies acquisition for **NBA, WNBA, NCAA_M, and NCAA_W**. WNBA is
an independent sixth Ma Golide league with isolated `localdata/wnba` state;
it is not aliased to NBA. EuroLeague and CBA retain their existing providers.

The acquisition layer provides bounded requests, retries/backoff, a bounded
redacted run-receipt list, response hashes, cooldowns for recent failures, and
atomic public-response caching. Secrets are never cache keys or
receipt fields. Canonical evidence remains DuckDB plus CSV.GZ and immutable
ledgers; raw caches and receipts are recovery aids, not a second sports store.

## What is retained

- Scoreboard identity, UTC tip, native status/format, final scores, linescores,
  records, venue, broadcasts, named current/opening books, and book payload
  provenance. Betting links and other navigation clutter are discarded.
- Event-summary team statistics, leaders, recent form, injuries, venue,
  predictor context, against-the-spread context, and one event-scoped article.
  Article `lastModified`/`published` is retained when available; payload fetch
  time is separately retained as `observed_utc`.
- ESPN BPI as independent prediction evidence. `homeTeam`/`awayTeam` are
  matched to provider team IDs, and both `gameProjection` and
  `teamChanceLoss` orientations are tested.
- Every accepted current book as a timestamped `BookSnapshot`, independent of
  whether a prediction exists. Opening numbers can lack an opening observation
  timestamp and therefore remain context, never closing-line-value evidence.

Summary enrichment is capped per run and recorded in provider health and the
acquisition telemetry. A missing summary, book, statistic, or article remains
a coverage gap rather than a synthesized value.

## Verified history observations

The deep-history source is ESPN's paginated core event index:

```text
https://sports.core.api.espn.com/v2/sports/basketball/leagues/{slug}/events
  ?dates={year}&limit=1000&page={page}
```

The index supplies event references. A bounded per-event site summary then
supplies validated status, identities, score, format, and available context.
The following are observations made on 2026-11-07, **not guarantees of
exhaustive regular-season coverage**:

| League | Nonempty index observation | Earlier empty query | Verified example / caveat |
| --- | --- | --- | --- |
| NBA | 1947: 19 indexed events | 1946: 0 | Event `1770402009`, 1947-04-02: Philadelphia Warriors 73–68 St. Louis Bombers. Final score exists; linescores, boxscore, and play-by-play do not. The old index is playoff-skewed. |
| WNBA | 1997: 112 indexed events | 1996: 0 | Event `170621006`, 1997-06-21, is the inaugural-game example. Its event format declares two halves. |
| NCAA_M | 2001: 242 indexed events | 2000: 0 | Returned metadata can disagree with the requested year/season; the count is not a complete-season total. |
| NCAA_W | 2001: 688 indexed events | 2000: 0 | Same metadata and completeness caveat as NCAA_M. |

For scale, the NBA 2025 index returned 1,412 events across two pages at
`limit=1000`. This is why history does not scan every calendar day blindly.
ESPN rejected the tested multi-day scoreboard ranges with HTTP 400; recent
backfill therefore uses individual dates.

Conservative acquisition floors are NBA 1947, WNBA 1997, NCAA_M 2001, and
NCAA_W 2001. An empty earlier query does **not** prove that no older data
exists anywhere in ESPN. These floors may be revised only after another
validated probe; they are not represented as absolute historical minima.

## Bounded backfill algorithm

Each write-enabled league run performs two independent jobs:

1. **Recent results:** at most seven individual past scoreboard dates. Public
   payloads use the durable cache for 24 hours while retaining their original
   observation timestamps.
2. **Deep history:** at most 25 indexed event summaries by default. The season
   index is cached for seven days. Historical summaries are sanitized into
   canonical evidence rather than retained indefinitely as large raw cache
   files. The durable cursor prevents normal refetching; an already-persisted
   event with both summary and history evidence is skipped.

The cursor is the `provider_health` row with provider `espn-history`. Its
coverage contains `next_season`, `next_event_index`, `last_index_year`, and
`floor_season`. It starts at the previous season and works backward. It
advances only after `DuckDBStore.upsert_game` succeeds, or when that exact
historical summary is already persisted. A provider/identity error leaves the
cursor on the failed event. A dry run never advances the deep cursor.

The workflow needs no new league/date/mode/write inputs. Its ordinary
three-hour league matrix runs this same bounded backfill. Scheduled cache
saves preserve public acquisition caches even on dry runs, but dry runs still
cannot create canonical sports state or advance deep history.

Inspect a persisted cursor read-only with:

```bash
python - <<'PY'
from ingestion.warehouse import connect
con = connect('localdata/nba/warehouse.duckdb')
print(con.execute("""
    select coverage_json from provider_health
    where provider = 'espn-history' and league = 'NBA'
""").fetchall())
con.close()
PY
```

Each operational backfill report includes recent dates, deep event counts,
coverage classes, the cursor, request telemetry, and acquired ledger dates.
Immutable per-date `backfill` ledgers are written only for dates whose games
were actually acquired in a write-enabled run.

## Native results and quarantine

Historical coverage is explicitly classified:

- `reconciled_periods`: declared native periods plus overtime, when present,
  reconcile to the official final score;
- `unreconciled_periods`: period values are retained but disagree with the
  final or violate format constraints;
- `final_score_only`: the event is final and has a score, but native
  linescores are absent or unusable;
- `non_final_provider_status`: an old payload still says in progress or
  another non-final state. It is retained as an anomaly, not converted to a
  final.

NBA quarters and NCAA halves follow their native formats. WNBA uses the
**event's declared format**: early two-half games remain H1/H2; later
four-quarter games remain Q1–Q4. A modern four-quarter assumption is never
used to relabel an early WNBA game. `PeriodKind.UNKNOWN` final-score-only rows
have no invented period results or FT/AOT classification.

Only final, reconciled results enter grading. Final-score-only rows increase
historical score coverage but do not supply quarter/half training labels,
frozen pre-tip books, or pre-tip predictions. Deep ingestion alone does not
establish learning, calibration, profitability, or eligible accumulators.

## Operator verification

```bash
python -m pytest -q
ruff check .
python -m ingestion.cli smoke --league NBA --date 2026-11-07
python -m ingestion.cli smoke --league WNBA --date 2026-11-07
python -m ingestion.cli smoke --league NCAA_M --date 2026-11-07
python -m ingestion.cli smoke --league NCAA_W --date 2026-11-07
```

Smoke is read-only and rejects `--write`. Its verification is
`LIVE_VERIFIED` only when the actual HTTP client succeeds; injected fixture
clients are labeled `FIXTURE_VERIFIED`. Scheduled DuckDB restoration, cursor
creation, ledger artifacts, and public-cache reuse still require inspection
of installed workflow runs. GitHub Actions caches are not archival backups;
retain canonical CSV.GZ and immutable ledgers according to the recovery plan.
