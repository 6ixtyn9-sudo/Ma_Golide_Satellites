# NKL settlement: first live run

**Session:** 2026-11-24 UTC, operator-side run (this sandbox has no egress).
**Question asked:** does the settlement chain work end to end against the site as
served — forecast names in, reconciling `ActualResult` out?

**Answer:** yes, for three games whose finals were already known from the club
listings. What follows is the run, the cost, and the two defects it found.

## What was run

`ingestion.providers.nkl.settle_forecast(client.get_text, played_on=…, home=…,
away=…, cache=…)`, where `client` is `ingestion.http.HttpClient()` and `cache` is
one dict reused across the three games. `Crawl-delay: 10` (from `robots.txt`)
honoured with a 10 s sleep between games.

## Result

| Date | Fixture | Settled | Requests | Final | Reconciles |
|---|---|---|---|---|---|
| 2026-10-01 | Kretinga v Neptunas 2 | **yes** | 2 | 76–91 | `(True, '')` |
| 2026-10-03 | Vilkaviskis v Kretinga | **yes** | 2 | 84–65 | `(True, '')` |
| 2026-10-10 | Kretinga v Ukmerges Stekas | **yes** | 1 | 95–99 | `(True, '')` |

All three finals match the club listings independently observed earlier: `125749`
76–91, `125760` 84–65, `125775` 95–99.

**Five requests, not four.** Games 1 and 3 share Kretinga's club page (game 3
costs one request for the match page alone); game 2 is Vilkaviskis at home, so it
needs its own club page. 2 + 2 + 1 = 5 is the cache working — a 4 would have meant
a page was not fetched when it was needed.

## Defects the run found

Both were reported by the code and fixed in the same session, in that order:

1. **`no row` was not a diagnosis.** The first run returned the same sentence for
   three games the listings plainly contain. `explain_miss()` (#221) now separates
   an empty page, a page of other dates, and an unmapped name — printing the row
   names exactly as the page spells them.
2. **A fixture printed three times was read as three fixtures.** Kretinga's page
   emits each played row once per layout variant, byte-identical, so the
   ambiguity rule refused every game on it. `collapse()` (#222) keeps one row per
   match id; identical repeats collapse, and contradictory data under one id still
   refuses. The two are deliberately not the same rule.

Neither was visible from a single-game test. The first was found because the run
printed its reasons; the second because the first printed them.

## What this run does *not* establish

- **Terms.** `nkl.lt` publishes `Crawl-delay: 10` in `robots.txt` and the delay is
  honoured, but no terms document has been read. Storing results needs an
  `OperatorAcceptance`; the `LITHUANIA_NKL` status is unchanged.
- **Overtime.** Totals 205, 202 and 194 (the season's highest) all reconciled with
  four periods, so either no OT games exist or NKL folds OT into the four columns.
  The five-column refusal is proven only against synthetic markup.
- **The listing's own score.** Each listing row carries a final and each match
  page carries another; they agreed on every game observed, and
  `settle_forecast` does not compare them. Whether that redundancy is a check or
  a trap is unknown, so no comparison was invented.
- **Row pairing across two pages.** Withdrawals, replayed fixtures and
  rescheduled dates are not covered by three games.
