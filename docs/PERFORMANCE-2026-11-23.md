# Paper record as of 2026-11-23

Computed from `localdata/*/auto_tickets_performance.json` on the merged tree.
Every number below is a field from those files; nothing is modelled or
projected.

| League | Settled positions | W | L | Stake (u) | Profit (u) | ROI | Hit | Brier (legs) | Pending | Unresolved legs |
|---|---|---|---|---|---|---|---|---|---|---|
| NBA | 19 | 5 | 14 | 19.0 | +0.795 | +4.18% | 26.3% | 0.2766 | 5 | 15 |
| NCAA_M | 14 | 6 | 8 | 14.0 | +8.864 | +63.32% | 42.9% | 0.2160 | 5 | 15 |
| **Total** | **33** | **11** | **22** | **33.0** | **+9.659** | **+29.27%** | **33.3%** | — | 10 | 30 |

Twenty-three of the twenty-five registered leagues have **zero settled
positions**. The two that do are the ESPN lanes. Every lane revived in this
period — B.League, ACB, LBA, KBL, BCL, BAL, BSL — has produced **no settled
position at all**, not "no pick today".

## Why "+29% ROI" is not a result

- **n = 33.** Hit rate 33.3% ± 16.1% (95%, binomial). The stake is 33 flat
  units, so ROI's standard error is dominated by per-bet variance: at a
  single-bet-like spread of ~0.95u the interval is ±32.4% (t = 1.77); with
  acca-heavy returns — and accas are 29 of the 33 settled positions — the
  interval is wider still (t < 1 at 2.5u spread). **The interval spans zero.**
- **The +63% is 14 bets in one league**, chosen after the fact from a
  twenty-five-league field. Its companions: NBA at +4.18% (noise) and a Brier
  of 0.2766, which is worse than a coin's 0.25 on a two-way leg.
- **Acca positions are not independent.** Legs repeat across accas, so the
  effective sample is smaller than 33 and the interval above is optimistic.
- **30 legs are still unresolved.** The settled subset is not a random sample
  of everything the model decided: it is the part that resolved first.

So "inconclusive" is the correct reading, and it is now a measurement rather
than an impression.

## What would make it conclusive

Pre-register the threshold before the games land, so it cannot be chosen to fit
the result. A defensible one for a bootstrapped single-user project:

1. **≥ 200 settled positions** across at least three independent leagues
   (not three variants of the same competition);
2. **positive closing-line value** on the settled subset — CLV survives small
   n far better than ROI, because each bet contributes one number;
3. ROI's 95% interval **excluding zero** at that n;
4. Brier **below 0.25** on two-way legs sustained over the same period.

Until all four hold, the honest label for every lane stays *can ingest*, not
*can grade against a market*, and certainly not *can bet*.

## Oddspedia (checked 2026-11-23)

`surebets` and `hot-bets` are fetchable and free. What they carry right now:
arbitrage rows on football three-way markets plus one Icelandic women's
basketball game; prices in **American** format; a headline "28.70% guaranteed
profit" row. Three reasons this is not a price source for this project:

- a 28.7% arb does not survive arbitrage — that shape is stale lines, limited
  stakes or free-bet artefacts, and treating it as a price would import an
  error directly into the edge calculation;
- it lists no competition we trade, so it cannot price our fixtures;
- automated reuse is governed by their terms, which we have not reviewed, so it
  cannot be labelled `CLEARED` — the same class as BetExplorer: at best
  `aggregate_consensus`, research-only.

Its one legitimate use would be a **market-open signal** ("is this round priced
anywhere yet"), and even that needs the terms read first.
