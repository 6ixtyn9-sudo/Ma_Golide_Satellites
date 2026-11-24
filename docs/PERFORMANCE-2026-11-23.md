# Paper record as of 2026-11-23

Computed from `localdata/*/auto_tickets_performance.json` on the merged tree.
Every number below is a field from those files; nothing is modelled or
projected.

| League | Settled positions | W | L | Stake (u) | Profit (u) | ROI | Hit | Brier (legs) | Pending | Unresolved legs |
|---|---|---|---|---|---|---|---|---|---|---|
| NBA | 19 | 5 | 14 | 19.0 | +0.795 | +4.18% | 26.3% | 0.2766 | 5 | 15 |
| NCAA_M | 14 | 6 | 8 | 14.0 | +8.864 | +63.32% | 42.9% | 0.2160 | 5 | 15 |
| **Total** | **33** | **11** | **22** | **33.0** | **+9.659** | **+29.27%** | **33.3%** | — | 10 | 30 |

Twenty-two of the twenty-four registered leagues have **zero settled
positions** (24 `auto_tickets_performance.json` files, 2 with a non-zero
`settled_positions` field: NBA 19, NCAA_M 14). The two that do are the ESPN
lanes. Every lane revived in this
period — B.League, ACB, LBA, KBL, BCL, BAL, BSL — has produced **no settled
position at all**, not "no pick today".

## Why "+29% ROI" is not a result

- **n = 33.** Hit rate 33.3% ± 16.1% (95%, binomial). ROI's interval cannot be
  pinned down without a variance assumption, and the assumption changes the
  width by a factor of three. The ladder, all spanning zero:
  - **±32.4%** — optimistic: every position treated as a single ~0.95u bet
    (t = 1.77).
  - **±62.4%** — the file-derivable model: 11 wins returned 42.659u on 33.0u
    staked, i.e. an average payout of 3.878u per winning position; with flat
    1u independent positions the single-position spread is
    3.878·√(0.333·0.667) = 1.83u, so the standard error of the mean is
    1.83/√33 = 0.32u and the ROI interval is ±62.4pp (z = 0.92).
  - **±96pp** — the acca-realistic end: bootstrapping 20 000 times over the
    odds recorded in the frozen decision receipts (43 positions, combined odds
    up to 14.94, outcomes redrawn to the observed 11/33) gives
    −17% … +175%.

  Every method agrees on the conclusion: **the interval spans zero**, and the
  narrow figure should never be quoted on its own.
- **The +63% is 14 bets in one league**, chosen after the fact from a
  twenty-four-league field. Its companions: NBA at +4.18% (noise) and a Brier
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
