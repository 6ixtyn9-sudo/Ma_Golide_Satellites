#!/usr/bin/env python3
"""League-agnostic counterfactual ticket replay harness for Ma_Golide_Satellites.

Calls the live ``forecasting.accas.build_accas`` engine directly over frozen
decision-time ``candidate_pool`` snapshots stored inside immutable
``auto_tickets_YYYY-MM-DD_decision.json`` receipts across all active lanes.

Usage:
  python3 scripts/replay_harness.py                              # status + live baseline
  python3 scripts/replay_harness.py --parity                     # verify live build_accas parity
  python3 scripts/replay_harness.py --since 2026-11-26           # post-cutover clean window
  python3 scripts/replay_harness.py --league nba,ncaa_m,ncaa_w   # specific lanes
  python3 scripts/replay_harness.py --mode cross-league --variant "legs=2,min_odds=3.0,max_odds=4.5,stake_mode=daily_pct,stake_frac=0.25"
  python3 scripts/replay_harness.py --ab live "legs=2,min_odds=3.0,max_odds=4.5"
  python3 scripts/replay_harness.py --kelly --variant "legs=3,min_odds=3.0,max_odds=4.5"
  python3 scripts/replay_harness.py --slots                      # rank slot + combined-odds table

Variant spec syntax (comma-separated key=value, or "live"):
  legs=2..6  limit=1..100  min_leg_odds=1.20  max_leg_odds=3.00  min_leg_prob=0.60
  price_weight=0.15  min_odds=3.0  max_odds=4.5  forbid_opposing=0|1  allow_provisional=0|1
  markets=moneyline|total|!spread  leagues=nba|ncaa_m|ncaa_w  include_derived=0|1
  stake_mode=flat|daily_pct|per_acca_pct|kelly  stake_frac=0.25  kelly_frac=0.25
  weights=3:2:1  skim=0.10
"""

from __future__ import annotations

import argparse
from collections import Counter
from pathlib import Path
import sys
from typing import Any, Dict, List, Sequence

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from forecasting.replay import (  # noqa: E402
    LOW_SAMPLE_DAYS,
    DayEvidence,
    ReplaySpec,
    check_parity,
    compare_variants,
    kelly_sweep,
    load_localdata_evidence,
    parse_variant_spec,
    replay,
)


def _fmt_pct(value: float | None) -> str:
    return "--" if value is None else f"{value * 100:+.1f}%"


def _fmt_rate(value: float | None) -> str:
    return "--" if value is None else f"{value * 100:.1f}%"


def _print_inventory(days: Sequence[DayEvidence], *, since: str | None) -> None:
    by_league: Dict[str, Counter] = {}
    for d in days:
        c = by_league.setdefault(d.league, Counter())
        c["receipts"] += 1
        if d.has_candidate_pool:
            c["with_pool"] += 1
        else:
            c["legacy_no_pool"] += 1
        if d.outcomes_by_fid:
            c["with_outcomes"] += 1
    scope = f"since {since}" if since else "all dates"
    print(f"REPLAY INVENTORY ({scope}) — {len(days)} receipt-days across {len(by_league)} active lane(s)")
    if not by_league:
        print("  (no immutable auto_tickets_*_decision.json receipts found)")
        return
    for league in sorted(by_league):
        c = by_league[league]
        print(
            f"  {league:18s} receipts={c['receipts']:3d}  "
            f"pool_snapshots={c['with_pool']:3d}  legacy_no_pool={c['legacy_no_pool']:3d}"
        )


def _print_summary(label: str, res: Dict[str, Any]) -> None:
    n = res["settled_days"]
    flag = f" [LOW SAMPLE: n={n} < {LOW_SAMPLE_DAYS}]" if n < LOW_SAMPLE_DAYS else ""
    legacy_note = (
        f" [LEGACY PUBLISHED-TICKETS SIZING ONLY: {res['legacy_sizing_days']}d]"
        if res["legacy_sizing_days"]
        else ""
    )
    print(f"\n[{label}] mode={res['mode']}  settled_days={n}  pending_days={res['pending_days']}  skipped_no_pool={res['skipped_legacy_no_pool']}{flag}{legacy_note}")
    if n == 0:
        print("  0 settled replayable days — wait for pool-equipped receipts to settle (or pass --since / --sizing-only-legacy).")
        return
    print(
        f"  Record: {res['wins']}W-{res['losses']}L-{res['pushes']}P ({_fmt_rate(res['hit_rate'])})  "
        f"Flat: {res['flat_profit']:+.2f}u / {res['flat_staked']:.0f}u ({_fmt_pct(res['flat_roi'])})"
    )
    print(
        f"  Growth: mean_log_g={res['mean_log_growth_per_day']:+.4f}/day  "
        f"Wealth={res['total_wealth']:.1f}% (bank={res['final_bank']:.1f}% + vault={res['skimmed_vault']:.1f}%)  "
        f"Peak={res['peak_wealth']:.1f}%  MaxDD={res['max_drawdown'] * 100:+.1f}%"
    )


def _print_slots_and_bands(res: Dict[str, Any]) -> None:
    if res["settled_days"] == 0:
        return
    print("\n  BY RANK SLOT (flat 1u):")
    for idx in sorted(res["slots"]):
        s = res["slots"][idx]
        dec = s["wins"] + s["losses"]
        roi = s["profit"] / s["staked"] if s["staked"] else None
        hit = s["wins"] / dec if dec else None
        print(
            f"    Slot #{idx:<2d}  {s['wins']:3d}W-{s['losses']:3d}L-{s['pushes']:2d}P "
            f"({_fmt_rate(hit):>6s})  flat={s['profit']:+7.2f}u ({_fmt_pct(roi):>7s})"
        )
    print("\n  BY COMBINED ODDS BAND (flat 1u):")
    for label, b in res["odds_bands"].items():
        if b["staked"] == 0:
            continue
        dec = b["wins"] + b["losses"]
        roi = b["profit"] / b["staked"] if b["staked"] else None
        hit = b["wins"] / dec if dec else None
        print(
            f"    {label:8s}  {b['wins']:3d}W-{b['losses']:3d}L-{b['pushes']:2d}P "
            f"({_fmt_rate(hit):>6s})  flat={b['profit']:+7.2f}u ({_fmt_pct(roi):>7s})"
        )


def _parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="League-agnostic counterfactual ticket replay harness (live build_accas path)."
    )
    parser.add_argument(
        "--localdata-root",
        type=Path,
        default=ROOT / "localdata",
        help="Root directory containing per-league localdata folders (default: localdata/).",
    )
    parser.add_argument(
        "--league",
        default="all",
        help="Comma-separated league slugs/names, or 'all' (default: all active lanes).",
    )
    parser.add_argument("--since", default=None, help="Include receipts on/after YYYY-MM-DD.")
    parser.add_argument("--until", default=None, help="Include receipts on/before YYYY-MM-DD.")
    parser.add_argument(
        "--mode",
        choices=("per-league", "cross-league"),
        default="per-league",
        help="Run build_accas per league or over merged daily cross-league pools.",
    )
    parser.add_argument(
        "--sizing-only-legacy",
        action="store_true",
        help="Allow staking-only replay on legacy receipts lacking candidate_pool (per-league only).",
    )
    parser.add_argument("--parity", action="store_true", help="Verify live build_accas parity vs frozen receipts.")
    parser.add_argument("--variant", default=None, help="Single variant spec to evaluate.")
    parser.add_argument("--ab", nargs=2, metavar=("SPEC_A", "SPEC_B"), help="Paired bootstrap A/B comparison.")
    parser.add_argument("--kelly", action="store_true", help="Daily bankroll deployment % sweep.")
    parser.add_argument("--slots", action="store_true", help="Print rank-slot and odds-band breakdown.")
    parser.add_argument("--bootstraps", type=int, default=2000, help="Bootstrap resamples for --ab.")
    return parser.parse_args(argv)


def main(argv: Sequence[str] | None = None) -> int:
    args = _parse_args(argv)
    leagues: List[str] | None = (
        None
        if args.league.strip().lower() == "all"
        else [item.strip() for item in args.league.split(",") if item.strip()]
    )
    days = load_localdata_evidence(
        args.localdata_root,
        leagues=leagues,
        since=args.since,
        until=args.until,
    )
    _print_inventory(days, since=args.since)

    if args.parity:
        report = check_parity(days)
        print(
            f"\n[PARITY] checked={report['checked']}  matched={report['matched']}  "
            f"mismatched={report['mismatched']}  skipped_legacy_no_pool={report['skipped_legacy_no_pool']}"
        )
        for item in report["mismatches"]:
            print(f"  MISMATCH {item['date']} {item['league']}: expected={item['expected']} rebuilt={item['rebuilt']}")
        return 1 if report["mismatched"] else 0

    if args.ab:
        spec_a = parse_variant_spec(args.ab[0])
        spec_b = parse_variant_spec(args.ab[1])
        cmp = compare_variants(
            days,
            spec_a,
            spec_b,
            mode=args.mode,
            allow_legacy_sizing_only=args.sizing_only_legacy,
            bootstraps=args.bootstraps,
        )
        _print_summary(f"A: {args.ab[0]}", cmp["a"])
        _print_summary(f"B: {args.ab[1]}", cmp["b"])
        status = cmp["status"]
        if status == "insufficient_data":
            print("\n[A/B] insufficient_data (0 settled replayable days).")
            return 0
        if status == "no_op":
            print(f"\n[A/B] NO-OP across all {cmp['n_days']} day(s) — identical tickets & sizing; bootstrap refused.")
            return 0
        flag = f" [LOW SAMPLE: n={cmp['n_days']} < {LOW_SAMPLE_DAYS}]" if cmp["low_sample"] else ""
        lo_g, hi_g = cmp["delta_log_growth_ci95"]
        lo_p, hi_p = cmp["delta_flat_profit_ci95"]
        print(
            f"\n[A/B] status={status}  days={cmp['n_days']}  selection_diff_days={cmp['selection_diff_days']}{flag}"
        )
        print(
            f"  Delta mean_log_g (B-A): {cmp['delta_mean_log_growth']:+.4f}/day  "
            f"95% CI [{lo_g:+.4f}, {hi_g:+.4f}]  P(B>A)={cmp['prob_b_beats_a_log_growth'] * 100:.1f}%"
        )
        print(
            f"  Delta flat P/L   (B-A): {cmp['delta_flat_profit']:+.2f}u  "
            f"95% CI [{lo_p:+.2f}u, {hi_p:+.2f}u]"
        )
        return 0

    spec = parse_variant_spec(args.variant) if args.variant else ReplaySpec()
    label = args.variant or "live"
    res = replay(
        days,
        spec,
        mode=args.mode,
        allow_legacy_sizing_only=args.sizing_only_legacy,
    )
    _print_summary(label, res)
    if args.slots:
        _print_slots_and_bands(res)
    if args.kelly:
        rows = kelly_sweep(
            days,
            spec,
            mode=args.mode,
            allow_legacy_sizing_only=args.sizing_only_legacy,
        )
        print("\n[KELLY / DAILY DEPLOYMENT SWEEP]")
        if not rows or rows[0]["settled_days"] == 0:
            print("  0 settled replayable days.")
        else:
            print(f"  {'Deploy/day':>10s} | {'Mean log g':>10s} | {'Wealth':>9s} | {'Peak':>9s} | {'MaxDD':>8s}")
            for row in rows:
                print(
                    f"  {row['daily_pct'] * 100:9.0f}% | "
                    f"{row['mean_log_growth_per_day']:+10.4f} | "
                    f"{row['total_wealth']:8.1f}% | "
                    f"{row['peak_wealth']:8.1f}% | "
                    f"{row['max_drawdown'] * 100:+7.1f}%"
                )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
