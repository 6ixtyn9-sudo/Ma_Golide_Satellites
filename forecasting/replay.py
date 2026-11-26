"""League-agnostic, live-faithful ticket replay and paired A/B harness.

Calls :func:`forecasting.accas.build_accas` directly over frozen decision-time
``candidate_pool`` snapshots stored inside ``auto_tickets_YYYY-MM-DD_decision.json``.
Never re-implements live selection rules: one code path, verified by ``check_parity``.

Doctrine
--------
- Selection counterfactuals use ONLY receipts with a frozen ``candidate_pool``
  snapshot (live-faithful pool). Legacy receipts without a pool snapshot are
  skipped by default so post-tip/re-scored data never masquerades as a live pool.
- Primary comparison metric is mean daily log growth (bank-independent), not raw
  final bank.
- Every summary reports ``n`` (settled bet-days) and flags ``n < 30`` as low-sample.
- Paired A/B diffs chosen tickets day-by-day first and refuses to bootstrap a no-op.
"""

from __future__ import annotations

from dataclasses import dataclass, replace
import json
import math
from pathlib import Path
import random
import re
from typing import Any, Dict, List, Optional, Sequence, Tuple

from ingestion.models import Game, schema_matches

from .accas import build_accas
from .operator import (
    TICKET_SCHEMA,
    _forecast_from_row,
    _grade_row,
    _is_derived_spread_ticket_row,
)

LOW_SAMPLE_DAYS = 30
RECEIPT_FILE_RE = re.compile(r"^auto_tickets_(\d{4}-\d{2}-\d{2})_decision\.json$")
WITHDRAWN_FILE_RE = re.compile(r"^auto_tickets_(\d{4}-\d{2}-\d{2})_withdrawn\.json$")
VALID_STAKE_MODES = frozenset({"flat", "daily_pct", "per_acca_pct", "kelly"})
ODDS_BANDS: Tuple[Tuple[str, float, float], ...] = (
    ("<2.0", 0.0, 2.0),
    ("2.0-3.0", 2.0, 3.0),
    ("3.0-4.0", 3.0, 4.0),
    ("4.0-6.0", 4.0, 6.0),
    (">=6.0", 6.0, float("inf")),
)


@dataclass(frozen=True)
class ReplaySpec:
    """Variant overrides forwarded directly to live ``build_accas`` + sizing."""

    legs: Optional[int] = None
    limit: Optional[int] = None
    min_leg_odds: Optional[float] = None
    max_leg_odds: Optional[float] = None
    min_leg_prob: Optional[float] = None
    price_weight: Optional[float] = None
    min_odds: Optional[float] = None
    max_odds: Optional[float] = None
    forbid_opposing: Optional[bool] = None
    allow_provisional: Optional[bool] = None
    markets: Tuple[str, ...] = ()
    exclude_markets: Tuple[str, ...] = ()
    leagues: Tuple[str, ...] = ()
    allow_single_fallback: bool = True
    include_derived: bool = False
    stake_mode: str = "flat"
    stake_frac: float = 0.25
    kelly_frac: float = 0.25
    weights: Tuple[float, ...] = ()
    skim: float = 0.0

    def __post_init__(self) -> None:
        if self.stake_mode not in VALID_STAKE_MODES:
            raise ValueError(
                f"invalid stake_mode {self.stake_mode!r}; expected one of {sorted(VALID_STAKE_MODES)}"
            )
        if not 0.0 < self.stake_frac <= 0.95:
            raise ValueError("stake_frac must be in (0, 0.95]")
        if not 0.0 < self.kelly_frac <= 1.0:
            raise ValueError("kelly_frac must be in (0, 1.0]")
        if not 0.0 <= self.skim < 0.95:
            raise ValueError("skim must be in [0, 0.95)")
        if self.legs is not None and not 2 <= self.legs <= 6:
            raise ValueError("legs must be between 2 and 6")
        if self.limit is not None and not 1 <= self.limit <= 100:
            raise ValueError("limit must be between 1 and 100")


def _parse_bool(raw: str) -> bool:
    lowered = raw.strip().lower()
    if lowered in {"1", "true", "yes", "on"}:
        return True
    if lowered in {"0", "false", "no", "off"}:
        return False
    raise ValueError(f"invalid boolean value: {raw!r}")


def parse_variant_spec(raw: str) -> ReplaySpec:
    """Parse ``'live'`` or comma-separated ``key=value`` variant spec."""
    text = raw.strip()
    if not text or text.lower() == "live":
        return ReplaySpec()
    kwargs: Dict[str, Any] = {}
    for token in (part.strip() for part in text.split(",") if part.strip()):
        if "=" not in token:
            raise ValueError(f"variant token must be key=value: {token!r}")
        key, val = (item.strip() for item in token.split("=", 1))
        if key in {"legs", "limit"}:
            kwargs[key] = int(val)
        elif key in {
            "min_leg_odds",
            "max_leg_odds",
            "min_leg_prob",
            "price_weight",
            "min_odds",
            "max_odds",
            "stake_frac",
            "kelly_frac",
            "skim",
        }:
            kwargs[key] = float(val)
        elif key in {"forbid_opposing", "allow_provisional", "allow_single_fallback", "include_derived"}:
            kwargs[key] = _parse_bool(val)
        elif key == "stake_mode":
            kwargs["stake_mode"] = val.lower()
        elif key == "weights":
            parts = [float(x) for x in val.split(":") if x.strip()]
            if not parts or any(x <= 0 for x in parts):
                raise ValueError(f"weights must be positive colon-separated numbers: {val!r}")
            kwargs["weights"] = tuple(parts)
        elif key == "markets":
            inc, exc = [], []
            for item in val.split("|"):
                item = item.strip().lower()
                if not item:
                    continue
                if item.startswith("!"):
                    exc.append(item[1:])
                else:
                    inc.append(item)
            kwargs["markets"] = tuple(inc)
            kwargs["exclude_markets"] = tuple(exc)
        elif key == "leagues":
            kwargs["leagues"] = tuple(
                item.strip().lower() for item in val.split("|") if item.strip()
            )
        else:
            raise ValueError(f"unknown variant key: {key!r}")
    return ReplaySpec(**kwargs)


@dataclass
class DayEvidence:
    """One league's immutable decision receipt + settled outcomes for one date."""

    date: str
    league: str
    league_slug: str
    decision_id: str
    frozen_at: str
    receipt: Dict[str, Any]
    candidate_pool: List[Dict[str, Any]]
    selection_policy: Dict[str, Any]
    outcomes_by_fid: Dict[str, str]

    @property
    def has_candidate_pool(self) -> bool:
        return bool(self.candidate_pool)


def _read_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def _withdrawn_ids(league_dir: Path) -> set[str]:
    ids: set[str] = set()
    state_path = league_dir / "auto_tickets_state.json"
    if state_path.exists():
        try:
            state = _read_json(state_path)
            for lock in (state.get("frozen_by_date") or {}).values():
                if isinstance(lock, dict) and lock.get("withdrawn") is True and lock.get("decision_id"):
                    ids.add(str(lock["decision_id"]))
        except (OSError, json.JSONDecodeError, AttributeError):
            pass
    for path in league_dir.glob("auto_tickets_*_withdrawn.json"):
        if not WITHDRAWN_FILE_RE.match(path.name):
            continue
        try:
            payload = _read_json(path)
            if isinstance(payload, dict) and payload.get("decision_id"):
                ids.add(str(payload["decision_id"]))
        except (OSError, json.JSONDecodeError):
            pass
    return ids


def _performance_outcomes_by_receipt(league_dir: Path) -> Dict[Tuple[str, str], Dict[str, str]]:
    """Load settled leg + candidate-pool outcomes from ``auto_tickets_performance.json``."""
    perf_path = league_dir / "auto_tickets_performance.json"
    if not perf_path.exists():
        return {}
    try:
        perf = _read_json(perf_path)
    except (OSError, json.JSONDecodeError):
        return {}
    by_key: Dict[Tuple[str, str], Dict[str, str]] = {}
    for day in perf.get("history") or ():
        if not isinstance(day, dict):
            continue
        key = (str(day.get("date") or ""), str(day.get("decision_id") or ""))
        outcomes = by_key.setdefault(key, {})
        pool_outcomes = day.get("candidate_pool_outcomes")
        if isinstance(pool_outcomes, dict):
            for fid, outcome in pool_outcomes.items():
                if outcome in {"win", "loss", "push"}:
                    outcomes[str(fid)] = str(outcome)
        for pos in day.get("positions") or ():
            if not isinstance(pos, dict):
                continue
            for leg in pos.get("legs") or ():
                if not isinstance(leg, dict):
                    continue
                fid = str(leg.get("forecast_id") or leg.get("pick_id") or "")
                outcome = leg.get("outcome") or leg.get("grade_status")
                if fid and outcome in {"win", "loss", "push"}:
                    outcomes[fid] = str(outcome)
    return by_key


def _discover_league_dirs(localdata_root: Path) -> List[Path]:
    if not localdata_root.exists():
        return []
    if any(RECEIPT_FILE_RE.match(p.name) for p in localdata_root.iterdir() if p.is_file()):
        return [localdata_root]
    dirs: List[Path] = []
    for child in sorted(localdata_root.iterdir()):
        if child.is_dir() and any(
            RECEIPT_FILE_RE.match(p.name) for p in child.iterdir() if p.is_file()
        ):
            dirs.append(child)
    return dirs


def load_localdata_evidence(
    localdata_root: Path,
    *,
    leagues: Optional[Sequence[str]] = None,
    since: Optional[str] = None,
    until: Optional[str] = None,
    warehouse_games_by_league: Optional[Dict[str, Dict[str, Game]]] = None,
) -> List[DayEvidence]:
    """Discover and load immutable decision receipts across all active lanes."""
    allowed_leagues = {item.strip().lower() for item in (leagues or ()) if item.strip()}
    evidence: List[DayEvidence] = []
    for league_dir in _discover_league_dirs(localdata_root):
        slug = league_dir.name.lower()
        withdrawn = _withdrawn_ids(league_dir)
        perf_outcomes = _performance_outcomes_by_receipt(league_dir)
        for path in sorted(league_dir.iterdir()):
            match = RECEIPT_FILE_RE.match(path.name)
            if not match:
                continue
            target_date = match.group(1)
            if since and target_date < since:
                continue
            if until and target_date > until:
                continue
            receipt = _read_json(path)
            if not isinstance(receipt, dict) or not schema_matches(
                receipt.get("schema_version"), TICKET_SCHEMA
            ):
                continue
            decision_id = str(receipt.get("decision_id") or "")
            if decision_id and decision_id in withdrawn:
                continue
            league_name = str(receipt.get("league") or slug.upper())
            if allowed_leagues and slug not in allowed_leagues and league_name.lower() not in allowed_leagues:
                continue
            pool = [
                row
                for row in (receipt.get("candidate_pool") or ())
                if isinstance(row, dict)
            ]
            policy = (
                dict(receipt["selection_policy"])
                if isinstance(receipt.get("selection_policy"), dict)
                else {}
            )
            outcomes = dict(perf_outcomes.get((target_date, decision_id), {}))
            games = None
            if warehouse_games_by_league:
                games = (
                    warehouse_games_by_league.get(league_name)
                    or warehouse_games_by_league.get(league_name.lower())
                    or warehouse_games_by_league.get(slug)
                )
            if games:
                for row in pool:
                    fid = str(row.get("forecast_id") or row.get("pick_id") or "")
                    if fid and fid not in outcomes:
                        grade = _grade_row(row, games)
                        if grade in {"win", "loss", "push"}:
                            outcomes[fid] = grade
                for collection in ("accas", "singles"):
                    for pos in receipt.get(collection) or ():
                        for leg in (pos.get("legs") or ()) if isinstance(pos, dict) else ():
                            if not isinstance(leg, dict):
                                continue
                            fid = str(leg.get("forecast_id") or leg.get("pick_id") or "")
                            if fid and fid not in outcomes:
                                grade = _grade_row(leg, games)
                                if grade in {"win", "loss", "push"}:
                                    outcomes[fid] = grade
            evidence.append(
                DayEvidence(
                    date=target_date,
                    league=league_name,
                    league_slug=slug,
                    decision_id=decision_id,
                    frozen_at=str(receipt.get("frozen_at") or ""),
                    receipt=receipt,
                    candidate_pool=pool,
                    selection_policy=policy,
                    outcomes_by_fid=outcomes,
                )
            )
    evidence.sort(key=lambda item: (item.date, item.league, item.decision_id))
    return evidence


def _derived_single_from_row(row: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    price = (row.get("book") or {}).get("selected_decimal_price")
    if not isinstance(price, (int, float)) or isinstance(price, bool) or price <= 1.0:
        return None
    prob = row.get("calibrated_probability")
    if prob is None:
        prob = row.get("raw_probability")
    return {
        "ticket_id": str(row.get("forecast_id") or row.get("pick_id") or ""),
        "ticket_type": "single",
        "leg_count": 1,
        "combined_decimal_odds": round(float(price), 4),
        "estimated_joint_win_probability": float(prob or 0.0),
        "legs": [row],
    }


def _row_selected_price(leg: Dict[str, Any]) -> Optional[float]:
    book = leg.get("book") or {}
    price = book.get("selected_decimal_price")
    if isinstance(price, (int, float)) and not isinstance(price, bool) and price > 1.0:
        return float(price)
    side_prices = book.get("side_prices")
    if isinstance(side_prices, dict):
        p = side_prices.get(str(leg.get("selection") or ""))
        if isinstance(p, (int, float)) and not isinstance(p, bool) and p > 1.0:
            return float(p)
    return None


def select_tickets_for_pool(
    pool_rows: Sequence[Dict[str, Any]],
    *,
    policy: Dict[str, Any],
    spec: ReplaySpec,
    derived_passthrough: bool = False,
) -> List[Dict[str, Any]]:
    """Build tickets via the live ``build_accas`` path from a decision-time pool."""
    native_rows: List[Dict[str, Any]] = []
    derived_rows: List[Dict[str, Any]] = []
    for row in pool_rows:
        league_val = str(row.get("league") or "").lower()
        market_val = str(row.get("market") or "").lower()
        if spec.leagues and league_val not in spec.leagues:
            continue
        if spec.markets and market_val not in spec.markets:
            continue
        if spec.exclude_markets and market_val in spec.exclude_markets:
            continue
        if _is_derived_spread_ticket_row(row):
            derived_rows.append(row)
        else:
            native_rows.append(row)

    native_forecasts = [_forecast_from_row(row) for row in native_rows]
    allow_provisional = (
        spec.allow_provisional
        if spec.allow_provisional is not None
        else bool(policy.get("allow_provisional", True))
    )
    forbid_opposing = (
        spec.forbid_opposing
        if spec.forbid_opposing is not None
        else bool(policy.get("forbid_conflicting_game_sides", False))
    )
    limit = spec.limit if spec.limit is not None else int(policy.get("limit") or 5)
    min_leg_odds = (
        spec.min_leg_odds if spec.min_leg_odds is not None else policy.get("min_leg_odds")
    )
    max_leg_odds = (
        spec.max_leg_odds if spec.max_leg_odds is not None else policy.get("max_leg_odds")
    )
    price_weight = (
        spec.price_weight if spec.price_weight is not None else policy.get("price_weight")
    )
    distinct_games = len({(fc.league, fc.game_id) for fc in native_forecasts})
    if spec.legs is not None:
        effective_legs: Optional[int] = spec.legs if distinct_games >= spec.legs else None
    else:
        requested = int(policy.get("requested_legs") or 3)
        live_eff = min(requested, distinct_games)
        effective_legs = live_eff if live_eff >= 2 else None

    tickets: List[Dict[str, Any]] = []
    if effective_legs is not None:
        accas = build_accas(
            native_forecasts,
            legs=effective_legs,
            limit=limit,
            allow_provisional=allow_provisional,
            require_disjoint_legs=bool(policy.get("require_disjoint_legs", True)),
            forbid_conflicting_game_sides=forbid_opposing,
            min_leg_odds=min_leg_odds,
            max_leg_odds=max_leg_odds,
            min_leg_probability=spec.min_leg_prob,
            price_weight=price_weight,
            min_combined_odds=spec.min_odds,
            max_combined_odds=spec.max_odds,
        )
        for acca in accas:
            tickets.append(
                {
                    "ticket_id": str(acca["acca_id"]),
                    "ticket_type": "acca",
                    "leg_count": int(acca["leg_count"]),
                    "combined_decimal_odds": float(acca["combined_decimal_odds"]),
                    "estimated_joint_win_probability": float(
                        acca["estimated_joint_win_probability"]
                    ),
                    "legs": list(acca["legs"]),
                }
            )

    if not tickets and spec.allow_single_fallback and len(native_forecasts) == 1:
        fc = native_forecasts[0]
        price = fc.decimal_price
        prob = fc.calibrated_probability
        lo = float(min_leg_odds) if min_leg_odds is not None else 1.0
        hi = float(max_leg_odds) if max_leg_odds is not None else float("inf")
        if (
            price is not None
            and prob is not None
            and lo <= price <= hi
            and (spec.min_leg_prob is None or prob >= spec.min_leg_prob)
            and (spec.min_odds is None or price >= spec.min_odds)
            and (spec.max_odds is None or price <= spec.max_odds)
        ):
            tickets.append(
                {
                    "ticket_id": fc.forecast_id,
                    "ticket_type": "single",
                    "leg_count": 1,
                    "combined_decimal_odds": round(float(price), 4),
                    "estimated_joint_win_probability": float(prob),
                    "legs": [fc.to_dict()],
                }
            )

    if derived_passthrough or spec.include_derived:
        for row in derived_rows:
            single = _derived_single_from_row(row)
            if single is not None:
                tickets.append(single)
    return tickets


def _receipt_ticket_signatures(receipt: Dict[str, Any]) -> List[Tuple[str, str, Tuple[str, ...]]]:
    sigs: List[Tuple[str, str, Tuple[str, ...]]] = []
    for collection, ttype in (("accas", "acca"), ("singles", "single")):
        for pos in receipt.get(collection) or ():
            if not isinstance(pos, dict):
                continue
            leg_ids = tuple(
                str(leg.get("forecast_id") or leg.get("pick_id") or "")
                for leg in (pos.get("legs") or ())
                if isinstance(leg, dict)
            )
            sigs.append((ttype, str(pos.get("ticket_id") or ""), leg_ids))
    return sigs


def _selected_ticket_signatures(tickets: Sequence[Dict[str, Any]]) -> List[Tuple[str, str, Tuple[str, ...]]]:
    sigs: List[Tuple[str, str, Tuple[str, ...]]] = []
    for ticket in tickets:
        leg_ids = tuple(
            str(leg.get("forecast_id") or leg.get("pick_id") or "")
            for leg in (ticket.get("legs") or ())
            if isinstance(leg, dict)
        )
        sigs.append((str(ticket["ticket_type"]), str(ticket["ticket_id"]), leg_ids))
    return sigs


def check_parity(days: Sequence[DayEvidence]) -> Dict[str, Any]:
    """Rebuild every pool-equipped receipt day through live ``build_accas`` and diff."""
    checked = 0
    matched = 0
    skipped_no_pool = 0
    mismatches: List[Dict[str, Any]] = []
    live_spec = ReplaySpec()
    for day in days:
        if not day.has_candidate_pool:
            skipped_no_pool += 1
            continue
        checked += 1
        rebuilt = select_tickets_for_pool(
            day.candidate_pool,
            policy=day.selection_policy,
            spec=live_spec,
            derived_passthrough=True,
        )
        expected_sig = _receipt_ticket_signatures(day.receipt)
        rebuilt_sig = _selected_ticket_signatures(rebuilt)
        if expected_sig == rebuilt_sig:
            matched += 1
        else:
            mismatches.append(
                {
                    "date": day.date,
                    "league": day.league,
                    "expected": expected_sig,
                    "rebuilt": rebuilt_sig,
                }
            )
    return {
        "checked": checked,
        "matched": matched,
        "mismatched": len(mismatches),
        "skipped_legacy_no_pool": skipped_no_pool,
        "mismatches": mismatches,
    }


def _grade_ticket(ticket: Dict[str, Any], outcomes_by_fid: Dict[str, str]) -> Dict[str, Any]:
    legs = ticket.get("legs") or []
    outcomes: List[Optional[str]] = []
    eff_odds = 1.0
    for leg in legs:
        fid = str(leg.get("forecast_id") or leg.get("pick_id") or "")
        outcome = outcomes_by_fid.get(fid)
        outcomes.append(outcome)
        if outcome != "push":
            price = _row_selected_price(leg)
            if price is not None:
                eff_odds *= price
    if any(o == "loss" for o in outcomes):
        status = "loss"
        eff = 0.0
    elif any(o is None for o in outcomes):
        status = "pending"
        eff = float(ticket.get("combined_decimal_odds") or eff_odds)
    elif outcomes and all(o == "push" for o in outcomes):
        status = "push"
        eff = 1.0
    else:
        status = "win"
        eff = eff_odds
    return {
        **ticket,
        "grade_status": status,
        "effective_decimal_odds": round(eff, 6),
        "leg_outcomes": outcomes,
    }


def _receipt_published_tickets(receipt: Dict[str, Any]) -> List[Dict[str, Any]]:
    tickets: List[Dict[str, Any]] = []
    for collection, ttype in (("accas", "acca"), ("singles", "single")):
        for pos in receipt.get(collection) or ():
            if not isinstance(pos, dict):
                continue
            tickets.append(
                {
                    "ticket_id": str(pos.get("ticket_id") or ""),
                    "ticket_type": ttype,
                    "leg_count": int(pos.get("leg_count") or len(pos.get("legs") or ())),
                    "combined_decimal_odds": float(pos.get("combined_decimal_odds") or 1.0),
                    "estimated_joint_win_probability": float(
                        pos.get("estimated_joint_win_probability")
                        or pos.get("estimated_win_probability")
                        or 0.0
                    ),
                    "legs": list(pos.get("legs") or ()),
                }
            )
    return tickets


def _allocate_stakes(
    tickets: Sequence[Dict[str, Any]],
    *,
    bank: float,
    spec: ReplaySpec,
) -> List[float]:
    n = len(tickets)
    if n == 0 or bank <= 0.0:
        return [0.0] * n
    if spec.stake_mode == "flat":
        return [1.0] * n
    if spec.stake_mode == "daily_pct":
        budget = bank * spec.stake_frac
        if spec.weights:
            raw_w = [
                spec.weights[i] if i < len(spec.weights) else spec.weights[-1]
                for i in range(n)
            ]
            total_w = sum(raw_w)
            return [budget * (w / total_w) for w in raw_w]
        return [budget / n] * n
    if spec.stake_mode == "per_acca_pct":
        raw = [bank * spec.stake_frac] * n
        cap = bank * 0.95
        total = sum(raw)
        if total > cap and total > 0:
            scale = cap / total
            return [s * scale for s in raw]
        return raw
    # stake_mode == "kelly": shrunk fractional Kelly capped at spec.stake_frac of bank/day
    fracs: List[float] = []
    for ticket in tickets:
        odds = float(ticket.get("combined_decimal_odds") or 0.0)
        p = float(ticket.get("estimated_joint_win_probability") or 0.0)
        if odds <= 1.0 or not 0.0 < p < 1.0:
            fracs.append(0.0)
            continue
        edge = p * odds - 1.0
        if edge <= 0.0:
            fracs.append(0.0)
            continue
        full_k = edge / (odds - 1.0)
        fracs.append(spec.kelly_frac * full_k)
    total_f = sum(fracs)
    if total_f > spec.stake_frac and total_f > 0:
        scale = spec.stake_frac / total_f
        fracs = [f * scale for f in fracs]
    return [bank * f for f in fracs]


def replay(
    days: Sequence[DayEvidence],
    spec: ReplaySpec,
    *,
    mode: str = "per-league",
    allow_legacy_sizing_only: bool = False,
    starting_bank: float = 100.0,
) -> Dict[str, Any]:
    """Replay ``spec`` over ``days`` and compute growth, drawdown, and slot stats."""
    if mode not in {"per-league", "cross-league"}:
        raise ValueError(f"invalid mode {mode!r}; expected 'per-league' or 'cross-league'")

    by_date: Dict[str, List[DayEvidence]] = {}
    skipped_no_pool = 0
    legacy_sizing_days = 0
    for day in days:
        if day.has_candidate_pool:
            by_date.setdefault(day.date, []).append(day)
        elif allow_legacy_sizing_only and mode == "per-league":
            legacy_sizing_days += 1
            by_date.setdefault(day.date, []).append(day)
        else:
            skipped_no_pool += 1

    bank = float(starting_bank)
    vault = 0.0
    peak_wealth = bank
    max_dd = 0.0
    wins = losses = pushes = 0
    flat_staked = 0.0
    flat_profit = 0.0
    pending_days = 0
    daily_records: List[Dict[str, Any]] = []
    slot_stats: Dict[int, Dict[str, Any]] = {}
    band_stats: Dict[str, Dict[str, Any]] = {
        label: {"wins": 0, "losses": 0, "pushes": 0, "staked": 0.0, "profit": 0.0}
        for label, _, _ in ODDS_BANDS
    }

    for target_date in sorted(by_date):
        entries = by_date[target_date]
        selected: List[Dict[str, Any]] = []
        merged_outcomes: Dict[str, str] = {}
        for entry in entries:
            merged_outcomes.update(entry.outcomes_by_fid)

        if mode == "cross-league":
            merged_pool: List[Dict[str, Any]] = []
            seen_fids: set[str] = set()
            for entry in entries:
                for row in entry.candidate_pool:
                    fid = str(row.get("forecast_id") or "")
                    if fid and fid not in seen_fids:
                        seen_fids.add(fid)
                        merged_pool.append(row)
            cross_policy = {
                "requested_legs": 3,
                "limit": 5,
                "allow_provisional": any(
                    bool(e.selection_policy.get("allow_provisional", False)) for e in entries
                ),
                "require_disjoint_legs": True,
                "forbid_conflicting_game_sides": True,
            }
            selected = select_tickets_for_pool(
                merged_pool,
                policy=cross_policy,
                spec=spec,
                derived_passthrough=False,
            )
        else:
            for entry in entries:
                if entry.has_candidate_pool:
                    league_tickets = select_tickets_for_pool(
                        entry.candidate_pool,
                        policy=entry.selection_policy,
                        spec=spec,
                        derived_passthrough=False,
                    )
                else:
                    league_tickets = _receipt_published_tickets(entry.receipt)
                for ticket in league_tickets:
                    ticket["league"] = entry.league
                selected.extend(league_tickets)

        if not selected:
            continue

        graded = [_grade_ticket(t, merged_outcomes) for t in selected]
        settled = [t for t in graded if t["grade_status"] in {"win", "loss", "push"}]
        if not settled:
            pending_days += 1
            continue

        stakes = _allocate_stakes(settled, bank=bank, spec=spec)
        day_staked = sum(stakes)
        day_returned = 0.0
        day_flat_pl = 0.0
        day_ticket_ids: List[str] = []
        for idx, (ticket, stake) in enumerate(zip(settled, stakes), 1):
            day_ticket_ids.append(str(ticket["ticket_id"]))
            status = ticket["grade_status"]
            eff = float(ticket["effective_decimal_odds"])
            if status == "win":
                wins += 1
                day_returned += stake * eff
                t_flat_pl = eff - 1.0
            elif status == "loss":
                losses += 1
                t_flat_pl = -1.0
            else:
                pushes += 1
                day_returned += stake
                t_flat_pl = 0.0
            flat_staked += 1.0
            flat_profit += t_flat_pl
            day_flat_pl += t_flat_pl

            slot = slot_stats.setdefault(
                idx, {"wins": 0, "losses": 0, "pushes": 0, "staked": 0.0, "profit": 0.0}
            )
            slot["wins" if status == "win" else "losses" if status == "loss" else "pushes"] += 1
            slot["staked"] += 1.0
            slot["profit"] += t_flat_pl

            comb_odds = float(ticket.get("combined_decimal_odds") or 1.0)
            for label, lo, hi in ODDS_BANDS:
                if lo <= comb_odds < hi:
                    b = band_stats[label]
                    b["wins" if status == "win" else "losses" if status == "loss" else "pushes"] += 1
                    b["staked"] += 1.0
                    b["profit"] += t_flat_pl
                    break

        wealth_before = bank + vault
        day_net = day_returned - day_staked
        if spec.skim > 0.0 and day_net > 0.0:
            skimmed = day_net * spec.skim
            vault += skimmed
            bank = max(1e-9, bank + day_net - skimmed)
        else:
            bank = max(1e-9, bank + day_net)
        wealth_after = bank + vault
        log_growth = (
            math.log(wealth_after / wealth_before)
            if wealth_before > 0.0 and wealth_after > 0.0
            else float("-inf")
        )
        peak_wealth = max(peak_wealth, wealth_after)
        dd = (wealth_after - peak_wealth) / peak_wealth if peak_wealth > 0 else 0.0
        max_dd = min(max_dd, dd)

        daily_records.append(
            {
                "date": target_date,
                "tickets": len(settled),
                "ticket_ids": tuple(day_ticket_ids),
                "flat_pl": day_flat_pl,
                "log_growth": log_growth,
                "bank_after": bank,
                "vault_after": vault,
                "wealth_after": wealth_after,
            }
        )

    settled_days = len(daily_records)
    mean_log_growth = (
        sum(d["log_growth"] for d in daily_records) / settled_days if settled_days else 0.0
    )
    return {
        "mode": mode,
        "settled_days": settled_days,
        "pending_days": pending_days,
        "skipped_legacy_no_pool": skipped_no_pool,
        "legacy_sizing_days": legacy_sizing_days,
        "low_sample": settled_days < LOW_SAMPLE_DAYS,
        "wins": wins,
        "losses": losses,
        "pushes": pushes,
        "hit_rate": wins / (wins + losses) if (wins + losses) else None,
        "flat_staked": flat_staked,
        "flat_profit": flat_profit,
        "flat_roi": flat_profit / flat_staked if flat_staked else None,
        "mean_log_growth_per_day": mean_log_growth,
        "starting_bank": starting_bank,
        "final_bank": bank,
        "skimmed_vault": vault,
        "total_wealth": bank + vault,
        "peak_wealth": peak_wealth,
        "max_drawdown": max_dd,
        "slots": slot_stats,
        "odds_bands": band_stats,
        "daily": daily_records,
    }


def compare_variants(
    days: Sequence[DayEvidence],
    spec_a: ReplaySpec,
    spec_b: ReplaySpec,
    *,
    mode: str = "per-league",
    allow_legacy_sizing_only: bool = False,
    bootstraps: int = 2000,
    seed: int = 42,
) -> Dict[str, Any]:
    """Paired day-block bootstrap A/B with mandatory NO-OP selection diff check."""
    res_a = replay(days, spec_a, mode=mode, allow_legacy_sizing_only=allow_legacy_sizing_only)
    res_b = replay(days, spec_b, mode=mode, allow_legacy_sizing_only=allow_legacy_sizing_only)
    map_a = {d["date"]: d for d in res_a["daily"]}
    map_b = {d["date"]: d for d in res_b["daily"]}
    all_dates = sorted(set(map_a) | set(map_b))

    selection_diff_days = sum(
        (map_a.get(d, {}).get("ticket_ids") or ()) != (map_b.get(d, {}).get("ticket_ids") or ())
        for d in all_dates
    )
    sizing_same = (
        spec_a.stake_mode == spec_b.stake_mode
        and spec_a.stake_frac == spec_b.stake_frac
        and spec_a.kelly_frac == spec_b.kelly_frac
        and spec_a.weights == spec_b.weights
        and spec_a.skim == spec_b.skim
    )
    if not all_dates:
        return {
            "status": "insufficient_data",
            "n_days": 0,
            "selection_diff_days": 0,
            "a": res_a,
            "b": res_b,
        }
    if selection_diff_days == 0 and sizing_same:
        return {
            "status": "no_op",
            "n_days": len(all_dates),
            "selection_diff_days": 0,
            "a": res_a,
            "b": res_b,
        }

    deltas_g = [
        float(map_b.get(d, {}).get("log_growth", 0.0))
        - float(map_a.get(d, {}).get("log_growth", 0.0))
        for d in all_dates
    ]
    deltas_pl = [
        float(map_b.get(d, {}).get("flat_pl", 0.0))
        - float(map_a.get(d, {}).get("flat_pl", 0.0))
        for d in all_dates
    ]
    n = len(all_dates)
    rng = random.Random(seed)
    boot_g: List[float] = []
    boot_pl: List[float] = []
    for _ in range(bootstraps):
        idxs = [rng.randrange(n) for _ in range(n)]
        boot_g.append(sum(deltas_g[i] for i in idxs) / n)
        boot_pl.append(sum(deltas_pl[i] for i in idxs))
    boot_g.sort()
    boot_pl.sort()
    lo_i = max(0, int(0.025 * bootstraps))
    hi_i = min(bootstraps - 1, int(0.975 * bootstraps))
    return {
        "status": "sizing_only" if selection_diff_days == 0 else "compared",
        "n_days": n,
        "low_sample": n < LOW_SAMPLE_DAYS,
        "selection_diff_days": selection_diff_days,
        "delta_mean_log_growth": sum(deltas_g) / n,
        "delta_log_growth_ci95": (boot_g[lo_i], boot_g[hi_i]),
        "prob_b_beats_a_log_growth": sum(x > 0 for x in boot_g) / bootstraps,
        "delta_flat_profit": sum(deltas_pl),
        "delta_flat_profit_ci95": (boot_pl[lo_i], boot_pl[hi_i]),
        "a": res_a,
        "b": res_b,
    }


def kelly_sweep(
    days: Sequence[DayEvidence],
    base_spec: ReplaySpec,
    *,
    mode: str = "per-league",
    allow_legacy_sizing_only: bool = False,
    fractions: Sequence[float] = (0.05, 0.10, 0.15, 0.20, 0.25, 0.30, 0.35, 0.40),
) -> List[Dict[str, Any]]:
    """Sweep daily bankroll deployment percentages for growth vs drawdown."""
    rows: List[Dict[str, Any]] = []
    for frac in fractions:
        spec = replace(base_spec, stake_mode="daily_pct", stake_frac=float(frac))
        res = replay(days, spec, mode=mode, allow_legacy_sizing_only=allow_legacy_sizing_only)
        rows.append(
            {
                "daily_pct": frac,
                "settled_days": res["settled_days"],
                "mean_log_growth_per_day": res["mean_log_growth_per_day"],
                "total_wealth": res["total_wealth"],
                "peak_wealth": res["peak_wealth"],
                "max_drawdown": res["max_drawdown"],
            }
        )
    return rows


__all__ = [
    "DayEvidence",
    "LOW_SAMPLE_DAYS",
    "ReplaySpec",
    "check_parity",
    "compare_variants",
    "kelly_sweep",
    "load_localdata_evidence",
    "parse_variant_spec",
    "replay",
    "select_tickets_for_pool",
]
