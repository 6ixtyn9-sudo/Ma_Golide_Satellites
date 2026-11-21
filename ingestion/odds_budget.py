"""Fail-closed, Git-durable usage reservations for The Odds API.

The provider's free allowance is shared by every Actions matrix runner.  A
process-local counter cannot protect that allowance: parallel runners do not
share memory and a failed publisher must not make the same price scan
unreserved again.  This module deliberately reserves the *maximum* cost before
credentials reach a runner, then keeps that reservation even if the runner
fails.  That is conservative by design.

Routine permits cover the three documented featured full-game markets from one
effective bookmaker group: ``h2h,totals,spreads``.  Native period-market calls
are intentionally outside this initial free-tier envelope.
"""

from __future__ import annotations

from copy import deepcopy
from datetime import datetime, timezone
import hashlib
import json
from pathlib import Path
from typing import Any, Iterable
from zoneinfo import ZoneInfo

BUDGET_SCHEMA = "magolide.theoddsapi_budget"
PERMIT_SCHEMA = "magolide.theoddsapi_permit"
PROVIDER = "theoddsapi"
BUDGET_TIMEZONE = "Africa/Johannesburg"

# 500 is the advertised Starter allowance.  The controller intentionally only
# reserves 360 credits, preserving 140 credits for unexpected external use,
# provider-side discrepancies, and an operator-approved credentialed smoke.
PROVIDER_MONTHLY_ALLOWANCE = 500
MONTHLY_RESERVATION_CAP = 360
HARD_RESERVE = PROVIDER_MONTHLY_ALLOWANCE - MONTHLY_RESERVATION_CAP
DAILY_RESERVATION_CAP = 12
ROUTINE_FULL_GAME_COST = 3
ROUTINE_MARKETS = ("h2h", "totals", "spreads")

# Direct, validated price-lane candidates only.  EuroLeague's documented sport
# key is not enough to promote it: its account/region/bookmaker coverage has
# not passed the required bounded live smoke.  Australia NBL remains
# probationary and has no compliant automated schedule contract.
ROUTINE_LEAGUE_PRIORITY = ("NCAA_W", "NBA", "NCAA_M", "WNBA")


class BudgetStateError(ValueError):
    """The durable reservation state is malformed or cannot be safely used."""


def now_utc_iso() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).strftime("%Y-%m-%dT%H:%M:%SZ")


def _parse_utc(value: str) -> datetime:
    if not isinstance(value, str) or not value:
        raise BudgetStateError("timestamp is missing")
    try:
        parsed = datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError as exc:
        raise BudgetStateError("timestamp is invalid") from exc
    if parsed.tzinfo is None:
        raise BudgetStateError("timestamp has no timezone")
    return parsed.astimezone(timezone.utc)


def _budget_period(now: datetime) -> tuple[str, str]:
    """Return the provider-month and daily-cap key in the operator timezone."""

    if now.tzinfo is None:
        raise ValueError("budget clock must be timezone-aware")
    local = now.astimezone(ZoneInfo(BUDGET_TIMEZONE))
    return local.strftime("%Y-%m"), local.strftime("%Y-%m-%d")


def empty_state(now: datetime) -> dict[str, Any]:
    month, _day = _budget_period(now)
    return {
        "schema_version": BUDGET_SCHEMA,
        "provider": PROVIDER,
        "budget_timezone": BUDGET_TIMEZONE,
        "quota_month": month,
        "provider_monthly_allowance": PROVIDER_MONTHLY_ALLOWANCE,
        "monthly_reservation_cap": MONTHLY_RESERVATION_CAP,
        "hard_reserve": HARD_RESERVE,
        "daily_reservation_cap": DAILY_RESERVATION_CAP,
        "routine_full_game_cost": ROUTINE_FULL_GAME_COST,
        "reservations": [],
        "last_provider_quota": {},
    }


def _as_int(value: Any, *, field: str, minimum: int = 0) -> int:
    if isinstance(value, bool):
        raise BudgetStateError(f"{field} must be an integer")
    try:
        parsed = int(value)
    except (TypeError, ValueError) as exc:
        raise BudgetStateError(f"{field} must be an integer") from exc
    if parsed < minimum:
        raise BudgetStateError(f"{field} is below its minimum")
    return parsed


def _validate_reservation(value: Any, *, month: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise BudgetStateError("reservation is not an object")
    required = {
        "permit_id",
        "run_id",
        "league",
        "target_date",
        "budget_date",
        "issued_utc",
        "maximum_credits",
        "markets",
    }
    if set(value) != required:
        raise BudgetStateError("reservation has an unknown schema")
    if not all(isinstance(value[key], str) and value[key] for key in required - {"markets", "maximum_credits"}):
        raise BudgetStateError("reservation has invalid identity fields")
    if value["league"] not in ROUTINE_LEAGUE_PRIORITY:
        raise BudgetStateError("reservation has an unapproved league")
    try:
        target = datetime.strptime(value["target_date"], "%Y-%m-%d").date()
        budget = datetime.strptime(value["budget_date"], "%Y-%m-%d").date()
    except ValueError as exc:
        raise BudgetStateError("reservation target/budget date is invalid") from exc
    if target.isoformat() != value["target_date"] or budget.isoformat() != value["budget_date"]:
        raise BudgetStateError("reservation target/budget date is invalid")
    if not value["budget_date"].startswith(month):
        raise BudgetStateError("reservation belongs to another quota month")
    _parse_utc(value["issued_utc"])
    if _as_int(value["maximum_credits"], field="reservation maximum credits") != ROUTINE_FULL_GAME_COST:
        raise BudgetStateError("reservation has an unsupported maximum cost")
    if tuple(value["markets"]) != ROUTINE_MARKETS:
        raise BudgetStateError("reservation has unsupported markets")
    return dict(value)


def validate_state(value: Any, now: datetime) -> dict[str, Any]:
    """Validate state, rolling safely to a fresh month when appropriate."""

    if value is None:
        return empty_state(now)
    if not isinstance(value, dict):
        raise BudgetStateError("budget state is not an object")
    expected = {
        "schema_version",
        "provider",
        "budget_timezone",
        "quota_month",
        "provider_monthly_allowance",
        "monthly_reservation_cap",
        "hard_reserve",
        "daily_reservation_cap",
        "routine_full_game_cost",
        "reservations",
        "last_provider_quota",
    }
    if set(value) != expected:
        raise BudgetStateError("budget state has an unknown schema")
    current_month, _day = _budget_period(now)
    if (
        value["schema_version"] != BUDGET_SCHEMA
        or value["provider"] != PROVIDER
        or value["budget_timezone"] != BUDGET_TIMEZONE
    ):
        raise BudgetStateError("budget state has invalid fields")
    # A structurally recognised prior-month ledger cannot authorise a current
    # request. Start a new period rather than carrying reservations over the
    # provider's first-of-month reset.
    if value["quota_month"] != current_month:
        return empty_state(now)
    if (
        _as_int(value["provider_monthly_allowance"], field="provider allowance")
        != PROVIDER_MONTHLY_ALLOWANCE
        or _as_int(value["monthly_reservation_cap"], field="monthly cap")
        != MONTHLY_RESERVATION_CAP
        or _as_int(value["hard_reserve"], field="hard reserve") != HARD_RESERVE
        or _as_int(value["daily_reservation_cap"], field="daily cap")
        != DAILY_RESERVATION_CAP
        or _as_int(value["routine_full_game_cost"], field="routine cost")
        != ROUTINE_FULL_GAME_COST
        or not isinstance(value["reservations"], list)
        or not isinstance(value["last_provider_quota"], dict)
    ):
        raise BudgetStateError("budget state has invalid fields")
    reservations = [_validate_reservation(item, month=current_month) for item in value["reservations"]]
    identities = {(item["league"], item["target_date"]) for item in reservations}
    if len(identities) != len(reservations):
        raise BudgetStateError("budget state has duplicate league/date reservations")
    ids = {item["permit_id"] for item in reservations}
    if len(ids) != len(reservations):
        raise BudgetStateError("budget state has duplicate permit identifiers")
    if sum(item["maximum_credits"] for item in reservations) > MONTHLY_RESERVATION_CAP:
        raise BudgetStateError("budget state exceeds the monthly reservation cap")
    quota = dict(value["last_provider_quota"])
    for key in ("x-requests-remaining", "x-requests-used", "x-requests-last"):
        if key in quota:
            quota[key] = _as_int(quota[key], field=key)
    if quota and not isinstance(quota.get("observed_utc"), str):
        raise BudgetStateError("provider quota observation has no timestamp")
    return {**value, "reservations": reservations, "last_provider_quota": quota}


def load_state(path: str | Path, now: datetime) -> dict[str, Any]:
    location = Path(path)
    if not location.exists():
        return empty_state(now)
    try:
        raw = json.loads(location.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise BudgetStateError("budget state is unreadable") from exc
    return validate_state(raw, now)


def write_state(path: str | Path, state: dict[str, Any]) -> None:
    location = Path(path)
    location.parent.mkdir(parents=True, exist_ok=True)
    temporary = location.with_name(location.name + ".tmp")
    temporary.write_text(json.dumps(state, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    temporary.replace(location)


def _due_matrix_rows(matrix: Any) -> list[dict[str, str]]:
    if not isinstance(matrix, dict) or not isinstance(matrix.get("include"), list):
        raise BudgetStateError("gate matrix has an invalid schema")
    candidates: list[dict[str, str]] = []
    for row in matrix["include"]:
        if not isinstance(row, dict):
            raise BudgetStateError("gate matrix has an invalid row")
        league = str(row.get("league") or "")
        target_date = str(row.get("target_date") or "")
        reasons = str(row.get("gate_reason") or "")
        if league not in ROUTINE_LEAGUE_PRIORITY:
            continue
        # Draft, settlement, uncertain-state, and results-only work must never
        # be enough to spend a price credit. The adaptive gate names the one
        # due reason that represents the dated ticket acquisition window.
        if "ticket_creation_or_acquisition_due" not in reasons.split("; "):
            continue
        try:
            target = datetime.strptime(target_date, "%Y-%m-%d").date()
        except ValueError as exc:
            raise BudgetStateError("gate matrix has an invalid target date") from exc
        candidates.append({"league": league, "target_date": target.isoformat()})
    priority = {league: index for index, league in enumerate(ROUTINE_LEAGUE_PRIORITY)}
    return sorted(candidates, key=lambda item: (priority[item["league"]], item["target_date"]))


def _permit_id(run_id: str, league: str, target_date: str) -> str:
    identity = f"{run_id}|{league}|{target_date}|{ROUTINE_FULL_GAME_COST}".encode("utf-8")
    return hashlib.sha256(identity).hexdigest()[:20]


def reserve(
    state: dict[str, Any],
    matrix: Any,
    *,
    run_id: str,
    now: datetime,
    automated_tick: bool,
) -> tuple[dict[str, Any], dict[str, dict[str, Any]]]:
    """Pre-reserve routine full-game credits and return runner permits.

    Manual dispatch deliberately receives no permit.  A low/malformed provider
    quota observation fails closed rather than assuming credits remain.
    """

    checked = validate_state(deepcopy(state), now)
    if not automated_tick:
        return checked, {}
    if not isinstance(run_id, str) or not run_id:
        raise BudgetStateError("run id is required")
    quota = checked["last_provider_quota"]
    remaining = quota.get("x-requests-remaining")
    if remaining is not None and (
        _as_int(remaining, field="x-requests-remaining") - ROUTINE_FULL_GAME_COST < HARD_RESERVE
    ):
        # Do not spend into the protected provider balance: a stale/local
        # reservation view must never turn a reported 141–142 credits into a
        # post-request balance below the 140-credit reserve.
        return checked, {}

    month, budget_date = _budget_period(now)
    assert checked["quota_month"] == month
    existing = {(item["league"], item["target_date"]) for item in checked["reservations"]}
    month_reserved = sum(item["maximum_credits"] for item in checked["reservations"])
    day_reserved = sum(
        item["maximum_credits"]
        for item in checked["reservations"]
        if item["budget_date"] == budget_date
    )
    permits: dict[str, dict[str, Any]] = {}
    issued = now.astimezone(timezone.utc).replace(microsecond=0).strftime("%Y-%m-%dT%H:%M:%SZ")
    for row in _due_matrix_rows(matrix):
        identity = (row["league"], row["target_date"])
        if identity in existing:
            continue
        if month_reserved + ROUTINE_FULL_GAME_COST > MONTHLY_RESERVATION_CAP:
            break
        if day_reserved + ROUTINE_FULL_GAME_COST > DAILY_RESERVATION_CAP:
            break
        permit = {
            "schema_version": PERMIT_SCHEMA,
            "permit_id": _permit_id(run_id, row["league"], row["target_date"]),
            "run_id": run_id,
            "league": row["league"],
            "target_date": row["target_date"],
            "issued_utc": issued,
            "maximum_credits": ROUTINE_FULL_GAME_COST,
            "markets": list(ROUTINE_MARKETS),
        }
        reservation = {
            "permit_id": permit["permit_id"],
            "run_id": run_id,
            "league": row["league"],
            "target_date": row["target_date"],
            "budget_date": budget_date,
            "issued_utc": issued,
            "maximum_credits": ROUTINE_FULL_GAME_COST,
            "markets": list(ROUTINE_MARKETS),
        }
        checked["reservations"].append(reservation)
        permits[f"{row['league']}|{row['target_date']}"] = permit
        existing.add(identity)
        month_reserved += ROUTINE_FULL_GAME_COST
        day_reserved += ROUTINE_FULL_GAME_COST
    return checked, permits


def validate_permit(value: Any, *, league: str, target_date: str) -> dict[str, Any]:
    """Validate an untrusted runner permit before a paid adapter request."""

    if not isinstance(value, dict):
        raise BudgetStateError("provider permit is not an object")
    expected = {
        "schema_version",
        "permit_id",
        "run_id",
        "league",
        "target_date",
        "issued_utc",
        "maximum_credits",
        "markets",
    }
    if set(value) != expected:
        raise BudgetStateError("provider permit has an unknown schema")
    if value.get("schema_version") != PERMIT_SCHEMA:
        raise BudgetStateError("provider permit has an unknown schema")
    if value.get("league") != league or value.get("target_date") != target_date:
        raise BudgetStateError("provider permit belongs to another league or date")
    try:
        permitted_date = datetime.strptime(str(value["target_date"]), "%Y-%m-%d").date()
    except ValueError as exc:
        raise BudgetStateError("provider permit target date is invalid") from exc
    if permitted_date.isoformat() != value["target_date"]:
        raise BudgetStateError("provider permit target date is invalid")
    if not all(isinstance(value[key], str) and value[key] for key in ("permit_id", "run_id", "issued_utc")):
        raise BudgetStateError("provider permit has invalid identity fields")
    _parse_utc(value["issued_utc"])
    if _as_int(value.get("maximum_credits"), field="provider permit maximum credits") != ROUTINE_FULL_GAME_COST:
        raise BudgetStateError("provider permit has an unsupported maximum cost")
    if tuple(value.get("markets") or ()) != ROUTINE_MARKETS:
        raise BudgetStateError("provider permit has unsupported markets")
    return dict(value)


def reconcile_provider_quota(
    state: dict[str, Any], quota_observations: Iterable[dict[str, Any]], *, now: datetime
) -> dict[str, Any]:
    """Persist the most conservative valid provider headers from completed jobs."""

    checked = validate_state(deepcopy(state), now)
    valid: list[dict[str, int]] = []
    for raw in quota_observations:
        if not isinstance(raw, dict):
            continue
        try:
            remaining = _as_int(raw.get("x-requests-remaining"), field="x-requests-remaining")
            used = _as_int(raw.get("x-requests-used"), field="x-requests-used")
            last = _as_int(raw.get("x-requests-last"), field="x-requests-last")
        except BudgetStateError:
            continue
        valid.append(
            {
                "x-requests-remaining": remaining,
                "x-requests-used": used,
                "x-requests-last": last,
            }
        )
    if not valid:
        return checked
    # Concurrent runners may see different header values. The lowest remaining
    # and highest used values are safe for the next reservation decision.
    checked["last_provider_quota"] = {
        "x-requests-remaining": min(item["x-requests-remaining"] for item in valid),
        "x-requests-used": max(item["x-requests-used"] for item in valid),
        "x-requests-last": max(item["x-requests-last"] for item in valid),
        "observed_utc": now.astimezone(timezone.utc).replace(microsecond=0).strftime(
            "%Y-%m-%dT%H:%M:%SZ"
        ),
    }
    return checked


__all__ = [
    "BUDGET_SCHEMA",
    "BUDGET_TIMEZONE",
    "DAILY_RESERVATION_CAP",
    "HARD_RESERVE",
    "MONTHLY_RESERVATION_CAP",
    "PERMIT_SCHEMA",
    "ROUTINE_FULL_GAME_COST",
    "ROUTINE_LEAGUE_PRIORITY",
    "ROUTINE_MARKETS",
    "BudgetStateError",
    "empty_state",
    "load_state",
    "now_utc_iso",
    "reconcile_provider_quota",
    "reserve",
    "validate_permit",
    "validate_state",
    "write_state",
]
