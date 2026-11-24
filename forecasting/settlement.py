"""Settled picks, read out of the ledgers the lanes already keep.

Why this exists
---------------
``forecasting.config_tuner`` analyses graded bets but was fed by hand. The lanes
already write graded picks: ``picks_YYYY-MM-DD.json`` (one list per slate, regraded
in place) and ``picks_audit_rolling.json`` (``{"settled_ledger": [...]}``). This
module turns those files into :class:`~forecasting.config_tuner.GradedBet` rows so
the tuning loop closes on real outcomes instead of a fixture somebody typed.

What it refuses to guess
------------------------
* A pick with no ``outcome`` is **pending**, not a loss. Pending rows are counted
  and skipped.
* A pick with no probability has nothing to calibrate against, so it is counted as
  ``no_probability`` rather than being given a confidence.
* The ledgers record ``market`` and ``period`` but **not** whether a moneyline pick
  was a banker or a robber -- the two were never distinguished at write time. Those
  rows are labelled ``MONEYLINE``, not assigned a kind they never had.
* The same settled pick can appear in a dated ledger and in the rolling audit. It is
  counted once; the duplicate is reported, because a double-counted pick quietly
  doubles its own influence on the gates.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, Iterable, List, Mapping, Optional, Tuple

import json

from forecasting.config_tuner import GradedBet

#: One list of picks per slate.
DATED_GLOB = "picks_????-??-??.json"

#: The rolling audit file, which nests its graded rows under ``settled_ledger``.
ROLLING_LEDGER = "picks_audit_rolling.json"

_WIN = "win"
_LOSS = "loss"
_PUSH = "push"

#: ``market``/``period`` as the ledger writes them -> the bet type the tuner labels.
_QUARTERS = {"q1", "q2", "q3", "q4"}


def bet_type_for(market: Any, period: Any) -> Optional[str]:
    """The label for a settled row, or ``None`` when the pair is not one we know.

    Labels follow the retired system's vocabulary. Moneylines are ``MONEYLINE``
    rather than ``BANKER``/``ROBBER`` because the ledger did not record which of the
    two it was; inventing one would corrupt every per-type count downstream.
    """

    market_text = str(market or "").strip().lower()
    period_text = str(period or "").strip().lower()

    if market_text == "highest_quarter":
        return "HIGHEST_QUARTER"
    if market_text == "moneyline":
        if period_text in _QUARTERS:
            return "SNIPER_MARGIN"
        if period_text in {"h1", "h2"}:
            return "FIRST_HALF_1X2"
        return "MONEYLINE"
    if market_text == "spread":
        if period_text in _QUARTERS:
            return "SNIPER_MARGIN"
        return "SPREAD"
    if market_text == "total":
        if period_text in _QUARTERS:
            return "SNIPER_OU"
        if period_text in {"h1", "h2"}:
            return "FIRST_HALF_OU"
        return "FT_OU"
    return None


def _probability(row: Mapping[str, Any]) -> Optional[float]:
    """``calibrated_probability`` first, else ``raw_probability`` -- the operator's order."""

    for field_name in ("calibrated_probability", "raw_probability"):
        value = row.get(field_name)
        if isinstance(value, bool) or not isinstance(value, (int, float)):
            continue
        if 0.0 <= float(value) <= 1.0:
            return float(value)
    return None


def _direction(row: Mapping[str, Any]) -> Optional[str]:
    selection = str(row.get("selection") or "").strip().upper()
    return selection if selection in {"OVER", "UNDER"} else None


def _outcome(row: Mapping[str, Any]) -> Optional[bool]:
    """``True`` win, ``False`` loss, ``None`` push... or raise ``KeyError`` when pending."""

    raw = str(row.get("outcome") or "").strip().lower()
    if raw == _WIN:
        return True
    if raw == _LOSS:
        return False
    if raw == _PUSH:
        return None
    raise KeyError("pending")


def _dedupe_key(row: Mapping[str, Any]) -> Tuple[str, ...]:
    return (
        str(row.get("game_id") or ""),
        str(row.get("market") or ""),
        str(row.get("period") or ""),
        str(row.get("selection") or ""),
        "" if row.get("line") is None else str(row.get("line")),
    )


@dataclass(frozen=True)
class SettlementSummary:
    """Graded bets plus an honest account of what was left out and why."""

    bets: Tuple[GradedBet, ...] = ()
    skipped: Mapping[str, int] = field(default_factory=dict)
    files: int = 0

    @property
    def total(self) -> int:
        return len(self.bets)

    @property
    def skipped_total(self) -> int:
        return sum(self.skipped.values())

    def skip_line(self) -> str:
        if not self.skipped:
            return "nothing skipped"
        return ", ".join(f"{reason} {count}" for reason, count in sorted(self.skipped.items()))


def _rows_from_file(path: Path) -> List[Mapping[str, Any]]:
    payload = json.loads(path.read_text())
    if isinstance(payload, dict):
        rows = payload.get("settled_ledger", payload.get("rows", []))
    else:
        rows = payload
    if not isinstance(rows, list):
        raise ValueError(f"{path}: expected a list of picks or a settled_ledger mapping")
    return [row for row in rows if isinstance(row, Mapping)]


def _ledger_files(targets: Iterable[Path | str]) -> Tuple[List[Path], List[str]]:
    files: List[Path] = []
    problems: List[str] = []
    for target in targets:
        path = Path(target)
        if path.is_dir():
            files.extend(sorted(path.glob(DATED_GLOB)))
            rolling = path / ROLLING_LEDGER
            if rolling.exists():
                files.append(rolling)
        elif path.is_file():
            files.append(path)
        else:
            problems.append(f"{path}: not found")
    return files, problems


def collect(*targets: Path | str) -> SettlementSummary:
    """Every settled pick under the given files or lanes, deduplicated.

    A missing path is reported in ``skipped["path_missing"]`` rather than raising:
    a lane with no history yet is a normal state, and the tuner handles an empty
    sample by changing nothing.
    """

    files, problems = _ledger_files(targets)
    skipped: Dict[str, int] = {}
    if problems:
        skipped["path_missing"] = len(problems)

    seen: set = set()
    bets: List[GradedBet] = []
    for path in files:
        try:
            rows = _rows_from_file(path)
        except (OSError, ValueError):
            skipped["unreadable_file"] = skipped.get("unreadable_file", 0) + 1
            continue
        for row in rows:
            key = _dedupe_key(row)
            if key in seen:
                skipped["duplicate"] = skipped.get("duplicate", 0) + 1
                continue
            seen.add(key)

            try:
                won = _outcome(row)
            except KeyError:
                skipped["pending"] = skipped.get("pending", 0) + 1
                continue

            probability = _probability(row)
            if probability is None:
                skipped["no_probability"] = skipped.get("no_probability", 0) + 1
                continue

            bet_type = bet_type_for(row.get("market"), row.get("period"))
            if bet_type is None:
                skipped["unmapped_market"] = skipped.get("unmapped_market", 0) + 1
                continue

            bets.append(
                GradedBet(
                    bet_type=bet_type,
                    league=str(row.get("league") or ""),
                    stated_confidence=probability * 100.0,
                    won=won,
                    direction=_direction(row),
                )
            )

    return SettlementSummary(tuple(bets), skipped, len(files))


__all__ = [
    "DATED_GLOB",
    "ROLLING_LEDGER",
    "SettlementSummary",
    "bet_type_for",
    "collect",
]
