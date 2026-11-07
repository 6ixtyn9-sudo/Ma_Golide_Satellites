"""Non-destructive, batched, idempotent Google Sheets writer.

Safety requirements implemented here (master prompt section 10):

- Dry-run is the default: ``plan_*`` functions return an :class:`UpdatePlan`
  with exact proposed cell changes and never touch a worksheet.  Writing
  requires ``allow_writes=True`` AND the caller's explicit opt-in env flag.
- ``clearContents()`` is never called; the whole used range is never
  rewritten.  Only cells inside the owned base columns (A..R) that actually
  change are written, batched per row.
- Columns S..V (t2-q1..t2-q4) and every column after V are protected: the
  plan cannot contain a change outside owned columns, and appended rows pad
  with blanks there.
- Rows are matched by canonical fixture identity (display date + home +
  away), never by row number alone.  Duplicate identities fail before any
  write.
- Formula guard: a target run whose current value is a formula (``=...``)
  is refused — acquisition never replaces formulas with static values.
- Idempotency: the second identical run produces an empty diff (and no new
  appends, because appended rows are matched on the next pass).
- Google calls are rate-limited (1.1s convention) with an injectable sleep.
"""

from __future__ import annotations

import logging
import time
from dataclasses import dataclass, field
from typing import Any, Callable, Dict, Iterable, List, Optional, Sequence, Tuple

from .identity import canonical_team_key
from .sheets_contract import (
    COL,
    OWNED_COLUMNS,
    ContractError,
    SheetContract,
    detect_contract,
)

logger = logging.getLogger(__name__)

RATE_LIMIT_DELAY = 1.1


def column_letter(col: int) -> str:
    """1 -> A, 26 -> Z, 27 -> AA."""
    out = ""
    while col > 0:
        col, rem = divmod(col - 1, 26)
        out = chr(65 + rem) + out
    return out


# ---------------------------------------------------------------------------
# Worksheet adapter (thin; lets tests use a fake without gspread)
# ---------------------------------------------------------------------------


class WorksheetLike:
    """Minimal worksheet surface the writer needs.

    Implementations: :class:`FakeWorksheet` (tests) and
    :class:`GspreadWorksheetAdapter` (production; gspread imported lazily).
    """

    def get_all_values(self) -> List[List[Any]]:  # pragma: no cover - interface
        raise NotImplementedError

    def batch_update_ranges(self, updates: List[Tuple[str, List[List[Any]]]]) -> None:
        """``[(A1_range, rows_matrix), ...]`` — one HTTP call per invocation
        is preferred; implementors may coalesce."""
        raise NotImplementedError

    def append_rows(self, rows: List[List[Any]]) -> None:
        raise NotImplementedError


class GspreadWorksheetAdapter(WorksheetLike):
    """Adapter for a real ``gspread`` worksheet (lazy import)."""

    def __init__(self, ws: Any) -> None:
        import gspread  # noqa: F401  (availability checked lazily)
        from gspread.utils import DateTimeOption, ValueRenderOption

        self._ws = ws
        self._value_render_option = ValueRenderOption.formula
        self._date_time_render_option = DateTimeOption.formatted_string

    def get_all_values(self) -> List[List[Any]]:
        # FORMULA is required for the no-overwrite guard; formatted date/time
        # strings keep row identity readable instead of returning serials.
        return self._ws.get_all_values(
            value_render_option=self._value_render_option,
            date_time_render_option=self._date_time_render_option,
        )

    def batch_update_ranges(self, updates: List[Tuple[str, List[List[Any]]]]) -> None:
        if not updates:
            return
        # gspread's batch_update takes a single body; coalesce per call.
        body = {
            "value_input_option": "USER_ENTERED",
            "data": [
                {
                    "range": rng,
                    "values": [[_jsonable(value) for value in row] for row in matrix],
                }
                for rng, matrix in updates
            ],
        }
        self._ws.spreadsheet.values_batch_update(body)

    def append_rows(self, rows: List[List[Any]]) -> None:
        if rows:
            serializable = [[_jsonable(value) for value in row] for row in rows]
            self._ws.append_rows(serializable, value_input_option="USER_ENTERED")


# ---------------------------------------------------------------------------
# Plan model
# ---------------------------------------------------------------------------


@dataclass
class CellChange:
    row: int
    col: int
    old: Any
    new: Any

    @property
    def a1(self) -> str:
        return f"{column_letter(self.col)}{self.row}"

    def to_dict(self) -> Dict[str, Any]:
        return {"cell": self.a1, "row": self.row, "col": self.col,
                "old": _jsonable(self.old), "new": _jsonable(self.new)}


def _jsonable(v: Any) -> Any:
    if hasattr(v, "isoformat"):
        return v.isoformat()
    if isinstance(v, (str, int, float, bool)) or v is None:
        return v
    return str(v)


@dataclass
class UpdatePlan:
    tab: str
    changes: List[CellChange] = field(default_factory=list)
    appends: List[List[Any]] = field(default_factory=list)
    skipped_formulas: List[str] = field(default_factory=list)
    notes: List[str] = field(default_factory=list)

    @property
    def is_noop(self) -> bool:
        return not self.changes and not self.appends

    def to_dict(self) -> Dict[str, Any]:
        return {
            "tab": self.tab,
            "changed_cells": len(self.changes),
            "appended_rows": len(self.appends),
            "changes": [c.to_dict() for c in self.changes],
            "skipped_formulas": self.skipped_formulas,
            "notes": self.notes,
        }


def _parse_ambiguous(raw: Any) -> set:
    """Date variants for a Date cell: day-first AND month-first readings.

    Legacy satellite display strings have unreliable day/month order
    (``11/01/2026`` in the reference workbook is 11 January, day-first,
    while US convention would read it as 1 November).  The prompt forbids
    inferring the order, so identity accepts BOTH interpretations plus
    unambiguous values (datetime/date objects, ISO strings).  Combined with
    exact home/away names this stays unique within a one-league-day
    satellite and never silently joins two different fixtures.
    """
    import datetime as _dt

    text = str(raw or "").strip()
    if not text:
        return set()
    if isinstance(raw, _dt.datetime):
        return {raw.date().isoformat()}
    if isinstance(raw, _dt.date):
        return {raw.isoformat()}
    try:
        return {_dt.date.fromisoformat(text[:10]).isoformat()}
    except ValueError:
        pass
    variants: set = set()
    parts = text.replace("/", ".").replace("-", ".").split(".")
    if len(parts) == 3 and all(p.isdigit() for p in parts):
        a, b, y = (int(p) for p in parts)
        if len(str(y)) == 4:
            for month, day in ((a, b), (b, a)):
                try:
                    variants.add(_dt.date(y, month, day).isoformat())
                except ValueError:
                    pass
    return variants


def _row_key(row_values: Sequence[Any], contract: SheetContract) -> Tuple[frozenset, str, str]:
    """Identity key from a sheet row: date variants + home + away.

    The date component is a SET of plausible interpretations (day-first and
    month-first) because legacy display order is unreliable; teams stay
    exact, so identity is still never team-only.
    """
    home = canonical_team_key(str(row_values[COL["Home"] - 1] or ""))
    away = canonical_team_key(str(row_values[COL["Away"] - 1] or ""))
    date_variants = frozenset(_parse_ambiguous(row_values[COL["Date"] - 1]))
    return (date_variants, home, away)


def _rendered_key(rendered: Dict[int, Any]) -> Tuple[frozenset, str, str]:
    date_val = rendered[COL["Date"]]
    if hasattr(date_val, "isoformat"):
        date_variants = frozenset({date_val.isoformat()})
    else:
        date_variants = frozenset(_parse_ambiguous(date_val))
    return (
        date_variants,
        canonical_team_key(str(rendered[COL["Home"]] or "")),
        canonical_team_key(str(rendered[COL["Away"]] or "")),
    )


def _values_equal(a: Any, b: Any) -> bool:
    """Sheet-tolerant equality: '' ~ None, numbers compare numerically."""
    a_s = "" if a is None else str(a).strip()
    b_s = "" if b is None else str(b).strip()
    if a_s == b_s:
        return True
    fa, fb = _to_float(a), _to_float(b)
    if fa is not None and fb is not None:
        return abs(fa - fb) < 1e-9
    return False


def _to_float(v: Any) -> Optional[float]:
    try:
        return float(str(v).strip())
    except (TypeError, ValueError):
        return None


# ---------------------------------------------------------------------------
# Planning
# ---------------------------------------------------------------------------


def plan_update(
    tab: str,
    worksheet: WorksheetLike,
    contract: SheetContract,
    rendered_rows: List[Dict[int, Any]],
    tab_kind: str,
) -> UpdatePlan:
    """Compute the exact non-destructive update plan for one clean tab.

    ``rendered_rows`` values are keyed by 1-based column index and cover
    ONLY owned columns (A..R).  Any attempt to render a protected column
    raises ``ContractError`` before any read/write happens.
    """
    for row in rendered_rows:
        bad = [c for c in row if c not in OWNED_COLUMNS]
        if bad:
            raise ContractError(
                f"{tab}: rendered row touches non-owned column(s) "
                f"{[column_letter(c) for c in bad]}; S:V and columns after V "
                "are protected"
            )

    all_values = worksheet.get_all_values()
    if not all_values:
        raise ContractError(f"{tab}: sheet has no header row")

    contract_live = detect_contract(all_values[0], tab)
    if contract_live.headers[:22] != contract.headers[:22]:
        raise ContractError(f"{tab}: header drift between snapshot and live sheet")

    # Index existing rows by (date_variants, home, away). Two rows match
    # when teams are equal AND their plausible date interpretations
    # intersect; an ambiguous match between two different existing rows
    # refuses to write.
    index: List[Tuple[frozenset, str, str, int]] = []
    for i, row_values in enumerate(all_values[1:], start=2):
        if not any(str(v or "").strip() for v in row_values):
            continue
        variants, home, away = _row_key(row_values, contract_live)
        if (home or away) and variants:
            for other in index:
                if other[1] == home and other[2] == away and (other[0] & variants):
                    raise ContractError(
                        f"{tab}: ambiguous duplicate rows {other[3]} and {i} "
                        f"for {home} vs {away}; refusing to write"
                    )
            index.append((variants, home, away, i))

    def _find(key: Tuple[frozenset, str, str]) -> Optional[int]:
        variants, home, away = key
        matches = [
            i for (v, h, a, i) in index
            if h == home and a == away and (v & variants)
        ]
        if len(matches) > 1:
            raise ContractError(
                f"{tab}: rendered identity matches multiple sheet rows "
                f"{matches}; refusing to choose one"
            )
        return matches[0] if matches else None

    plan = UpdatePlan(tab=tab)
    seen_new: set = set()
    claimed_rows: Dict[int, Tuple[frozenset, str, str]] = {}
    width = len(all_values[0]) if all_values else 22

    for rendered in rendered_rows:
        key = _rendered_key(rendered)
        if key in seen_new:
            raise ContractError(f"{tab}: duplicate rendered identity {key}")
        seen_new.add(key)

        existing_idx = _find(key)
        if existing_idx is not None and existing_idx in claimed_rows:
            raise ContractError(
                f"{tab}: sheet row {existing_idx} is claimed by both "
                f"{claimed_rows[existing_idx]} and {key}; ambiguous legacy "
                "date/identity, refusing to write"
            )
        if existing_idx is not None:
            claimed_rows[existing_idx] = key
        if existing_idx is None:
            # Append: owned values + blanks up to the sheet width so
            # protected columns on NEW rows start empty (nothing to
            # preserve there yet).
            row_out = ["" for _ in range(max(width, 22))]
            for col, val in sorted(rendered.items()):
                row_out[col - 1] = val
            plan.appends.append(row_out)
            continue

        current = all_values[existing_idx - 1]
        changed: Dict[int, Any] = {}
        for col, val in sorted(rendered.items()):
            old = current[col - 1] if col - 1 < len(current) else None
            if not _values_equal(old, val):
                changed[col] = val
        if not changed:
            continue

        # Contiguous runs of changed cells; refuse runs containing formulas.
        cols = sorted(changed)
        runs: List[List[int]] = [[cols[0]]]
        for c in cols[1:]:
            if c == runs[-1][-1] + 1:
                runs[-1].append(c)
            else:
                runs.append([c])
        for run in runs:
            if any(
                str(current[c - 1] if c - 1 < len(current) else "").strip().startswith("=")
                for c in run
            ):
                plan.skipped_formulas.append(
                    f"row {existing_idx} cols {column_letter(run[0])}-{column_letter(run[-1])}"
                )
                plan.notes.append(
                    f"row {existing_idx}: formula in owned columns preserved; "
                    "values not overwritten"
                )
                continue
            for c in run:
                plan.changes.append(
                    CellChange(
                        row=existing_idx,
                        col=c,
                        old=current[c - 1] if c - 1 < len(current) else None,
                        new=changed[c],
                    )
                )
    return plan


def apply_plan(
    plan: UpdatePlan,
    worksheet: WorksheetLike,
    allow_writes: bool,
    sleep: Callable[[float], None] = time.sleep,
    rate_limit_delay: float = RATE_LIMIT_DELAY,
) -> Dict[str, int]:
    """Execute a plan. Refuses unless ``allow_writes`` is True.

    Batching: one batched range-update call per affected row (runs already
    grouped), one call for appends; a rate-limit sleep between calls.
    """
    if not allow_writes:
        raise PermissionError(
            "allow_writes=False: dry run only. Set ALLOW_SHEET_WRITES=true and "
            "pass the write flag explicitly to enable sheet writes."
        )
    counts = {"changed_cells": 0, "appended_rows": 0, "calls": 0}
    if plan.changes:
        by_row: Dict[int, List[CellChange]] = {}
        for ch in plan.changes:
            by_row.setdefault(ch.row, []).append(ch)
        updates: List[Tuple[str, List[List[Any]]]] = []
        for row, changes in sorted(by_row.items()):
            col_min = min(c.col for c in changes)
            col_max = max(c.col for c in changes)
            matrix = [[None] * (col_max - col_min + 1)]
            lookup = {c.col: c.new for c in changes}
            for j, c in enumerate(range(col_min, col_max + 1)):
                matrix[0][j] = lookup.get(c, "")
            rng = f"{column_letter(col_min)}{row}:{column_letter(col_max)}{row}"
            updates.append((rng, matrix))
        worksheet.batch_update_ranges(updates)
        counts["calls"] += 1
        sleep(rate_limit_delay)
        counts["changed_cells"] = len(plan.changes)
    if plan.appends:
        worksheet.append_rows(plan.appends)
        counts["calls"] += 1
        sleep(rate_limit_delay)
        counts["appended_rows"] = len(plan.appends)
    return counts


# ---------------------------------------------------------------------------
# Fake worksheet (tests + dry-run harness)
# ---------------------------------------------------------------------------


class FakeWorksheet(WorksheetLike):
    """In-memory worksheet supporting values, formulas and notes."""

    def __init__(self, headers: Sequence[str], rows: Optional[List[List[Any]]] = None) -> None:
        self.headers = list(headers)
        self._rows: List[List[Any]] = [list(r) for r in (rows or [])]
        self.notes: Dict[Tuple[int, int], str] = {}
        self.update_calls: List[Tuple[str, List[List[Any]]]] = []
        self.append_calls: List[List[List[Any]]] = []

    def get_all_values(self) -> List[List[Any]]:
        return [list(self.headers)] + [list(r) for r in self._rows]

    def batch_update_ranges(self, updates: List[Tuple[str, List[List[Any]]]]) -> None:
        self.update_calls.extend(updates)
        for rng, matrix in updates:
            self._apply_range(rng, matrix)

    def _apply_range(self, rng: str, matrix: List[List[Any]]) -> None:
        import re

        m = re.match(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", rng)
        if not m:
            raise ValueError(f"bad range {rng}")
        c1 = _letter_index(m.group(1))
        r1 = int(m.group(2))
        for i, row_vals in enumerate(matrix):
            sheet_row = r1 + i  # 1-based; row 1 is the header row
            if sheet_row == 1:
                # Header rewrite (used by the data store table bootstrap).
                for j, val in enumerate(row_vals):
                    if val is not None and j < len(self.headers):
                        self.headers[j] = val
                continue
            row_idx = sheet_row - 2  # 0-based index into data rows
            while len(self._rows) <= row_idx:
                self._rows.append([""] * max(len(self.headers), 22))
            for j, val in enumerate(row_vals):
                if val is not None:
                    col_idx = c1 - 1 + j
                    while len(self._rows[row_idx]) <= col_idx:
                        self._rows[row_idx].append("")
                    self._rows[row_idx][col_idx] = val

    def append_rows(self, rows: List[List[Any]]) -> None:
        self.append_calls.append([list(r) for r in rows])
        for r in rows:
            row = list(r)
            while len(row) < max(len(self.headers), 22):
                row.append("")
            self._rows.append(row)

    def rows(self) -> List[List[Any]]:
        return [list(r) for r in self._rows]


def _letter_index(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n
