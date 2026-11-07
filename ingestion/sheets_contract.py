"""Binding Google Sheets cell contract for the base 22-column prefix.

The 22 base columns are positionally and semantically binding (master
prompt section 6).  This module:

- validates a sheet's header row against the contract and FAILS SAFELY on
  duplicates, shifts or missing base headers;
- exposes 1-based column indexes for every base field;
- renders canonical records into exact cell values for ``UpcomingClean``
  and ``ResultsClean`` under their different field semantics.

Acquisition owns ONLY columns A..R (1..18).  Columns S..V (``t2-q1``..
``t2-q4``) and every column after V are existing model/output territory:
they are never rendered, cleared or rewritten here.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date as date_cls
from datetime import time as time_cls
from typing import Dict, List, Optional, Sequence, Tuple

from .identity import render_local
from .models import (
    BookSnapshot,
    FinalStatus,
    Game,
    League,
    Prediction,
    fmt_prob,
    fmt_score_pair,
    fmt_total,
    pred_from_probabilities,
    round_half_up,
)

BASE_HEADERS: List[str] = [
    "League",
    "Game Type",
    "Home",
    "Away",
    "Date",
    "Time",
    "Prob %",
    "Pred",
    "Pred Score",
    "Avg",
    "Odds",
    "Q1",
    "Q2",
    "Q3",
    "Q4",
    "OT",
    "Status",
    "FT Score",
    "t2-q1",
    "t2-q2",
    "t2-q3",
    "t2-q4",
]

# 1-based column positions (A=1 .. V=22).
COL: Dict[str, int] = {name: i + 1 for i, name in enumerate(BASE_HEADERS)}

# Columns acquisition OWNS on both clean tabs (A..R = 1..18).
OWNED_COLUMNS: Tuple[int, ...] = tuple(range(1, 19))
# Columns that must never be touched by acquisition (S..V = 19..22 plus all
# columns after V — the existing model/output territory).
PROTECTED_PREFIX_COLUMNS: Tuple[int, ...] = tuple(range(19, 23))

UPCOMING_TAB = "UpcomingClean"
RESULTS_TAB = "ResultsClean"


class ContractError(RuntimeError):
    """Sheet header/layout drift; the writer must refuse to run."""


@dataclass(frozen=True)
class SheetContract:
    """Validated header row for one clean tab."""

    headers: Tuple[str, ...]
    tab_kind: str  # "upcoming" | "results"

    @property
    def base_ok(self) -> bool:
        return tuple(self.headers[:22]) == tuple(BASE_HEADERS)

    def col(self, name: str) -> int:
        try:
            return COL[name]
        except KeyError:
            raise ContractError(f"unknown base column: {name!r}")

    def extra_headers(self) -> List[str]:
        return list(self.headers[22:])


def detect_contract(headers: Sequence[str], tab_kind: str) -> SheetContract:
    """Validate a header row; raise ContractError on any drift.

    Rules:
    - The first 22 headers must equal BASE_HEADERS positionally (trimmed;
      compared case-insensitively because gspread round-trips are stable but
      hand edits happen).
    - An unsuffixed base header appearing AGAIN anywhere beyond its expected
      position is ambiguous drift and fails (e.g. a second 'Q1').
    - Columns after V are allowed and preserved, whatever they contain.
    """
    normalized = [str(h or "").strip() for h in headers]
    base = [h.strip().lower() for h in BASE_HEADERS]
    first22 = [h.lower() for h in normalized[:22]]

    if len(normalized) < 22:
        raise ContractError(
            f"{tab_kind}: only {len(normalized)} headers found; the 22-column "
            "base prefix is required"
        )
    mismatches = [
        (i + 1, BASE_HEADERS[i], normalized[i])
        for i in range(22)
        if first22[i] != base[i]
    ]
    if mismatches:
        raise ContractError(
            f"{tab_kind}: base header drift at column(s) "
            + ", ".join(f"{pos}={found!r} (expected {want!r})" for pos, want, found in mismatches)
        )

    # Ambiguity guard: a base header name repeated after position 22.
    lower_set = set(base)
    for idx, h in enumerate(normalized[22:], start=23):
        if h.lower() in lower_set:
            raise ContractError(
                f"{tab_kind}: duplicate unsuffixed base header {h!r} at column "
                f"{idx}; ambiguous contract, refusing to write"
            )
    return SheetContract(headers=tuple(normalized), tab_kind=tab_kind)


# ---------------------------------------------------------------------------
# Row rendering
# ---------------------------------------------------------------------------


def _identity_cells(game: Game, tz_name: str) -> Dict[int, object]:
    local_date, local_time = render_local(game.scheduled_utc, tz_name)
    return {
        COL["League"]: League.display_value(game.league),
        COL["Game Type"]: game.phase or "League",
        COL["Home"]: game.home_name,
        COL["Away"]: game.away_name,
        COL["Date"]: local_date,
        COL["Time"]: local_time,
    }


def _prediction_cells(pred: Optional[Prediction]) -> Dict[int, object]:
    """Render the coherent G:J bundle; blanks when prediction missing."""
    if pred is None:
        return {
            COL["Prob %"]: "",
            COL["Pred"]: "",
            COL["Pred Score"]: "",
            COL["Avg"]: "",
        }
    display = pred.display()
    avg_raw = display["Avg"]
    return {
        COL["Prob %"]: display["Prob %"] or "",
        COL["Pred"]: display["Pred"] if display["Pred"] is not None else "",
        COL["Pred Score"]: display["Pred Score"] or "",
        # Numeric at one decimal (workbook shows 232.8-style values); raw
        # precision lives in canonical storage.
        COL["Avg"]: round_half_up(float(avg_raw), 1) if avg_raw else "",
    }


def render_upcoming_row(
    game: Game,
    prediction: Optional[Prediction],
    odds_decimal: Optional[float],
    quarter_lines: Dict[str, Optional[float]],
    full_game_line: Optional[float],
    tz_name: str,
) -> Dict[int, object]:
    """Cells for one ``UpcomingClean`` row (columns A..R only).

    Semantics (section 6): Q1..Q4 and FT Score carry **bookmaker total
    lines**; Avg carries the model expected total; OT/Status stay blank for
    upcoming games.
    """
    cells = _identity_cells(game, tz_name)
    cells.update(_prediction_cells(prediction))
    cells[COL["Odds"]] = odds_decimal if odds_decimal is not None else ""
    for q in ("Q1", "Q2", "Q3", "Q4"):
        value = quarter_lines.get(q)
        cells[COL[q]] = float(value) if value is not None else ""
    cells[COL["OT"]] = ""
    cells[COL["Status"]] = ""
    cells[COL["FT Score"]] = (
        float(full_game_line) if full_game_line is not None else ""
    )
    return cells


def render_result_row(
    game: Game,
    frozen_prediction: Optional[Prediction],
    odds_decimal: Optional[float],
    tz_name: str,
) -> Dict[int, object]:
    """Cells for one ``ResultsClean`` row (columns A..R only).

    Semantics (section 6): Q1..Q4/OT carry **actual team-pair scores**
    ``home - away``; Status is FT/AOT; FT Score is the actual final;
    Prob %/Pred/Pred Score/Avg/Odds are the FROZEN pre-tip values.
    """
    cells = _identity_cells(game, tz_name)
    cells.update(_prediction_cells(frozen_prediction))
    cells[COL["Odds"]] = odds_decimal if odds_decimal is not None else ""

    actual = game.actual
    if actual is None or actual.period_kind.value != "quarters":
        for q in ("Q1", "Q2", "Q3", "Q4"):
            cells[COL[q]] = ""
        cells[COL["OT"]] = ""
        cells[COL["Status"]] = ""
        cells[COL["FT Score"]] = ""
        return cells

    for i, q in enumerate(("Q1", "Q2", "Q3", "Q4")):
        h = actual.home_quarters[i] if i < len(actual.home_quarters) else None
        a = actual.away_quarters[i] if i < len(actual.away_quarters) else None
        cells[COL[q]] = _pair(h, a)
    cells[COL["OT"]] = _pair(actual.home_overtime, actual.away_overtime) or ""
    status = game.final_status()
    cells[COL["Status"]] = status.value if status is not None else ""
    cells[COL["FT Score"]] = _pair(actual.home_final, actual.away_final) or ""
    return cells


def _pair(home: Optional[float], away: Optional[float]) -> Optional[str]:
    """``home - away`` team-pair string with integer display (e.g. ``25 - 24``)."""
    if home is None or away is None:
        return None
    return f"{round_half_up(home, 0):.0f} - {round_half_up(away, 0):.0f}"


def blank_pair_ok(game: Game) -> bool:
    """A result row without official periods stays quarantined (blank)."""
    return game.actual is None or game.status is not FinalStatus.FT
