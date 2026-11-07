"""Derivations: H2H, recent form, and standings layouts.

All derivations are generated from OFFICIAL canonical results (Games with
ActualResult), never from legacy raw text blocks (master prompt section 6).

Tier-2 clean schema (observed in the reference workbook, binding):

    Date, Home, Away,
    Q1H, Q1A, Q2H, Q2A, Q3H, Q3A, Q4H, Q4A,
    OTH, OTA, Winner, TotalH, TotalA, FT, SumMismatch

- ``CleanH2H_i``: previous meetings between the two upcoming teams.
- ``CleanRecentHome_i``: recent games of the upcoming HOME team, including
  its home AND away games (the suffix names the upcoming side, not a venue
  filter).
- Newest-first ordering; no duplicates; the upcoming game itself is
  excluded; no future games.
- ``Winner`` is H/A relative to the ROW's home/away orientation.
- ``FT`` is the row's ``home-away`` string.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List, Optional, Sequence, Tuple

from .identity import normalize_team_name
from .models import Game, GameStatus, PeriodKind, StandingsRow, round_half_up

TIER2_HEADERS: List[str] = [
    "Date", "Home", "Away",
    "Q1H", "Q1A", "Q2H", "Q2A", "Q3H", "Q3A", "Q4H", "Q4A",
    "OTH", "OTA", "Winner", "TotalH", "TotalA", "FT", "SumMismatch",
]

DEFAULT_H2H_DEPTH = 10
DEFAULT_RECENT_DEPTH = 15


def _t2_row(game: Game) -> Optional[List]:
    """One canonical result -> one Tier-2 row (row orientation preserved)."""
    actual = game.actual
    if actual is None or actual.period_kind is not PeriodKind.QUARTERS:
        return None
    if actual.home_final is None or actual.away_final is None:
        return None

    def _int_or_blank(v):
        return round_half_up(v, 0) if v is not None else ""

    winner = "H" if actual.home_final > actual.away_final else (
        "A" if actual.away_final > actual.home_final else ""
    )
    sum_home = sum(v for v in actual.home_quarters if v is not None) + (
        actual.home_overtime or 0
    )
    sum_away = sum(v for v in actual.away_quarters if v is not None) + (
        actual.away_overtime or 0
    )
    mismatch = ""
    if abs(sum_home - actual.home_final) > 0 or abs(sum_away - actual.away_final) > 0:
        mismatch = f"{sum_home - actual.home_final:g}/{sum_away - actual.away_final:g}"

    local = game.utc_datetime.date().isoformat()
    row = [local, game.home_name, game.away_name]
    for i in range(4):
        h = actual.home_quarters[i] if i < len(actual.home_quarters) else None
        a = actual.away_quarters[i] if i < len(actual.away_quarters) else None
        row.extend([_int_or_blank(h), _int_or_blank(a)])
    row.extend([
        _int_or_blank(actual.home_overtime),
        _int_or_blank(actual.away_overtime),
        winner,
        round_half_up(actual.home_final, 0),
        round_half_up(actual.away_final, 0),
        f"{round_half_up(actual.home_final, 0):.0f}-{round_half_up(actual.away_final, 0):.0f}",
        mismatch,
    ])
    return row


def _eligible_history(
    games: Sequence[Game], before_utc: str
) -> List[Game]:
    """Official FINAL games that finished before ``before_utc``, newest first.

    Quarantined (unreconciled/live/abandoned) games are excluded — they must
    never enter H2H/recent history.
    """
    eligible: List[Game] = []
    for g in games:
        if g.status is not GameStatus.FINAL or g.actual is None:
            continue
        if g.scheduled_utc >= before_utc:
            continue  # includes the upcoming game itself and anything after
        if g.actual.home_final is None or g.actual.away_final is None:
            continue
        ok, _why = g.actual.reconciles()
        if not ok:
            continue  # periods contradict the final: quarantine, never guess
        eligible.append(g)
    return sorted(eligible, key=lambda g: g.scheduled_utc, reverse=True)


def _dedupe(games: Sequence[Game]) -> List[Game]:
    seen: set = set()
    out: List[Game] = []
    for g in games:
        key = (g.scheduled_utc, normalize_team_name(g.home_name),
               normalize_team_name(g.away_name))
        if key in seen:
            continue
        seen.add(key)
        out.append(g)
    return out


def derive_h2h(
    history: Sequence[Game], home: str, away: str, before_utc: str,
    depth: int = DEFAULT_H2H_DEPTH,
) -> List[List]:
    """Previous meetings between the two teams, newest first."""
    nh, na = normalize_team_name(home), normalize_team_name(away)
    rows: List[List] = []
    for g in _dedupe(_eligible_history(history, before_utc)):
        gh, ga = normalize_team_name(g.home_name), normalize_team_name(g.away_name)
        if {gh, ga} == {nh, na} and gh != ga:
            row = _t2_row(g)
            if row is not None:
                rows.append(row)
        if len(rows) >= depth:
            break
    return rows


def derive_recent(
    history: Sequence[Game], team: str, before_utc: str,
    depth: int = DEFAULT_RECENT_DEPTH,
) -> List[List]:
    """Recent games for one upcoming side, at BOTH venues, newest first."""
    nt = normalize_team_name(team)
    rows: List[List] = []
    for g in _dedupe(_eligible_history(history, before_utc)):
        gh, ga = normalize_team_name(g.home_name), normalize_team_name(g.away_name)
        if nt in (gh, ga) and gh != ga:
            row = _t2_row(g)
            if row is not None:
                rows.append(row)
        if len(rows) >= depth:
            break
    return rows


def bundle_rows(
    history: Sequence[Game], home: str, away: str, before_utc: str,
    h2h_depth: int = DEFAULT_H2H_DEPTH, recent_depth: int = DEFAULT_RECENT_DEPTH,
) -> Dict[str, List[List]]:
    """All three Tier-2 bundles for upcoming game index ``i``."""
    return {
        "H2H": derive_h2h(history, home, away, before_utc, h2h_depth),
        "RecentHome": derive_recent(history, home, before_utc, recent_depth),
        "RecentAway": derive_recent(history, away, before_utc, recent_depth),
    }


# ---------------------------------------------------------------------------
# Standings layout (legacy sections + repeated headers)
# ---------------------------------------------------------------------------

STANDINGS_HEADERS: List[str] = ["Position", "Team name", "GP", "W", "L", "PF", "PA", "PCT"]


def standings_layout(
    rows: Sequence[StandingsRow], sections: Optional[Sequence[str]] = None
) -> List[List]:
    """Render the legacy Standings layout: section label rows + headers.

    ``loadStandings`` scans for section labels (East/West/...), repeated
    header rows, then team rows.  PF/PA are written CUMULATIVE (the legacy
    loader computes ``(PF-PA)/GP`` itself); PCT is the 0..1 fraction.
    """
    out: List[List] = []
    by_section: Dict[str, List[StandingsRow]] = {}
    for r in rows:
        by_section.setdefault(r.section or "League", []).append(r)
    wanted = list(sections) if sections else list(by_section)
    for section in wanted:
        out.append([section])
        out.append(list(STANDINGS_HEADERS))
        for r in sorted(by_section.get(section, []), key=lambda x: (x.position is None, x.position or 0)):
            out.append([
                r.position if r.position is not None else "",
                r.team_name,
                r.games_played if r.games_played is not None else "",
                r.wins if r.wins is not None else "",
                r.losses if r.losses is not None else "",
                round_half_up(r.points_for, 0) if r.points_for is not None else "",
                round_half_up(r.points_against, 0) if r.points_against is not None else "",
                round_half_up(r.win_pct, 3) if r.win_pct is not None else "",
            ])
    return out
