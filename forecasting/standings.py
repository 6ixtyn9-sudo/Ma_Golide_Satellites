"""Standings derived from our own reconciled finals, not from a table we scrape.

Why this exists
---------------
``forecasting/magolide_score.py`` is the only surface in this repository that can
rate a fixture with zero results history: a standings-difference scorer ported
from the retired spreadsheet. It needs ``rank``, ``pct``, ``net_rtg``,
``home_pct``, ``away_pct`` and ``streak`` per club, and nothing supplies them for
the fixtures-only lanes. ESPN and Euroleague publish standings for their own
competitions; the shadow card takes a JSON that no process generates. So the
scorer has been runnable in principle and unrun in practice.

For a lane that settles its own games there is no need to ask anybody: a win is a
win and a point is a point. This module turns the finals already in the store into
the rows the scorer consumes, which keeps the numbers traceable to games we
decided ourselves rather than to a table whose update lag nobody measures.

What it derives, and what it deliberately does not
--------------------------------------------------
Derived, all arithmetic over reconciled finals:

* games, wins, losses, win percentage;
* points for/against and their per-game difference;
* home and away records, win percentage in each;
* current streak (signed: ``+3`` is three wins, ``-2`` two losses) and the last
  ten games' win percentage;
* a rank, ordered by win percentage, then by points difference.

``net_rtg`` here is a **points-per-game difference, not a possession-adjusted
rating**. There is no possession data in this corpus, so calling it a rating
would be a claim the inputs cannot support; the name is kept only because it is
the field the scorer reads.

Refused rather than guessed: a final whose periods do not reconcile with its own
total is skipped (it cannot be graded, so it cannot count), and a game whose two
sides carry the same club name is skipped rather than counted twice. A club with
no finals has no row -- it is absent from the result, which is honest, rather than
present at 0-0 as though it had played.

Two routes, one scorer surface
-----------------------------
:func:`standings_from_finals` derives a table from finals this project graded. It
makes no request, writes nothing, and knows nothing about rankings the competition
itself publishes; a league tiebreak rule that differs from "win percentage, then
points difference" will produce a different order, and ``SOURCE`` names it.

:func:`from_published` quotes a competition's own table instead, once
``ingestion/providers/betexplorer_standings.py`` has read it. It keeps the
published rank -- the NKL table captured on 2026-11-25 is *not* ordered by win
percentage -- and it takes only those fields a table can state: ``pct``,
``net_rtg``, and the home/away splits. A streak cannot be quoted from a table
whose form glyphs are undecoded, so it arrives only from ``derived`` rows for the
same clubs, which is why the two routes take a ``Mapping`` and not a flag. Rows
from this route are labelled ``PUBLISHED_SOURCE`` by whoever stores them; the
scorer row itself carries no source field, so the label must be carried by the
caller rather than inferred later.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Dict, Iterable, List, Mapping, Optional, Sequence

from ingestion.identity import canonical_team_name, parse_utc
from ingestion.models import Game, GameStatus
from ingestion.models import StandingsRow as PublishedStandingsRow

from .magolide_score import StandingsRow

SOURCE = "derived_from_reconciled_finals"

#: The label for rows that came from a competition's own published table rather
#: than from finals we graded. Kept separate so a reader can tell arithmetic we
#: did from arithmetic we quoted.
PUBLISHED_SOURCE = "betexplorer_published_table"

#: How many recent games the ``l10`` fields describe.
LAST_N = 10


@dataclass
class _Tally:
    team: str
    games: int = 0
    wins: int = 0
    losses: int = 0
    points_for: float = 0.0
    points_against: float = 0.0
    home_games: int = 0
    home_wins: int = 0
    away_games: int = 0
    away_wins: int = 0
    #: Results oldest first, as ``True`` for a win: the scorer needs the current
    #: streak and the last ten, and both read from the end.
    recent: List[bool] = field(default_factory=list)


def _eligible(game: Game) -> bool:
    """A final we can count: graded, past, and not a team playing itself."""

    if game.status is not GameStatus.FINAL or game.actual is None:
        return False
    if game.home_name.strip() == game.away_name.strip():
        return False
    ok, _ = game.actual.reconciles()
    if not ok:
        return False
    return game.actual.home_final is not None and game.actual.away_final is not None


def _record(tally: _Tally, *, won: bool, points_for: float, points_against: float,
            at_home: bool) -> None:
    tally.games += 1
    tally.wins += int(won)
    tally.losses += int(not won)
    tally.points_for += points_for
    tally.points_against += points_against
    if at_home:
        tally.home_games += 1
        tally.home_wins += int(won)
    else:
        tally.away_games += 1
        tally.away_wins += int(won)
    tally.recent.append(won)


def _row(tally: _Tally, rank: int) -> StandingsRow:
    played = tally.games
    pct = tally.wins / played if played else None
    diff = (tally.points_for - tally.points_against) / played if played else None
    recent = tally.recent[-LAST_N:] if tally.recent else []
    return StandingsRow(
        team=tally.team,
        rank=float(rank),
        pct=None if pct is None else round(pct, 4),
        net_rtg=None if diff is None else round(diff, 2),
        home_pct=(
            round(tally.home_wins / tally.home_games, 4) if tally.home_games else None
        ),
        away_pct=(
            round(tally.away_wins / tally.away_games, 4) if tally.away_games else None
        ),
        streak=_streak(tally.recent),
        l10_wins=float(sum(recent)),
        l10_losses=float(len(recent) - sum(recent)),
    )


def _streak(recent: Sequence[bool]) -> Optional[int]:
    """Signed current streak, read backwards from the most recent game.

    ``recent`` is oldest-first, so the current run starts at the *last* element.
    Reading from the front would report the streak the season opened with, which
    is the one number here a reader would never double-check.
    """

    if not recent:
        return None
    current = recent[-1]
    length = 0
    for won in reversed(recent):
        if won != current:
            break
        length += 1
    return length if current else -length


def standings_from_finals(
    games: Iterable[Game], *, as_of_utc: Optional[str] = None
) -> List[StandingsRow]:
    """One row per club that has played an eligible final, ranked.

    ``as_of_utc`` excludes games that had not tipped or had not finished by that
    instant, so a caller can rebuild the table as it stood before a given game --
    which is what keeps a backtest honest. Without it every stored final counts.
    """

    cutoff = parse_utc(as_of_utc) if as_of_utc else None
    ordered = sorted(games, key=lambda game: (parse_utc(game.scheduled_utc), game.game_id))
    tallies: Dict[str, _Tally] = {}

    for game in ordered:
        if not _eligible(game):
            continue
        if cutoff is not None and parse_utc(game.scheduled_utc) >= cutoff:
            continue
        home = tallies.setdefault(game.home_name.strip(), _Tally(game.home_name.strip()))
        away = tallies.setdefault(game.away_name.strip(), _Tally(game.away_name.strip()))
        home_points = float(game.actual.home_final)
        away_points = float(game.actual.away_final)
        home_won = home_points > away_points
        _record(
            home, won=home_won, points_for=home_points, points_against=away_points,
            at_home=True,
        )
        _record(
            away, won=not home_won, points_for=away_points, points_against=home_points,
            at_home=False,
        )

    def order(tally: _Tally) -> tuple:
        played = tally.games
        pct = tally.wins / played if played else 0.0
        diff = (tally.points_for - tally.points_against) / played if played else 0.0
        return (-pct, -diff, tally.team)

    return [_row(tally, index + 1) for index, tally in enumerate(sorted(tallies.values(), key=order))]


#: A published split record, exactly as a table writes it: ``8-1``.
_RECORD_RE = re.compile(r"^(?P<wins>\d{1,3})-(?P<losses>\d{1,3})$")


def _split_pct(record: Optional[str]) -> Optional[float]:
    """Win percentage from a published ``W-L`` split, or ``None``.

    ``0-0`` is a club that has not played in that role yet, so it yields ``None``
    rather than ``0.0``: an empty record must not read as a losing one.
    """

    if not record:
        return None
    match = _RECORD_RE.match(str(record).strip())
    if match is None:
        return None
    wins, losses = int(match.group("wins")), int(match.group("losses"))
    played = wins + losses
    if played <= 0:
        return None
    return round(wins / played, 4)


def from_published(
    rows: Iterable[PublishedStandingsRow],
    *,
    derived: Optional[Mapping[str, StandingsRow]] = None,
) -> List[StandingsRow]:
    """Scorer rows from a competition's own published table.

    ``ingestion.providers.betexplorer_standings`` reads the table; this turns it
    into the seven numbers ``forecasting/magolide_score.py`` consumes. Every value
    is arithmetic on what the table states -- ``pct = W / MP`` and
    ``net_rtg = (PF - PA) / MP`` -- and nothing is inferred.

    **Rank is quoted, never recomputed.** The published table is ordered by the
    competition's own tiebreak rules, which are not win percentage: the NKL table
    captured on 2026-11-25 lists a 0.7647 club above a 0.8125 one and a 0.8571
    club fifth. Recomputing the order from ``pct`` would silently disagree with
    the table this row cites, so ``position`` passes through untouched.

    ``streak`` and the last-ten fields cannot come from a published table at all:
    the ``Form`` glyphs are undecoded (see the provider docstring). When a lane
    also settles its own games, pass ``derived`` -- the output of
    :func:`standings_from_finals` -- and those two fields are filled from finals
    we graded ourselves while everything else stays quoted. A club with no derived
    history keeps ``None`` rather than a zero.

    A row without a name, or without games played, is skipped: a club that has not
    played has no rate, and a 0.0 would read as a losing season.
    """

    by_derivation = dict(derived or {})
    seen: set[str] = set()
    out: List[StandingsRow] = []
    for row in rows:
        name = canonical_team_name(str(row.team_name or "").strip())
        if not name or name in seen:
            continue
        played = row.games_played
        if played is None or played <= 0:
            continue
        seen.add(name)

        pct = row.win_pct
        if pct is None and row.wins is not None:
            pct = round(row.wins / played, 4)
        net_rtg: Optional[float] = None
        if row.points_for is not None and row.points_against is not None:
            net_rtg = round((float(row.points_for) - float(row.points_against)) / played, 2)

        graded = by_derivation.get(name) or by_derivation.get(str(row.team_name).strip())
        out.append(
            StandingsRow(
                team=name,
                rank=float(row.position) if row.position is not None else None,
                pct=pct,
                net_rtg=net_rtg,
                home_pct=_split_pct(row.home_record),
                away_pct=_split_pct(row.away_record),
                streak=graded.streak if graded is not None else None,
                l10_wins=graded.l10_wins if graded is not None else None,
                l10_losses=graded.l10_losses if graded is not None else None,
            )
        )
    return out
