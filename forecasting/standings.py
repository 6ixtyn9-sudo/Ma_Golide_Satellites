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

This module makes no request, writes nothing, and knows nothing about rankings the
competition itself publishes. A league tiebreak rule that differs from "win
percentage, then points difference" will produce a different order than the
official table, and the ``source`` field says which one this is.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, Iterable, List, Optional, Sequence

from ingestion.identity import parse_utc
from ingestion.models import Game, GameStatus

from .magolide_score import StandingsRow

SOURCE = "derived_from_reconciled_finals"

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
