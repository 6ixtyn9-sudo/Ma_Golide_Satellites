#!/usr/bin/env python3
"""Write a MAGOLIDE_SCORE research section into a lane's dated pick and ticket cards.

Why a script and not the operator
---------------------------------
The operator's cards are built from forecasts, and these four lanes have none:
Greece, WCBA and WNBL never generate a forecast, and NKL's stops at
``price_not_named_bookmaker_quote``. Rather than change how a forecast is made --
a decision that owns calibration, tickets and the bank -- this writes a section
*beside* those cards and says, on every line, that it is research.

The section is fenced by two markers and replaced in place on every run, so a
lane can be re-scored without the card growing a second copy. Creating a card that
did not exist is allowed: "this lane has no picks" is exactly the state this is
meant to change, and an empty section is still an honest zero.

What it costs
-------------
One fixtures page (to find a match page), one match page (to mint the token), and
one fragment per sub-table. Every request goes through
``betexplorer_standings.paced``, whose floor is the host's measured ten seconds.

Nothing here writes a stake, a receipt, a bank row or a probability.
"""

from __future__ import annotations

import argparse
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Dict, Iterable, List, Mapping, Optional, Sequence, Tuple

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from forecasting import research_slate  # noqa: E402
from forecasting.research_slate import render_research_section, research_rows  # noqa: E402
from ingestion.identity import parse_utc  # noqa: E402
from ingestion.league_registry import league_spec, parse_league  # noqa: E402
from ingestion.models import BookSnapshot, Game, GameStatus, StandingsRow  # noqa: E402
from ingestion.providers import betexplorer, betexplorer_standings  # noqa: E402

#: Fences. Chosen so no other reader in the project mistakes them for a section it
#: owns: ``picks_*.txt`` is scanned for the compact marker and
#: ``auto_tickets_*.txt`` for ``PENDING`` and two audit headers, and none of those
#: strings appears here.
MARKER_START = "=== MAGOLIDE_SCORE RESEARCH-ONLY — AGGREGATE/UNCALIBRATED ==="
MARKER_END = "=== END MAGOLIDE_SCORE RESEARCH-ONLY ==="

CARD_KINDS = {"picks": "picks_{date}.txt", "tickets": "auto_tickets_{date}.txt"}


def upsert_section(text: str, section: Sequence[str]) -> str:
    """Replace this script's fenced block, or append it once.

    A card is rewritten by the publisher whenever it runs, so the block is not
    durable state: it is a view. Replacing rather than appending is what keeps a
    lane scored five times in a day from printing the slate five times.
    """

    block = [MARKER_START, *section, MARKER_END]
    lines = text.splitlines()
    if MARKER_START in lines and MARKER_END in lines:
        start = lines.index(MARKER_START)
        end = lines.index(MARKER_END, start)
        return "\n".join([*lines[:start], *block, *lines[end + 1 :]]) + "\n"
    body = text.rstrip("\n")
    joined = "\n".join([body, "", *block]) if body else "\n".join(block)
    return joined + "\n"


def card_path(localdata: Path, state_key: str, kind: str, target_date: str) -> Path:
    template = CARD_KINDS[kind]
    return localdata / state_key / template.format(date=target_date)


def write_section(
    localdata: Path,
    *,
    state_key: str,
    target_date: str,
    section: Sequence[str],
    kinds: Iterable[str] = ("picks", "tickets"),
) -> Dict[str, str]:
    """Upsert the block into each card, reporting what was written."""

    written: Dict[str, str] = {}
    for kind in kinds:
        path = card_path(localdata, state_key, kind, target_date)
        path.parent.mkdir(parents=True, exist_ok=True)
        existing = path.read_text(encoding="utf-8") if path.is_file() else ""
        path.write_text(upsert_section(existing, section), encoding="utf-8")
        written[kind] = str(path)
    return written


def slate_games(
    games: Sequence[Game], target_date: str, *, tz_offset_hours: int = 0
) -> List[Game]:
    """Fixtures scheduled on ``target_date`` in the lane's wall clock.

    The cards are dated in the lane's local day while ``scheduled_utc`` is UTC, so
    a naive UTC comparison would put a late tip on the wrong card. The offset is a
    parameter rather than a table: the caller knows the lane, and inventing a
    default timezone here would be a guess printed as a fact.
    """

    out: List[Game] = []
    for game in games:
        if game.status is not GameStatus.SCHEDULED:
            continue
        local = parse_utc(game.scheduled_utc).timestamp() + tz_offset_hours * 3600
        if datetime.fromtimestamp(local, tz=timezone.utc).strftime("%Y-%m-%d") == target_date:
            out.append(game)
    return out


def fetch_published(
    league: str,
    read_text,
    *,
    match_url: Optional[str] = None,
    delay: Optional[float] = None,
) -> Tuple[List[StandingsRow], Dict[str, object]]:
    """Published standings for a lane, finding a match page when not given one.

    The fixtures page is read only to obtain a ``(slug, match_id)`` pair: the
    standings token exists on match pages and nowhere else, so a lane whose
    fixtures say nothing is reported rather than guessed at.
    """

    league_key = betexplorer._require_league(league)
    if match_url is None:
        fixture_html = read_text(betexplorer.FIXTURES_URLS[league_key])
        links = betexplorer.match_links(fixture_html)
        if not links:
            raise betexplorer_standings.IdentityError(
                f"{league_key}: fixtures page carried no match links, so no match "
                "page exists to mint a standings token"
            )
        slug, match_id = links[0]
        match_url = betexplorer_standings.match_url_for(league_key, slug, match_id)

    result = betexplorer_standings.fetch_standings(
        league_key, read_text, match_url=match_url, delay=delay
    )
    report = dict(result.report)
    report["match_url"] = match_url
    return result.rows, report


def build_section(
    league: str,
    games: Sequence[Game],
    published: Sequence[StandingsRow],
    snapshots: Mapping[str, Sequence[BookSnapshot]],
    *,
    target_date: str,
) -> List[str]:
    """Rows and slips for one lane's slate, or an explicit empty block."""

    ordered: List[BookSnapshot] = [
        snapshot for game in games for snapshot in snapshots.get(game.game_id, ())
    ]
    rows = research_rows(league, list(games), list(published), snapshots=ordered)
    section = render_research_section(rows, date_label=target_date, league=league)
    if section:
        return section
    return [
        "",
        f"{research_slate.RESEARCH_LABEL} — {league} {target_date} — 0 fixture(s)",
        "- No scheduled fixture for this lane on this date, so there is nothing to score.",
        "- This is an honest zero, not a failed read: the table was read and no game was listed.",
    ]


def _read_text_factory(delay: float):
    """A paced transport for live runs.

    The pacing lives here, once, and ``fetch_published`` is then called with
    ``delay=0.0``: the inner ``PacedReader`` it builds would otherwise sleep a
    second time for the same request.
    """

    from ingestion.http import HttpClient

    client = HttpClient()

    def read(url: str) -> str:
        return client.get_text(url)

    return betexplorer_standings.paced(read, delay=delay)


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--league", required=True)
    parser.add_argument("--date", required=True, help="YYYY-MM-DD, the card's date")
    parser.add_argument("--localdata", default="localdata")
    parser.add_argument("--warehouse", default=None)
    parser.add_argument(
        "--delay",
        type=float,
        default=betexplorer_standings.CRAWL_DELAY_SECONDS,
        help="seconds between requests (the host's measured floor is 10)",
    )
    args = parser.parse_args(argv)

    league = parse_league(args.league)
    spec = league_spec(league)
    localdata = Path(args.localdata)

    from ingestion.duckdb_store import DuckDBStore

    store = DuckDBStore(args.warehouse or "localdata/warehouse.duckdb")
    try:
        games = slate_games(store.list_games(league=league), args.date)
        snapshots = {game.game_id: store.get_book_snapshots(game.game_id) for game in games}
        read = _read_text_factory(args.delay)
        published, report = fetch_published(league.value, read, delay=0.0)
    finally:
        store.close()

    section = build_section(
        league.value, games, published, snapshots, target_date=args.date
    )
    written = write_section(
        localdata, state_key=spec.state_key, target_date=args.date, section=section
    )
    print(
        f"{league.value} {args.date}: {len(games)} fixture(s), "
        f"{len(published)} published row(s), requests={report.get('requests')}, "
        f"subs={report.get('subs_read')}, mismatch={report.get('split_mismatch')}"
    )
    for kind, path in written.items():
        print(f"  {kind}: {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
