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
import time
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Dict, Iterable, List, Mapping, Optional, Sequence, Tuple

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from forecasting import research_slate  # noqa: E402
from forecasting.config_tuner import TUNED_PATH, load_tuned  # noqa: E402
from forecasting.research_slate import render_research_section, research_rows  # noqa: E402
from ingestion.http import HttpError  # noqa: E402
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


@dataclass(frozen=True)
class Fixture:
    """The four fields a research row reads, without inventing a canonical Game.

    ``forecasting.research_slate.research_rows`` touches only ``game_id``,
    ``home_name``, ``away_name`` and ``scheduled_utc``, so a fixture taken straight
    from a published page needs no store, no season, and no status.
    """

    game_id: str
    home_name: str
    away_name: str
    scheduled_utc: str


def slate_from_page(
    league: str, target_date: str, read_text, *, now_utc: str
) -> Tuple[List[Fixture], Dict[str, List[BookSnapshot]], str, int]:
    """The day's slate and its aggregate prices, straight from the fixtures page.

    One read serves two purposes: the page lists the fixtures *and* the match pages
    whose token mints the standings table. ``betexplorer.fetch_odds`` then parses
    the same bytes for prices, so a lane with no warehouse is still scoreable --
    which is the state these four lanes are actually in.

    Returns ``(fixtures, snapshots_by_game, fixtures_html, reads)``.
    """

    league_key = betexplorer._require_league(league)
    fixtures_url = betexplorer.FIXTURES_URLS[league_key]
    fixtures_html = read_text(fixtures_url)
    reads = 1

    def cached(url: str) -> str:
        nonlocal reads
        if url == fixtures_url:
            return fixtures_html
        reads += 1
        return read_text(url)

    snapshots = betexplorer.fetch_odds(
        league_key, target_date, text_fetch=cached, now_utc=now_utc
    )
    by_game: Dict[str, List[BookSnapshot]] = {}
    for snapshot in snapshots:
        by_game.setdefault(snapshot.game_id, []).append(snapshot)

    fixtures = [
        Fixture(
            game_id=game_id,
            home_name=str(rows[0].home_name or ""),
            away_name=str(rows[0].away_name or ""),
            scheduled_utc=str(rows[0].scheduled_utc or ""),
        )
        for game_id, rows in by_game.items()
    ]
    return fixtures, by_game, fixtures_html, reads


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
    fixtures_html: Optional[str] = None,
    delay: Optional[float] = None,
) -> Tuple[List[StandingsRow], Dict[str, object]]:
    """Published standings for a lane, finding a match page when not given one.

    The fixtures page is read only to obtain a ``(slug, match_id)`` pair: the
    standings token exists on match pages and nowhere else, so a lane whose
    fixtures say nothing is reported rather than guessed at.
    """

    league_key = betexplorer._require_league(league)
    if match_url is None:
        page = fixtures_html
        if page is None:
            page = read_text(betexplorer.FIXTURES_URLS[league_key])
        links = betexplorer.match_links(page)
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


def failure_section(league: str, target_date: str, reason: str) -> List[str]:
    """The block a failed run leaves behind, replacing whatever was there.

    A read that never happened must not leave an earlier "no fixtures" line in
    place: that sentence claims the table was read and the slate was empty, and on
    this run neither was true. The reason is carried verbatim, because the useful
    part of a 429 is that it was a 429.
    """

    return [
        "",
        f"{research_slate.RESEARCH_LABEL} — {league} {target_date} — READ FAILED",
        "- No read succeeded, so this lane has no research rows for this date.",
        f"- Reason: {reason}",
        "- The earlier block for this date was replaced, not kept: a stale "
        "\"no fixtures\" line would claim a read that did not happen.",
    ]


def build_section(
    league: str,
    games: Sequence[Game],
    published: Sequence[StandingsRow],
    snapshots: Mapping[str, Sequence[BookSnapshot]],
    *,
    target_date: str,
    bet_config=None,
) -> List[str]:
    """Rows and slips for one lane's slate, or an explicit empty block."""

    ordered: List[BookSnapshot] = [
        snapshot for game in games for snapshot in snapshots.get(game.game_id, ())
    ]
    rows = research_rows(
        league, list(games), list(published), snapshots=ordered, bet_config=bet_config
    )
    section = render_research_section(rows, date_label=target_date, league=league)
    if section:
        return section
    return [
        "",
        f"{research_slate.RESEARCH_LABEL} — {league} {target_date} — 0 fixture(s)",
        "- No scheduled fixture for this lane on this date, so there is nothing to score.",
        "- This is an honest zero, not a failed read: the table was read and no game was listed.",
    ]


def _read_text_factory(delay: float, stamp_path=None):
    """A paced transport for live runs.

    Without a stamp path this paces within one process. With one it also honours
    the gap *between* processes, which is the case that produced a 429: a
    one-request diagnostic, then this script, seconds apart. The pacing lives here,
    once, and ``fetch_published`` is then called with ``delay=0.0`` -- the inner
    reader it builds would otherwise sleep again for the same request.
    """

    from ingestion.http import HttpClient

    client = HttpClient()

    def read(url: str) -> str:
        return client.get_text(url)

    if stamp_path is None:
        return betexplorer_standings.paced(read, delay=delay)
    return betexplorer_standings.cooldown(read, stamp_path, delay=delay)


def _utc_now() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def read_slate(args, league, read, now_utc: str):
    """One attempt at the whole read: fixtures, prices, table.

    Kept separate from :func:`main` so a retry is a retry of *the read* and not of
    the writing: a card is only touched once a read has either succeeded or run out
    of attempts.
    """

    if args.from_store:
        from ingestion.duckdb_store import DuckDBStore

        store = DuckDBStore(args.warehouse)
        try:
            games = slate_games(
                store.list_games(league=league),
                args.date,
                tz_offset_hours=args.tz_offset_hours,
            )
            snapshots = {
                game.game_id: store.get_book_snapshots(game.game_id) for game in games
            }
        finally:
            store.close()
        published, report = fetch_published(league.value, read, delay=0.0)
        return games, snapshots, published, report

    # Default: the fixtures page is both the slate and the route to a match page,
    # so a lane with no warehouse is still scoreable.
    games, snapshots, fixtures_html, page_reads = slate_from_page(
        league.value, args.date, read, now_utc=now_utc
    )
    published, report = fetch_published(
        league.value, read, fixtures_html=fixtures_html, delay=0.0
    )
    report["requests"] = int(report.get("requests") or 0) + page_reads
    return games, snapshots, published, report


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--league", required=True)
    parser.add_argument("--date", required=True, help="YYYY-MM-DD, the card's date")
    parser.add_argument("--localdata", default="localdata")
    parser.add_argument(
        "--from-store",
        action="store_true",
        help="score the warehouse's own games instead of the fixtures page",
    )
    parser.add_argument("--warehouse", default="localdata/warehouse.duckdb")
    parser.add_argument(
        "--tz-offset-hours",
        type=int,
        default=0,
        help="lane wall-clock offset from UTC, for --from-store slate selection",
    )
    parser.add_argument(
        "--bet-config",
        default=str(TUNED_PATH),
        help=(
            "tuned bet configuration the card should run under; the shipped baseline "
            "is always loaded first, so a missing file means defaults"
        ),
    )
    parser.add_argument(
        "--delay",
        type=float,
        default=betexplorer_standings.CRAWL_DELAY_SECONDS,
        help="seconds between requests (the host's smallest safe spacing)",
    )
    parser.add_argument(
        "--retries",
        type=int,
        default=1,
        help="retries after a 429; a throttle is transient, a wrong header is not",
    )
    parser.add_argument(
        "--retry-wait",
        type=float,
        default=45.0,
        help="seconds to wait before a retry",
    )
    args = parser.parse_args(argv)

    league = parse_league(args.league)
    spec = league_spec(league)
    localdata = Path(args.localdata)
    stamp_path = localdata / ".betexplorer_last_request"
    read = _read_text_factory(args.delay, stamp_path)
    now_utc = _utc_now()

    def fail(reason: str) -> int:
        section = failure_section(league.value, args.date, reason)
        written = write_section(
            localdata, state_key=spec.state_key, target_date=args.date, section=section
        )
        print(f"{league.value} {args.date}: READ FAILED — {reason}")
        for kind, path in written.items():
            print(f"  {kind}: {path}")
        return 1

    attempts = 0
    while True:
        try:
            games, snapshots, published, report = read_slate(args, league, read, now_utc)
            break
        except HttpError as exc:
            if exc.status == 429 and attempts < args.retries:
                attempts += 1
                print(
                    f"  HTTP 429 — waiting {args.retry_wait:.0f}s, "
                    f"retry {attempts}/{args.retries}"
                )
                time.sleep(args.retry_wait)
                continue
            return fail(
                f"{type(exc).__name__}: {exc}"
                + (f" (after {attempts + 1} attempt(s))" if attempts else "")
            )
        except (betexplorer_standings.IdentityError, OSError, ValueError) as exc:
            # A shape this reader refuses is not transient: no retry.
            return fail(f"{type(exc).__name__}: {exc}")

    section = build_section(
        league.value,
        games,
        published,
        snapshots,
        target_date=args.date,
        bet_config=load_tuned(tuned_path=args.bet_config),
    )
    written = write_section(
        localdata, state_key=spec.state_key, target_date=args.date, section=section
    )
    print(
        f"{league.value} {args.date}: {len(games)} fixture(s), "
        f"{len(published)} published row(s), requests={report.get('requests')}, "
        f"subs={report.get('subs_read')}, mismatch={report.get('split_mismatch')}, "
        f"waited={getattr(read, 'waited', 0.0):.1f}s"
    )
    for kind, path in written.items():
        print(f"  {kind}: {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
