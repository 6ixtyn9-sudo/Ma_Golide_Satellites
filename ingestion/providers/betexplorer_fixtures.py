"""Schedule discovery from the BetExplorer fixtures table.

Why this exists
---------------
Ten competitions had a price page but no schedule route, so a price could never
attach to a canonical game.  This adapter closes that loop for the leagues whose
fixtures table is the only public route they have: the same single page yields
the slate (teams, tip, provider match code) that :mod:`betexplorer` already
reads for its representative average.

What it is, precisely
---------------------
* **Source.** ``FIXTURES_URLS[league]`` — the league's fixtures page, fetched
  once per run.  Robots permits ``/basketball/`` paths; season navigation stays
  path-based and never uses a disallowed ``?stage=``/``?page=``/``?year=``
  query.  Source terms are unreviewed (``SourcePermission.UNRESOLVED``), so the
  quotes this page carries are recorded as an aggregate consensus and are
  research-only, never ticket-eligible.
* **Clock.** Every competition is served in the site's own CET/CEST wall clock,
  not the league's local time, and the page carries no ISO attribute to read.
  The label is therefore parsed as a naive ``Europe/Paris`` time and converted
  to UTC.  This is a reconstruction, and it is recorded as such: the game
  evidence says so, so no downstream reader can mistake it for provider truth.
* **Identity.** The stable 8-character match code in the row href.  Names are
  display only: this table serves ASCII-transliterated names with sponsor
  suffixes, colliding city fragments and a `` W`` suffix on women's clubs, so a
  name join would be a guess.
* **Missing kickoff.** Rows whose datetime cell is ``&nbsp;`` have no time the
  server ever transmits.  They are counted and reported, never dated by
  assumption and never invented as games.
* **Results.** This page serves no final score, so the lane has no settlement
  route.  :func:`last_gaps` reports that absence by name rather than letting a
  lane look healthy while it can never settle.

Nothing here promotes a page average to a named-book quote, and nothing here
writes a price: prices stay in :mod:`betexplorer`, keyed by the same
``betexplorer:<match_id>`` identity this module assigns.
"""

from __future__ import annotations

from datetime import datetime, timezone
from typing import Any, Callable, Dict, List, Optional, Tuple

from ..identity import to_utc_iso
from ..models import CoverageGap, CoverageState, Game, GameStatus, League
from . import IdentityError
from .betexplorer import FIXTURES_URLS, SITE_TZ, _rows, _text_fetch

PROVIDER = "betexplorer_fixtures"
SCHEDULE_SOURCE = "betexplorer_fixtures"

#: Competitions this adapter owns the schedule for.  Greece is the men's league
#: whose board only ever served prices; the rest are the four 2026-11-24
#: additions.
LANES: Tuple[str, ...] = (
    "CHINA_WCBA_W",
    "LITHUANIA_NKL",
    "AUSTRALIA_WNBL_W",
    "ISRAEL_LIGA_LEUMIT",
    "GREECE_BASKET_LEAGUE",
)

_LAST_GAPS: List[CoverageGap] = []
_LAST_PROBE: Optional[Dict[str, Any]] = None


def is_enabled() -> bool:
    return True


def last_gaps() -> List[CoverageGap]:
    return list(_LAST_GAPS)


def last_probe() -> Optional[Dict[str, Any]]:
    return dict(_LAST_PROBE) if _LAST_PROBE else None


def _reset() -> None:
    _LAST_GAPS.clear()
    global _LAST_PROBE
    _LAST_PROBE = None


def _gap(required_field: str, detail: str, game_id: Optional[str] = None) -> None:
    _LAST_GAPS.append(
        CoverageGap(
            game_id=game_id,
            required_field=required_field,
            reason=CoverageState.MISSING,
            provider=PROVIDER,
            detail=detail[:500],
        )
    )


def _require_lane(league: League | str) -> str:
    key = league.value if isinstance(league, League) else str(league).strip().upper()
    if key not in LANES:
        raise ValueError(f"{PROVIDER} is not the schedule source for {key!r}")
    return key


def _season_label(local_dt: datetime) -> str:
    """A season label derived from the fixture's own month.

    The page never states the season on a row.  July-and-later belongs to the
    season that started that calendar year; earlier months belong to the season
    that started the previous year.  This is a label for grouping, not a claim
    about provider metadata.
    """

    start = local_dt.year if local_dt.month >= 7 else local_dt.year - 1
    return f"{start}-{start + 1}"


def fetch_slate(
    league: League | str,
    date_yyyymmdd: str,
    *,
    text_fetch: Optional[Callable[[str], Any]] = None,
    fetch_factory: Optional[Callable[[Any], Any]] = None,
    now_utc: Optional[str] = None,
) -> List[Game]:
    """Return fixtures for ``date_yyyymmdd`` plus the next scheduled fixture day.

    The window spans the site-local (Paris) target date plus the earliest later
    own-dated fixture day on the same page (preview-only downstream).  Every
    returned :class:`Game` stamps ``evidence["slate_date"]`` with its provider
    Paris date and ``evidence["slate_role"]`` as ``"target"`` or ``"preview"``.
    Own-dated rows beyond the preview day are counted in
    ``last_probe()["dropped_out_of_window"]`` rather than discarded silently.
    """

    key = _require_lane(league)
    _reset()
    target = str(date_yyyymmdd).replace("-", "")
    datetime.strptime(target, "%Y%m%d")
    observed = (
        datetime.fromisoformat(str(now_utc).replace("Z", "+00:00"))
        if now_utc
        else datetime.now(timezone.utc)
    )
    if observed.tzinfo is None:
        raise ValueError("now_utc must be timezone-aware")

    url = FIXTURES_URLS[key]
    read_text = _text_fetch(fetch_factory, text_fetch)
    html = read_text(url)
    if not isinstance(html, str) or "table-main" not in html:
        raise IdentityError(f"{PROVIDER} {key}: fixtures page did not return a results table")

    parsed = _rows(html, observed)
    raw_rows = html.count("<tr>")
    unparsed = max(0, raw_rows - len(parsed))

    # Find the earliest own-dated fixture strictly after target from the same
    # parsed rows (calendar property of the competition page, not tick time).
    preview_yyyymmdd: Optional[str] = None
    for local_dt, _match_id, _slug, _home, _away, _odds, raw_label in parsed:
        if not raw_label or local_dt is None:
            continue
        row_yyyymmdd = local_dt.strftime("%Y%m%d")
        if row_yyyymmdd > target and (
            preview_yyyymmdd is None or row_yyyymmdd < preview_yyyymmdd
        ):
            preview_yyyymmdd = row_yyyymmdd

    preview_iso = (
        f"{preview_yyyymmdd[:4]}-{preview_yyyymmdd[4:6]}-{preview_yyyymmdd[6:]}"
        if preview_yyyymmdd
        else None
    )

    games: List[Game] = []
    seen: set[str] = set()
    seen_undated: set[str] = set()
    blank_kickoff = 0
    blank_kickoff_with_date = 0
    target_games = 0
    preview_games = 0
    slate_priced = 0
    preview_priced = 0
    page_priced = 0
    dropped_out_of_window = 0
    dropped_past_target = 0
    next_fixture_date: Optional[str] = None
    for local_dt, match_id, _slug, home, away, odds, raw_label in parsed:
        if not raw_label:
            # A row whose own datetime cell is blank has no served kickoff
            # time. Never promote an inherited previous-row clock into a Game;
            # emit one named gap per undated fixture carrying the page's round
            # date when a preceding row established one.
            blank_kickoff += 1
            if local_dt is not None:
                blank_kickoff_with_date += 1
            if match_id not in seen_undated:
                seen_undated.add(match_id)
                matchup = f"{home.strip()} vs {away.strip()}"
                if local_dt is not None:
                    fixture_date = local_dt.strftime("%Y-%m-%d")
                    detail = (
                        f"{key}: fixture {match_id} ({matchup}) on {fixture_date} "
                        "has no kickoff time served yet; not dated by assumption "
                        "(becomes a game on a later refresh once the host "
                        "publishes the tip)"
                    )
                else:
                    detail = (
                        f"{key}: fixture {match_id} ({matchup}) has no kickoff "
                        "date or time served yet; not dated by assumption"
                    )
                _gap(
                    "fixtures_kickoff_missing",
                    detail,
                    game_id=f"betexplorer:{match_id}",
                )
            continue
        if local_dt is None:
            continue
        if odds:
            page_priced += 1
        row_yyyymmdd = local_dt.strftime("%Y%m%d")
        row_iso = local_dt.strftime("%Y-%m-%d")
        if row_yyyymmdd >= target:
            if next_fixture_date is None or row_iso < next_fixture_date:
                next_fixture_date = row_iso
        if row_yyyymmdd < target:
            dropped_past_target += 1
            continue
        if row_yyyymmdd == target:
            is_preview = False
        elif preview_yyyymmdd is not None and row_yyyymmdd == preview_yyyymmdd:
            is_preview = True
        else:
            dropped_out_of_window += 1
            continue
        if odds:
            if is_preview:
                preview_priced += 1
            else:
                slate_priced += 1
        if match_id in seen:
            continue
        seen.add(match_id)
        if is_preview:
            preview_games += 1
        else:
            target_games += 1
        scheduled_utc = to_utc_iso(local_dt)
        games.append(
            Game(
                game_id=f"{'betexplorer'}:{match_id}",
                league=League(key),
                season=_season_label(local_dt),
                # The table publishes no club identifier.  The match code is
                # the only stable key, and it is already part of the game id;
                # team ids stay empty rather than inventing a franchise.
                home_team_id="",
                away_team_id="",
                home_name=home.strip(),
                away_name=away.strip(),
                scheduled_utc=scheduled_utc,
                status=GameStatus.SCHEDULED,
                provider_ids={
                    "betexplorer": match_id,
                    "betexplorer_slug": _slug,
                },
                evidence={
                    "tip_clock": "site CET/CEST wall clock, no ISO attribute served",
                    "tip_timezone_assumed": SITE_TZ,
                    "slate_date": row_iso,
                    "slate_role": "preview" if is_preview else "target",
                    "identity": "provider match code from the row href",
                    "name_join": "not used",
                    "source_permission": "unresolved",
                    "price_provenance": "aggregate_consensus",
                },
            )
        )

    games.sort(key=lambda game: (game.scheduled_utc, game.game_id))
    global _LAST_PROBE
    _LAST_PROBE = {
        "league": key,
        "url": url,
        "target_date": target,
        "slate_window_mode": "next_fixture_day",
        "rows": raw_rows,
        "parsed_rows": len(parsed),
        "blank_kickoff_rows": blank_kickoff,
        "blank_kickoff_rows_with_date": blank_kickoff_with_date,
        "unparsed_rows": unparsed,
        "slate_games": len(games),
        "target_games": target_games,
        "preview_date": preview_iso,
        "preview_games": preview_games,
        # Scope matters: the page prices a handful of rows per day, so a
        # whole-page count would report a slate as priced when its own rows
        # carry nothing.
        "slate_priced_rows": slate_priced,
        "preview_priced_rows": preview_priced,
        "page_priced_rows": page_priced,
        "dropped_out_of_window": dropped_out_of_window,
        "dropped_out_of_window_reason": "dated_rows_beyond_next_fixture_day",
        "dropped_past_target": dropped_past_target,
        "next_fixture_date": next_fixture_date,
    }

    if games and not (slate_priced + preview_priced):
        # Not an error: the page was read and had no quote.  It is a fact the
        # run must state, because "no odds" and "no page" look identical in a
        # report that only prints counts.
        _gap(
            "fixtures_price_coverage",
            f"{key}: {len(games)} fixture(s) in window starting {target} "
            f"(target={target_games}, preview={preview_games}) and none "
            "carried a both-sided quote",
        )
    return games
