"""BetExplorer's published league table, read from a pre-match page's widget.

Why this exists
---------------
``forecasting/magolide_score.py`` rates a fixture from standings differences and
needs ``rank``, ``pct``, ``net_rtg``, ``home_pct`` and ``away_pct`` per club. For
the aggregate-only lanes -- Greece, WCBA, WNBL, NKL -- nothing supplied them, so
the research score could not be produced and those lanes emitted no pick rows at
all. This module reads the table the site already publishes for exactly those
competitions.

Evidence status (see ``docs/PROVIDERS.md``)
------------------------------------------
``LIVE_VERIFIED 2026-11-25 (LITHUANIA_NKL)``. The served capture behind every
constant here:

* a **match page** ``/basketball/lithuania/nkl/jurbarkas-vytis-vdu/McCqxqs3/``
  answers 200 and carries a stats widget, ``<div class="glib-stats-content"
  id="glib-stats">``, whose menu links are the only route to a table;
* that widget writes three table tabs of its own --
  ``/basketball/lithuania/nkl/standings/?table=table&table_sub=&ts=<token>
  &dcheck=0&as-ajax=1&l=en&event_context=<match_id>`` plus ``table_sub=overall``,
  ``table_sub=home`` and ``table_sub=away``, and a sibling ``table=form`` tab;
* the ``table_sub=overall`` link, fetched with ``Referer`` set to the match page,
  answers **200 / 5,268 bytes** and its body is a real server-rendered table,
  ``<table id="table-type-1" class="stats-table stats-main table-1">``;
* its head is ``# Team MP W L TP Pts Form`` and its first body row is
  ``Telsiai 16 15 1 1449:1186 15``.

Read from those bytes, not assumed:

* cells carry a ``col_*`` class token (``col_rank``, ``col_participant_name``,
  ``col_matches_played``, ``col_wins``, ``col_losses``, ``col_goals``,
  ``col_points``, ``col_form``), so columns are found by class rather than by
  position -- ``Pts`` appearing between ``TP`` and ``Form`` moved nothing;
* a club id rides on the row (``glib-participant-IHmmS3rt``) and in the team
  anchor's ``onclick``, while the anchor *text* is the display name;
* the rank cell's ``title`` names the stage, observed as
  ``Promotion - NKL (Winners stage: )`` -- a stage table is a subset, so the
  label is carried on every row instead of being flattened away.

What is deliberately NOT read
-----------------------------
The ``Form`` cell is left alone. Only one form glyph was visible in the capture,
so which class means a win and which means a loss is **not** established; a
coin-flipped streak would be worse than no streak. ``forecasting/standings.py``
derives the streak from finals we graded ourselves, and that stays the source.

Values are arithmetic on published numbers only: ``win_pct = W / MP`` and
``points_for``/``points_against`` are the cumulative ``PF:PA`` totals verbatim.
``net_rtg`` is computed downstream as ``(PF - PA) / MP`` -- a points-per-game
difference, never a possession-adjusted rating. A row whose core cells are
missing or self-contradictory (``W + L != MP``, ``W > MP``) is **skipped and
counted**, never repaired: a table we fixed is no longer a table we read.

Cost and pacing
---------------
One match page plus one fragment per sub-table (three by default) per
competition. The token is minted by the match page and is cached for the run that
read it -- never across runs, because a stale token is an unmeasured failure. The
host's tolerance was measured at roughly one request per ten seconds (a 2-second
spacing drew 429 and an empty body on the second request), so reads go through
:func:`paced`, which sleeps *between* requests and exposes an injectable sleep so
no test ever waits.

Nothing here is a ticket input. The table is aggregate-published, the prices in
the same lane are the site average, and the rows are research-scoped.
"""

from __future__ import annotations

import re
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Dict, List, Optional, Tuple

from ..identity import canonical_team_name
from ..models import StandingsRow
from . import IdentityError
from .betexplorer import FIXTURES_URLS, _require_league
from .nkl import PacedReader

PROVIDER = "betexplorer"
SOURCE = "betexplorer_published_table"

#: The evidence label a caller may attach to anything derived from these rows.
EVIDENCE = (
    "LIVE_VERIFIED (NKL table fragment captured 2026-11-25; published league "
    "table, aggregate source, research-scoped)"
)

BASE = "https://www.betexplorer.com/basketball"

#: Sub-tables the match page offers, in the order the widget lists them.
TABLE_SUBS = ("overall", "home", "away")

#: The floor, from measurements rather than taste: 2 s between requests drew 429
#: on the second (2026-11-25), 10 s drew 429 on the fourth (2026-11-26, after three
#: answered), 12 s and 20 s both completed four requests. Fifteen is the smallest
#: spacing that has never failed here.
CRAWL_DELAY_SECONDS = 15.0

#: The table element the fragment answers with. Any ``stats-table`` will do: the
#: capture wrote ``class="stats-table stats-main table-1"``, and the head is
#: validated separately, so a class reorder is not a parse failure.
_TABLE_RE = re.compile(
    r"<table[^>]*class=\"[^\"]*stats-table[^\"]*\"[^>]*>(?P<body>.*?)</table>", re.S
)
_ROW_RE = re.compile(r"<tr[^>]*>(?P<row>.*?)</tr>", re.S)
_CELL_RE = re.compile(r"<td(?P<attrs>[^>]*)>(?P<body>.*?)</td>", re.S)
_TH_RE = re.compile(r"<th(?P<attrs>[^>]*)>(?P<text>.*?)</th>", re.S)
_DATA_TYPE_RE = re.compile(r'data-type="(?P<kind>[^"]*)"')
_TAG_RE = re.compile(r"<[^>]+>")
_ANCHOR_TEXT_RE = re.compile(r"<a\b[^>]*>(?P<text>[^<]*)</a>", re.S)
_CLUB_ID_RE = re.compile(r"glib-participant-(?P<id>[A-Za-z0-9]{6,})")
_TITLE_RE = re.compile(r"title=\"(?P<title>[^\"]*)\"")
_SCORE_SPLIT_RE = re.compile(r"^\s*(?P<pf>\d{1,4})\s*:\s*(?P<pa>\d{1,4})\s*$")

#: The contract is the **body cell classes**, not the head's labels. WCBA serves
#: ``# | Group A | MP | W | L | TP | PCT | Form`` (2026-11-26) where the capture
#: served ``# | Team | MP | W | L | TP | Pts | Form``: the team column is labelled
#: with its group, and the points column carries a percentage instead of a count.
#: Both pages write the same ``col_*`` classes on their cells, so those are what a
#: table must carry to be read.
REQUIRED_CELL_CLASSES = ("col_matches_played", "col_wins", "col_losses")
TEAM_CELL_CLASSES = ("col_participant_name", "col_name")

#: The tab links a match page writes for itself. The token lives inside these and
#: nowhere else on the page, which is why the match page is read first.
_TABLE_TAB_RE = re.compile(
    r"(?P<path>/basketball/[a-z0-9-]+/[a-z0-9-]+/standings/\?"
    r"[^\s\"'<>]*?table=table&(?P<query>[^\s\"'<>]*))"
)


def competition_prefix(league: str) -> str:
    """``.../lithuania/nkl`` for a league key, derived from its fixtures URL."""

    league_key = _require_league(league)
    url = FIXTURES_URLS[league_key]
    suffix = "/fixtures/"
    if not url.endswith(suffix):  # pragma: no cover - registry invariant
        raise IdentityError(f"{PROVIDER}: {league_key} has no /fixtures/ base URL")
    return url[: -len(suffix)]


def match_url_for(league: str, slug: str, match_id: str) -> str:
    """Rebuild a match page URL from a fixtures row's slug and id.

    The fixtures table gives both, and the match page is the only thing that can
    mint a standings token, so this is the one URL shape a caller never has to
    invent.
    """

    if not slug or not match_id:
        raise IdentityError(f"{PROVIDER}: match page needs both slug and match_id")
    return f"{competition_prefix(league)}/{slug}/{match_id}/"


def _text(fragment: str) -> str:
    return re.sub(r"\s+", " ", _TAG_RE.sub(" ", fragment)).strip()


def _cell_class(attrs: str) -> str:
    found = re.search(r"class=\"(?P<class>[^\"]*)\"", attrs)
    return found.group("class") if found else ""


def _clean_number(text: str) -> Optional[int]:
    match = re.match(r"^\s*(?P<number>\d{1,4})\s*\.?\s*$", _text(text))
    if not match:
        return None
    return int(match.group("number"))


@dataclass
class ParsedTable:
    """One fragment, parsed, with the club ids and the stage label it carried."""

    sub: str
    rows: List[StandingsRow] = field(default_factory=list)
    team_ids: Dict[str, str] = field(default_factory=dict)
    stage: Optional[str] = None
    report: Dict[str, object] = field(default_factory=dict)

    def __len__(self) -> int:
        return len(self.rows)


def _columns(head: str) -> List[str]:
    """The head's labels, in order -- reported, never validated against."""

    return [text for text in (_text(cell.group("text")) for cell in _TH_RE.finditer(head)) if text]


def _team_label(head: str) -> Optional[str]:
    """The team column's own label. A group table puts its group here.

    Found by the same classes the body uses, so ``Group A`` and ``Team`` are the
    same finding: the column the clubs live in.
    """

    for cell in _TH_RE.finditer(head):
        attrs = cell.group("attrs")
        kinds = set(_cell_class(attrs).split())
        typed = _DATA_TYPE_RE.search(attrs)
        if typed:
            kinds.add(typed.group("kind"))
        if kinds & (set(TEAM_CELL_CLASSES) | {"col_participant_name"}):
            return _text(cell.group("text"))
    return None


def parse_standings_table(html: str, sub: str = "overall") -> ParsedTable:
    """Parse the standings fragment into canonical rows.

    Every ``stats-table`` in the body is read, in order, because a competition may
    publish one table per group (WCBA labels its team column ``Group A``): the
    fragment is a group table, not a league table, and pretending otherwise would
    drop half the competition. The group label rides on each row's ``section`` when
    the rank cell does not name a stage.

    Raises ``IdentityError`` when the body carries no standings table at all -- an
    empty body (the shape the competition-level route returned) or a fixtures table
    -- and when tables are present but yield no row. Guessing at a table that is
    not there is what produced a fixtures table being read as standings once
    already, and the head's labels are *not* used as a gate: ``Group A``/``PCT`` is
    a shape this reader must accept, not refuse.
    """

    tables = list(_TABLE_RE.finditer(html or ""))
    if not tables:
        raise IdentityError(
            f"{PROVIDER}: no standings table in the {sub} fragment "
            f"({len(html or '')} bytes)"
        )

    parsed = ParsedTable(sub=sub, report={"tables": len(tables)})
    skipped: Dict[str, int] = {}
    groups: List[str] = []
    columns: List[str] = []
    seen_teams: set[str] = set()

    def skip(reason: str) -> None:
        skipped[reason] = skipped.get(reason, 0) + 1

    for index, table in enumerate(tables):
        body = table.group("body")
        head_end = body.find("</thead>")
        head, rows_html = (body[:head_end], body[head_end:]) if head_end >= 0 else ("", body)
        if not columns:
            columns = _columns(head)
        label = _team_label(head)
        group_label = label if label and label not in {"Team", "#", ""} else None
        if group_label and group_label not in groups:
            groups.append(group_label)

        for row in _ROW_RE.finditer(rows_html):
            # ``group(0)``, not ``group("row")``: the club id rides on the <tr> tag
            # itself (``class="odd glib-participant-IHmmS3rt"``).
            markup = row.group(0)
            cells = list(_CELL_RE.finditer(row.group("row")))
            if not cells:
                skip("no_cells")
                continue
            by_class: Dict[str, str] = {}
            rank_title: Optional[str] = None
            for cell in cells:
                attrs = cell.group("attrs")
                classes = _cell_class(attrs).split()
                for token in classes:
                    if token.startswith("col_"):
                        by_class.setdefault(token, cell.group("body"))
                if "col_rank" in classes:
                    title = _TITLE_RE.search(attrs)
                    rank_title = title.group("title") if title else None

            team_cell = by_class.get("col_participant_name") or by_class.get("col_name")
            if team_cell is None:
                # A stage/group separator row: it has a label, not a club.
                skip("no_team_cell")
                continue
            anchor = _ANCHOR_TEXT_RE.search(team_cell)
            name = _text(anchor.group("text") if anchor else team_cell)
            if not name:
                skip("empty_team_name")
                continue

            ranks = _clean_number(by_class.get("col_rank", ""))
            played = _clean_number(by_class.get("col_matches_played", ""))
            wins = _clean_number(by_class.get("col_wins", ""))
            losses = _clean_number(by_class.get("col_losses", ""))
            if played is None or wins is None or losses is None:
                skip("missing_core_cells")
                continue
            if played <= 0:
                skip("no_games_played")
                continue
            if wins + losses != played or wins > played or losses > played:
                skip("inconsistent_record")
                continue

            points_for: Optional[float] = None
            points_against: Optional[float] = None
            total = _SCORE_SPLIT_RE.match(_text(by_class.get("col_goals", "")))
            if total is not None:
                points_for = float(total.group("pf"))
                points_against = float(total.group("pa"))

            club_id = _CLUB_ID_RE.search(markup)
            canonical = canonical_team_name(name)
            if canonical in seen_teams:
                # The same club in two tables is one club; two groups overlap only
                # if the page repeated a row, and counting it twice would double
                # every total it contributes.
                skip("duplicate_team")
                continue
            seen_teams.add(canonical)
            section = rank_title or group_label
            if rank_title:
                parsed.stage = parsed.stage or rank_title
            parsed.rows.append(
                StandingsRow(
                    team_name=canonical,
                    position=ranks,
                    games_played=played,
                    wins=wins,
                    losses=losses,
                    points_for=points_for,
                    points_against=points_against,
                    win_pct=round(wins / played, 4),
                    # Streak and last-10 stay unset: the Form glyph is not decoded
                    # here (see the module docstring), so there is nothing honest to
                    # put in them.
                    section=section,
                    source=SOURCE,
                )
            )
            if club_id is not None:
                parsed.team_ids[canonical] = club_id.group("id")

    parsed.report.update(
        {
            "sub": sub,
            "columns": columns,
            "groups": groups,
            "rows_parsed": len(parsed.rows),
            "rows_skipped": sum(skipped.values()),
            "skip_reasons": skipped,
            "stage": parsed.stage,
            "has_points_column": any(column in {"Pts", "PCT"} for column in columns),
        }
    )
    if not parsed.rows:
        raise IdentityError(
            f"{PROVIDER}: {sub} fragment produced zero rows from {len(tables)} "
            f"table(s), head {columns or ['<none>']}, skipped: {skipped or 'none'}"
        )
    return parsed


def find_standings_links(match_html: str) -> Dict[str, str]:
    """The ``table_sub`` -> absolute fragment URL map a match page publishes.

    ``table_sub=`` appears empty on the wrapper link and explicit on the submenu
    links; the explicit ones are what get collected. A missing ``overall`` entry
    means this page has no table widget -- the caller reports that rather than
    constructing a token it never received.
    """

    links: Dict[str, str] = {}
    for match in _TABLE_TAB_RE.finditer(match_html or ""):
        path = match.group("path").replace("&amp;", "&")
        query = match.group("query").replace("&amp;", "&")
        sub = re.search(r"table_sub=(?P<sub>[a-z]*)", query)
        if sub is None or not sub.group("sub"):
            continue
        links.setdefault(sub.group("sub"), "https://www.betexplorer.com" + path)
    return links


def _record(wins: int, losses: int) -> str:
    return f"{wins}-{losses}"


def merge_splits(
    overall: ParsedTable,
    splits: Dict[str, ParsedTable],
) -> Tuple[List[StandingsRow], List[str]]:
    """Attach ``home_record``/``away_record`` from the home and away sub-tables.

    A club's record is attached only when the two sub-tables exist **and** add up
    to the overall record (home wins plus away wins equals total wins, likewise
    losses). Two fragments read seconds apart can come from different moments --
    a game finishing between the reads is enough -- and a home/away split that
    does not reconstruct the row beside it is the visible symptom. Such a club is
    left unset and named in the returned ``mismatched`` list, so the caller can
    report the gap instead of publishing a split that contradicts its own total.
    """

    by_team: Dict[str, Dict[str, StandingsRow]] = {}
    for sub, table in splits.items():
        for row in table.rows:
            by_team.setdefault(row.team_name, {})[sub] = row

    mismatched: List[str] = []
    for row in overall.rows:
        home = by_team.get(row.team_name, {}).get("home")
        away = by_team.get(row.team_name, {}).get("away")
        if home is None or away is None:
            continue
        cells = (home.wins, home.losses, away.wins, away.losses, row.wins, row.losses)
        if any(value is None for value in cells):
            continue
        if home.wins + away.wins != row.wins or home.losses + away.losses != row.losses:
            mismatched.append(row.team_name)
            continue
        row.home_record = _record(home.wins, home.losses)
        row.away_record = _record(away.wins, away.losses)
    return overall.rows, mismatched


@dataclass
class StandingsResult:
    """Everything one competition read produced, with the cost of producing it."""

    league: str
    rows: List[StandingsRow] = field(default_factory=list)
    team_ids: Dict[str, str] = field(default_factory=dict)
    stage: Optional[str] = None
    report: Dict[str, object] = field(default_factory=dict)

    def __len__(self) -> int:
        return len(self.rows)


def fetch_standings(
    league: str,
    read_text: Callable[[str], str],
    *,
    match_url: Optional[str] = None,
    slug: Optional[str] = None,
    match_id: Optional[str] = None,
    subs: Tuple[str, ...] = TABLE_SUBS,
    delay: Optional[float] = None,
    sleep: Optional[Callable[[float], None]] = None,
) -> StandingsResult:
    """Read one competition's published table from a pre-match page.

    ``read_text`` is a plain ``str -> str`` transport (the same shape
    ``betexplorer.fetch_odds`` accepts), so tests hand it a dict and no test
    touches the network. Requests are paced *between* calls; the count is
    reported so a receipt can state what a read cost.

    The token is used for this call only. It is not returned for reuse and not
    cached anywhere: a token that outlives the page that minted it is a silent
    failure waiting to be misread as "the source stopped publishing".
    """

    league_key = _require_league(league)
    target = match_url or match_url_for(league_key, slug or "", match_id or "")
    paced = PacedReader(read_text, delay=CRAWL_DELAY_SECONDS if delay is None else delay,
                        sleep=sleep)

    match_html = paced(target)
    links = find_standings_links(match_html)
    if "overall" not in links:
        raise IdentityError(
            f"{PROVIDER} {league_key}: match page {target} published no standings "
            f"table links ({len(match_html or '')} bytes, "
            f"tabs seen: {sorted(links) or 'none'})"
        )

    # One token per run: the fragment URLs on a page all carry the same value, so
    # a mismatch means the page was assembled from two pages and neither token can
    # be trusted.
    token_values = {
        found.group("ts")
        for url in links.values()
        for found in [re.search(r"[?&]ts=(?P<ts>[A-Za-z0-9]+)", url)]
        if found
    }
    if len(token_values) > 1:
        raise IdentityError(
            f"{PROVIDER} {league_key}: match page carried {len(token_values)} "
            "distinct standings tokens; refusing to choose"
        )

    parsed: Dict[str, ParsedTable] = {}
    for sub in subs:
        url = links.get(sub)
        if url is None:
            continue
        parsed[sub] = parse_standings_table(paced(url), sub=sub)

    if "overall" not in parsed:
        raise IdentityError(f"{PROVIDER} {league_key}: overall table was not requested")

    overall = parsed.pop("overall")
    rows, split_mismatch = merge_splits(overall, parsed)
    team_ids = dict(overall.team_ids)
    for table in parsed.values():
        for name, club_id in table.team_ids.items():
            team_ids.setdefault(name, club_id)

    return StandingsResult(
        league=league_key,
        rows=rows,
        team_ids=team_ids,
        stage=overall.stage,
        report={
            "provider": PROVIDER,
            "league": league_key,
            "verification": EVIDENCE,
            "requests": paced.calls,
            "subs_read": sorted(parsed) + ["overall"],
            "subs_missing": [sub for sub in subs if sub not in parsed and sub != "overall"],
            "token_seen": bool(token_values),
            "stage": overall.stage,
            "rows": len(rows),
            "split_mismatch": split_mismatch,
            "parse": overall.report,
            "role": "research_scoped_never_a_named_book_price",
        },
    )


class CooldownTransport:
    """A transport that honours the host's floor *between processes*.

    ``PacedReader`` spaces the requests of one run. It cannot see a run that
    happened two seconds earlier in another process, which is exactly how this
    host answered 429 on 2026-11-25: a one-request diagnostic, then the script,
    seconds apart. A stamp on disk is the only thing both processes share.

    The stamp is written in ``finally``: a request that failed still spent the
    host's tolerance, so the next caller must wait just the same. ``delay <= 0``
    disables both the wait and the write, which is what tests use.
    """

    def __init__(
        self,
        read_text: Callable[[str], str],
        stamp_path,
        *,
        delay: Optional[float] = None,
        sleep: Optional[Callable[[float], None]] = None,
        now: Optional[Callable[[], float]] = None,
    ) -> None:
        self._read_text = read_text
        self._stem = Path(stamp_path)
        self._delay = max(0.0, float(CRAWL_DELAY_SECONDS if delay is None else delay))
        self._sleep = sleep or time.sleep
        self._now = now or time.time
        self.calls = 0
        self.waited = 0.0

    def _last_request(self) -> Optional[float]:
        try:
            return float(self._stem.read_text().strip())
        except (OSError, ValueError):
            return None

    def __call__(self, url: str) -> str:
        if self._delay:
            last = self._last_request()
            if last is not None:
                remaining = self._delay - (self._now() - last)
                if remaining > 0:
                    self._sleep(remaining)
                    self.waited += remaining
        self.calls += 1
        try:
            return self._read_text(url)
        finally:
            if self._delay:
                try:
                    self._stem.parent.mkdir(parents=True, exist_ok=True)
                    self._stem.write_text(f"{self._now():.3f}", encoding="utf-8")
                except OSError:  # pragma: no cover - a read-only workspace
                    pass


def cooldown(
    read_text: Callable[[str], str],
    stamp_path,
    *,
    delay: Optional[float] = None,
    sleep: Optional[Callable[[float], None]] = None,
    now: Optional[Callable[[], float]] = None,
) -> CooldownTransport:
    """The transport a cross-process run should pass in."""

    return CooldownTransport(read_text, stamp_path, delay=delay, sleep=sleep, now=now)


def paced(read_text: Callable[[str], str], *, delay: Optional[float] = None,
          sleep: Optional[Callable[[float], None]] = None) -> PacedReader:
    """The reader a standings run should pass in, so the delay cannot be forgotten."""

    return PacedReader(
        read_text, delay=CRAWL_DELAY_SECONDS if delay is None else delay, sleep=sleep
    )


__all__ = [
    "CRAWL_DELAY_SECONDS",
    "CooldownTransport",
    "cooldown",
    "EVIDENCE",
    "PROVIDER",
    "REQUIRED_CELL_CLASSES",
    "SOURCE",
    "TABLE_SUBS",
    "TEAM_CELL_CLASSES",
    "ParsedTable",
    "StandingsResult",
    "competition_prefix",
    "fetch_standings",
    "find_standings_links",
    "match_url_for",
    "merge_splits",
    "paced",
    "parse_standings_table",
]
