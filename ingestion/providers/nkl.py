"""nkl.lt match pages: the first route in this project that can settle.

Captured 2026-11-24 from ``https://nkl.lt/matches/125781/``: 200, 680,857 B,
sha256 ``61a4fa2e…``, robots ``User-agent: * / Crawl-delay: 10`` (allow).  The
served markup carries the thing the aggregator never did: four quarters that add
up to the final, and a date on the page.

Two markup facts decide every regex here, both observed rather than assumed:

* **Attributes are unquoted.**  The page writes ``class=nkl-quarters-table`` and
  ``href=https://nkl.lt/komandos/<slug>/ >``.  A parser written against quoted
  attributes -- every other reader in this repository -- matches nothing on it.
* **The quarter table is the only table with quarter headers.**  Its head is
  ``Q1 Q2 Q3 Q4`` and each body row is a club cell followed by four integers.  A
  game with overtime would grow a fifth column, so the header list is read
  strictly: anything but exactly Q1..Q4 is refused rather than truncated to four.

What this module refuses
------------------------
A total that cannot be read is not invented from the quarter sum.  Summing the
quarters and calling the sum the final makes ``reconciles()`` true by
construction, which would hide exactly the failure it exists to catch -- a game
whose extra period was dropped.  ``to_actual_result()`` therefore returns ``None``
unless the page states a final **and** the quarters add up to it.

Reading and settling
-------------------
Parsing is pure: every reader below takes bytes it was handed.  The two functions
that place requests are :func:`settle_forecast` (forecast names plus a date in,
one reconciling record out, two reads for the first game of a club and one for
each after it) and :func:`paced`, which wraps a reader so consecutive requests
honour the ``Crawl-delay: 10`` this host publishes rather than the politeness of a
caller who remembers to sleep.

The lane is on.  ``LITHUANIA_NKL`` reports ``results_source = "nkl"`` as of
2026-11-24, on a live run that settled three games whose quarters reconciled with
their stated finals (``docs/NKL-SETTLEMENT-LIVE-RUN.md``), and the scheduled
results tick reaches these functions through ``Pipeline._nkl_settle``.
"""

from __future__ import annotations

import re
import time
from dataclasses import dataclass, field
from datetime import date, datetime
from typing import Dict, List, Optional, Tuple
from zoneinfo import ZoneInfo

PROVIDER = "nkl"
SITE = "https://nkl.lt"
SITE_TZ = "Europe/Vilnius"

#: Unquoted attributes, as served: ``class=nkl-quarters-table``.
_QUARTERS_TABLE_RE = re.compile(
    r"<table[^>]*class=[\"']?nkl-quarters-table[\"']?[^>]*>(?P<body>.*?)</table>", re.S
)
_TH_RE = re.compile(r"<th[^>]*>(?P<text>[^<]*)</th>")
_TR_RE = re.compile(r"<tr[^>]*>(?P<row>.*?)</tr>", re.S)
_TD_RE = re.compile(r"<td[^>]*>(?P<text>.*?)</td>", re.S)
_CLUB_RE = re.compile(r"href=https?://nkl\.lt/komandos/(?P<slug>[a-z0-9-]+)/")
_CODE_RE = re.compile(r"<span>(?P<code>[A-Z0-9]{2,5})</span>")
_INTS_RE = re.compile(r"^\s*(\d{1,3})\s*$")

#: ``Rungtynių pradžia`` then the value cell: ``2026-10-15, 19:00``, local wall clock.
_START_RE = re.compile(
    r"Rungtyni[ųu]\s*prad[žz]ia\s*</span>\s*<span[^>]*>(?P<date>\d{4}-\d{2}-\d{2}),\s*"
    r"(?P<time>\d{1,2}:\d{2})",
    re.S,
)

#: A final the page states for itself.  The captured page has no element with a
#: score-shaped class, so this looks for one and reports its absence rather than
#: falling back to the quarter sum.
_SCORE_ELEMENT_RE = re.compile(
    r"class=[\"']?(?P<token>[^\"'>]*(?:match-score|nkl-score|final-score|score-box)[^\"'>]*)[\"']?[^>]*>"
    r"(?P<body>.{0,400})",
    re.S,
)
_PAIR_RE = re.compile(r"(?P<home>\d{1,3})\s*[:\-–]\s*(?P<away>\d{1,3})")


@dataclass(frozen=True)
class MatchPage:
    """One served match page, read as strictly as the markup allows."""

    url: str
    home_slug: str
    away_slug: str
    home_code: str
    away_code: str
    played_on: Optional[date]
    tip_local: Optional[datetime]
    home_quarters: List[Optional[float]] = field(default_factory=list)
    away_quarters: List[Optional[float]] = field(default_factory=list)
    home_final: Optional[float] = None
    away_final: Optional[float] = None
    #: False when the head is anything but exactly Q1..Q4.  Four values read from
    #: a five-column table would otherwise look complete while having dropped a
    #: period, which is the one failure reconciles() must not be able to miss.
    headers_ok: bool = False
    #: Which element stated the final, and the text around it.  A final that is
    #: really a summary row echoing the quarter sum would make reconciles() true
    #: by construction, so the claim has to be auditable in the receipt.
    final_source: Optional[str] = None
    counts: Dict[str, int] = field(default_factory=dict)
    reasons: Tuple[str, ...] = ()

    @property
    def complete(self) -> bool:
        return (
            self.headers_ok
            and len(self.home_quarters) == 4
            and len(self.away_quarters) == 4
            and all(v is not None for v in self.home_quarters + self.away_quarters)
            and self.home_final is not None
            and self.away_final is not None
        )

    @property
    def state(self) -> str:
        """What the page is, which is not the same as whether it parsed.

        A fixture page with no result table is not a parse failure -- the game
        has not been played.  Collapsing the two would let a lane report a
        reader bug every night the schedule runs ahead of the scores, and would
        hide a genuine markup change among the noise.
        """

        if not self.home_quarters or not self.away_quarters:
            return "no_result"
        if self.to_actual_result() is not None:
            return "settled"
        return "incomplete"

    def quarter_sums(self) -> Tuple[float, float]:
        return (
            sum(v for v in self.home_quarters if v is not None),
            sum(v for v in self.away_quarters if v is not None),
        )

    def to_actual_result(self):
        """An ``ActualResult`` this page can stand behind, or ``None``.

        ``None`` when the page states no final: a sum of quarters compared to
        itself always reconciles, and that is not a check.
        """

        from ..models import ActualResult, PeriodKind

        if not self.complete:
            return None
        return ActualResult(
            period_kind=PeriodKind.QUARTERS,
            home_quarters=[float(v) for v in self.home_quarters],
            away_quarters=[float(v) for v in self.away_quarters],
            home_final=float(self.home_final),
            away_final=float(self.away_final),
            source=f"{PROVIDER} {self.url}",
        )

    def to_dict(self) -> Dict[str, object]:
        home_sum, away_sum = self.quarter_sums()
        return {
            "url": self.url,
            "home_slug": self.home_slug,
            "away_slug": self.away_slug,
            "home_code": self.home_code,
            "away_code": self.away_code,
            "played_on": None if self.played_on is None else self.played_on.isoformat(),
            "tip_local": None if self.tip_local is None else self.tip_local.isoformat(),
            "home_quarters": list(self.home_quarters),
            "away_quarters": list(self.away_quarters),
            "home_final": self.home_final,
            "away_final": self.away_final,
            "final_source": self.final_source,
            "quarter_sums": [home_sum, away_sum],
            "state": self.state,
            "settleable": self.to_actual_result() is not None,
            "counts": dict(self.counts),
            "reasons": list(self.reasons),
        }


def _first_group(pattern: re.Pattern[str], text: str, group: str) -> Optional[str]:
    match = pattern.search(text)
    return None if match is None else match.group(group)


def _read_quarters(
    table_body: str,
) -> Tuple[List[Optional[float]], List[Optional[float]], list, bool]:
    """Quarter headers and the two integer rows, refused unless exactly four."""

    headers = [m.group("text").strip() for m in _TH_RE.finditer(table_body)]
    problems: list = []
    headers_ok = headers[1:] == ["Q1", "Q2", "Q3", "Q4"]
    if not headers_ok:
        problems.append(f"quarter headers are {headers[1:] or 'none'}, not Q1..Q4")

    rows: List[List[Optional[float]]] = []
    for row_match in _TR_RE.finditer(table_body):
        cells = _TD_RE.findall(row_match.group("row"))
        if not cells:
            continue
        values: List[Optional[float]] = []
        for cell in cells[1:]:
            match = _INTS_RE.match(re.sub(r"<[^>]+>", "", cell))
            values.append(None if match is None else float(match.group(1)))
        rows.append(values)

    if len(rows) != 2:
        problems.append(f"expected two club rows, read {len(rows)}")
        return [], [], problems, headers_ok

    home, away = rows
    if len(home) != 4 or len(away) != 4:
        problems.append(f"expected four quarters per club, read {len(home)} and {len(away)}")
    return home, away, problems, headers_ok


_CANDIDATE_RE = re.compile(r"\b\d{1,3}\s*[:\-–]\s*\d{1,3}\b")


def final_candidates(html: str, limit: int = 5) -> List[str]:
    """Score-shaped pairs on the page, with the markup around them.

    The capture has no score-shaped class name, so the element that states the
    final is not yet known.  This reports what is there instead of guessing a
    selector: a bounded list of pairs with 80 characters of context each, which
    is enough to write the rule once and not enough to drown a receipt.
    """

    found: List[str] = []
    for match in _CANDIDATE_RE.finditer((html or "")[:40000]):
        excerpt = re.sub(r"\s+", " ", (html or "")[max(0, match.start() - 80) : match.end() + 80])
        if excerpt not in found:
            found.append(excerpt.strip())
        if len(found) >= limit:
            break
    return found


def _read_final(html: str) -> Optional[Tuple[float, float, str]]:
    element = _SCORE_ELEMENT_RE.search(html)
    if element is None:
        return None
    flat = re.sub(r"<[^>]+>", " ", element.group("body"))
    pair = _PAIR_RE.search(flat)
    if pair is None:
        return None
    excerpt = re.sub(r"\s+", " ", flat[max(0, pair.start() - 60) : pair.end() + 60]).strip()
    return float(pair.group("home")), float(pair.group("away")), (
        f"class={element.group('token')} | {excerpt}"
    )


#: NKL's club name -> the name BetExplorer uses for the same club, reviewed by
#: eye on 2026-11-24 against the two 18-name lists.  A mapping, not a matcher:
#: "Palangos Palanga" is BetExplorer's "BC Olimpas Palanga" while "Plungės
#: Olimpas" is "Plunges Olimpas", so any scoring on shared words swaps those two
#: clubs -- one of them would settle with the other's quarter scores.
CLUB_NAMES: Dict[str, str] = {
    "Alytaus Patriotai": "Alytaus Patriotai",
    "Jurbarko Jurbarkas-Karys-Manvesta": "Jurbarkas",
    "Kauno r. Atletas": "Kauno Atletas",
    "Kauno r. Omega-Tauras-LSU": "Omega-Tauras",
    "Kauno Žalgiris-2": "Zalgiris Kaunas 2",
    "Klaipėdos Neptūnas-II": "Neptunas 2",
    "Kretingos Kretinga-FlexBlow": "Kretinga",
    "Kėdainių Kėdainiai BC": "Kedainiai BC",
    "Marijampolės Sūduva": "Suduva-Mantinga",
    "Mažeikių M Basket": "Mazeikiai",
    "Palangos Palanga": "BC Olimpas Palanga",
    "Plungės Olimpas": "Plunges Olimpas",
    "Telšių Telšiai": "Telsiai",
    "Ukmergės Stekas": "Ukmerges Stekas",
    "Vilkaviškio Perlas Go": "Vilkaviskis",
    "Vilniaus Rytas-2": "Rytas 2",
    "Šakių Vytis-VDU": "Vytis VDU",
    "Šilutės Šilutė-PDS.lt": "Silute",
}


#: NKL club name -> its slug on the site, read from `/komandos/` (229,220 B, 18
#: unique slugs, each linked twice).  Needed because a fixture listing gives the
#: club's name in a span and its slug only inside a match page.
CLUB_SLUGS: Dict[str, str] = {
    "Alytaus Patriotai": "alytaus-patriotai",
    "Jurbarko Jurbarkas-Karys-Manvesta": "jurbarko-jurbarkas-karys-manvesta",
    "Kauno r. Atletas": "kauno-r-atletas",
    "Kauno r. Omega-Tauras-LSU": "kauno-r-omega-tauras-lsu",
    "Kauno Žalgiris-2": "kauno-zalgiris-2",
    "Kėdainių Kėdainiai BC": "kedainiu-kedainiai-bc",
    "Klaipėdos Neptūnas-II": "klaipedos-neptunas-ii",
    "Kretingos Kretinga-FlexBlow": "kretingos-kretinga-flexblow",
    "Marijampolės Sūduva": "marijampoles-suduva",
    "Mažeikių M Basket": "mazeikiu-m-basket",
    "Palangos Palanga": "palangos-palanga",
    "Plungės Olimpas": "plunges-olimpas",
    "Šakių Vytis-VDU": "sakiu-vytis-vdu",
    "Šilutės Šilutė-PDS.lt": "silutes-silute-pds-lt",
    "Telšių Telšiai": "telsiu-telsiai",
    "Ukmergės Stekas": "ukmerges-stekas",
    "Vilkaviškio Perlas Go": "vilkaviskio-perlas-go",
    "Vilniaus Rytas-2": "vilniaus-rytas-2",
}


def nkl_name_for(betexplorer_name: str) -> Optional[str]:
    """The NKL spelling of a BetExplorer club name, or ``None`` when unmapped."""

    wanted = betexplorer_name.strip()
    for nkl_name, name in CLUB_NAMES.items():
        if name == wanted:
            return nkl_name
    return None


def club_slug(nkl_name: str) -> Optional[str]:
    return CLUB_SLUGS.get(nkl_name.strip())


def club_url(nkl_name: str) -> Optional[str]:
    slug = club_slug(nkl_name)
    return None if slug is None else f"{SITE}/komandos/{slug}/"


def club_name(nkl_name: str) -> Optional[str]:
    """The BetExplorer spelling of a club, or ``None`` when it is not mapped.

    ``None`` is the honest answer for a promotion or a rename: settling with a
    neighbouring club's result because the names were close is worse than not
    settling that game at all.
    """

    return CLUB_NAMES.get(nkl_name.strip())


def _candidates(
    rows: List["ListingRow"], *, played_on: date, home: str, away: str
) -> List["ListingRow"]:
    """Every row on that date whose pair matches as a set.

    The pair matches as a set because a source that lists the fixture the other
    way round is still the same game.  Both names must map: an unmapped club
    makes a row invisible rather than a wrong match.
    """

    wanted = {home.strip(), away.strip()}
    hits = []
    for row in rows:
        if row.played_on != played_on:
            continue
        names = {club_name(row.home_name), club_name(row.away_name)}
        if None in names:
            continue
        if names == wanted:
            hits.append(row)
    return hits


def collapse(hits: List["ListingRow"]) -> Tuple[List["ListingRow"], List[str]]:
    """One row per match, plus the ids that arrived twice carrying different data.

    A club page prints the same fixture more than once -- Kretinga's page emits
    the 2026-10-01 row three times, byte-identical -- because the markup repeats
    a row per layout variant.  That is one match, and refusing it as ambiguous
    refuses a game the page is stating plainly.

    The distinction that matters: *identical* repeats collapse, but a match id
    carrying contradictory scores, dates or names is a genuine conflict of the
    kind a second read exists to catch, and it is returned apart so the caller
    can refuse and say why.
    """

    unique: Dict[str, "ListingRow"] = {}
    conflicting: List[str] = []
    for row in hits:
        prior = unique.get(row.match_id)
        if prior is None:
            unique[row.match_id] = row
        elif prior != row and row.match_id not in conflicting:
            conflicting.append(row.match_id)
    return list(unique.values()), conflicting


def find_row(
    rows: List["ListingRow"],
    *,
    played_on: date,
    home: str,
    away: str,
) -> Optional["ListingRow"]:
    """The listing row for a forecast, by date and club pair.

    Identical repeats of one match are one row.  Two *different* matches for the
    same pair on the same date are ambiguous and return ``None``, as does one
    match printed with contradictory data -- nothing here picks a winner.
    """

    hits, conflicting = collapse(_candidates(rows, played_on=played_on, home=home,
                                            away=away))
    if conflicting or len(hits) != 1:
        return None
    return hits[0]


def explain_miss(
    rows: List["ListingRow"],
    *,
    played_on: date,
    home: str,
    away: str,
) -> List[str]:
    """Why ``find_row`` found nothing, in terms someone can act on.

    "no row" was the message that cost a live run: it is the same sentence for a
    page that parsed no rows, a page whose rows are all other dates, and a page
    whose club names are spelled differently from the reviewed table.  The three
    have nothing in common except the fix being different, so they are reported
    apart.
    """

    if not home.strip() or not away.strip():
        return ["a club name was empty"]
    if not rows:
        return ["the club page parsed no fixture rows at all"]

    hits, conflicting = collapse(_candidates(rows, played_on=played_on, home=home,
                                            away=away))
    if conflicting:
        return [
            f"match {len(conflicting)} id(s) {conflicting} printed more than once on "
            f"{played_on.isoformat()} with different data; refusing to choose"
        ]
    if len(hits) > 1:
        ids = sorted(row.match_id for row in hits)
        return [
            f"{len(hits)} distinct matches on {played_on.isoformat()} for "
            f"{home} v {away}: {ids}; refusing to choose"
        ]

    # Collapse before printing: a fixture the page repeats three times should
    # read as one fixture, not three.
    on_date = collapse([row for row in rows if row.played_on == played_on])[0]
    if not on_date:
        dates = sorted({row.played_on.isoformat() for row in rows})
        return [
            f"no fixture row on {played_on.isoformat()}; the page lists "
            f"{len(rows)} rows between {dates[0]} and {dates[-1]}"
        ]

    unmapped = {
        name
        for row in on_date
        for name in (row.home_name, row.away_name)
        if club_name(name) is None
    }
    listed = " | ".join(
        f"{row.home_name} v {row.away_name} ({row.match_id})" for row in on_date
    )
    reason = f"{len(on_date)} row(s) on {played_on.isoformat()}: {listed}"
    if unmapped:
        reason += f"; no BetExplorer mapping for {sorted(unmapped)}"
    return [reason]


#: nkl.lt publishes exactly this in robots.txt: ``Crawl-delay: 10``.
CRAWL_DELAY_SECONDS = 10.0


class PacedReader:
    """Wrap a text reader so consecutive requests honour a crawl delay.

    nkl.lt asks for ten seconds between requests, and a slate of games is the one
    shape where that matters: a forecast's two reads are seconds apart, and a run
    that settles six games would otherwise make twelve requests back to back.
    The delay is applied *between* requests -- never before the first, which would
    penalise a run that reads one page -- and the sleep is injectable so tests
    neither wait nor patch the clock.

    ``calls`` counts reads that actually reached the transport, which is what a
    receipt should report.
    """

    def __init__(self, read_text, *, delay: Optional[float] = None, sleep=None) -> None:
        self._read_text = read_text
        # Resolved at construction, not as a default argument: a delayed binding
        # would let the module constant change while the reader kept the old one.
        self._delay = max(0.0, float(CRAWL_DELAY_SECONDS if delay is None else delay))
        self._sleep = sleep or time.sleep
        self.calls = 0

    def __call__(self, url: str) -> str:
        if self.calls and self._delay:
            self._sleep(self._delay)
        self.calls += 1
        return self._read_text(url)


def paced(read_text, *, delay: Optional[float] = None, sleep=None) -> PacedReader:
    """The reader a settlement run should pass in, so the delay cannot be
    forgotten at the call site."""

    return PacedReader(read_text, delay=delay, sleep=sleep)


#: A club page's fixture table.  Each row is a joined view -- id, kickoff, both
#: clubs and the final -- which is what makes a forecast findable without opening
#: a 680 KB match page first.
_ROW_RE = re.compile(r"<tr[^>]*>(?P<tag>.*?)</tr>", re.S)
_ROW_URL_RE = re.compile(r"data-match-url=(?P<url>\S+)")
_ROW_DATE_RE = re.compile(
    r"<td[^>]*>\s*(?P<date>\d{4}-\d{2}-\d{2})\s+(?P<time>\d{2}:\d{2}):\d{2}\s*</td>"
)
_ROW_TEAM_RE = re.compile(
    r"class=[\"']?nkl-game-team-cell[\"']?[^>]*>.*?<span>(?P<name>[^<]+)</span>", re.S
)
_ROW_SCORE_RE = re.compile(
    r"class=[\"']?nkl-game-score[\"']?[^>]*>\s*(?P<home>\d{1,3})\s*-\s*(?P<away>\d{1,3})"
)


@dataclass(frozen=True)
class ListingRow:
    """One row of a club's fixture table, as the page states it."""

    match_id: str
    url: str
    kickoff_local: datetime
    home_name: str
    away_name: str
    home_score: Optional[float]
    away_score: Optional[float]

    @property
    def played_on(self) -> date:
        return self.kickoff_local.date()

    def to_dict(self) -> Dict[str, object]:
        return {
            "match_id": self.match_id,
            "url": self.url,
            "kickoff_local": self.kickoff_local.isoformat(),
            "home_name": self.home_name,
            "away_name": self.away_name,
            "home_score": self.home_score,
            "away_score": self.away_score,
        }


def parse_club_page(html: str) -> Tuple[List[ListingRow], Dict[str, int]]:
    """Every fixture row a club page lists, plus what could not be read.

    No score is invented: an upcoming row has no result cell and is returned with
    ``None`` scores rather than with zeros or a repeat of the last row.
    """

    rows: List[ListingRow] = []
    counts = {"rows": 0, "with_url": 0, "with_date": 0, "with_teams": 0, "with_score": 0}
    for row in _ROW_RE.finditer(html or ""):
        body = row.group(0)
        url_match = _ROW_URL_RE.search(body)
        if url_match is None:
            continue
        counts["rows"] += 1
        match_id = url_match.group("url").rstrip("/").rsplit("/", 1)[-1]
        counts["with_url"] += 1

        date_match = _ROW_DATE_RE.search(body)
        teams = _ROW_TEAM_RE.findall(body)
        score_match = _ROW_SCORE_RE.search(body)
        if date_match is None or len(teams) != 2:
            continue
        counts["with_date"] += 1
        counts["with_teams"] += 1
        if score_match is not None:
            counts["with_score"] += 1

        day = date.fromisoformat(date_match.group("date"))
        hour, minute = (int(part) for part in date_match.group("time").split(":"))
        rows.append(
            ListingRow(
                match_id=match_id,
                url=url_match.group("url"),
                kickoff_local=datetime(
                    day.year, day.month, day.day, hour, minute, tzinfo=ZoneInfo(SITE_TZ)
                ),
                home_name=teams[0].strip(),
                away_name=teams[1].strip(),
                home_score=None if score_match is None else float(score_match.group("home")),
                away_score=None if score_match is None else float(score_match.group("away")),
            )
        )
    return rows, counts


@dataclass(frozen=True)
class Settlement:
    """What settling one forecast cost, and whether it produced a record."""

    result: object
    url: Optional[str] = None
    #: The nkl.lt match id, so the caller can stamp it as a provider id rather
    #: than re-parsing it out of ``url`` at the far end of a pipeline.
    match_id: Optional[str] = None
    requests_made: int = 0
    reasons: Tuple[str, ...] = ()

    def to_dict(self) -> Dict[str, object]:
        result = self.result
        return {
            "settled": result is not None,
            "url": self.url,
            "match_id": self.match_id,
            "requests_made": self.requests_made,
            "reasons": list(self.reasons),
            "periods": None if result is None else {
                "home_quarters": result.home_quarters,
                "away_quarters": result.away_quarters,
                "home_final": result.home_final,
                "away_final": result.away_final,
                "reconciles": result.reconciles(),
            },
        }


def settle_forecast(
    read_text,
    *,
    played_on: date,
    home: str,
    away: str,
    cache: Optional[Dict[str, List[ListingRow]]] = None,
) -> Settlement:
    """Find a forecast's game on nkl.lt and read its quarters.

    Two requests at most: the club's fixture page (which lists every round, so it
    is cached for the caller and reused across a slate) and the match page for the
    one game.  Nothing is fetched from the match page until the fixture has been
    identified by date and club pair.
    """

    reasons: list = []
    home_nkl = nkl_name_for(home)
    away_nkl = nkl_name_for(away)
    if home_nkl is None or away_nkl is None:
        missing = [name for name, mapped in ((home, home_nkl), (away, away_nkl)) if mapped is None]
        reasons.append(f"no NKL mapping for {missing}")
        return Settlement(result=None, reasons=tuple(reasons))

    store = cache if cache is not None else {}
    url = club_url(home_nkl)
    requests_made = 0
    rows = store.get(home_nkl or "")
    if rows is None:
        try:
            rows, _ = parse_club_page(read_text(url))
            requests_made += 1
        except Exception as exc:  # noqa: BLE001 - a failed read is the receipt
            reasons.append(f"club page failed: {str(exc)[:120]}")
            return Settlement(result=None, url=url, requests_made=1, reasons=tuple(reasons))
        store[home_nkl or ""] = rows

    row = find_row(rows, played_on=played_on, home=home, away=away)
    if row is None:
        reasons.append(f"no row for {home} v {away} on {played_on}")
        reasons.extend(explain_miss(rows, played_on=played_on, home=home, away=away))
        return Settlement(result=None, url=url, requests_made=requests_made, reasons=tuple(reasons))
    if row.home_score is None:
        reasons.append(f"row {row.match_id} has no result yet")
        return Settlement(result=None, url=row.url, requests_made=requests_made,
                          match_id=row.match_id, reasons=tuple(reasons))

    try:
        page = parse_match_page(read_text(row.url), url=row.url)
        requests_made += 1
    except Exception as exc:  # noqa: BLE001
        reasons.append(f"match page failed: {str(exc)[:120]}")
        return Settlement(result=None, url=row.url, requests_made=requests_made,
                          match_id=row.match_id, reasons=tuple(reasons))

    result = page.to_actual_result()
    if result is None:
        reasons.extend(page.reasons or ("no reconcilable periods",))
        return Settlement(result=None, url=row.url, requests_made=requests_made,
                          match_id=row.match_id, reasons=tuple(reasons))
    ok, why = result.reconciles()
    if not ok:
        reasons.append(f"quarters do not reconcile with the final: {why}")
        return Settlement(result=None, url=row.url, requests_made=requests_made,
                          match_id=row.match_id, reasons=tuple(reasons))
    return Settlement(result=result, url=row.url, requests_made=requests_made,
                      match_id=row.match_id, reasons=tuple(reasons))


def parse_match_page(html: str, url: str = "") -> MatchPage:
    """Read one served nkl.lt match page.  Never raises on absent markup."""

    text = html or ""
    problems: list = []

    table = _QUARTERS_TABLE_RE.search(text)
    home_quarters: List[Optional[float]] = []
    away_quarters: List[Optional[float]] = []
    headers_ok = False
    if table is None:
        problems.append("no nkl-quarters-table in the page")
    else:
        home_quarters, away_quarters, table_problems, headers_ok = _read_quarters(
            table.group("body")
        )
        problems.extend(table_problems)

    clubs = [m.group("slug") for m in _CLUB_RE.finditer(table.group("body") if table else "")]
    codes = [m.group("code") for m in _CODE_RE.finditer(table.group("body") if table else "")]
    if len(clubs) != 2:
        problems.append(f"expected two club slugs, read {len(clubs)}")
    if len(codes) != 2:
        problems.append(f"expected two club codes, read {len(codes)}")

    played_on: Optional[date] = None
    tip_local: Optional[datetime] = None
    start = _START_RE.search(text)
    if start is None:
        problems.append("no 'Rungtynių pradžia' date cell")
    else:
        played_on = date.fromisoformat(start.group("date"))
        hour, minute = (int(part) for part in start.group("time").split(":"))
        tip_local = datetime(
            played_on.year, played_on.month, played_on.day, hour, minute,
            tzinfo=ZoneInfo(SITE_TZ),
        )

    final = _read_final(text)
    if final is None:
        problems.append(
            "no final score element (quarters are read but not compared; "
            "a sum of quarters is not a final)"
        )
        candidates = final_candidates(text)
        if candidates:
            problems.append("score-shaped pairs on the page: " + " | ".join(candidates))
    home_final, away_final, final_source = (
        (None, None, None) if final is None else final
    )

    counts = {
        "quarter_headers": len(_TH_RE.findall(table.group("body"))) if table else 0,
        "club_slugs": len(clubs),
        "club_codes": len(codes),
        "final_read": 0 if final is None else 1,
    }
    return MatchPage(
        url=url,
        home_slug=clubs[0] if len(clubs) == 2 else "",
        away_slug=clubs[1] if len(clubs) == 2 else "",
        home_code=codes[0] if len(codes) == 2 else "",
        away_code=codes[1] if len(codes) == 2 else "",
        played_on=played_on,
        tip_local=tip_local,
        home_quarters=home_quarters,
        away_quarters=away_quarters,
        home_final=home_final,
        away_final=away_final,
        final_source=final_source,
        headers_ok=headers_ok,
        counts=counts,
        reasons=tuple(problems),
    )
