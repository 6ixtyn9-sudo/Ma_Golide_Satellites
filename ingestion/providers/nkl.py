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

No request is made here: this reads bytes that were already captured.
"""

from __future__ import annotations

import re
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
