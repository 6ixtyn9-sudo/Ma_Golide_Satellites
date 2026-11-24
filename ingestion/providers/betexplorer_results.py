"""The aggregator's own results pages: a reader that counts before it believes.

Why this exists
---------------
``GREECE_BASKET_LEAGUE``, ``CHINA_WCBA_W``, ``LITHUANIA_NKL`` and
``AUSTRALIA_WNBL_W`` each publish fixtures and prices but no result this project
can read: Esake's reader is close but not proven, and the three others either
refuse every score-bearing namespace (NKL) or answer with an authenticated
gateway (WNBL).  BetExplorer serves a results page for the same league slug the
fixtures page already uses -- path-only, no query string, robots-allowed -- and
a 2026-11-24 capture found 34, 25 and 140 distinct matches on the three pages.

So the finals exist.  What this module does **not** claim is that they have been
read: every rule below is derived from the *fixtures* row markup this project
already parses (``in-match`` link, two ``<span>`` names, a
``table-main__datetime`` cell), because that is the only markup on this host that
has been captured in full.  The results page is therefore read with the same
rules plus a score cell, and every row that does not satisfy all of them is
counted rather than guessed at:

* no match code, not exactly two club names, or not exactly one score-shaped pair
  → the row is skipped and the reason is counted;
* a clock is not a score: ``20:00`` in a kickoff cell is a time, and a pair is
  only read as a final when **both** sides are at least :data:`MIN_FINAL_POINTS`,
  which no clock can be.  Forfeits and other freak finals fall outside that rule
  and are deliberately excluded: a missed final is recoverable, an invented one
  is not;
* no date → no ``played_on``, and the row is not settled by this module anyway.

Nothing here writes and nothing here decides that a lane may settle.  It returns
rows and counts; the registry stays ``source_unavailable`` until a reader has been
replayed against captured bytes and a final has been matched to a forecast.

Provenance, as the operator decided it
--------------------------------------
These are **aggregator results, matched by provider code** -- not federation
finals.  Every row carries :data:`SOURCE` so no downstream record can quietly
promote an aggregator's number to a federation's.  The site's terms prohibit
scraping even where robots.txt allows the path; the operator's decision on
2026-11-24 was to keep reading path-only pages for this paper-only record, with
attribution on every row and no per-game hammering.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from typing import Dict, List, Optional, Tuple

PROVIDER = "betexplorer_results"

#: Attribution and provenance in one string: this is never a federation final.
SOURCE = "BetExplorer — aggregator result, matched by provider code"

#: A basketball final's smaller side.  A clock's hour is at most 23, so no time
#: can satisfy this; a forfeit (20:0) is excluded on purpose.
MIN_FINAL_POINTS = 25

#: Rows are the unit here too, and for the same reason: a page is not a row.
_ROW_RE = re.compile(r"<tr>(?:(?!</tr>).)*?</tr>", re.S)
#: The match link, exactly as the fixtures parser requires it.  Requiring the
#: class is what keeps a duplicate bare link from reading as a second match.
_LINK_RE = re.compile(
    r'<a href="/basketball/[^"]+/(?P<slug>[a-z0-9-]+)/(?P<match_id>[A-Za-z0-9]{8})/"'
    r' class="in-match">'
)
_SPANS_RE = re.compile(r"<span>(?:<strong>)?(?P<name>[^<]+?)(?:</strong>)?</span>")
_DATETIME_RE = re.compile(r'class="table-main__datetime"[^>]*>(?P<label>[^<]*)</td>')
#: The score's own cell, when the page has one.  The fallback scans the row.
_RESULT_CELL_RE = re.compile(r'class="table-main__result"[^>]*>(?P<label>.*?)</td>', re.S)
_CLOCK_PAIR_RE = re.compile(r"(?<!\d)(\d{1,2}):(\d{2})(?!\d)")
_TAG_RE = re.compile(r"<[^>]+>")
_ABSOLUTE_RE = re.compile(
    r"^(?P<day>\d{1,2})\.(?P<month>\d{1,2})\.(?P<year>\d{4})?\s+(?P<hour>\d{1,2}):(?P<minute>\d{2})$"
)


def _text(fragment: str) -> str:
    plain = _TAG_RE.sub(" ", fragment or "")
    for entity, char in (("&nbsp;", " "), ("&#160;", " "), ("&#xa0;", " "), ("&amp;", "&")):
        plain = plain.replace(entity, char)
    return " ".join(plain.split())


def _score_pair(text: str) -> Optional[Tuple[int, int]]:
    """The one pair in ``text`` where both sides are too big to be a clock."""

    candidates: List[Tuple[int, int]] = []
    for left, right in _CLOCK_PAIR_RE.findall(text):
        a, b = int(left), int(right)
        if a >= MIN_FINAL_POINTS and b >= MIN_FINAL_POINTS:
            candidates.append((a, b))
    return candidates[0] if len(candidates) == 1 else None


def _played_on(label: str) -> Optional[date]:
    match = _ABSOLUTE_RE.match(label.strip())
    if not match:
        return None
    if not match.group("year"):
        return None
    try:
        return date(int(match.group("year")), int(match.group("month")), int(match.group("day")))
    except ValueError:
        return None


@dataclass(frozen=True)
class ResultRow:
    match_id: str
    slug: str
    home: str
    away: str
    home_score: int
    away_score: int
    played_on: Optional[date]
    source: str = SOURCE


@dataclass(frozen=True)
class ResultPage:
    rows: Tuple[ResultRow, ...]
    skipped: Tuple[Dict[str, str], ...]
    counts: Dict[str, int]

    def to_dict(self) -> Dict[str, object]:
        return {
            "counts": dict(self.counts),
            "rows": [
                {
                    "match_id": row.match_id,
                    "slug": row.slug,
                    "home": row.home,
                    "away": row.away,
                    "home_score": row.home_score,
                    "away_score": row.away_score,
                    "played_on": row.played_on.isoformat() if row.played_on else None,
                    "source": row.source,
                }
                for row in self.rows
            ],
            "skipped": [dict(item) for item in self.skipped],
        }


def parse_results(html: str) -> ResultPage:
    """Read a saved results page into rows, counting every row it could not read.

    An empty ``rows`` from a page with no ``<tr>`` at all is a failure to read,
    not a page without finals, so the counts carry ``rows`` beside ``read`` and a
    caller can tell the two apart.
    """

    if not isinstance(html, str) or not html.strip():
        raise ValueError(f"{PROVIDER}: empty body, so the page was not read")

    counts = {
        "rows": 0,
        "read": 0,
        "without_code": 0,
        "names_not_two": 0,
        "score_not_one": 0,
        "undated": 0,
        "duplicate_match_id": 0,
    }
    rows: List[ResultRow] = []
    skipped: List[Dict[str, str]] = []
    seen: set[str] = set()

    for row in _ROW_RE.findall(html):
        counts["rows"] += 1
        link = _LINK_RE.search(row)
        if link is None:
            counts["without_code"] += 1
            continue
        match_id = link.group("match_id")
        if match_id in seen:
            counts["duplicate_match_id"] += 1
            continue

        names = [name.strip() for name in _SPANS_RE.findall(row) if name.strip()]
        if len(names) != 2:
            counts["names_not_two"] += 1
            skipped.append({"match_id": match_id, "reason": f"club names not two ({len(names)})"})
            continue

        cell = _RESULT_CELL_RE.search(row)
        score = _score_pair(_text(cell.group("label"))) if cell else None
        if score is None:
            score = _score_pair(_text(_DATETIME_RE.sub(" ", row)))
        if score is None:
            counts["score_not_one"] += 1
            skipped.append({"match_id": match_id, "reason": "no single score-shaped pair"})
            continue

        stamp = _DATETIME_RE.search(row)
        played_on = _played_on(stamp.group("label")) if stamp else None
        if played_on is None:
            counts["undated"] += 1

        seen.add(match_id)
        counts["read"] += 1
        rows.append(
            ResultRow(
                match_id=match_id,
                slug=link.group("slug"),
                home=names[0],
                away=names[1],
                home_score=score[0],
                away_score=score[1],
                played_on=played_on,
            )
        )

    return ResultPage(rows=tuple(rows), skipped=tuple(skipped), counts=counts)


__all__ = [
    "MIN_FINAL_POINTS",
    "PROVIDER",
    "ResultPage",
    "ResultRow",
    "SOURCE",
    "parse_results",
]
