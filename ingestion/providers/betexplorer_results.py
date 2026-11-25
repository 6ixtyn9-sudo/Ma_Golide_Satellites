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
read end to end: the row shape was captured on 2026-11-24 and the rules below are
written against it --

```html
<td class="h-text-left"><a data-test="11015187"
    href="/basketball/australia/wnbl-women/perth-lynx-canberra-capitals/EaZEfXoI/"
    class="in-match"><span><strong>Perth Lynx W</strong></span> -
    <span>Canberra Capitals W</span></a></td>
<td class="h-text-center"><a href=".../EaZEfXoI/">113:90</a></td>
```

The first replay read **zero** rows for all three leagues: the fixtures parser
requires ``href`` immediately before ``class="in-match"``, and this page writes
``data-test`` in between, so attribute order -- not content -- was the whole
failure.  Every row that does not satisfy every rule is counted, never guessed:

* no match code, not exactly two club names, or not exactly one score-shaped pair
  → the row is skipped and the reason is counted;
* a clock is not a score: ``20:00`` in a kickoff cell is a time, and a pair is
  only read as a final when **both** sides are at least :data:`MIN_FINAL_POINTS`,
  which no clock can be.  Forfeits and other freak finals fall outside that rule
  and are deliberately excluded: a missed final is recoverable, an invented one
  is not;
* no date → no ``played_on``, and the row is not settled by this module anyway.
  The results table groups rows under a round header rather than dating each one,
  so in practice every row from this page is undated: the date has to come from
  the forecast the result is matched against, not from the result page.

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
#: Any anchor, then its href read out of the tag.  The fixtures parser requires
#: ``href`` immediately before ``class="in-match"`` and the results page writes
#: ``<a data-test="11015187" href="..." class="in-match">`` -- attribute order is
#: not a contract, so the tag is matched whole and its href read separately.
#: Reading the tag's *content* is what separates the two clubs from the score link.
_ANCHOR_RE = re.compile(r"<a\b[^>]*>", re.IGNORECASE)
_HREF_RE = re.compile(r'href="(?P<href>/basketball/[^"]+/(?P<code>[A-Za-z0-9]{8})/)"')
_IN_MATCH = 'class="in-match"'

_SPANS_RE = re.compile(r"<span>(?:<strong>)?(?P<name>[^<]+?)(?:</strong>)?</span>")
_DATETIME_RE = re.compile(r'class="table-main__datetime"[^>]*>(?P<label>[^<]*)</td>')
#: The score's own cell, when the page has one.  The fallback scans the row.
_RESULT_CELL_RE = re.compile(r'class="table-main__result"[^>]*>(?P<label>.*?)</td>', re.S)
#: Two or three digits either side: a basketball final reaches three (113:90),
#: a clock never does.  Bounding it at two would silently skip every century.
_PAIR_RE = re.compile(r"(?<!\d)(\d{1,3}):(\d{1,3})(?!\d)")
_TAG_RE = re.compile(r"<[^>]+>")
_ABSOLUTE_RE = re.compile(
    r"^(?P<day>\d{1,2})\.(?P<month>\d{1,2})\.(?P<year>\d{4})?\s+(?P<hour>\d{1,2}):(?P<minute>\d{2})$"
)


def _text(fragment: str) -> str:
    plain = _TAG_RE.sub(" ", fragment or "")
    for entity, char in (("&nbsp;", " "), ("&#160;", " "), ("&#xa0;", " "), ("&amp;", "&")):
        plain = plain.replace(entity, char)
    return " ".join(plain.split())


def _match_anchor(row: str) -> Optional[Tuple[re.Match, str]]:
    """The ``in-match`` anchor and the markup inside it.

    The anchor holds both clubs; the total sits in the next cell.  So names come
    from the anchor's content and the score comes from the row.
    """

    for tag in _ANCHOR_RE.finditer(row):
        if _IN_MATCH not in tag.group(0):
            continue
        href = _HREF_RE.search(tag.group(0))
        if href is None:
            continue
        close = row.find("</a>", tag.end())
        inner = inner_start(row, tag.end(), close)
        return href, inner
    return None


def inner_start(row: str, start: int, close: int) -> str:
    """The anchor's content, or the rest of the row when it is never closed."""

    return row[start : close if close > start else len(row)]


def _score_pair(text: str) -> Optional[Tuple[int, int]]:
    """The one pair in ``text`` where both sides are too big to be a clock."""

    candidates: List[Tuple[int, int]] = []
    for left, right in _PAIR_RE.findall(text):
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
        found = _match_anchor(row)
        if found is None:
            counts["without_code"] += 1
            continue
        link, inner = found
        match_id = link.group("code")
        if match_id in seen:
            counts["duplicate_match_id"] += 1
            continue

        names = [name.strip() for name in _SPANS_RE.findall(inner) if name.strip()]
        if len(names) != 2:
            counts["names_not_two"] += 1
            skipped.append({"match_id": match_id, "reason": f"club names not two ({len(names)})"})
            continue

        cell = _RESULT_CELL_RE.search(row)
        score = _score_pair(_text(cell.group("label"))) if cell else None
        if score is None:
            # The total sits in a plain unclassed cell as a second link to the
            # same match, so it is found by its shape rather than by a class --
            # and it is inside that link, which is why the row is read whole.
            score = _score_pair(_text(row))
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
                slug=link.group("href").rstrip("/").rsplit("/", 2)[-2],
                home=names[0],
                away=names[1],
                home_score=score[0],
                away_score=score[1],
                played_on=played_on,
            )
        )

    return ResultPage(rows=tuple(rows), skipped=tuple(skipped), counts=counts)


# ---------------------------------------------------------------------------
# The match page: a final and the split the page states beside it
# ---------------------------------------------------------------------------
#
# The results index gives totals.  The match page for the same code states the
# split directly under the final, as a parenthesised colon list on 2026-11-25:
#
#     24.11.2026 - 12:00
#     93:107
#     (15:25, 30:28, 17:27, 31:27)
#
# 15+30+17+31 = 93 and 25+28+27+27 = 107, so this is a record
# ``ActualResult.reconciles()`` accepts -- no quarter labels, no table cells, no
# script, and no name join: the code in the URL is the code the fixtures table
# already assigns.  That matters because the route was once recorded as closed
# "by measurement": the probe looked for ``<td>`` cells and ``Q1..Q4`` labels and
# reported ``td_cells 0, period_markers 0`` on a page that carries the split in
# neither shape.  Those counters measured the probe's vocabulary, not the page.
_SPLIT_RE = re.compile(
    r"\(\s*(?P<body>\d{1,3}\s*:\s*\d{1,3}(?:\s*,\s*\d{1,3}\s*:\s*\d{1,3}){2,})\s*\)"
)
_REGULATION_PERIODS = 4
#: One entry of the split: home on the left, away on the right.
_QUAD_ENTRY_RE = re.compile(r"(\d{1,3})\s*:\s*(\d{1,3})")


@dataclass(frozen=True)
class MatchPeriods:
    """One finished match, with the split the page states next to its total."""

    home_final: int
    away_final: int
    regulation: Tuple[Tuple[int, int], ...]
    overtime: Tuple[Tuple[int, int], ...] = ()

    @property
    def home_quarters(self) -> Tuple[float, ...]:
        return tuple(float(home) for home, _ in self.regulation)

    @property
    def away_quarters(self) -> Tuple[float, ...]:
        return tuple(float(away) for _, away in self.regulation)

    @property
    def home_overtime(self) -> Optional[float]:
        return float(sum(home for home, _ in self.overtime)) if self.overtime else None

    @property
    def away_overtime(self) -> Optional[float]:
        return float(sum(away for _, away in self.overtime)) if self.overtime else None

    @property
    def matches(self) -> bool:
        """Whether the split adds up to the total the same page states."""

        home = sum(home for home, _ in self.regulation) + sum(h for h, _ in self.overtime)
        away = sum(away for _, away in self.regulation) + sum(a for _, a in self.overtime)
        return (home, away) == (self.home_final, self.away_final)


_CODE_RE = re.compile(r"[A-Za-z0-9]{8}")
_SLUG_RE = re.compile(r"[a-z0-9-]+")


def match_page_url(league: str, slug: str, code: str) -> str:
    """The game's own page, rebuilt from the league's verified fixtures path.

    The fixtures row already carries both halves of this URL: the slug it links
    to and the 8-character code in the same href.  Rebuilding it here means
    settlement needs no name join and no second index read -- the code the
    fixture was stored under is the code whose page is fetched.
    """

    from .betexplorer import FIXTURES_URLS

    base = FIXTURES_URLS.get(str(league).upper())
    if base is None:
        raise ValueError(f"{league!r} has no verified fixtures path")
    if not base.endswith("fixtures/"):
        raise ValueError(f"unexpected fixtures path for {league}: {base}")
    if not _SLUG_RE.fullmatch(slug or "") or not _CODE_RE.fullmatch(code or ""):
        raise ValueError(f"refusing a malformed match page identity: {slug!r}/{code!r}")
    url = base[: -len("fixtures/")] + f"{slug}/{code}/"
    if "?" in url:  # pragma: no cover - defensive: query shapes are disallowed
        raise ValueError(f"refusing a query-shaped URL: {url}")
    return url


def parse_match_page(html: str) -> Optional[MatchPeriods]:
    """The final and split on one match page, or ``None``.

    Refuses -- rather than guesses -- when the page states no split, states more
    than one (a page is one match, so two splits mean the read is wrong), states
    fewer than four periods, or states a split that does not add up to the
    total beside it.  A wrong settlement is unrecoverable; a missing one is not.
    """

    plain = _text(html or "")
    splits = list(_SPLIT_RE.finditer(plain))
    if len(splits) != 1:
        return None
    split = splits[0]
    # The final sits immediately above the split: the last clock-shaped pair
    # (a time's hour is at most 23, so it cannot pass MIN_FINAL_POINTS).
    finals = [
        (int(left), int(right))
        for left, right in _PAIR_RE.findall(plain[: split.start()])
        if int(left) >= MIN_FINAL_POINTS and int(right) >= MIN_FINAL_POINTS
    ]
    if not finals:
        return None
    home_final, away_final = finals[-1]

    periods = [
        (int(home), int(away)) for home, away in _QUAD_ENTRY_RE.findall(split.group("body"))
    ]
    if len(periods) < _REGULATION_PERIODS:
        return None
    result = MatchPeriods(
        home_final=home_final,
        away_final=away_final,
        regulation=tuple(periods[:_REGULATION_PERIODS]),
        overtime=tuple(periods[_REGULATION_PERIODS:]),
    )
    return result if result.matches else None


__all__ = [
    "MIN_FINAL_POINTS",
    "MatchPeriods",
    "match_page_url",
    "parse_match_page",
    "PROVIDER",
    "ResultPage",
    "ResultRow",
    "SOURCE",
    "parse_results",
]
