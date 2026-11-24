"""The Greek federation's own results pages: a reader, not yet a settlement lane.

Why this exists
---------------
``GREECE_BASKET_LEAGUE`` acquires fixtures and research-only prices, but its
``results_source`` is ``source_unavailable``, so nothing it forecasts can ever be
graded. The federation (Ε.Σ.Α.Κ.Ε.) publishes finals itself, on pages the
2026-11-24 probe validated to 10/10 coverage against the lane's own fixtures.

What is verified here, and what is not
--------------------------------------
Verified, from the probe's verbatim snippets and its 10/10 join:

* game identity is the hex in ``EsakegameView?idgame=<HEX>`` (refetched: same id,
  same score, different byte count);
* a finished game's score sits in a ``esake-program-game-final-score`` block as
  ``<span>100&nbsp;-&nbsp;64</span>``;
* the row carries a day, a Greek month label and a ``HH:MM`` tip.

Captured 2026-11-24: the container is ``<div class="esake-program-game">``, its
final-score block puts each club's span inside a ``col-lg-5`` column and the
total in the middle one, and the toggle handler ``.closest('.esake-program-game')``
in a script attribute repeats the container name twice per row — which is the
string that made a bare-token split cut every row in half before it was found.

**Not** verified: how an unplayed row is marked, and whether the totals arrive as
one ``100 - 64`` span (as captured) or two spans. This reader is therefore built
to fail closed and to *count* what it cannot resolve rather than guessing:

* no row container recognised → :class:`IdentityError`, never an empty page;
* a row without a ``final-score`` block is simply not a final (it may be
  unplayed, or the markup may have moved) and is counted as such;
* a row whose club spans are not exactly two is not a final — ambiguity is a
  count, not a coin flip;
* an unresolvable month label or missing season context leaves the date unknown,
  and an unknown date cannot settle anything.

A live read on 2026-11-24 settled the question this module was built to answer:
one request, 133,247 B, sha256 ``8a2ebdb9…``, and counts of containers 7, rows 7,
``with_final`` 7, ``unresolved_date`` 0 — seven games, seven identities, seven
totals, seven dates.

Nothing in this module writes, and nothing in it decides that a lane may settle.
The registry stays on ``source_unavailable`` for a reason this reader does not
answer: ``ActualResult.reconciles()`` needs Q1–Q4 (or H1/H2) that sum to the
final, and a final alone is not a settlement record. The page carries quarters in
its hidden per-quarter row, but that row is not parsed here, and a quarter parser
written from collapsed text would be a guess wearing a date.

Terms, as the operator decided them
-----------------------------------
The site's «Όροι Χρήσης» licenses copying for personal, non-commercial use with
source credit and contains no automated-access prohibition. On 2026-11-24 the
operator accepted that scope for this system's paper record, so the reader keeps
the cheap half of the bargain: every row carries :data:`SOURCE`, every receipt
carries the URL and the sha256 of the bytes it came from, and no request is made
per game.

Politeness
----------
Four pages, spaced, is a whole run (:data:`REQUEST_BUDGET`,
:data:`MIN_INTERVAL_SECONDS`). The federation's pages answered in 4–18 seconds
during the probe; a reader that bursts would be asking a slow host to carry our
impatience.
"""

from __future__ import annotations

import hashlib
import re
import unicodedata
from dataclasses import dataclass
from datetime import date
from typing import Dict, List, Optional, Tuple

from ..http import HttpClient
from . import IdentityError

PROVIDER = "esake"

#: Attribution travels with the data: the terms that permit copying ask for it.
SOURCE = "Ε.Σ.Α.Κ.Ε. esake.gr — Stoiximan GBL results"

SITE = "https://www.esake.gr"
RESULTS_PATH = "/el/action/EsakeResults"
GAME_PATH = "/el/action/EsakegameView"

#: Championship id for Stoiximan GBL 2026-27, read from the page's own select.
CHAMPIONSHIP_GBL = "184645B9"
#: Phase «Α Φάση – Κανονική Περίοδος» (regular season).
PHASE_REGULAR = "00000001"

#: One series page plus one game page is plenty to learn whether this reads.
REQUEST_BUDGET = 4
#: The host is slow (4–18s per page in the probe). Space requests like it matters.
MIN_INTERVAL_SECONDS = 3.0

REQUEST_HEADERS = {
    "User-Agent": "MaGolideProbe/2026-11-24 (+paper record; 4 requests per run)",
    "Accept-Encoding": "identity",
}


def results_url(
    series: Optional[int] = None,
    *,
    championship: str = CHAMPIONSHIP_GBL,
    season: str = PHASE_REGULAR,
) -> str:
    """The matchday page: one page per round, no pagination, no per-game call."""

    url = f"{SITE}{RESULTS_PATH}?idchampionship={championship}&idseason={season}&idteam="
    if series is not None:
        url += f"&series={int(series)}"
    return url


def game_url(idgame: str) -> str:
    """A single game's page, keyed by the identity the results page publishes."""

    return f"{SITE}{GAME_PATH}?idgame={str(idgame).upper()}&mode=2"


def default_fetch():
    """A text client that says who it is and takes its time.

    The host answered in 4–18 seconds during the probe and publishes no crawl
    rules; the polite reading of that is a small number of spaced requests, not
    a burst with retries.
    """

    client = HttpClient(
        timeout=45.0,
        retries=1,
        user_agent=REQUEST_HEADERS["User-Agent"],
        default_headers={"Accept-Encoding": REQUEST_HEADERS["Accept-Encoding"]},
    )
    return lambda url: client.get_text(url)


def checksum(body: str) -> str:
    """sha256 of the bytes as served, so a receipt names what it actually read."""

    return hashlib.sha256(body.encode("utf-8", "surrogatepass")).hexdigest()


# --- parsing ---------------------------------------------------------------

#: A row container, and only a container: the class attribute is required.
#:
#: The served page carries the bare strings ``esake-program-game`` in three
#: places that are not row starts -- ``esake-program-game-row-1`` and
#: ``esake-program-game-final-score`` inside the row, and, decisively, the
#: toggle handler ``.closest('.esake-program-game')`` in a script attribute.
#: Splitting on the bare token therefore cut every row at its own toggle, which
#: is why a 139,358 B page of seven games counted fourteen containers, seven of
#: them with no identity and none of them with a score.  Requiring ``class="``
#: excludes the script literal; the lookahead excludes the row's own parts.
_ROW_RE = re.compile(r'class="[^"]*esake-program-game(?![\w-])[^"]*"', re.IGNORECASE)
_LINK_RE = re.compile(r"EsakegameView\?idgame=([0-9A-Fa-f]{8})", re.IGNORECASE)
_SPAN_TAG_RE = re.compile(r"<span\b[^>]*>|</span\s*>", re.IGNORECASE)
_TAG_RE = re.compile(r"<[^>]+>")
_GAP = r"(?:&nbsp;|&#160;|&#xa0;|\u00a0|\s)*"
_SCORE_RE = re.compile(rf"(?<!\d)(\d{{1,3}}){_GAP}[-–]{_GAP}(\d{{1,3}})(?!\d)")
_TIME_RE = re.compile(r"\b([01]?\d|2[0-3]):([0-5]\d)\b")
_DAY_MONTH_RE = re.compile(r"\b(\d{1,2})\s+([^\W\d_]{3,})", re.UNICODE)
#: The same pair with no whitespace at all, for cells that break inside a word.
_UNSPACED_DATE_RE = re.compile(r"(?<!\d)(\d{1,2})([^\W\d_]{3,})", re.UNICODE)

#: Month labels, matched on 4 letters first because ΙΟΥΝ/ΙΟΥΛ share a prefix.
#: Only «Νοε» is byte-verified; the rest are the conventional Greek abbreviations
#: and are treated as unknown rather than guessed when they do not match.
_MONTHS_4 = {"ΙΟΥΝ": 6, "ΙΟΥΛ": 7}
_MONTHS_3 = {
    "ΙΑΝ": 1,
    "ΦΕΒ": 2,
    "ΜΑΡ": 3,
    "ΑΠΡ": 4,
    "ΜΑΙ": 5,
    "ΜΑΪ": 5,
    "ΜΑΗ": 5,
    "ΑΥΓ": 8,
    "ΣΕΠ": 9,
    "ΟΚΤ": 10,
    "ΝΟΕ": 11,
    "ΔΕΚ": 12,
}

_FINISHED_CLASS = "esake-program-game-final-score"
#: The row's own clock cell, captured 2026-11-24:
#: ``<div class="esake-program-game-info one-col"><img ...clock.svg...>Σαβ 21 Νοε - 15:00</div>``.
#: Reading the date from *here* rather than from the whole row is what keeps an
#: unrelated day in some other part of the page from being read as this game's.
_DATE_CELL_CLASS = "esake-program-game-info"
#: The per-quarter row, hidden behind the toggle.  It repeats both club names and
#: carries four numbers per club, so reading club spans from the whole container
#: would find four names and call every finished game ambiguous.
_HIDDEN_CLASS = "esake-program-game-hidden-row"


def _fold(text: str) -> str:
    """Upper-case and strip Greek accents, so «Νοε» == «ΝΟΕ» == «Νοέ»."""

    decomposed = unicodedata.normalize("NFD", str(text or ""))
    stripped = "".join(ch for ch in decomposed if not unicodedata.combining(ch))
    return stripped.upper().strip()


def month_from_label(label: str) -> Optional[int]:
    """Greek month label to 1–12, or ``None`` when it is not one we can read."""

    folded = _fold(label)
    if not folded:
        return None
    if len(folded) >= 4 and folded[:4] in _MONTHS_4:
        return _MONTHS_4[folded[:4]]
    return _MONTHS_3.get(folded[:3])


def _text(fragment: str) -> str:
    """Visible text of an HTML fragment: tags dropped, entities resolved."""

    plain = _TAG_RE.sub(" ", fragment)
    for entity, char in (
        ("&nbsp;", " "),
        ("&#160;", " "),
        ("&#xa0;", " "),
        ("&amp;", "&"),
        ("&quot;", '"'),
        ("&#39;", "'"),
    ):
        plain = plain.replace(entity, char)
    return " ".join(plain.split())


def _looks_like_when(text: str) -> bool:
    """True for a day/month/tip cell, False for a club name.

    The digit requirement is what keeps this from eating clubs: «ΜΑΡΟΥΣΙ» starts
    with a month abbreviation, «21 Νοε» is a date. Nothing here decides anything
    on its own -- dropping a real club only makes the row ambiguous, and an
    ambiguous row is counted rather than settled.
    """

    stripped = str(text or "").strip()
    if _TIME_RE.fullmatch(stripped):
        return True
    if stripped.isdigit():
        return True
    if not any(ch.isdigit() for ch in stripped):
        return False
    return any(month_from_label(token) is not None for token in stripped.split())


def _spans(fragment: str) -> List[str]:
    """The contents of every outermost span, in document order.

    Depth, not a non-greedy match: a club cell can nest a span inside a span, and
    ``<span>(.*?)</span>`` stops at the *inner* close tag -- which trims the name
    to whatever followed it rather than reading the whole cell.
    """

    found: List[str] = []
    depth = 0
    start = 0
    for tag in _SPAN_TAG_RE.finditer(fragment):
        if tag.group(0)[1] != "/":
            if depth == 0:
                start = tag.end()
            depth += 1
        else:
            depth = max(0, depth - 1)
            if depth == 0:
                found.append(fragment[start : tag.start()])
    return found


def _club_spans(chunk: str) -> List[str]:
    """Club names in document order, from the row's own spans."""

    names: List[str] = []
    for raw in _spans(chunk):
        visible = _text(raw)
        if not visible or not any(ch.isalpha() for ch in visible):
            continue
        # A span can be a score, a when-cell or a club; only a club is a name.
        if _SCORE_RE.fullmatch(visible.strip()):
            continue
        if _looks_like_when(visible):
            continue
        names.append(visible)
    return names


def _final_block(chunk: str) -> str:
    """The row's ``final-score`` block alone: both clubs and the total.

    Empty when the row has no block, which is how an unplayed game reads -- or
    how a game reads if its markup moved, which is the same answer: not a final.
    """

    at = chunk.find(_FINISHED_CLASS)
    if at < 0:
        return ""
    end = chunk.find(_HIDDEN_CLASS, at)
    return chunk[at : end if end > at else at + 4000]


def _club_names(chunk: str) -> List[str]:
    """The two clubs, preferring the block that also carries the total.

    Two page shapes serve this and only one is byte-verified: the captured row
    puts each club's span *inside* the final-score columns, while an earlier
    derivation put them in sibling ``esake-program-team`` cells.  Both are read
    the same way -- block first, and the container only when the block does not
    yield exactly two names, with the hidden per-quarter row always cut away.
    """

    block = _final_block(chunk)
    if block:
        names = _club_spans(block)
        if len(names) == 2:
            return names
    visible = chunk
    at = visible.find(_HIDDEN_CLASS)
    if at >= 0:
        visible = visible[:at]
    return _club_spans(visible)


def _score(chunk: str) -> Optional[Tuple[int, int]]:
    """The final, and only from a ``final-score`` block.

    A bare ``NN-NN`` elsewhere on the page is not evidence of anything: a date
    cell or a bracket label can wear that shape. No block, no final.
    """

    block = _final_block(chunk)
    if not block:
        return None
    match = _SCORE_RE.search(_text(block))
    if not match:
        return None
    return int(match.group(1)), int(match.group(2))


def _tip(chunk: str) -> Optional[str]:
    match = _TIME_RE.search(_text(chunk))
    return f"{match.group(1).zfill(2)}:{match.group(2)}" if match else None


def _date_cell_text(chunk: str) -> str:
    """The visible text of the row's first info cell, which is the clock cell."""

    at = chunk.find(_DATE_CELL_CLASS)
    if at < 0:
        return ""
    opened = chunk.find(">", at)
    start = opened + 1 if opened >= 0 else at
    end = chunk.find("</div>", start)
    return _text(chunk[start : end if end > start else start + 400])


def date_probe(chunk: str) -> Dict[str, object]:
    """What the date reader sees, for when the page and the reader disagree.

    The whole-row read and the clock-cell read are reported side by side with the
    matches each produced, because "the page shows a date and the reader says
    None" has two very different causes -- text the reader never received, and a
    pattern that did not match text it did receive -- and one call should tell
    them apart.
    """

    cell = _date_cell_text(chunk)
    text = _text(chunk)
    return {
        "cell_text": cell[:160],
        "cell_matches": [list(pair) for pair in _DAY_MONTH_RE.findall(cell)],
        "cell_result": _day_and_month(cell) if cell else None,
        "text_head": text[:160],
        "text_matches": [list(pair) for pair in _DAY_MONTH_RE.findall(text)][:5],
        "result": _day_and_month(chunk),
    }


def _day_and_month(chunk: str) -> Tuple[Optional[int], Optional[int]]:
    """The day and the Greek month, reading around the page's own line breaks.

    The served date cell wraps inside the month name -- the raw bytes carry
    newlines and indentation between «Νοε» and «μβρίου» -- so a regex that
    demands ``21<space>Νοεμβρίου`` finds nothing on a page that plainly shows a
    date.  Whitespace inside a word is a rendering artifact, so the second pass
    removes all of it: «Σαβ 21 Νοεμβρίου 19:00» becomes «Σαβ21Νοεμβρίου19:00»
    and the day still bounds itself because a month name is letters.
    """

    # The clock cell first: it is the row's own date, and a smaller input cannot
    # be broken by anything else on the page.  The whole row is the fallback.
    for scope in (_date_cell_text(chunk), _text(chunk)):
        if not scope:
            continue
        for day_text, label in _DAY_MONTH_RE.findall(scope):
            month = month_from_label(label)
            if month is not None:
                return int(day_text), month
        squeezed = "".join(scope.split())
        for day_text, label in _UNSPACED_DATE_RE.findall(squeezed):
            month = month_from_label(label)
            if month is not None:
                return int(day_text), month
    return None, None


def _resolve_date(day: Optional[int], month: Optional[int], season_start_year: Optional[int]):
    """A calendar date, or ``None`` — never a guess at the year.

    A European season runs from late summer into the next calendar year, so the
    caller supplies the start year and months before August roll into +1. Without
    that year there is no date, and without a date nothing can settle.
    """

    if day is None or month is None or season_start_year is None:
        return None
    year = season_start_year if month >= 8 else season_start_year + 1
    try:
        return date(year, month, day)
    except ValueError:
        return None


@dataclass(frozen=True)
class ResultRow:
    """One game as this page published it. No quarters: a total is not a split."""

    idgame: str
    home: Optional[str]
    away: Optional[str]
    home_score: Optional[int]
    away_score: Optional[int]
    played_on: Optional[date]
    tip_local: Optional[str]
    url: str
    source: str = SOURCE

    @property
    def has_result(self) -> bool:
        """The page named both clubs and both totals, date or no date."""

        return (
            self.home is not None
            and self.away is not None
            and self.home_score is not None
            and self.away_score is not None
        )

    @property
    def is_final(self) -> bool:
        """A result that can be attached to a calendar day, and therefore graded."""

        return self.has_result and self.played_on is not None


@dataclass(frozen=True)
class ResultPage:
    rows: Tuple[ResultRow, ...]
    finals: Tuple[ResultRow, ...]
    incomplete: Tuple[Dict[str, str], ...]
    counts: Dict[str, int]

    def to_dict(self) -> Dict[str, object]:
        return {
            "counts": dict(self.counts),
            "finals": [
                {
                    "idgame": row.idgame,
                    "home": row.home,
                    "away": row.away,
                    "home_score": row.home_score,
                    "away_score": row.away_score,
                    "played_on": row.played_on.isoformat() if row.played_on else None,
                    "tip_local": row.tip_local,
                    "url": row.url,
                    "source": row.source,
                }
                for row in self.finals
            ],
            "incomplete": [dict(item) for item in self.incomplete],
        }


def parse_results(html: str, *, season_start_year: Optional[int] = None) -> ResultPage:
    """Read one matchday page. Raises when the page carries no game identity.

    An empty list of finals from a page that *was* read is a legitimate answer
    (a round not yet played); a page with no recognisable row is not, because it
    says nothing about the round and must not read as "no games".
    """

    if not isinstance(html, str) or not html.strip():
        raise IdentityError(f"{PROVIDER}: empty body, so the round was not read")

    chunks = _ROW_RE.split(html)[1:]
    links = _LINK_RE.findall(html)
    if not chunks:
        raise IdentityError(
            f"{PROVIDER}: no row container recognised ({len(links)} game link(s) "
            "present) — the markup moved and this reader cannot vouch for the page"
        )
    if not links:
        raise IdentityError(f"{PROVIDER}: page carries no EsakegameView game identity")

    rows: List[ResultRow] = []
    incomplete: List[Dict[str, str]] = []
    counts = {
        "containers": len(chunks),
        "rows": 0,
        "with_final": 0,
        "without_final": 0,
        "names_ambiguous": 0,
        "unresolved_date": 0,
        "duplicate_idgame": 0,
        "no_identity": 0,
    }
    seen: set[str] = set()

    for chunk in chunks:
        found = _LINK_RE.search(chunk)
        if not found:
            counts["no_identity"] += 1
            continue
        idgame = found.group(1).upper()
        if idgame in seen:
            counts["duplicate_idgame"] += 1
            continue
        seen.add(idgame)

        names = _club_names(chunk)
        score = _score(chunk)
        day, month = _day_and_month(chunk)
        played_on = _resolve_date(day, month, season_start_year)
        tip = _tip(chunk)

        home = names[0] if len(names) == 2 else None
        away = names[1] if len(names) == 2 else None
        if len(names) != 2:
            counts["names_ambiguous"] += 1
            incomplete.append(
                {"idgame": idgame, "reason": f"club names not exactly two ({len(names)})"}
            )

        row = ResultRow(
            idgame=idgame,
            home=home,
            away=away,
            home_score=score[0] if score else None,
            away_score=score[1] if score else None,
            played_on=played_on,
            tip_local=tip,
            url=game_url(idgame),
        )
        counts["rows"] += 1

        if score is None:
            counts["without_final"] += 1
            incomplete.append({"idgame": idgame, "reason": "no final-score block in the row"})
        if row.has_result and played_on is None:
            counts["unresolved_date"] += 1
            incomplete.append(
                {
                    "idgame": idgame,
                    "reason": "result present but the date could not be read "
                    f"(day={day!r} month={month!r} season_start_year={season_start_year!r})",
                }
            )
        if row.is_final:
            counts["with_final"] += 1
        rows.append(row)

    finals = tuple(row for row in rows if row.is_final)
    return ResultPage(rows=tuple(rows), finals=finals, incomplete=tuple(incomplete), counts=counts)


__all__ = [
    "CHAMPIONSHIP_GBL",
    "GAME_PATH",
    "MIN_INTERVAL_SECONDS",
    "PHASE_REGULAR",
    "PROVIDER",
    "REQUEST_BUDGET",
    "REQUEST_HEADERS",
    "RESULTS_PATH",
    "ResultPage",
    "ResultRow",
    "SITE",
    "SOURCE",
    "checksum",
    "date_probe",
    "default_fetch",
    "game_url",
    "month_from_label",
    "parse_results",
    "results_url",
]
