"""Capture the period table for finished games whose lane has no results route.

Why this is a capture tool and not a settlement adapter
------------------------------------------------------
Five fixtures-table lanes (WCBA Women, Lithuania NKL, WNBL Women, Israel Liga
Leumit, Greece) can acquire fixtures and research-only prices, but they cannot
settle, so they cannot forecast either.  The forecast corpus is not "finals":

    forecasting/model.py:409   game.actual is not None and actual.reconciles()[0]

and ``ActualResult.reconciles`` refuses ``PeriodKind.UNKNOWN``.  A final score
alone — 89-75 with no quarters — is therefore **not** enough to enter training.
The lanes need the period table, which lives on each finished game's own match
page, not on the fixtures page (all 270 score cells on the captured fixtures
page are empty).

Writing that parser blind would be guessing at markup, and this project has
already paid for guessing.  So this tool does the one thing that cannot be
guessed: it fetches a bounded number of match pages, saves each body verbatim
with its sha256, and reports neutral shape signals.  The parser is then written
against real served bytes, and a test is written against the body that is
committed for it.

It deliberately produces **no results**.  There is no ``ActualResult`` here, no
partial score, and no "best effort" number: a page that was captured is
evidence, and a page that was not captured is a named gap.

Robots
------
Every URL this tool builds is a path on ``/basketball/`` derived from the
league's already-verified fixtures slug, with no query string.  The
season-navigation query shapes (``?stage=``, ``?year=``, ``?month=``,
``?page=``, ``?match=``) are robots-disallowed and are never used here.

Usage
-----
    python scripts/probe_betexplorer_results.py --league GREECE_BASKET_LEAGUE --max 6
    python scripts/probe_betexplorer_results.py --league CHINA_WCBA_W --date 2026-11-24

Then hand the saved bodies back for the parser. Each is a served page, not a
reconstruction.
"""

from __future__ import annotations

import argparse
from datetime import datetime, timedelta, timezone
import hashlib
import json
from pathlib import Path
import re
import sys
import time

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from ingestion.providers.betexplorer import BASE, FIXTURES_URLS, SITE_TZ, _rows  # noqa: E402

#: A game three hours past tip is certainly over; anything closer may be live,
#: and a live page's numbers are not a final.
LIVE_GUARD = timedelta(hours=3)

#: Minimum spacing between match-page requests.  The source answers 429 to fast
#: bursts; this tool is a one-off capture, so it can afford to be slow.
MIN_INTERVAL_SECONDS = 3.0

#: Hard ceiling on captured games per run, regardless of ``--max``.
MAX_CAPTURE = 12

#: Match pages get a longer gap than fixtures rows: three of them, a few seconds
#: apart, drew a 429 from this host on 2026-11-24.  Being throttled is the host
#: telling us the interval, and the answer is to take it, not to retry.
MATCH_INTERVAL_SECONDS = 10.0

#: A score in a text cell, and the word the scoreboard tends to be named after.
_COLON_SCORE_RE = re.compile(r"\b\d{1,3}:\d{1,3}\b")
_SCOREBOARD_RE = re.compile(r"scoreboard", re.IGNORECASE)

#: The authority this tool acts under when a page is behind a consent gate.
#: Carried in every capture: a gate passed without a named authority is the sort
#: of quiet decision this project keeps out of its pipelines.
AGE_GATE_AUTHORISATION = (
    "Operator authorisation, 2026-11-24: the operator states they meet this site's 18+ "
    "requirement, so the age self-declaration is theirs and travels with every capture "
    "taken past the gate. It is not a robots exception: path-only URLs, no query shapes, "
    "no per-game burst, and the spacing rules are unchanged."
)

#: Seconds between runs against this host.  Three match pages a few seconds apart
#: drew a 429 on 2026-11-24; a pause inside one run does not help when the runs are
#: the burst, so the stamp lives on disk and the next run honours it.
COOLDOWN_SECONDS = 30.0

#: An age gate is not a match page, and it is 110 KB of modal and banners that
#: looks exactly like a captured page to everything downstream.
_GATE_MARKERS = (
    'id="ageTitle"',
    "age__modal",
    "idade obrigat",
    "age verification",
    "age-verification",
)


_GATE_WRITE_RES = (
    re.compile(r"document\.cookie\s*=\s*[\"'`](?P<value>[^\"'`]+)"),
    re.compile(r"(?:local|session)Storage\.setItem\(\s*[\"'](?P<key>[^\"']+)[\"']"),
)

#: Identifiers that merely contain "age" -- a CSS class, an element id, a vendor
#: name. Reported separately from the writes because they are not a way in, and a
#: list that mixes the two invites replaying a class name.
_GATE_MENTION_RE = re.compile(r"[\"'](?P<key>[A-Za-z0-9_-]*age[A-Za-z0-9_-]*)[\"']")


def gate_hints(body: str) -> list[str]:
    """What a gate writes when its accept button is pressed.

    A gate is passed by reproducing what it records, and guessing the name is how
    a project ends up arguing with a host. These are read from the page's own
    script, in document order: only cookie values and storage keys, because those
    are the only things that are writes.
    """

    hints: list[str] = []
    for pattern in _GATE_WRITE_RES:
        for match in pattern.finditer(body or ""):
            value = match.group(match.lastindex or 0).strip()
            if value and value not in hints:
                hints.append(value)
    return hints[:12]


def gate_mentions(body: str) -> list[str]:
    """Names that look gate-shaped but record nothing -- context, not a way in."""

    mentions: list[str] = []
    for match in _GATE_MENTION_RE.finditer(body or ""):
        value = match.group("key").strip()
        if value and value not in mentions:
            mentions.append(value)
    return mentions[:12]


_COOKIE_WRITE_RE = re.compile(
    r"document\.cookie\s*=\s*[\"'`](?P<value>[^\"'`]+)[\"'`]"
)
_STORAGE_WRITE_RE = re.compile(
    r"(?:local|session)Storage\.setItem\(\s*[\"'](?P<key>[^\"']+)[\"']\s*,"
)


def gate_cookies(body: str) -> list[tuple[str, str]]:
    """The cookies a gate writes on accept, as (name, value), in page order.

    These are the only thing an HTTP client can replay: a cookie travels in a
    header. A storage key cannot, which is why ``gate_storage_keys`` exists --
    the difference is the difference between passing the gate and pretending to.
    """

    found: list[tuple[str, str]] = []
    for match in _COOKIE_WRITE_RE.finditer(body or ""):
        pair = match.group("value").split(";", 1)[0].strip()
        name, _, value = pair.partition("=")
        name = name.strip()
        if not name:
            continue
        item = (name, value.strip())
        if item not in found:
            found.append(item)
    return found


def gate_storage_keys(body: str) -> list[str]:
    """Storage keys the accept button sets -- readable, not replayable."""

    keys: list[str] = []
    for match in _STORAGE_WRITE_RE.finditer(body or ""):
        key = match.group("key").strip()
        if key and key not in keys:
            keys.append(key)
    return keys


def cooldown_wait(previous: float | None, now: float, seconds: float) -> float:
    """Seconds still owed before this run may touch the host.

    The stamp records the earliest instant the *next* run may act, so a run that
    had to wait does not bank the wait and let the one after it fire early.
    """

    if previous is None:
        return 0.0
    return max(0.0, (float(previous) + float(seconds)) - float(now))


class HostCooldown:
    """A stamp on disk, so separate runs cannot add up to a burst."""

    def __init__(self, path: Path, seconds: float = COOLDOWN_SECONDS):
        self.path = Path(path)
        self.seconds = float(seconds)

    def previous(self) -> float | None:
        try:
            return float(self.path.read_text(encoding="utf-8").strip())
        except (OSError, ValueError):
            return None

    def wait(self, *, now: float | None = None, sleep=time.sleep) -> float:
        current = time.time() if now is None else float(now)
        delay = cooldown_wait(self.previous(), current, self.seconds)
        if delay > 0:
            sleep(delay)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.path.write_text(f"{current + delay:.3f}\n", encoding="utf-8")
        return delay


def gate_reason(body: str) -> str | None:
    """Why this body is a gate rather than the page that was asked for."""

    lowered = (body or "").lower()
    for marker in _GATE_MARKERS:
        if marker.lower() in lowered:
            return f"age gate ({marker})"
    return None

_INT_CELL_RE = re.compile(r">\s*(\d{1,3})\s*<")
_INT_TOKEN_RE = re.compile(r"\b\d{1,3}\b")


def match_url(league: str, event_id: str) -> str:
    """Build the game's own page from the league's verified fixtures path."""

    base = FIXTURES_URLS[league]
    if not base.endswith("fixtures/"):
        raise ValueError(f"unexpected fixtures path for {league}: {base}")
    url = base[: -len("fixtures/")] + f"{event_id}/"
    if "?" in url:  # pragma: no cover - defensive: query shapes are disallowed
        raise ValueError(f"refusing a query-shaped URL: {url}")
    return url


def results_url(league: str) -> str:
    """The league's own results page: same slug, ``results/`` instead of ``fixtures/``."""

    base = FIXTURES_URLS[league]
    if not base.endswith("fixtures/"):
        raise ValueError(f"unexpected fixtures path for {league}: {base}")
    url = base[: -len("fixtures/")] + "results/"
    if "?" in url:  # pragma: no cover - defensive: query shapes are disallowed
        raise ValueError(f"refusing a query-shaped URL: {url}")
    return url


#: A match link is ``/basketball/<country>/<league>/[<match-slug>/]<CODE>/``; the
#: slug is optional, because ``match_url`` builds the short form and the served
#: pages carry the long one.
_CODE_LINK_RE = re.compile(
    r"/basketball/[a-z-]+/[a-z0-9-]+/(?:[a-z0-9-]+/)?([0-9A-Za-z]{8})/"
)
_CODE_HREF_RE = re.compile(
    r'href="(?P<href>/basketball/[a-z-]+/[a-z0-9-]+/(?:[a-z0-9-]+/)?'
    r'(?P<code>[0-9A-Za-z]{8})/)"'
)
_COLON_SCORE_RE = re.compile(r"\b(\d{1,3})\s*:\s*(\d{1,3})\b")


def index_signals(body: str) -> dict[str, object]:
    """What the results page offers, as counts rather than as a parse.

    A lane whose fixtures table carries only future rows has no finished game to
    open, so the results page is the only way in.  This reports whether the page
    is a table, how many match codes it carries, and how many colon-separated
    score strings appear -- enough to decide whether a parser is worth writing,
    without writing one blind.
    """

    codes = _CODE_LINK_RE.findall(body)
    return {
        "chars": len(body),
        "has_table": "table-main" in body,
        "match_codes": len(codes),
        "distinct_match_codes": len(set(codes)),
        "colon_scores": len(_COLON_SCORE_RE.findall(body)),
        "first_codes": sorted(set(codes))[:8],
        "has_no_matches_hint": "No matches found" in body,
    }


#: How much of the first match's markup to keep.  Enough to hold a whole table row
#: and the two headers above it, small enough to read in one go.  The first
#: capture spent 600 of these on the page's tab bar and pushed the row it was
#: meant to show out of the printed excerpt, so the lead-in is smaller than the
#: tail: the match code belongs near the start of the file, not past its middle.
WINDOW_CHARS = 2200
WINDOW_LEAD_IN = 300


def first_match_window(body: str) -> str:
    """The bytes around the first match code: the row, not the whole page.

    A count says the page carries 34 codes; it does not say what a row looks like.
    This is the smallest artifact that does, and it is saved with the page so the
    parser can be written against captured markup instead of a description of it.
    """

    found = _CODE_LINK_RE.search(body)
    if not found:
        return ""
    start = max(0, found.start() - WINDOW_LEAD_IN)
    return body[start : start + WINDOW_CHARS]


def capture_index(league: str, *, out_dir: Path, read_text) -> dict[str, object]:
    """Save the league's results page verbatim and report its shape."""

    url = results_url(league)
    out_dir.mkdir(parents=True, exist_ok=True)
    record: dict[str, object] = {"league": league, "url": url}
    try:
        body = read_text(url)
    except Exception as exc:  # noqa: BLE001 - a failed read is a receipt
        record.update({"status": "fetch_failed", "detail": str(exc)[:200]})
        body = ""
    if body:
        path = out_dir / "results-index.html"
        path.write_text(body, encoding="utf-8")
        window = first_match_window(body)
        window_path = out_dir / "results-index-window.txt"
        if window:
            window_path.write_text(window, encoding="utf-8")
        record.update(
            {
                "status": "captured",
                "bytes": len(body.encode("utf-8")),
                "sha256": hashlib.sha256(body.encode("utf-8")).hexdigest(),
                "path": str(path),
                "window_chars": len(window),
                "window_path": str(window_path) if window else None,
                **index_signals(body),
            }
        )
    index_path = out_dir / "results-index.json"
    payload = {
        "schema_version": "magolide.results_probe.v1",
        "mode": "results_index",
        "captured_utc": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "note": (
            "A served page saved verbatim. No score is parsed and no ActualResult "
            "is produced; this answers what the page carries, not what a game was."
        ),
        **record,
    }
    index_path.write_text(json.dumps(payload, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    payload["index_path"] = str(index_path)
    return payload


def finished_candidates(
    league: str, html: str, *, now: datetime, limit: int
) -> list[dict[str, str]]:
    """Rows whose Paris tip is far enough in the past to be a final."""

    if not isinstance(html, str) or "table-main" not in html:
        raise ValueError(f"{league}: fixtures page did not return a results table")
    cutoff = now - LIVE_GUARD
    output: list[dict[str, str]] = []
    seen: set[str] = set()
    for local_dt, match_id, _slug, home, away, _odds in _rows(html, now):
        if local_dt >= cutoff or match_id in seen:
            continue
        seen.add(match_id)
        output.append(
            {
                "event_id": match_id,
                "home_name": home.strip(),
                "away_name": away.strip(),
                "paris_tip": local_dt.isoformat(),
                "url": match_url(league, match_id),
            }
        )
    output.sort(key=lambda row: row["paris_tip"], reverse=True)
    return output[:limit]


#: Quarter and period markers a match page would use if it carried them.
_PERIOD_RE = re.compile(r"\b(?:[1-4]Q|Q[1-4]|OT)\b")


#: The league's own tab paths.  Each is eight letters, so the code pattern matches
#: them: the first `--match 1` run fetched `/.../nkl/fixtures/` and reported
#: `event_id: fixtures`, which is the kind of "success" that teaches nothing.
_NAV_CODES = frozenset({"fixtures", "results", "summary", "stats"})


def absolute_url(url_or_path: str) -> str:
    """The host's absolute URL for a path it published as a relative href.

    A captured results page writes ``/basketball/...`` without a scheme, and a
    client handed that answers ``Invalid URL ... No scheme supplied``.  That is a
    tool fault with a misleading shape -- it looks like the host refusing -- so the
    path is made absolute here, once, rather than at each call site.
    """

    if url_or_path.startswith(("http://", "https://")):
        return url_or_path
    # ``BASE`` already ends in ``/basketball``, and so does every href here.
    site = BASE[: -len("/basketball")]
    return site + ("" if url_or_path.startswith("/") else "/") + url_or_path


def _is_match_code(code: str) -> bool:
    """A real code mixes case or digits; a nav word is eight lowercase letters."""

    if code in _NAV_CODES:
        return False
    return any(char.isdigit() or char.isupper() for char in code)


def index_matches(body: str) -> list[dict[str, str]]:
    """Distinct matches from a captured results page: code plus the page's own URL.

    The URL comes from the href rather than being rebuilt, so a match page is
    opened at the address the host published for it.
    """

    found: dict[str, str] = {}
    for match in _CODE_HREF_RE.finditer(body):
        code = match.group("code")
        if code in found or not _is_match_code(code):
            continue
        found[code] = match.group("href")
    return [{"code": code, "url": url} for code, url in found.items()]


def index_codes(body: str) -> list[str]:
    """Just the codes, in document order."""

    return [match["code"] for match in index_matches(body)]


def capture_match_pages(
    league: str,
    *,
    matches: list[dict[str, str]],
    out_dir: Path,
    read_text,
    sleep=time.sleep,
    limit: int = 1,
    cooldown: "HostCooldown | None" = None,
    gate_accept: bool = False,
    gate_read_text=None,
) -> dict[str, object]:
    """Fetch the league's own match pages for the first ``limit`` codes.

    The results index gives finals and nothing else, and a final does not
    reconcile to a quarter record.  A match page is where period scores would be,
    so this answers "can this host settle?" for the price of one request per
    league rather than for the price of a design decision.
    """

    wanted = matches[: max(0, min(int(limit), MAX_CAPTURE))]
    out_dir.mkdir(parents=True, exist_ok=True)
    waited = cooldown.wait(sleep=sleep) if cooldown is not None else 0.0
    captured: list[dict[str, object]] = []
    throttled = False
    requests_made = 0
    for position, match in enumerate(wanted):
        if position:
            sleep(MATCH_INTERVAL_SECONDS)
        code = match["code"]
        url = absolute_url(match.get("url") or match_url(league, code))
        try:
            requests_made += 1
            body = read_text(url)
        except Exception as exc:  # noqa: BLE001 - a failed capture is a receipt
            detail = str(exc)[:200]
            captured.append({"event_id": code, "url": url, "status": "fetch_failed",
                             "detail": detail})
            # 429 is the host naming its own limit. Stopping is the whole point of
            # a receipt that carries the status: retrying would spend the budget
            # arguing with a host that already answered.
            if "429" in detail:
                throttled = True
                break
            continue
        path = out_dir / f"match_{code}.html"
        path.write_text(body, encoding="utf-8")
        gate = gate_reason(body)
        passed_by = None
        if gate is not None and gate_accept:
            cookies = gate_cookies(body)
            if cookies and gate_read_text is not None:
                header = "; ".join(f"{name}={value}" for name, value in cookies)
                try:
                    requests_made += 1
                    second = gate_read_text(url, header)
                except Exception as exc:  # noqa: BLE001 - the retry is evidence too
                    gate = f"{gate}; replay failed: {str(exc)[:120]}"
                else:
                    if gate_reason(second) is None:
                        body = second
                        path.write_text(body, encoding="utf-8")
                        gate = None
                        passed_by = [name for name, _ in cookies]
                    else:
                        gate = f"{gate}; still gated after replaying {len(cookies)} cookie(s)"
            elif cookies:
                gate = f"{gate}; replay requested without a fetcher"
            elif not cookies:
                keys = gate_storage_keys(body)
                gate = (
                    f"{gate}; no cookie to replay"
                    + (f", storage keys {keys}" if keys else "")
                    + " -- needs a browser, not a header"
                )
        if gate is not None:
            # Saved, because it is evidence of what the host served, but never
            # counted as a match page: 110 KB of age modal reads as a captured
            # page to every check that only looks at size.
            captured.append(
                {
                    "event_id": code,
                    "url": url,
                    "status": "gated",
                    "gate": gate,
                    "bytes": len(body.encode("utf-8")),
                    "sha256": hashlib.sha256(body.encode("utf-8")).hexdigest(),
                    "path": str(path),
                }
            )
            continue
        text = re.sub(r"<[^>]+>", " ", body)
        captured.append(
            {
                "event_id": code,
                "url": url,
                "status": "captured",
                "bytes": len(body.encode("utf-8")),
                "sha256": hashlib.sha256(body.encode("utf-8")).hexdigest(),
                "path": str(path),
                **({"gate_passed_by": passed_by} if passed_by else {}),
                "period_markers": len(_PERIOD_RE.findall(text)),
                "period_examples": sorted(set(_PERIOD_RE.findall(text)))[:6],
                # What a scoreboard looks like when it is not a Q1..Q4 table:
                # a pair of numbers in a text cell, and the word the block is
                # usually named after.
                "colon_scores": len(_COLON_SCORE_RE.findall(text)),
                "scoreboard_mentions": len(_SCOREBOARD_RE.findall(body)),
                **shape_signals(body),
            }
        )

    payload = {
        "schema_version": "magolide.results_probe.v1",
        "mode": "match_pages",
        "league": league,
        "captured_utc": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "requested": len(wanted),
        "stopped_on_429": throttled,
        "gated": sum(1 for page in captured if page["status"] == "gated"),
        "cooldown_seconds_waited": round(waited, 3),
        "requests_made": requests_made,
        "gate_accept": bool(gate_accept),
        "age_gate_authorisation": AGE_GATE_AUTHORISATION,
        "note": (
            "Match pages saved verbatim. period_markers counts Q1..Q4/OT-shaped "
            "tokens, which is what a reconciliation record would need and what a "
            "results index does not carry."
        ),
        "pages": captured,
    }
    path = out_dir / "match-pages.json"
    path.write_text(json.dumps(payload, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    payload["receipt_path"] = str(path)
    return payload


def shape_signals(body: str) -> dict[str, object]:
    """Neutral counts, so the parser can be written against the real page.

    Nothing here interprets a number as a score.  These exist to tell the next
    reader which markup the served page actually uses.
    """

    integer_cells = _INT_CELL_RE.findall(body)
    return {
        "chars": len(body),
        "td_cells": body.count("<td"),
        "span_cells": body.count("<span"),
        "integer_cells": len(integer_cells),
        "integer_tokens": len(_INT_TOKEN_RE.findall(body)),
        "first_integer_cells": integer_cells[:16],
        "has_scoreboard_hint": any(
            hint in body
            for hint in ("match-scoreboard", "scoreboard", "js-scoreboard", "Period")
        ),
    }


def capture(
    league: str,
    *,
    max_games: int,
    date: str | None,
    out_dir: Path,
    now: datetime,
    read_text,
    sleep=time.sleep,
) -> dict[str, object]:
    """Fetch up to ``max_games`` finished match pages and save them verbatim."""

    limit = max(1, min(int(max_games), MAX_CAPTURE))
    fixtures_html = read_text(FIXTURES_URLS[league])
    if date:
        # A dated capture is not a date-bounded request: the fixtures page is
        # season-wide, and the date only filters which rows are worth fetching.
        now = datetime.strptime(date.replace("-", ""), "%Y%m%d").replace(
            hour=23, minute=59, tzinfo=timezone.utc
        )
    candidates = finished_candidates(league, fixtures_html, now=now, limit=limit)
    out_dir.mkdir(parents=True, exist_ok=True)

    captured: list[dict[str, object]] = []
    for index, row in enumerate(candidates):
        if index:
            sleep(MIN_INTERVAL_SECONDS)
        try:
            body = read_text(row["url"])
        except Exception as exc:  # noqa: BLE001 - a failed capture is a receipt, not a crash
            captured.append({**row, "status": "fetch_failed", "detail": str(exc)[:200]})
            continue
        digest = hashlib.sha256(body.encode("utf-8")).hexdigest()
        path = out_dir / f"{row['event_id']}.html"
        path.write_text(body, encoding="utf-8")
        captured.append(
            {
                **row,
                "status": "captured",
                "bytes": len(body.encode("utf-8")),
                "sha256": digest,
                "path": str(path),
                **shape_signals(body),
            }
        )

    index_path = out_dir / "index.json"
    payload = {
        "schema_version": "magolide.results_probe.v1",
        "league": league,
        "captured_utc": now.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "site_timezone": SITE_TZ,
        "live_guard_hours": LIVE_GUARD.total_seconds() / 3600,
        "requested": limit,
        "candidates": len(candidates),
        "captured": sum(1 for row in captured if row["status"] == "captured"),
        "note": (
            "Bodies are served pages saved verbatim. No score is parsed and no "
            "ActualResult is produced: a final alone cannot enter the forecast "
            "corpus, which requires actual.reconciles()."
        ),
        "games": captured,
    }
    index_path.write_text(json.dumps(payload, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    payload["index_path"] = str(index_path)
    return payload


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument(
        "--league",
        choices=sorted(FIXTURES_URLS),
        help=(
            "lane to probe; required for anything that fetches or replays a league, "
            "and not needed for --gate-hints, which reads a saved file"
        ),
    )
    parser.add_argument("--max", type=int, default=6, help="match pages to capture (hard cap 12)")
    parser.add_argument("--date", help="only rows tipping on/before this YYYY-MM-DD")
    parser.add_argument("--out", default="results_probe", help="directory for saved bodies")
    parser.add_argument("--now", help="override the observation instant (UTC ISO)")
    parser.add_argument(
        "--gate-hints",
        metavar="PATH",
        help=(
            "print what a saved gate page writes when its accept button is pressed "
            "(a cookie or a storage key), with no request at all"
        ),
    )
    parser.add_argument(
        "--gate-accept",
        action="store_true",
        help=(
            "when a page comes back as an age gate, replay the cookie its own accept "
            "button writes (one extra request, never a loop); a gate that only sets "
            "localStorage is reported as needing a browser rather than retried"
        ),
    )
    parser.add_argument(
        "--cooldown",
        type=float,
        default=COOLDOWN_SECONDS,
        help=(
            "seconds this run must wait since the last one touched the host; "
            "0 disables the on-disk stamp"
        ),
    )
    parser.add_argument(
        "--match",
        type=int,
        default=0,
        metavar="N",
        help=(
            "fetch N match pages for the first codes of the captured results "
            "index -- the only place period scores could be (hard cap 12)"
        ),
    )
    parser.add_argument(
        "--read",
        action="store_true",
        help=(
            "replay the reader against an already-captured results-index.html: "
            "no network, no request, just what the saved bytes say"
        ),
    )
    parser.add_argument(
        "--index",
        action="store_true",
        help=(
            "capture the league's own results page instead of match pages -- the "
            "entry point for a lane whose fixtures table holds only future rows"
        ),
    )
    args = parser.parse_args(argv)

    from ingestion.providers.betexplorer import _text_fetch

    now = (
        datetime.fromisoformat(args.now.replace("Z", "+00:00"))
        if args.now
        else datetime.now(timezone.utc)
    )
    if args.gate_hints:
        body = Path(args.gate_hints).read_text(encoding="utf-8", errors="replace")
        print(
            json.dumps(
                {
                    "path": args.gate_hints,
                    "gate": gate_reason(body),
                    "hints": gate_hints(body),
                    "mentions": gate_mentions(body),
                    "replayable": bool(gate_cookies(body)),
                    "age_gate_authorisation": AGE_GATE_AUTHORISATION,
                },
                indent=2,
                sort_keys=True,
            )
        )
        return 0 if gate_hints(body) else 1

    if args.league is None:
        parser.error(
            "--league is required to fetch or replay a league; only --gate-hints "
            "reads a saved page and needs no lane"
        )

    if args.match:
        from ingestion.providers.betexplorer import _text_fetch

        index_path = Path(args.out) / args.league / "results-index.html"
        if not index_path.exists():
            parser.error(f"no captured results index at {index_path}; run --index first")
        cooldown = (
            HostCooldown(Path(args.out) / ".host-cooldown", args.cooldown)
            if args.cooldown > 0
            else None
        )
        gate_read_text = None
        if args.gate_accept:
            from ingestion.http import HttpClient

            def gate_read_text(url: str, cookie: str) -> str:
                return HttpClient().get_text(url, headers={"Cookie": cookie})

        payload = capture_match_pages(
            args.league,
            matches=index_matches(index_path.read_text(encoding="utf-8")),
            out_dir=Path(args.out) / args.league,
            read_text=_text_fetch(None, None),
            limit=args.match,
            cooldown=cooldown,
            gate_accept=args.gate_accept,
            gate_read_text=gate_read_text,
        )
        print(json.dumps(payload, indent=2, sort_keys=True))
        return 0 if any(page["status"] == "captured" for page in payload["pages"]) else 1

    if args.read:
        path = Path(args.out) / args.league / "results-index.html"
        if not path.exists():
            parser.error(f"no captured page at {path}; run --index first")
        from ingestion.providers.betexplorer_results import parse_results

        page = parse_results(path.read_text(encoding="utf-8"))
        payload = {
            "schema_version": "magolide.results_probe.v1",
            "mode": "results_read",
            "league": args.league,
            "path": str(path),
            **page.to_dict(),
        }
        print(json.dumps(payload, indent=2, sort_keys=True, ensure_ascii=False))
        return 0 if page.rows else 1

    if args.index:
        payload = capture_index(
            args.league,
            out_dir=Path(args.out) / args.league,
            read_text=_text_fetch(None, None),
        )
        print(json.dumps(payload, indent=2, sort_keys=True))
        return 0 if payload.get("status") == "captured" else 1

    payload = capture(
        args.league,
        max_games=args.max,
        date=args.date,
        out_dir=Path(args.out) / args.league,
        now=now,
        read_text=_text_fetch(None, None),
    )
    print(json.dumps(payload, indent=2, sort_keys=True))
    return 0 if payload["captured"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
