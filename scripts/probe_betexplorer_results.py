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
import subprocess
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


#: How close a cookie write has to sit to the gate's own markup to be called
#: related to it. Chosen to cover one inline handler and its callbacks, not the
#: page: a language switcher that clears ``lang`` is a cookie write too, and
#: replaying it is how a tool convinces itself it passed a gate.
GATE_PROXIMITY_CHARS = 3000

_ACCEPT_CONTROL_RE = re.compile(
    r"[\"'](?P<id>[A-Za-z0-9_-]*age[A-Za-z0-9_-]*(?:yes|accept|confirm|ok|enter|btn)"
    r"[A-Za-z0-9_-]*)[\"']",
    re.IGNORECASE,
)


def gate_anchors(body: str) -> list[int]:
    """Offsets of the gate's own markup: its markers and its accept controls."""

    lowered = (body or "").lower()
    positions = [lowered.find(marker.lower()) for marker in _GATE_MARKERS]
    positions = [position for position in positions if position >= 0]
    positions.extend(match.start() for match in _ACCEPT_CONTROL_RE.finditer(body or ""))
    return sorted(positions)


def gate_cookie_writes(body: str) -> list[dict[str, object]]:
    """Every cookie the page writes, with how far it sits from the gate.

    Distance is the whole point.  The 2026-11-24 gate page writes
    ``lang=; expires=Thu, 01 Jan 1970`` (a language reset) and a bare ``mode``;
    both are cookie writes, neither is consent, and replaying them produced a
    second identical 110 KB gate.  A write is a hint only when it is near the
    gate that is supposed to set it.
    """

    anchors = gate_anchors(body)
    writes: list[dict[str, object]] = []
    seen: dict[tuple[str, str], int] = {}
    for match in _COOKIE_WRITE_RE.finditer(body or ""):
        pair = match.group("value").split(";", 1)[0].strip()
        name, _, value = pair.partition("=")
        name = name.strip()
        if not name:
            continue
        distance = min((abs(match.start() - anchor) for anchor in anchors), default=None)
        key = (name, value.strip())
        if key in seen:
            continue
        seen[key] = len(writes)
        writes.append(
            {
                "name": name,
                "value": value.strip(),
                "raw": match.group("value").strip(),
                "distance": distance,
                "near_gate": distance is not None and distance <= GATE_PROXIMITY_CHARS,
            }
        )
    return writes


def gate_cookies(body: str) -> list[tuple[str, str]]:
    """The cookies a gate writes on accept, as (name, value), in page order.

    These are the only thing an HTTP client can replay: a cookie travels in a
    header. A storage key cannot, which is why ``gate_storage_keys`` exists --
    the difference is the difference between passing the gate and pretending to.
    Only writes near the gate count; see ``gate_cookie_writes``.
    """

    return [
        (str(write["name"]), str(write["value"]))
        for write in gate_cookie_writes(body)
        if write["near_gate"]
    ]


def gate_script(body: str) -> dict[str, object]:
    """What the accept handler does, read out of the page itself.

    A saved body is enough to answer this when the handler is inline. It is not
    when the gate ships its logic in an external bundle, and the honest answer
    then is the list of bundles plus no excerpts -- not a plausible guess at a
    key name.
    """

    # A button carries the handler in one of three places: an inline onclick, a
    # listener attached by id in a later script, or a delegated handler in a
    # bundle. The first occurrence of the id is the markup -- the 2026-11-24 page
    # proved that reading only it shows banners and no behaviour at all -- so
    # every occurrence is reported, with whether it looks like code.
    handler_windows: list[dict[str, object]] = []
    controls: list[str] = []
    for match in _ACCEPT_CONTROL_RE.finditer(body or ""):
        control = match.group("id")
        if control not in controls:
            controls.append(control)

    for control in controls[:3]:
        for at in [m.start() for m in re.finditer(re.escape(control), body or "")][:3]:
            window = body[max(0, at - 200) : at + 900]
            handler_windows.append(
                {
                    "control": control,
                    "offset": at,
                    "excerpt": re.sub(r"\s+", " ", window).strip(),
                    "looks_like_code": bool(
                        re.search(r"addEventListener|onclick|\$\s*\(|function\s*\(|=>", window)
                    ),
                }
            )

    blocks = []
    for index, match in enumerate(
        re.finditer(r"<script(?![^>]*\bsrc=)[^>]*>(?P<text>.*?)</script>", body or "", re.S)
    ):
        text = match.group("text")
        writes = bool(re.search(r"document\.cookie|localStorage|sessionStorage", text))
        blocks.append(
            {
                "index": index,
                "bytes": len(text),
                "mentions_controls": [control for control in controls if control in text],
                "writes": writes,
                # The real page attached nothing by id, and two small inline
                # blocks do write cookies. If the handler selects by class
                # instead of by id, the id search cannot see it -- so a writing
                # block is shown rather than summarised.
                **({"excerpt": re.sub(r"\s+", " ", text).strip()[:600]} if writes else {}),
            }
        )

    scripts = []
    for match in re.finditer(r"<script[^>]+src=[\"'](?P<src>[^\"']+)[\"']", body or ""):
        src = match.group("src")
        if src not in scripts:
            scripts.append(src)
    return {
        "accept_handlers": handler_windows,
        "external_scripts": scripts[:20],
        "inline_blocks": blocks[:24],
        "inline_script_blocks": len(blocks),
        # The question the gate hides: is there a match page in this body at all?
        # Same vocabulary as the capture receipt, so the numbers compare.
        "shape": {
            **shape_signals(body),
            "period_markers": len(_PERIOD_RE.findall(re.sub(r"<[^>]+>", " ", body or ""))),
            "colon_scores": len(_COLON_SCORE_RE.findall(re.sub(r"<[^>]+>", " ", body or ""))),
        },
    }


def gate_storage_keys(body: str) -> list[str]:
    """Storage keys the accept button sets -- readable, not replayable."""

    keys: list[str] = []
    for match in _STORAGE_WRITE_RE.finditer(body or ""):
        key = match.group("key").strip()
        if key and key not in keys:
            keys.append(key)
    return keys


def probe_build() -> dict[str, object]:
    """Which copy of this script produced the receipt.

    Three sessions in a row were spent reading output from a checkout that was
    one merge behind: the flags existed in the repository and not in the file
    being run. A receipt that does not name its own build cannot tell that
    apart from a host answering badly, so this asks git and reports what it
    finds, or says plainly that it could not tell.
    """

    root = Path(__file__).resolve().parent.parent
    try:
        commit = subprocess.run(
            ["git", "-C", str(root), "rev-parse", "--short", "HEAD"],
            capture_output=True,
            text=True,
            timeout=5,
            check=False,
        ).stdout.strip()
        dirty = bool(
            subprocess.run(
                ["git", "-C", str(root), "status", "--porcelain", "--", str(Path(__file__))],
                capture_output=True,
                text=True,
                timeout=5,
                check=False,
            ).stdout.strip()
        )
    except (OSError, subprocess.SubprocessError):
        return {"script": Path(__file__).name, "git_commit": None, "modified": None}
    return {
        "script": Path(__file__).name,
        "git_commit": commit or None,
        "modified": dirty,
    }


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

#: The gate's own vocabulary, as a minified bundle would spell it. Used to look
#: for the handler in the asset the page loads, since the page itself attaches
#: nothing by id.
_SET_COOKIE_RE = re.compile(
    r"set_cookie\(\s*[\"'](?P<name>[^\"']+)[\"']\s*,\s*(?P<value>[^,)]{0,80})"
)

_GATE_ASSET_TOKENS = (
    "age__btn--yes",
    "ageYes",
    "ageModal",
    "ageTitle",
    "age__modal",
    "ageOverlay",
)


def bundle_map(js: str) -> dict[str, object]:
    """Look for the gate inside a script the page loads.

    The served page holds the modal's markup and nothing that acts on it, so the
    behaviour is either in a bundle or nowhere. This reports where the gate's own
    tokens appear and every ``document.cookie`` assignment nearby -- the second
    list is the one that can name a consent key, and it is evidence from the
    code, not a guess at a name.
    """

    text = js or ""
    windows: list[dict[str, object]] = []
    for token in _GATE_ASSET_TOKENS:
        for match in list(re.finditer(re.escape(token), text))[:2]:
            windows.append(
                {
                    "token": token,
                    "offset": match.start(),
                    "excerpt": re.sub(
                        r"\s+", " ", text[max(0, match.start() - 300) : match.start() + 500]
                    ).strip(),
                }
            )
    # The gate does not write raw cookies: it calls the bundle's own set_cookie
    # helper. The only raw assignment in 192 KB is the helper's own body, so a
    # reader that stops at document.cookie= reports one write and misses the key.
    anchors = [match.start() for match in re.finditer(r"init_age_wall|ageYes|ageOverlay", text)]
    helper_calls = []
    for match in _SET_COOKIE_RE.finditer(text):
        distance = min((abs(match.start() - anchor) for anchor in anchors), default=None)
        helper_calls.append(
            {
                "name": match.group("name"),
                "value": re.sub(r"\s+", " ", match.group("value")).strip()[:60],
                "offset": match.start(),
                "near_gate": distance is not None and distance <= 4000,
            }
        )

    wall = re.search(r"init_age_wall", text)
    wall_region = None
    if wall is not None:
        wall_region = re.sub(
            r"\s+", " ", text[wall.start() : wall.start() + 3000]
        ).strip()

    writes = []
    for match in list(re.finditer(r"document\.cookie\s*=\s*", text))[:8]:
        writes.append(
            {
                "offset": match.start(),
                "excerpt": re.sub(r"\s+", " ", text[match.start() : match.start() + 160]).strip(),
            }
        )
    return {
        "bytes": len(text.encode("utf-8")),
        "sha256": hashlib.sha256(text.encode("utf-8")).hexdigest(),
        "windows": windows[:6],
        "cookie_writes": writes,
        "set_cookie_calls": helper_calls[:12],
        "gate_keys": [call["name"] for call in helper_calls if call["near_gate"]],
        "age_wall_region": wall_region,
        "tokens_found": [token for token in _GATE_ASSET_TOKENS if token in text],
    }


def cookie_reader(client):
    """Adapt a client's ``get_text`` to ``(url, cookie) -> body``.

    The naive version of this passes the cookie string as ``get_text``'s second
    positional argument, which is ``headers``, and a request that never leaves
    the process fails with "dictionary update sequence element #0 has length 1".
    The 2026-11-24 receipt recorded that as ``fetch_failed`` beside ``gated: 0``,
    which at a glance reads as a gate passed. It was a request that never
    happened. This is the one place the header is built.
    """

    def read(url: str, cookie: str) -> str:
        return client.get_text(url, headers={"Cookie": cookie})

    return read


def gate_bundle_report(
    url: str,
    *,
    out_dir: Path,
    read_text,
    cooldown: "HostCooldown | None" = None,
    sleep=time.sleep,
) -> dict[str, object]:
    """One bounded request for one static asset, then read it offline."""

    out_dir.mkdir(parents=True, exist_ok=True)
    waited = cooldown.wait(sleep=sleep) if cooldown is not None else 0.0
    payload: dict[str, object] = {
        "schema_version": "magolide.results_probe.v1",
        "mode": "gate_bundle",
        "tool": probe_build(),
        "url": url,
        "cooldown_seconds_waited": round(waited, 3),
        "requests_made": 1,
        "age_gate_authorisation": AGE_GATE_AUTHORISATION,
        "note": (
            "A public static asset saved verbatim and read for the gate's own code. "
            "A key found here is evidence; a key guessed from here would not be."
        ),
    }
    try:
        body = read_text(url)
    except Exception as exc:  # noqa: BLE001 - a failed read is the receipt
        payload["status"] = "fetch_failed"
        payload["detail"] = str(exc)[:200]
        return payload
    name = re.sub(r"[^A-Za-z0-9._-]+", "_", url.rsplit("/", 1)[-1])[:60] or "gate-bundle.js"
    path = out_dir / f"gate-bundle-{name}"
    path.write_text(body, encoding="utf-8")
    payload.update({"status": "captured", "path": str(path), **bundle_map(body)})
    return payload


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
        "tool": probe_build(),
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
    cookies: list[tuple[str, str]] | None = None,
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
            if cookies:
                # A key read from the page's own bundle, or supplied by the
                # operator. Named in the receipt so a pass is never mistaken for
                # something this tool worked out by itself.
                body = read_text(url, "; ".join(f"{name}={value}" for name, value in cookies))
            else:
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
            # Not `cookies`: that name holds the caller's header for every
            # request in this loop, and reassigning it here would swap the
            # operator's cookie for the gate's on the next page.
            gate_cookie_pairs = gate_cookies(body)
            if gate_cookie_pairs and gate_read_text is not None:
                header = "; ".join(f"{name}={value}" for name, value in gate_cookie_pairs)
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
                        passed_by = [name for name, _ in gate_cookie_pairs]
                    else:
                        gate = f"{gate}; still gated after replaying {len(gate_cookie_pairs)} cookie(s)"
            elif gate_cookie_pairs:
                gate = f"{gate}; replay requested without a fetcher"
            elif not gate_cookie_pairs:
                keys = gate_storage_keys(body)
                stray = [write["name"] for write in gate_cookie_writes(body) if not write["near_gate"]]
                if stray:
                    gate = (
                        f"{gate}; no cookie written near the accept control "
                        f"(unrelated writes: {stray}) -- replaying those is not consent"
                    )
                else:
                    gate = (
                        f"{gate}; no cookie to replay"
                        + (f", storage keys {keys}" if keys else "")
                        + " -- needs a browser, not a header"
                    )
        text = re.sub(r"<[^>]+>", " ", body)
        shape = shape_signals(body)
        period_markers = len(_PERIOD_RE.findall(text))
        # "Gated" requires the modal AND no content.  BetExplorer ships the modal
        # in the markup of every page and then hides it client-side, so marker
        # presence alone is not a refusal: treating it as one would make a page
        # that carries a table uncapturable forever.  Zero-vs-nonzero, no
        # threshold to tune.
        refused = gate is not None and shape["td_cells"] == 0 and period_markers == 0
        if refused:
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
        captured.append(
            {
                "event_id": code,
                "url": url,
                "status": "captured",
                "bytes": len(body.encode("utf-8")),
                "sha256": hashlib.sha256(body.encode("utf-8")).hexdigest(),
                "path": str(path),
                **({"gate_passed_by": passed_by} if passed_by else {}),
                # The modal was in the HTML and the page still carried content.
                # Named rather than hidden: it is how this host ships.
                **({"gate_markup": gate} if gate else {}),
                "period_markers": period_markers,
                "period_examples": sorted(set(_PERIOD_RE.findall(text)))[:6],
                # What a scoreboard looks like when it is not a Q1..Q4 table:
                # a pair of numbers in a text cell, and the word the block is
                # usually named after.
                "colon_scores": len(_COLON_SCORE_RE.findall(text)),
                "scoreboard_mentions": len(_SCOREBOARD_RE.findall(body)),
                **shape,
            }
        )

    # gated: 0 is not good news on its own. On 2026-11-24 a request that never
    # left the process reported fetch_failed beside gated: 0, which reads at a
    # glance as "the gate was passed". The summary says which it is, in one line,
    # so nobody has to hold three counts in their head to learn nothing happened.
    counts = {
        status: sum(1 for page in captured if page["status"] == status)
        for status in ("captured", "gated", "fetch_failed")
    }
    summary = (
        f"{counts['captured']} captured, {counts['gated']} gated, "
        f"{counts['fetch_failed']} failed of {len(wanted)} requested"
        + (f"; {counts['fetch_failed']} request(s) never returned a body" if counts["fetch_failed"] else "")
    )
    payload = {
        "schema_version": "magolide.results_probe.v1",
        "tool": probe_build(),
        "mode": "match_pages",
        "summary": summary,
        "league": league,
        "captured_utc": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "requested": len(wanted),
        "stopped_on_429": throttled,
        "gated": sum(1 for page in captured if page["status"] == "gated"),
        "cooldown_seconds_waited": round(waited, 3),
        "requests_made": requests_made,
        "gate_accept": bool(gate_accept),
        "cookies_sent": [name for name, _ in (cookies or [])],
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
        "--gate-script",
        metavar="PATH",
        help=(
            "read a saved gate page and print what its accept control does (inline handler "
            "excerpts, external bundles), with no request"
        ),
    )
    parser.add_argument(
        "--cookie",
        action="append",
        metavar="NAME=VALUE",
        help=(
            "send this cookie on each match page request (repeatable) -- for a key read "
            "from the page's own bundle; names are recorded in the receipt"
        ),
    )
    parser.add_argument(
        "--bundle-read",
        metavar="PATH",
        help="re-read an already-saved gate bundle, with no request at all",
    )
    parser.add_argument(
        "--gate-bundle",
        metavar="URL",
        help=(
            "fetch one static script the gate page loads and look for the gate's own "
            "code in it -- one request, honouring the cooldown"
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
    if args.gate_script:
        body = Path(args.gate_script).read_text(encoding="utf-8", errors="replace")
        print(
            json.dumps(
                {
                    "tool": probe_build(),
                    "path": args.gate_script,
                    "gate": gate_reason(body),
                    **gate_script(body),
                },
                indent=2,
                sort_keys=True,
            )
        )
        return 0

    if args.bundle_read:
        body = Path(args.bundle_read).read_text(encoding="utf-8", errors="replace")
        print(
            json.dumps(
                {
                    "schema_version": "magolide.results_probe.v1",
                    "mode": "gate_bundle_read",
                    "tool": probe_build(),
                    "path": args.bundle_read,
                    "requests_made": 0,
                    **bundle_map(body),
                },
                indent=2,
                sort_keys=True,
            )
        )
        return 0

    if args.gate_bundle:
        from ingestion.providers.betexplorer import _text_fetch

        report = gate_bundle_report(
            args.gate_bundle,
            out_dir=Path(args.out),
            read_text=_text_fetch(None, None),
            cooldown=(
                HostCooldown(Path(args.out) / ".host-cooldown", args.cooldown)
                if args.cooldown > 0
                else None
            ),
        )
        print(json.dumps(report, indent=2, sort_keys=True))
        return 0 if report.get("status") == "captured" else 1

    if args.gate_hints:
        body = Path(args.gate_hints).read_text(encoding="utf-8", errors="replace")
        print(
            json.dumps(
                {
                    "tool": probe_build(),
                    "path": args.gate_hints,
                    "gate": gate_reason(body),
                    "hints": [
                        write["raw"] for write in gate_cookie_writes(body) if write["near_gate"]
                    ],
                    "unrelated_writes": [
                        {"raw": write["raw"], "distance": write["distance"]}
                        for write in gate_cookie_writes(body)
                        if not write["near_gate"]
                    ],
                    "storage_keys": gate_storage_keys(body),
                    "mentions": gate_mentions(body),
                    "replayable": bool(gate_cookies(body)),
                    "age_gate_authorisation": AGE_GATE_AUTHORISATION,
                },
                indent=2,
                sort_keys=True,
            )
        )
        # Exit 0 means "there is a cookie to replay", which is the only answer
        # that lets the next command run. Finding writes that belong to something
        # else is not a way in, and saying so with 0 would invite the 429.
        return 0 if gate_cookies(body) else 1

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
        supplied: list[tuple[str, str]] = []
        for item in args.cookie or []:
            name, _, value = item.partition("=")
            if not name.strip():
                parser.error(f"--cookie needs NAME=VALUE, got {item!r}")
            supplied.append((name.strip(), value.strip()))

        gate_read_text = None
        supplied_reader = None
        if args.gate_accept or supplied:
            from ingestion.http import HttpClient

            gate_read_text = cookie_reader(HttpClient())
            supplied_reader = gate_read_text

        payload = capture_match_pages(
            args.league,
            matches=index_matches(index_path.read_text(encoding="utf-8")),
            out_dir=Path(args.out) / args.league,
            read_text=supplied_reader or _text_fetch(None, None),
            limit=args.match,
            cooldown=cooldown,
            gate_accept=args.gate_accept,
            gate_read_text=gate_read_text,
            cookies=supplied or None,
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
            "tool": probe_build(),
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
