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

from ingestion.providers.betexplorer import FIXTURES_URLS, SITE_TZ, _rows  # noqa: E402

#: A game three hours past tip is certainly over; anything closer may be live,
#: and a live page's numbers are not a final.
LIVE_GUARD = timedelta(hours=3)

#: Minimum spacing between match-page requests.  The source answers 429 to fast
#: bursts; this tool is a one-off capture, so it can afford to be slow.
MIN_INTERVAL_SECONDS = 3.0

#: Hard ceiling on captured games per run, regardless of ``--max``.
MAX_CAPTURE = 12

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
    parser.add_argument("--league", required=True, choices=sorted(FIXTURES_URLS))
    parser.add_argument("--max", type=int, default=6, help="match pages to capture (hard cap 12)")
    parser.add_argument("--date", help="only rows tipping on/before this YYYY-MM-DD")
    parser.add_argument("--out", default="results_probe", help="directory for saved bodies")
    parser.add_argument("--now", help="override the observation instant (UTC ISO)")
    args = parser.parse_args(argv)

    from ingestion.providers.betexplorer import _text_fetch

    now = (
        datetime.fromisoformat(args.now.replace("Z", "+00:00"))
        if args.now
        else datetime.now(timezone.utc)
    )
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
