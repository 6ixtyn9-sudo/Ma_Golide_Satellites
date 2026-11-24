"""Read the Greek federation's finals, and answer the one question names cannot.

Why a probe and not an adapter yet
----------------------------------
``GREECE_BASKET_LEAGUE`` cannot settle, so it cannot be graded, so no config for
it can ever legitimately carry a version. The federation publishes the finals
itself, and the 2026-11-24 external probe validated the route to 10/10 coverage.
Two things still stand between that evidence and a settlement lane:

1. **The page's real markup.** The parser in ``ingestion/providers/esake.py`` is
   built from the probe's verbatim snippets, and its container classes are a
   guess that fails closed. One real fetch settles it.
2. **Identity.** The federation names clubs in Greek (``ΚΟΛΟΣΣΟΣ``,
   ``ΠΡΟΜΗΘΕΑΣ``); this lane's canonical games carry the board's Latin names.
   Guessing a transliteration per club is exactly the invention this project
   refuses.

This tool therefore fetches a bounded number of matchday pages, saves each body
verbatim with its sha256, parses them, and then asks the question that needs no
club names at all: **does a game's own tip (to the minute, in Europe/Athens)
identify it uniquely among the finals on that date?** If it does, the join is
falsifiable and cheap. If it does not, the answer is "no", and that is worth
knowing before anyone writes an adapter.

It writes no results: no ``ActualResult``, no partial score, no guessed club.
A page that was read is evidence; a page that was not read is a named failure.

Politeness
----------
Four requests, spaced by default. The host is slow and publishes no crawl rules;
being quick at its expense would be the wrong kind of clever.

Usage
-----
    python scripts/probe_esake_results.py --series 8
    python scripts/probe_esake_results.py --series 7 --series 8 --warehouse localdata/greece_basket_league/warehouse.duckdb

Then hand back the saved bodies so the parser can be written against them.
"""

from __future__ import annotations

import argparse
from datetime import datetime, timezone
from pathlib import Path
import json
import sys
import time
from zoneinfo import ZoneInfo

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from ingestion.models import League  # noqa: E402
from ingestion.providers import esake  # noqa: E402

ATHENS = ZoneInfo("Europe/Athens")
LEAGUE = League.GREECE_BASKET_LEAGUE


def _athens_tip(scheduled_utc: str) -> tuple[str, str] | None:
    """The game's own wall clock in Athens: (YYYY-MM-DD, HH:MM)."""

    text = str(scheduled_utc or "").replace("Z", "+00:00")
    try:
        moment = datetime.fromisoformat(text)
    except ValueError:
        return None
    if moment.tzinfo is None:
        moment = moment.replace(tzinfo=timezone.utc)
    local = moment.astimezone(ATHENS)
    return local.strftime("%Y-%m-%d"), local.strftime("%H:%M")


def _stored_games(warehouse: Path) -> list[dict]:
    """Canonical games already in the lane's warehouse, if the runner has one."""

    if not warehouse.is_file():
        return []
    from ingestion.duckdb_store import DuckDBStore

    store = DuckDBStore(str(warehouse), persist=False)
    try:
        games = store.list_games(league=LEAGUE)
    finally:
        store.close()
    return [
        {
            "game_id": game.game_id,
            "scheduled_utc": game.scheduled_utc,
            "home": game.home_name,
            "away": game.away_name,
            "settled": game.actual is not None,
        }
        for game in games
    ]


def _join(games: list[dict], finals: list[esake.ResultRow]) -> dict:
    """Can a final be attached to a stored game by date + tip, unambiguously?"""

    by_tip: dict[tuple[str, str], list[esake.ResultRow]] = {}
    by_date: dict[str, int] = {}
    without_tip = 0
    for final in finals:
        if final.played_on is None or final.tip_local is None:
            without_tip += 1
            continue
        day = final.played_on.isoformat()
        by_tip.setdefault((day, final.tip_local), []).append(final)
        by_date[day] = by_date.get(day, 0) + 1

    report = {
        "stored_games": len(games),
        "stored_unsettled": sum(1 for game in games if not game["settled"]),
        "finals": len(finals),
        "finals_without_tip": without_tip,
        "attachable": 0,
        "ambiguous": 0,
        "same_date_no_tip": 0,
        "unmatched": 0,
        "examples": [],
    }
    for game in games:
        if game["settled"]:
            continue
        tip = _athens_tip(game["scheduled_utc"])
        if tip is None:
            report["unmatched"] += 1
            continue
        candidates = by_tip.get(tip, [])
        if len(candidates) == 1:
            report["attachable"] += 1
            if len(report["examples"]) < 5:
                report["examples"].append(
                    {
                        "game_id": game["game_id"],
                        "athens_tip": f"{tip[0]} {tip[1]}",
                        "idgame": candidates[0].idgame,
                        "score": f"{candidates[0].home_score}-{candidates[0].away_score}",
                    }
                )
        elif candidates:
            report["ambiguous"] += 1
            if len(report["examples"]) < 5:
                report["examples"].append(
                    {
                        "game_id": game["game_id"],
                        "athens_tip": f"{tip[0]} {tip[1]}",
                        "candidates": [c.idgame for c in candidates],
                        "note": "refused: two games at the same minute",
                    }
                )
        elif by_date.get(tip[0]):
            report["same_date_no_tip"] += 1
        else:
            report["unmatched"] += 1
    return report


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--series",
        type=int,
        action="append",
        default=None,
        help="matchday number, repeatable (omit for the page's own default round)",
    )
    parser.add_argument(
        "--season-start-year",
        type=int,
        default=2026,
        help="season start year, used to turn day+month into a date (default 2026)",
    )
    parser.add_argument(
        "--warehouse",
        default="localdata/greece_basket_league/warehouse.duckdb",
        help="lane warehouse for the date+tip join (skipped when absent)",
    )
    parser.add_argument("--outdir", default="/tmp/esake", help="where bodies are saved")
    parser.add_argument(
        "--spacing",
        type=float,
        default=esake.MIN_INTERVAL_SECONDS,
        help="seconds between requests",
    )
    parser.add_argument("--max-requests", type=int, default=esake.REQUEST_BUDGET)
    parser.add_argument("--quiet", action="store_true")
    args = parser.parse_args(argv)

    series_list = args.series if args.series else [None]
    outdir = Path(args.outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    fetch = esake.default_fetch()

    pages: list[dict] = []
    finals: list[esake.ResultRow] = []
    requests = 0
    for series in series_list:
        if requests >= args.max_requests:
            break
        if requests:
            time.sleep(max(0.0, args.spacing))
        url = esake.results_url(series)
        requests += 1
        try:
            body = fetch(url)
        except Exception as exc:  # noqa: BLE001 - a probe reports, it does not retry forever
            pages.append({"url": url, "series": series, "status": "failed", "error": str(exc)})
            continue

        label = "default" if series is None else f"series-{series}"
        path = outdir / f"esake_results_{label}.html"
        path.write_text(body, encoding="utf-8")
        digest = esake.checksum(body)
        (outdir / f"esake_results_{label}.sha256").write_text(digest + "\n", encoding="utf-8")

        try:
            page = esake.parse_results(body, season_start_year=args.season_start_year)
        except Exception as exc:  # noqa: BLE001 - an unreadable page is a result too
            pages.append(
                {
                    "url": url,
                    "series": series,
                    "status": "unparsed",
                    "bytes": len(body.encode("utf-8", "surrogatepass")),
                    "sha256": digest,
                    "saved": str(path),
                    "error": f"{type(exc).__name__}: {exc}",
                }
            )
            continue

        finals.extend(page.finals)
        pages.append(
            {
                "url": url,
                "series": series,
                "status": "parsed",
                "bytes": len(body.encode("utf-8", "surrogatepass")),
                "sha256": digest,
                "saved": str(path),
                "counts": page.counts,
            }
        )

    report = {
        "status": "ok" if any(page["status"] == "parsed" for page in pages) else "failed",
        "provider": esake.PROVIDER,
        "source": esake.SOURCE,
        "requests": requests,
        "pages": pages,
        "finals": len(finals),
        "join": _join(_stored_games(Path(args.warehouse)), finals),
        "note": (
            "No results are written. A final is only usable once the join is "
            "unambiguous; ambiguous minutes are refused, not resolved by name."
        ),
    }

    receipt = outdir / "esake_probe_receipt.json"
    receipt.write_text(json.dumps(report, indent=2, ensure_ascii=False), encoding="utf-8")
    if not args.quiet:
        print(json.dumps(report, indent=2, ensure_ascii=False))
    return 0 if report["status"] == "ok" else 1


if __name__ == "__main__":
    raise SystemExit(main())
