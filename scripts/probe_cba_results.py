"""Find the round space of CBA's data hub, and record what each round answers.

Why this exists
---------------
``CHINA_WCBA_W``'s documented results URL was read twice and the two reads
disagree. On 2026-11-24 an external probe got the whole 2026-27 season from
``cbamatch/games/matchschedules?competitionId=56&seasonId=2286`` (297,504 B,
integer ``gameId``, scores, ``gameStatus``). Re-read from an egress host at
2026-11-24T14:02Z the identical URL answers::

    HTTP 500, 179 B
    {"error":"Internal Server Error","message":"param name: roundId value not exists"}

So the endpoint now demands a ``roundId``. That is the whole of what the error
says: it names the missing parameter and says nothing about the values it
accepts. Guessing a single ``roundId`` would be invention, so this probe
enumerates a **bounded** range, records every answer verbatim, and reports which
rounds returned a body that actually carries games.

It writes no results. A body that was read is evidence; a round that answered
nothing is a named failure. Nothing here is trusted as a score until a parser has
been written against a saved body and replayed against it.

Politeness
----------
One request per round, spaced by ``--delay`` seconds (default 2). The host
publishes no crawl rules of its own, so the default is deliberately slower than
it needs to be. ``--rounds`` caps the total, and the cap is enforced.

Usage
-----
    python scripts/probe_cba_results.py --rounds 1-40
    python scripts/probe_cba_results.py --rounds 1-10 --delay 5 --quiet

Then hand back the saved bodies and ``cba_probe_receipt.json`` so the parser can
be written against served bytes.
"""

from __future__ import annotations

import argparse
import hashlib
import json
from datetime import datetime, timezone
from pathlib import Path
import sys
import time

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import requests  # noqa: E402

HOST = "https://www.cba.net.cn"
PATH = "/datahub/cbamatch/games/matchschedules"
COMPETITION_ID = 56
SEASON_ID = 2286
#: The host publishes no robots.txt for this path; an absent file is an allow
#: state under RFC 9309 and is reported as absence, never as permission.
REQUEST_BUDGET = 40

HEADERS = {
    "Accept": "application/json",
    "User-Agent": "MaGolide/1.0 (paper-only record; bounded probe)",
}


def rounds_url(round_id: int, *, competition_id: int = COMPETITION_ID,
               season_id: int = SEASON_ID) -> str:
    """The one documented URL with the parameter the host says it needs."""

    return (
        f"{HOST}{PATH}?competitionId={int(competition_id)}"
        f"&seasonId={int(season_id)}&roundId={int(round_id)}"
    )


def parse_rounds(spec: str) -> list[int]:
    """``1-40`` or ``1,3,5`` or a single value; always bounded and ordered."""

    rounds: list[int] = []
    for part in str(spec).replace(" ", "").split(","):
        if not part:
            continue
        if "-" in part[1:]:
            start, _, end = part.partition("-")
            rounds.extend(range(int(start), int(end) + 1))
        else:
            rounds.append(int(part))
    unique = sorted({round_id for round_id in rounds if round_id > 0})
    if not unique:
        raise ValueError(f"no rounds in {spec!r}")
    if len(unique) > REQUEST_BUDGET:
        raise ValueError(
            f"{len(unique)} rounds exceeds the bounded budget of {REQUEST_BUDGET}"
        )
    return unique


def summarise(body: str) -> dict:
    """What this body is, decided from the body alone -- never from a guess."""

    text = body or ""
    result = {
        "bytes": len(text.encode("utf-8")),
        "sha256": hashlib.sha256(text.encode("utf-8")).hexdigest(),
        "looks_like_json": False,
        "game_count": 0,
        "error": "",
        "head": text[:200],
    }
    try:
        payload = json.loads(text)
    except ValueError:
        return result
    result["looks_like_json"] = True
    if not isinstance(payload, dict):
        return result
    if payload.get("error"):
        result["error"] = str(payload.get("message") or payload.get("error"))[:200]
        return result
    result["game_count"] = _largest_game_list(payload, depth=0)
    return result


def _largest_game_list(node: object, *, depth: int) -> int:
    """The biggest list of objects anywhere in the body, without naming a key.

    The host's envelope is not known: the 2026-11-24 read was a 500, so ``data``,
    ``list``, ``rows`` and ``result`` are all guesses. Counting the largest list
    of objects is a measurement of the body rather than an assumption about it,
    which is the whole point of a probe.
    """

    if depth > 4:
        return 0
    if isinstance(node, list):
        if node and all(isinstance(item, dict) for item in node):
            return len(node)
        return max((_largest_game_list(item, depth=depth + 1) for item in node), default=0)
    if isinstance(node, dict):
        return max(
            (_largest_game_list(value, depth=depth + 1) for value in node.values()),
            default=0,
        )
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--rounds", default="1-40", help="e.g. 1-40 or 1,3,7")
    parser.add_argument("--competition-id", type=int, default=COMPETITION_ID)
    parser.add_argument("--season-id", type=int, default=SEASON_ID)
    parser.add_argument("--outdir", default="/tmp/cba", help="where bodies are saved")
    parser.add_argument("--delay", type=float, default=2.0, help="seconds between requests")
    parser.add_argument("--timeout", type=float, default=30.0)
    parser.add_argument("--quiet", action="store_true")
    args = parser.parse_args(argv)

    try:
        rounds = parse_rounds(args.rounds)
    except ValueError as exc:
        print(f"refused: {exc}", file=sys.stderr)
        return 2

    outdir = Path(args.outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    receipt = {
        "provider": "cba_datahub",
        "league": "CHINA_WCBA_W",
        "path": PATH,
        "competition_id": args.competition_id,
        "season_id": args.season_id,
        "started_utc": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "client": HEADERS["User-Agent"],
        "requests": [],
    }

    with requests.Session() as session:
        session.headers.update(HEADERS)
        for index, round_id in enumerate(rounds):
            if index:
                time.sleep(max(args.delay, 0.0))
            url = rounds_url(
                round_id,
                competition_id=args.competition_id,
                season_id=args.season_id,
            )
            record: dict = {"round_id": round_id, "url": url}
            try:
                response = session.get(url, timeout=args.timeout)
                record["status"] = int(response.status_code)
                record["content_type"] = response.headers.get("Content-Type", "")
                body = response.text
                record.update(summarise(body))
                (outdir / f"round-{round_id:03d}.json").write_text(body, encoding="utf-8")
            except requests.RequestException as exc:  # transport failure is a datum
                record["status"] = None
                record["error"] = f"{type(exc).__name__}: {exc}"[:200]
            receipt["requests"].append(record)
            if not args.quiet:
                print(
                    f"round {round_id:3d}  status={record.get('status')}  "
                    f"bytes={record.get('bytes', 0):>7}  games={record.get('game_count', 0):>4}"
                    f"  {record.get('error', '')[:60]}"
                )

    receipt["finished_utc"] = datetime.now(timezone.utc).isoformat(timespec="seconds")
    receipt["requests_made"] = len(receipt["requests"])
    receipt["rounds_with_games"] = [
        item["round_id"] for item in receipt["requests"] if item.get("game_count")
    ]
    receipt_path = outdir / "cba_probe_receipt.json"
    receipt_path.write_text(json.dumps(receipt, indent=2), encoding="utf-8")

    print(f"\nrounds scanned: {receipt['requests_made']}")
    print(f"rounds with games: {receipt['rounds_with_games'] or 'none'}")
    print(f"receipt: {receipt_path}")
    return 0 if receipt["rounds_with_games"] else 1


if __name__ == "__main__":
    sys.exit(main())
