#!/usr/bin/env python3
"""Which basketball leagues does the priced API actually carry?

The runners cannot answer this: ``SPORT_KEYS`` in
:mod:`ingestion.providers.theoddsapi` maps only the leagues the registry has
already admitted, so a league absent from that map is indistinguishable from a
league the provider does not carry. This script asks the provider, using its
documented ``GET /v4/sports`` catalogue, and prints the answer next to our own
league names.

It never guesses a sport key: it reads the provider's list and matches it
against the registry. Reasoning about which key *should* exist is exactly how
an invented contract gets written, so the output distinguishes "provider lists
it and we map it", "provider lists it and we do not", and "provider does not
list it" -- three different follow-ups.

Credential handling: the key is read from the environment, sent only in the
query string, and never printed or written. The saved body is the response,
not the request.

Usage::

    THE_ODDS_API_KEY=... python scripts/probe_odds_api_coverage.py
    ... --out probes/odds-coverage-2026-11-24

Every request is recorded in ``LEDGER.tsv`` with the same fields as the lane
ledger (timestamp UTC | slug | url | status | content-type | bytes | sha256 |
body | note). Content type is recorded as ``-`` because the client does not
expose headers; that gap is stated rather than filled with a guess.
"""

from __future__ import annotations

import argparse
import hashlib
from datetime import datetime, timezone
import os
from pathlib import Path
import sys
from typing import Any, Dict, Iterable, List, Optional, Tuple

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from ingestion.http import HttpClient  # noqa: E402
from ingestion.league_registry import SPECS  # noqa: E402
from ingestion.models import League  # noqa: E402
from ingestion.providers.theoddsapi import ENV_KEY, SPORT_KEYS  # noqa: E402

CATALOGUE = "https://api.the-odds-api.com/v4/sports/"


def catalogue_url(api_key: str, all_sports: bool = True) -> str:
    params = "?all=true" if all_sports else "?"
    return f"{CATALOGUE}{params}apiKey={api_key}"


def redact_url(url: str) -> str:
    """Keep the request line usable in a ledger without leaking the key."""

    head, _, _ = url.partition("apiKey=")
    return f"{head}apiKey=<redacted>"


def basketball_keys(payload: Any) -> List[Dict[str, Any]]:
    """Rows of the documented catalogue that describe basketball.

    The catalogue is a list of objects carrying ``key``/``group``/``title``/
    ``active``. Anything not shaped that way is skipped rather than coerced.
    """

    if not isinstance(payload, list):
        return []
    rows = []
    for row in payload:
        if not isinstance(row, dict):
            continue
        key = str(row.get("key") or "")
        if not key:
            continue
        group = str(row.get("group") or "")
        if "basketball" in key.lower() or group.strip().lower() == "basketball":
            rows.append(
                {
                    "key": key,
                    "group": group,
                    "title": str(row.get("title") or ""),
                    "active": bool(row.get("active")),
                    "has_outrights": bool(row.get("has_outrights")),
                }
            )
    return sorted(rows, key=lambda row: row["key"])


def _registry_tokens(league: League) -> List[str]:
    value = league.value.lower()
    tokens = [value, value.replace("_", "")]
    return sorted({token for token in tokens if token})


def classify(key: str, mapped: Dict[str, str], leagues: Iterable[League]) -> Tuple[str, Optional[str]]:
    """Return (verdict, league) for one provider key.

    ``mapped`` is the live ``SPORT_KEYS`` map, so a league the provider carries
    but the registry has not admitted is visible as such instead of absent.
    """

    if key in mapped.values():
        league_name = next(name for name, sport_key in mapped.items() if sport_key == key)
        return "already mapped", league_name

    slug = key[len("basketball_"):] if key.startswith("basketball_") else key
    slug_tokens = {slug, slug.replace("_", "")}
    for league in leagues:
        if slug_tokens & set(_registry_tokens(league)):
            return "listed, unmapped", league.value
    return "listed, no league", None


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--out",
        default=None,
        help="evidence directory (default: probes/odds-coverage-<utc date>)",
    )
    args = parser.parse_args(argv)

    api_key = os.environ.get(ENV_KEY, "").strip()
    if not api_key:
        print(f"{ENV_KEY} is not set; nothing probed.", file=sys.stderr)
        return 2

    out_dir = Path(
        args.out
        or f"probes/odds-coverage-{datetime.now(timezone.utc).strftime('%Y-%m-%d')}"
    )
    out_dir.mkdir(parents=True, exist_ok=True)

    url = catalogue_url(api_key)
    client = HttpClient(timeout=20.0, retries=1)
    status = 0
    body = ""
    note = ""
    try:
        body = client.get_text(url)
        status = client.last_status or 200
    except Exception as exc:  # fail closed: record, never invent a catalogue
        status = getattr(exc, "status", None) or 0
        note = type(exc).__name__
    digest = hashlib.sha256(body.encode("utf-8")).hexdigest() if body else "-"
    body_path = out_dir / f"sports-{status}.json"
    if body:
        body_path.write_text(body, encoding="utf-8")

    stamp = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    ledger = out_dir / "LEDGER.tsv"
    new_file = not ledger.exists()
    with ledger.open("a", encoding="utf-8") as handle:
        if new_file:
            handle.write(
                "timestamp_utc\tslug\turl\tstatus\tcontent_type\tbytes\tsha256\tbody\tnode\n"
            )
        handle.write(
            "\t".join(
                [
                    stamp,
                    "odds_catalogue",
                    redact_url(url),
                    str(status),
                    "-",
                    str(len(body.encode("utf-8"))),
                    digest,
                    body_path.name if body else "-",
                    note or "",
                ]
            )
            + "\n"
        )

    print(f"{stamp}  GET {redact_url(url)}  -> {status}, {len(body.encode('utf-8'))} B")
    if client.last_quota:
        print(f"quota headers: {client.last_quota}")
    if not body:
        print("no body: nothing to classify. Do not read this as 'no coverage'.", file=sys.stderr)
        return 1

    import json

    try:
        payload = json.loads(body)
    except ValueError as exc:
        print(f"body is not JSON ({exc}); saved as evidence, not parsed.", file=sys.stderr)
        return 1

    rows = basketball_keys(payload)
    leagues = [spec.league for spec in SPECS]
    print(f"\nbasketball keys in the catalogue: {len(rows)}\n")
    header = f"{'provider key':34} {'active':6} {'verdict':18} league"
    print(header)
    print("-" * len(header))
    for row in rows:
        verdict, league = classify(row["key"], SPORT_KEYS, leagues)
        print(
            f"{row['key']:34} {str(row['active']):6} {verdict:18} "
            f"{league or '-'}  {row['title'][:34]}"
        )

    print(
        "\n'listed, unmapped' means the provider carries the league and the "
        "registry has not admitted it: that needs a credentialed odds smoke and "
        "a price policy, not an adapter guess.\n"
        "'listed, no league' means a basketball competition we do not trade.\n"
        "A league missing from this table is the only case that reads as "
        "'provider does not carry it'."
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
