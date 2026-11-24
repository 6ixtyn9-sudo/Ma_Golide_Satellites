"""Render a Ma Golide SHADOW score card for a fixtures-only league.

Input is one JSON document (path given as ``--input``):

    {
      "league": "GREECE_BASKET_LEAGUE",
      "date": "2026-11-24",
      "games": [
        {"home": "Panathinaikos BC", "away": "AEK Athens",
         "scheduled_utc": "2026-11-24T17:15:00Z", "game_id": "opt-1"}
      ],
      "standings": {
        "Panathinaikos BC": {"rank": 1, "pct": 0.82, "net_rtg": 8.1,
                             "home_pct": 0.90, "away_pct": 0.75, "streak": 4}
      }
    }

The standings map is optional and may be partial: missing teams fall back to
the retired script's neutral defaults and score RISKY on purpose.  Output is a
text card under ``--outdir`` (default ``localdata/<league state key>/``),
named ``shadow_scores_<date>.txt``.

Everything the card contains is descriptive shadow evidence: no prices, no
tickets, no settlement, no ROI.  See ``forecasting/magolide_score.py``.
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from forecasting.magolide_score import (  # noqa: E402
    StandingsRow,
    render_shadow_card,
    shadow_lane_leagues,
    shadow_rows,
)
from ingestion.models import Game, League  # noqa: E402


def _league(value: str) -> League:
    try:
        return League[value.upper()]
    except KeyError as exc:
        eligible = ", ".join(spec.league.value for spec in shadow_lane_leagues())
        raise SystemExit(
            f"unknown league {value!r}; shadow-capable lanes: {eligible}"
        ) from exc


def _games(payload: dict, league: League) -> list[Game]:
    games: list[Game] = []
    for i, row in enumerate(payload.get("games") or [], start=1):
        games.append(
            Game(
                game_id=str(row.get("game_id") or f"shadow-{i}"),
                league=league,
                season=str(row.get("season") or "unknown"),
                home_team_id=str(row.get("home") or ""),
                away_team_id=str(row.get("away") or ""),
                home_name=str(row.get("home") or ""),
                away_name=str(row.get("away") or ""),
                scheduled_utc=str(row.get("scheduled_utc") or ""),
            )
        )
    return games


def _standings(payload: dict) -> dict[str, StandingsRow]:
    out: dict[str, StandingsRow] = {}
    for team, row in (payload.get("standings") or {}).items():
        out[team] = StandingsRow(
            team=team,
            rank=row.get("rank"),
            pct=row.get("pct"),
            net_rtg=row.get("net_rtg"),
            home_pct=row.get("home_pct"),
            away_pct=row.get("away_pct"),
            streak=row.get("streak"),
            l10_wins=row.get("l10_wins"),
            l10_losses=row.get("l10_losses"),
        )
    return out


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", required=True, help="fixtures/standings JSON path")
    parser.add_argument(
        "--outdir",
        default=None,
        help="output directory (default: localdata/<league state key>/)",
    )
    args = parser.parse_args()

    payload = json.loads(Path(args.input).read_text())
    league = _league(str(payload.get("league")))
    league_key = next(
        spec.state_key for spec in shadow_lane_leagues() if spec.league is league
    )
    date_label = str(payload.get("date") or "unknown-date")

    rows = shadow_rows(league, _games(payload, league), _standings(payload))
    card = render_shadow_card(league.value, rows, date_label=date_label)

    outdir = (
        Path(args.outdir)
        if args.outdir
        else Path("localdata") / league_key
    )
    outdir.mkdir(parents=True, exist_ok=True)
    target = outdir / f"shadow_scores_{date_label}.txt"
    target.write_text(card)
    print(f"wrote {target} ({len(rows)} shadow row(s))")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
