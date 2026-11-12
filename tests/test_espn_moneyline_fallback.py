"""ESPN legacy flat moneyline odds must never cross home/away sides."""

from __future__ import annotations

import pytest

from ingestion.models import Game, League, Market
from ingestion.providers.espn import book_snapshots_from_event

GAME_ID = "401909978"
OBSERVED_UTC = "2026-11-12T10:56:42Z"


def _game() -> Game:
    return Game(
        game_id=GAME_ID,
        league=League.NBA,
        season="2027",
        home_team_id="8",
        away_team_id="9",
        home_name="Minnesota Timberwolves",
        away_name="Dallas Mavericks",
        scheduled_utc="2026-11-13T01:00:00Z",
        provider_ids={"espn": GAME_ID},
    )


def _moneyline_snapshots(offer: dict) -> list:
    event = {
        "id": GAME_ID,
        "competitions": [
            {
                "status": {"type": {"state": "pre"}},
                "odds": [offer],
            }
        ],
    }
    return [
        snapshot
        for snapshot in book_snapshots_from_event(event, _game(), observed_utc=OBSERVED_UTC)
        if snapshot.market is Market.MONEYLINE
    ]


def test_home_flat_moneyline_is_not_reused_as_away_price():
    offer = {
        "provider": {"id": "1006", "displayName": "Draft Kings"},
        "homeTeamOdds": {"moneyLine": -216},
    }

    (snapshot,) = _moneyline_snapshots(offer)

    assert snapshot.home_price_decimal == pytest.approx(1 + 100 / 216)
    assert snapshot.away_price_decimal is None


def test_away_flat_moneyline_fallback_still_uses_away_price():
    offer = {
        "provider": {"id": "1006", "displayName": "Draft Kings"},
        "homeTeamOdds": {"moneyLine": -216},
        "awayTeamOdds": {"moneyLine": 179},
    }

    (snapshot,) = _moneyline_snapshots(offer)

    assert snapshot.home_price_decimal == pytest.approx(1 + 100 / 216)
    assert snapshot.away_price_decimal == pytest.approx(1 + 179 / 100)
