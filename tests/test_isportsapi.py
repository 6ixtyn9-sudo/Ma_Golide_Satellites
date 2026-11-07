"""Hermetic iSportsAPI Chinese CBA contract tests (no network)."""

from __future__ import annotations

from copy import deepcopy

import pytest

from ingestion.models import GameStatus, League, Market, Period, PeriodKind
from ingestion.providers import IdentityError
from ingestion.providers import isportsapi

OBSERVED = "2026-11-07T08:00:00Z"


def _row(match_id="100", *, status=-1):
    return {
        "matchId": match_id,
        "leagueId": "321",
        "leagueName": "CBA",
        "quarterCount": 4,
        "matchTime": 1893456000 if status == 0 else 1794062100,
        "status": status,
        "homeId": "10",
        "homeName": "Beijing Ducks",
        "awayId": "20",
        "awayName": "Shanghai Sharks",
        "homeScore": 110,
        "awayScore": 105,
        "homeFirstQuarterScore": 25,
        "awayFirstQuarterScore": 22,
        "homeSecondQuarterScore": 30,
        "awaySecondQuarterScore": 24,
        "homeThirdQuarterScore": 28,
        "awayThirdQuarterScore": 30,
        "homeFourthQuarterScore": 27,
        "awayFourthQuarterScore": 29,
        "overTimeCount": 0,
        "leagueSeason": "26-27",
        "matchType": 1,
    }


def _payloads():
    return {
        "league/basic": {"code": 0, "data": [
            {"leagueId": "111", "leagueName": "National Basketball Association", "leagueShortName": "NBA"},
            {"leagueId": "321", "leagueName": "Chinese Basketball Association", "leagueShortName": "CBA"},
            {"leagueId": "999", "leagueName": "CBA Club Cup", "leagueShortName": "CBACUP"},
        ]},
        "schedule": {"code": 0, "data": [_row("100"), _row("101", status=0)]},
        "bookmaker": {"code": 0, "data": [
            {"companyIdEu": 3, "companyName": "Interwetten", "companyIdMainSpread": 0, "companyIdMainTotal": 0},
        ]},
        "odds/fulltime": {"code": 0, "data": {
            "moneyLine": ["101,3,3.00,1.34,3.20,1.23"],
            "spread": ["101,8,-7,0.90,0.90,-6,0.90,0.90,-99,9.00,9.00"],
            "total": ["101,11,158.5,0.90,0.90,156,0.90,0.90,999,9.00,9.00"],
        }},
        "odds/half": {"code": 0, "data": {
            "spread": ["101,8,-4,0.90,0.90,-3.5,0.86,0.86"],
            "total": ["101,11,78.5,0.90,0.90,77,0.83,0.90"],
        }},
        "odds/quarter": {"code": 0, "data": {
            "spread": [",".join(
                ["101", "8", "-1.5", "0.80", "0.90"]
                + [""] * 9
                + [""] * 3 + ["-2", "0.85", "0.85"] + [""] * 6
            )],
            "total": [],
        }},
    }


def _fetch(payloads):
    def fetch(url):
        for path, payload in payloads.items():
            if f"/{path}" in url or (path == "schedule" and "/schedule?" in url):
                return deepcopy(payload)
        raise AssertionError(f"unexpected URL {url}")
    return fetch


def _games():
    return isportsapi.discover_games("20261107", _fetch(_payloads()))


def test_discovers_exact_cba_and_reconciles_quarters():
    games = _games()
    assert [game.game_id for game in games] == ["100", "101"]
    final = games[0]
    assert final.league is League.CBA
    assert final.status is GameStatus.FINAL
    assert final.actual.period_kind is PeriodKind.QUARTERS
    assert final.actual.reconciles() == (True, "")
    assert final.provider_ids == {"isportsapi": "100"}
    assert games[1].actual is None


def test_overtime_is_summed_without_touching_regulation_quarters():
    payloads = _payloads()
    row = payloads["schedule"]["data"][0]
    row.update({
        "homeScore": 130, "awayScore": 125, "overTimeCount": 2,
        "homeFirstOverTimeScore": 9, "awayFirstOverTimeScore": 8,
        "homeSecondOverTimeScore": 11, "awaySecondOverTimeScore": 12,
    })
    game = isportsapi.fetch_recent_results("20261107", _fetch(payloads))[0]
    assert game.actual.home_overtime == 20
    assert game.actual.away_overtime == 20
    assert game.actual.overtime_periods == 2
    assert game.final_status().value == "AOT"
    assert game.actual.reconciles() == (True, "")


def test_incomplete_final_is_gapped_not_backfilled():
    payloads = _payloads()
    payloads["schedule"]["data"][0]["homeThirdQuarterScore"] = None
    games = _games_with_payloads(payloads)
    assert [game.game_id for game in games] == ["101"]
    assert isportsapi.last_gaps()[0].required_field == "schedule_or_result"


def _games_with_payloads(payloads):
    return isportsapi.discover_games("20261107", _fetch(payloads))


def test_ambiguous_cba_identity_fails_closed():
    payloads = _payloads()
    payloads["league/basic"]["data"].append(
        {"leagueId": "322", "leagueName": "Chinese Basketball Association", "leagueShortName": "CBA"}
    )
    with pytest.raises(IdentityError, match="refusing to guess"):
        _games_with_payloads(payloads)


def test_other_leagues_are_rejected():
    with pytest.raises(ValueError, match="only for Chinese CBA"):
        isportsapi.discover_games("20261107", _fetch(_payloads()), League.NBA)


def test_result_identity_is_checked():
    games = _games()
    scheduled = games[1]
    with pytest.raises(IdentityError, match="not a complete final"):
        isportsapi.fetch_game_result("101", _fetch(_payloads()), scheduled)
    final = isportsapi.fetch_game_result("100", _fetch(_payloads()))
    assert final.actual.reconciles()[0]


def test_book_prices_use_prematch_fields_and_native_periods(monkeypatch):
    monkeypatch.setenv("ISPORTS_API_KEY", "fixture-key")
    snapshots = isportsapi.fetch_odds(
        "CBA", "20261107", lambda _provider: _fetch(_payloads()), now_utc=OBSERVED
    )
    assert len(snapshots) == 7
    by_market_period = {(s.market, s.period): s for s in snapshots}
    moneyline = by_market_period[Market.MONEYLINE, Period.FULL_GAME]
    assert moneyline.line is None
    assert moneyline.book == "Interwetten"  # companyIdEu 3, not spread ID 3 (Crown)
    assert moneyline.home_price_decimal == pytest.approx(4.20)
    assert moneyline.away_price_decimal == pytest.approx(2.23)
    spread = by_market_period[Market.SPREAD, Period.FULL_GAME]
    assert spread.line == -6
    assert spread.home_price_decimal == pytest.approx(1.90)
    total = by_market_period[Market.TOTAL, Period.FULL_GAME]
    assert total.line == 156  # not the 999 in-play field
    assert by_market_period[Market.SPREAD, Period.H1].line == -3.5
    assert by_market_period[Market.SPREAD, Period.Q1].line == -1.5
    assert by_market_period[Market.SPREAD, Period.Q2].line == -2  # opening fallback
    assert all(snapshot.pre_match for snapshot in snapshots)


def test_moneyline_directory_failure_does_not_erase_other_books(monkeypatch):
    monkeypatch.setenv("ISPORTS_API_KEY", "fixture-key")
    payloads = _payloads()
    payloads["bookmaker"] = {"code": 401, "message": "unauthorised", "data": []}
    snapshots = isportsapi.fetch_odds(
        "CBA", "20261107", lambda _provider: _fetch(payloads), now_utc=OBSERVED
    )
    assert len(snapshots) == 6
    assert not any(snapshot.market is Market.MONEYLINE for snapshot in snapshots)
    assert any(gap.required_field == "moneyline_bookmakers" for gap in isportsapi.last_gaps())


def test_schedule_smoke_without_credentials_is_explicit_and_read_only(monkeypatch):
    monkeypatch.delenv("ISPORTS_API_KEY", raising=False)
    result = isportsapi.schedule_smoke(
        "20261107", lambda *_args, **_kwargs: pytest.fail("network"), League.CBA
    )
    assert result["verification"] == "UNAUTHORIZED"
    assert result["enabled"] is False


def test_provider_error_is_not_treated_as_empty_coverage():
    payloads = _payloads()
    payloads["schedule"] = {"code": 401, "message": "unauthorised", "data": []}
    with pytest.raises(IdentityError, match="provider code"):
        _games_with_payloads(payloads)
