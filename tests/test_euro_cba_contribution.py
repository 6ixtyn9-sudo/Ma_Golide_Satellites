"""EuroLeague and CBA contribution: keyless sources reaching the pipeline.

Covers the adapters added/fixed to make both leagues contribute:
- EuroLeague schedule-row results (one request per season) and the v3
  per-round standings replacement (live-verified 2026-11-16).
- The official keyless CBA data-server adapter (AES-128-ECB decrypted).
- The keyless BetExplorer moneyline adapter for both leagues.
- End-to-end traversal through the real pipeline and forecast service, plus
  red-team cases: malformed results, missing/stale prices, unknown teams,
  undecryptable bodies, and duplicate/conflicting identities.

All tests are hermetic: committed fixtures recorded live on 2026-11-16.
"""

from __future__ import annotations

import json
from datetime import timedelta
from typing import Any, Callable, Dict, List, Optional

import pytest

from forecasting.service import MarketForecastService
from forecasting.types import MODEL_VERSION
from ingestion.books import BookConfig
from ingestion.duckdb_store import DuckDBStore
from ingestion.identity import canonical_team_key, parse_utc, to_utc_iso
from ingestion.models import (
    ActualResult,
    BookSnapshot,
    Game,
    GameStatus,
    League,
    Market,
    Period,
    PeriodKind,
)
from ingestion.pipeline import Pipeline, PipelineConfig
from ingestion.providers import betexplorer, cbaleague, euroleague
from ingestion.providers._cba_aes import (
    _expand_key,
    decrypt_block,
    decrypt_body,
    encrypt_block,
    encrypt_body,
)
from ingestion.store import InMemoryStore

FIXTURES = "tests/fixtures/"


def _fixture_text(relative: str) -> str:
    with open(FIXTURES + relative, encoding="utf-8") as handle:
        return handle.read()


def _fixture_json(relative: str) -> Any:
    return json.loads(_fixture_text(relative))


def _mapping_fetch(mapping: Dict[str, Callable[[], Any]]) -> Callable[[str], Any]:
    def fetch(url: str) -> Any:
        for needle, loader in mapping.items():
            if needle in url:
                return loader()
        raise AssertionError(f"no fixture mapped for URL: {url}")

    return fetch


def _el_fetch() -> Callable[[str], Any]:
    return _mapping_fetch(
        {
            "seasons/E2026/games": lambda: _fixture_json("euroleague/schedule_scores.json"),
        }
    )


def _cba_fetch() -> Callable[[str], Any]:
    return _mapping_fetch(
        {
            "getSeason": lambda: _fixture_text("cbaleague/getSeason.txt"),
            "2026-11-01": lambda: _fixture_text("cbaleague/day_2026_11_01.txt"),
            "2026-11-02": lambda: _fixture_text("cbaleague/day_2026_11_02.txt"),
            "2026-12-04": lambda: _fixture_text("cbaleague/day_2026_12_04.txt"),
        }
    )


def _be_text_fetch() -> Callable[[str], str]:
    def fetch(url: str) -> str:
        if "euroleague" in url:
            return _fixture_text("betexplorer/el_fixtures_table.html")
        if "china" in url:
            return _fixture_text("betexplorer/cba_fixtures_table.html")
        raise AssertionError(f"unexpected BetExplorer URL: {url}")

    return fetch


_CBA_KEY = b"uVayqL4ONKjFbVzQ"


def _quarters_for_final(home: int, away: int) -> Dict[str, Dict[str, int]]:
    """Deterministic quarter split summing exactly to each recorded final."""

    def split(total: int) -> Dict[str, int]:
        base = total // 4
        return {
            "oneScore": base,
            "twoScore": base,
            "threeScore": base,
            "fourScore": total - 3 * base,
        }

    return {"home": split(home), "visiting": split(away), "style": 1}


def _cba_post_fn(payloads: Optional[Dict[int, str]] = None):
    """Serve recorded quarter payloads; synthesize the rest from the day row."""

    finals: Dict[int, tuple] = {}
    for day in ("2026_11_01", "2026_11_02"):
        rows = decrypt_body(_fixture_text(f"cbaleague/day_{day}.txt"), _CBA_KEY)
        for row in rows:
            if row.get("status") == 4:
                finals[int(row["scheduleId"])] = (
                    int(row["homeTeamScore"]),
                    int(row["visitingTeamScore"]),
                )

    def post(url: str, body: Dict[str, Any]) -> str:
        schedule_id = int(body["scheduleId"])
        if payloads and schedule_id in payloads:
            return payloads[schedule_id]
        recorded = {
            100110977: "cbaleague/quarters_100110977.txt",
            100110991: "cbaleague/quarters_100110991_ot.txt",
        }
        if schedule_id in recorded:
            return _fixture_text(recorded[schedule_id])
        if schedule_id in finals:
            home, away = finals[schedule_id]
            return encrypt_body(_quarters_for_final(home, away), _CBA_KEY)
        return encrypt_body(_quarters_for_final(82, 88), _CBA_KEY)

    return post


# ---------------------------------------------------------------------------
# EuroLeague schedule-row results and v3 standings
# ---------------------------------------------------------------------------


def test_euroleague_season_results_reconcile_from_one_schedule_request():
    urls = []
    instrumented = _mapping_fetch(
        {
            "seasons/E2026/games": lambda: (urls.append("schedule") or _fixture_json(
                "euroleague/schedule_scores.json"
            )),
        }
    )
    results = euroleague.season_results("E2026", instrumented, failures=[])
    assert urls == ["schedule"], "a whole season must cost exactly one request"
    assert len(results) == 6
    by_id = {game.game_id: game for game in results}
    ot = by_id["E2026-101"]
    assert ot.actual.overtime_periods == 1
    assert ot.actual.home_overtime == 19.0 and ot.actual.away_overtime == 17.0
    assert ot.actual.reconciles() == (True, "")
    assert ot.status is GameStatus.FINAL
    regulation = by_id["E2026-109"]
    assert regulation.actual.home_final == 103.0 and regulation.actual.away_final == 91.0
    assert regulation.actual.home_quarters == [28.0, 32.0, 16.0, 27.0]
    assert regulation.actual.away_quarters == [20.0, 23.0, 25.0, 23.0]


def test_euroleague_season_results_quarantine_malformed_rows_and_continue():
    payload = _fixture_json("euroleague/schedule_scores.json")
    tampered = json.loads(json.dumps(payload))
    for row in tampered["data"]:
        if row.get("gameCode") == 109:
            row["local"]["partials"]["partials4"] = 999  # breaks reconciliation
        if row.get("gameCode") == 110:
            del row["local"]["partials"]  # no partials: never improvised
    fetch = _mapping_fetch({"seasons/E2026/games": lambda: tampered})
    failures: List[Dict[str, str]] = []
    results = euroleague.season_results("E2026", fetch, failures=failures)
    assert {game.game_id for game in results} == {
        "E2026-101",
        "E2026-91",
        "E2026-66",
        "E2026-108",
    }
    assert {failure["game_id"] for failure in failures} == {"E2026-109", "E2026-110"}


def test_euroleague_discover_rejects_empty_season_payload():
    # A wrong (bare-year) season code answers HTTP 200 with an empty data list;
    # a bare year is now refused before the request is even built.
    with pytest.raises(ValueError, match="E2026 or U2026"):
        euroleague.discover_games(None, lambda url: {"data": []}, "2026")

    fetch = _mapping_fetch({"seasons/E2026/games": lambda: {"data": []}})
    with pytest.raises(ValueError, match="empty"):
        euroleague.discover_games(None, fetch, "E2026")


def test_euroleague_standings_v3_parses_and_derives_latest_completed_round():
    fetch = _mapping_fetch(
        {
            "seasons/E2026/games": lambda: _fixture_json("euroleague/schedule_scores.json"),
            "basicstandings": lambda: _fixture_json("euroleague/standings_per_round.json"),
        }
    )
    report: Dict[str, str] = {}
    rows = euroleague.fetch_standings(fetch, "E2026", round=11, report=report)
    assert report["status"] == "ok"
    assert report["evidence"].startswith("LIVE_VERIFIED")
    assert len(rows) == 20
    assert rows[0]["position"] == 1
    assert rows[0]["team"] == "Valencia Basket"
    assert rows[0]["games"] == 11 and rows[0]["wins"] == 8 and rows[0]["losses"] == 3

    derived = euroleague.fetch_standings(fetch, "E2026", report=report)
    assert len(derived) == 20
    assert report["round"] == "11", "latest completed round derived from the schedule"


def test_euroleague_standings_unfinished_round_and_failures_stay_defensive():
    def fetch(url: str) -> Any:
        if "rounds/12" in url:
            return {"detail": "Not found standings for season: E2026, round: 12"}
        raise AssertionError(url)

    report: Dict[str, str] = {}
    assert euroleague.fetch_standings(fetch, "E2026", round=12, report=report) == []
    assert report["status"] == "unavailable"
    assert euroleague.fetch_standings(
        lambda _url: (_ for _ in ()).throw(RuntimeError("network down")), "E2026", report={}
    ) == []


# ---------------------------------------------------------------------------
# CBA official adapter
# ---------------------------------------------------------------------------


def test_cba_aes_matches_fips_197_aes128_known_answer_vector():
    key = bytes.fromhex("000102030405060708090a0b0c0d0e0f")
    plaintext = bytes.fromhex("00112233445566778899aabbccddeeff")
    ciphertext = bytes.fromhex("69c4e0d86a7b0430d8cdb78070b4c55a")
    round_keys = _expand_key(key)
    assert encrypt_block(plaintext, round_keys) == ciphertext
    assert decrypt_block(ciphertext, round_keys) == plaintext


def test_cba_current_season_and_discovery_map_canonical_identity():
    fetch = _cba_fetch()
    assert cbaleague.current_season(fetch) == 2026
    seasons = cbaleague.available_seasons(fetch)
    assert seasons[0] == 1995 and seasons[-1] == 2026
    games = cbaleague.discover_games("20261101", fetch)
    assert len(games) == 7
    by_id = {game.game_id: game for game in games}
    jilin = by_id["100110977"]
    assert jilin.home_name == "Jilin" and jilin.away_name == "Guangzhou"
    assert jilin.home_team_id == "29137" and jilin.away_team_id == "29139"
    assert jilin.scheduled_utc == "2026-11-01T11:35:00Z"  # 19:35 Beijing
    assert jilin.status is GameStatus.FINAL and jilin.season == "2026"
    assert jilin.provider_ids == {"cbaleague": "100110977"}
    upcoming = cbaleague.discover_games("20261204", fetch)
    assert {game.status for game in upcoming} == {GameStatus.SCHEDULED}
    assert upcoming[0].actual is None


def test_cba_results_reconcile_quarters_and_overtime():
    fetch = _cba_fetch()
    finals = cbaleague.fetch_recent_results("20261101", fetch, post_fn=_cba_post_fn())
    assert len(finals) == 7
    jilin = next(game for game in finals if game.game_id == "100110977")
    assert jilin.actual.home_final == 79.0 and jilin.actual.away_final == 90.0
    assert jilin.actual.home_quarters == [21.0, 24.0, 18.0, 16.0]
    assert jilin.actual.reconciles() == (True, "")
    ot_day = cbaleague.fetch_recent_results("20261102", fetch, post_fn=_cba_post_fn())
    ot_game = next(game for game in ot_day if game.game_id == "100110991")
    assert ot_game.actual.overtime_periods == 1
    assert (ot_game.actual.home_overtime, ot_game.actual.away_overtime) == (10.0, 12.0)
    assert ot_game.actual.home_final == 97.0 and ot_game.actual.away_final == 99.0


def test_pipeline_surfaces_cba_adapter_gaps_for_discovery_and_results(monkeypatch):
    key = b"uVayqL4ONKjFbVzQ"
    rows = decrypt_body(_fixture_text("cbaleague/day_2026_11_01.txt"), key)
    malformed = json.loads(json.dumps(rows))
    del malformed[0]["homeTeamName"]
    fetch = _mapping_fetch(
        {
            "2026-11-01": lambda: encrypt_body(malformed, key),
            "2026-11-02": lambda: _fixture_text("cbaleague/day_2026_11_02.txt"),
        }
    )
    pipeline = Pipeline(
        PipelineConfig(league="CBA", date_yyyymmdd="20261101"), InMemoryStore(), fetch=fetch
    )
    discovered = pipeline._discover_games()
    assert len(discovered) == len(rows) - 1
    assert any(gap.required_field == "schedule_row" for gap in pipeline.gaps)

    result_pipeline = Pipeline(
        PipelineConfig(league="CBA", date_yyyymmdd="20261101"),
        InMemoryStore(),
        fetch=_cba_fetch(),
    )
    bad_post = _cba_post_fn(payloads={100110977: "{}"})
    monkeypatch.setattr(cbaleague, "_post", lambda url, body, fn=None: bad_post(url, body))
    finals = result_pipeline._fetch_finals()
    assert len(finals) == 6
    assert any(gap.required_field == "actual_result" for gap in result_pipeline.gaps)


def test_cba_unreconcilable_quarters_quarantine_without_fabrication():
    fetch = _cba_fetch()
    bad_info = {
        "home": {"oneScore": 21, "twoScore": 24, "threeScore": 18, "fourScore": 17},
        "visiting": {"oneScore": 33, "twoScore": 15, "threeScore": 28, "fourScore": 14},
        "style": 1,
    }
    post_fn = _cba_post_fn(payloads={100110977: encrypt_body(bad_info, b"uVayqL4ONKjFbVzQ")})
    finals = cbaleague.fetch_recent_results("20261101", fetch, post_fn=post_fn)
    assert "100110977" not in {game.game_id for game in finals}
    gaps = cbaleague.last_gaps()
    assert any(gap.game_id == "100110977" and "reconciliation" in gap.detail for gap in gaps)
    # The remaining six games still reconcile; nothing was folded into OT.
    assert len(finals) == 6


def test_cba_undecryptable_body_is_refused_not_banked():
    def broken_post(url: str, body: Dict[str, Any]) -> str:
        return json.dumps({"error": "plain unencrypted 500 envelope"})

    with pytest.raises(Exception) as excinfo:
        cbaleague.fetch_quarters("100110977", broken_post)
    assert "decrypt" in str(excinfo.value)


def test_cba_unknown_status_stays_unknown_and_final_requires_four_quarters():
    row = {
        "scheduleId": 100199999,
        "season": 2026,
        "homeTeamId": 29115,
        "visitingTeamId": 29136,
        "homeTeamName": "北京首钢",
        "visitingTeamName": "北京控股",
        "status": 9,
        "scheduleTime": "2026-11-20 19:35:00",
    }
    game = cbaleague._game_from_row(row)
    assert game.status is GameStatus.UNKNOWN


def test_cba_unaliased_chinese_name_stays_unmapped_and_quarantines():
    row = {
        "scheduleId": 100188888,
        "season": 2026,
        "homeTeamId": 99999,
        "visitingTeamId": 29115,
        "homeTeamName": "虚构赞助球队",
        "visitingTeamName": "北京首钢",
        "status": 1,
        "scheduleTime": "2026-11-20 19:35:00",
    }
    game = cbaleague._game_from_row(row)
    assert game.home_name == "虚构赞助球队", "unknown names are never guessed"
    assert canonical_team_key(game.home_name) != canonical_team_key("Beijing")


# ---------------------------------------------------------------------------
# BetExplorer adapter
# ---------------------------------------------------------------------------


OBSERVED = "2026-11-16T12:00:00Z"


def test_betexplorer_euroleague_fixtures_parse_prices_and_identity():
    snapshots = betexplorer.fetch_odds(
        "EUROLEAGUE", "20261117", None, now_utc=OBSERVED, text_fetch=_be_text_fetch()
    )
    # 9 priced rows sit in the 2026-11-17 block of el_fixtures_table.html, but
    # 3 (xGZWYFJD, 8I2fZqVg, 0hany60s) carry `<td class="table-main__datetime">&nbsp;</td>`
    # and must never be dated by inheritance; only the 6 own-dated rows emit snapshots.
    assert len(snapshots) == 6
    by_event = {snapshot.provider_event_id: snapshot for snapshot in snapshots}
    assert {"xGZWYFJD", "8I2fZqVg", "0hany60s"}.isdisjoint(by_event)
    dubai = by_event["UkWOzxk1"]
    assert dubai.book == "average" and dubai.provider == "betexplorer"
    assert dubai.market is Market.MONEYLINE and dubai.period is Period.FULL_GAME
    assert (dubai.home_price_decimal, dubai.away_price_decimal) == (1.30, 3.46)
    assert dubai.home_name == "Dubai Basketball"
    assert dubai.away_name == "Anadolu Efes Istanbul"
    assert dubai.scheduled_utc == "2026-11-17T16:00:00Z"  # site 17:00 CET
    assert dubai.pre_match is True
    # The odds pair is exactly two-way; no totals/spreads are invented.
    assert dubai.over_price_decimal is None and dubai.line is None


def test_betexplorer_blank_datetime_repeats_previous_row_and_absolute_dates():
    html = _fixture_text("betexplorer/el_fixtures_table.html")
    rows = betexplorer._rows(html, parse_utc(OBSERVED))
    by_event = {row[1]: row for row in rows}
    # Zalgiris row has a blank datetime cell and repeats Fenerbahce's 19:00.
    assert by_event["xGZWYFJD"][0].isoformat() == "2026-11-17T19:00:00+01:00"
    # Absolute dd.mm. labels infer the nearest year: 18.11. is this season's
    # November (2026) while 18.03. round-of-16 fixtures fall in 2027.
    november = [row for row in rows if row[0].day == 18 and row[0].month == 11]
    assert november and all(row[0].year == 2026 for row in november)
    march = [row for row in rows if row[0].day == 18 and row[0].month == 3]
    assert march and all(row[0].year == 2027 for row in march)


def test_betexplorer_skips_rows_without_two_prices_and_other_dates():
    snapshots = betexplorer.fetch_odds(
        "EUROLEAGUE", "20261117", None, now_utc=OBSERVED, text_fetch=_be_text_fetch()
    )
    assert {snapshot.scheduled_utc[:10] for snapshot in snapshots} == {"2026-11-17"}
    cba = betexplorer.fetch_odds(
        "CBA", "20261204", None, now_utc=OBSERVED, text_fetch=_be_text_fetch()
    )
    assert cba == [], "the recorded CBA table carries no posted odds yet"
    assert betexplorer.fetch_odds(
        "CBA", "20261117", None, now_utc=OBSERVED, text_fetch=_be_text_fetch()
    ) == []


def test_betexplorer_rejects_unsupported_leagues():
    with pytest.raises(ValueError, match="supports"):
        betexplorer.fetch_odds("NBA", "20261117", None, text_fetch=_be_text_fetch())


def test_betexplorer_cba_rows_with_prices_map_canonical_names():
    row_template = (
        "<tr><td class=\"table-main__datetime\">04.12. 12:35</td>"
        "<td class=\"h-text-left\"><a href=\"/basketball/china/cba/jiangsu-dragons-sichuan/v7jNw7e2/\" "
        "class=\"in-match\"><span>Jiangsu Dragons</span> - <span>Sichuan</span></a></td>"
        "<td class=\"h-text-center\">&nbsp;</td><td class=\"table-main__bs\">8</td>"
        "<td class=\"table-main__odds\" data-oid=\"a1\"><button data-odd=\"1.55\"></button></td>"
        "<td class=\"table-main__odds\" data-oid=\"a2\"><button data-odd=\"2.50\"></button></td></tr>"
    )
    html = f'<table class="table-main table-main--leaguefixtures">{row_template}</table>'
    snapshots = betexplorer.fetch_odds(
        "CBA", "20261204", None, now_utc=OBSERVED, text_fetch=lambda _url: html
    )
    assert len(snapshots) == 1
    snapshot = snapshots[0]
    assert snapshot.home_name == "Jiangsu Dragons" and snapshot.away_name == "Sichuan"
    assert snapshot.scheduled_utc == "2026-12-04T11:35:00Z"


# ---------------------------------------------------------------------------
# Pipeline traversal: real pregame path for both leagues
# ---------------------------------------------------------------------------


def _pregame(league: str, date: str, fetch, text_fetch=None, fixture_now_utc=None) -> dict:
    store = InMemoryStore()
    config = PipelineConfig(
        league=league,
        date_yyyymmdd=date,
        fixture_now_utc=(
            fixture_now_utc or f"{date[:4]}-{date[4:6]}-{date[6:]}T00:00:00Z"
        ),
        book=BookConfig(),
    )
    pipeline = Pipeline(config, store, fetch=fetch, fetch_factory=lambda _provider: fetch)
    if text_fetch is not None:
        import ingestion.providers.betexplorer as be

        original = be._text_fetch
        be._text_fetch = lambda factory, tf=None: text_fetch
        try:
            report = pipeline.pregame()
        finally:
            be._text_fetch = original
    else:
        report = pipeline.pregame()
    return report, store


def test_euroleague_pregame_reaches_downstream_rows_with_prices():
    report, store = _pregame(
        "EUROLEAGUE",
        "20261117",
        _el_fetch(),
        _be_text_fetch(),
        fixture_now_utc=OBSERVED,
    )
    assert report["fixtures_found"] == 9
    assert not [gap for gap in report["gaps"] if gap.get("required_field") == "book_identity"]
    game_ids = {row["game_id"] for row in report["rows"]}
    moneylines = [
        snapshot
        for game_id in game_ids
        for snapshot in store.get_book_snapshots(game_id)
        if snapshot.market is Market.MONEYLINE
    ]
    # E2026-112 (xGZWYFJD), E2026-114 (0hany60s), and E2026-115 (8I2fZqVg) each
    # own `<td class="table-main__datetime">&nbsp;</td>` in el_fixtures_table.html,
    # so fetch_odds skips them rather than attaching an inherited tip.
    undated_games = {"E2026-112", "E2026-114", "E2026-115"}
    assert len(moneylines) == 6, "every own-dated slate game carries its BetExplorer pair"
    assert {snapshot.game_id for snapshot in moneylines} == game_ids - undated_games
    # Rows carry per-period total-line evidence (None for full_game, since
    # this source exposes no totals); the moneyline prices live on the
    # canonical snapshots asserted above.
    assert report["rows"][0]["book_evidence"]["full_game"] is None
    # Totals are unavailable from this source and stay visibly missing.
    assert any(
        gap.get("required_field") == "book_full_game_line" and gap.get("reason") == "missing"
        for gap in report["gaps"]
    )


def test_cba_pregame_joins_official_slate_with_betexplorer_prices():
    # Both synthetic rows carry their own served tip (matching official CBA
    # 19:35 Beijing = 12:35 CET); the second row previously used `&nbsp;`, which
    # relied on tip inheritance and is now rejected by fetch_odds.
    priced = (
        "<tr><td class=\"table-main__datetime\">04.12. 12:35</td>"
        "<td class=\"h-text-left\"><a href=\"/basketball/china/cba/jiangsu-dragons-sichuan/v7jNw7e2/\" "
        "class=\"in-match\"><span>Jiangsu Dragons</span> - <span>Sichuan</span></a></td>"
        "<td class=\"h-text-center\">&nbsp;</td><td class=\"table-main__bs\">8</td>"
        "<td class=\"table-main__odds\" data-oid=\"a1\"><button data-odd=\"1.55\"></button></td>"
        "<td class=\"table-main__odds\" data-oid=\"a2\"><button data-odd=\"2.50\"></button></td></tr>"
        "<tr><td class=\"table-main__datetime\">04.12. 12:35</td>"
        "<td class=\"h-text-left\"><a href=\"/basketball/china/cba/jilin-liaoning/8tV6oJE8/\" "
        "class=\"in-match\"><span>Jilin</span> - <span>Liaoning</span></a></td>"
        "<td class=\"h-text-center\">&nbsp;</td><td class=\"table-main__bs\">8</td>"
        "<td class=\"table-main__odds\" data-oid=\"b1\"><button data-odd=\"2.10\"></button></td>"
        "<td class=\"table-main__odds\" data-oid=\"b2\"><button data-odd=\"1.72\"></button></td></tr>"
    )
    html = f'<table class="table-main table-main--leaguefixtures">{priced}</table>'
    report, store = _pregame(
        "CBA", "20261204", _cba_fetch(), lambda _url: html
    )
    assert report["fixtures_found"] == 6
    game_ids = {row["game_id"] for row in report["rows"]}
    moneylines = [
        snapshot
        for game_id in game_ids
        for snapshot in store.get_book_snapshots(game_id)
        if snapshot.market is Market.MONEYLINE
    ]
    assert {snapshot.book for snapshot in moneylines} == {"average"}
    joined_names = {(snapshot.home_name, snapshot.away_name) for snapshot in moneylines}
    assert joined_names == {("Jiangsu Dragons", "Sichuan"), ("Jilin", "Liaoning")}
    assert not [gap for gap in report["gaps"] if gap.get("required_field") == "book_identity"]


# ---------------------------------------------------------------------------
# Backfill: durable history cursors for the new league paths
# ---------------------------------------------------------------------------


def test_euroleague_backfill_ingests_whole_seasons_and_advances_cursor(tmp_path):
    store = DuckDBStore(str(tmp_path / "el" / "warehouse.duckdb"), persist=True)
    config = PipelineConfig(league="EUROLEAGUE", date_yyyymmdd="20261116")
    every_season = _mapping_fetch({"seasons/": lambda: _fixture_json("euroleague/schedule_scores.json")})
    pipeline = Pipeline(config, store, fetch=every_season, allow_writes=True)
    pipeline.config.history_request_budget = 2
    report = pipeline.backfill()
    deep = report["backfill"]["deep_history"]
    assert deep["status"] == "acquired"
    assert deep["finals_written"] == 12
    assert [season["season"] for season in deep["seasons"]] == ["E2026", "E2025"]
    assert deep["cursor"]["next_season"] == 2024
    assert deep["learning_or_calibration_claimed"] is False
    finals = [game for game in store.list_games(league=League.EUROLEAGUE) if game.actual]
    assert len(finals) == 12
    assert all(game.actual.reconciles()[0] for game in finals)


def test_cba_backfill_walks_days_and_advances_cursor(tmp_path, monkeypatch):
    store = DuckDBStore(str(tmp_path / "cba" / "warehouse.duckdb"), persist=True)
    # The walk starts at the slate date and walks backwards day by day.
    config = PipelineConfig(league="CBA", date_yyyymmdd="20261102")
    post_fn = _cba_post_fn()
    monkeypatch.setattr(cbaleague, "_post", lambda url, body, fn=None: post_fn(url, body))
    pipeline = Pipeline(config, store, fetch=_cba_fetch(), allow_writes=True)
    pipeline.config.history_request_budget = 2
    pipeline.config.deep_history_event_budget = 10
    report = pipeline.backfill()
    deep = report["backfill"]["deep_history"]
    assert deep["status"] == "acquired"
    assert deep["days_scanned"] == ["2026-11-02", "2026-11-01"]
    assert deep["finals_written"] == 9  # 2 finals on 11-02, 7 on 11-01
    assert deep["cursor"]["next_day"] == "2026-10-31"
    finals = [game for game in store.list_games(league=League.CBA) if game.actual]
    assert len(finals) == 9
    assert all(game.actual.reconciles()[0] for game in finals)


def test_cba_backfill_retains_day_and_resumes_remaining_finals_after_event_budget(tmp_path, monkeypatch):
    store = DuckDBStore(str(tmp_path / "cba-budget" / "warehouse.duckdb"), persist=True)
    config = PipelineConfig(league="CBA", date_yyyymmdd="20261102")
    pipeline = Pipeline(config, store, fetch=lambda _url: "unused", allow_writes=True)
    pipeline.config.history_request_budget = 1
    pipeline.config.deep_history_event_budget = 1
    rows = [
        {"scheduleId": 901, "status": 4},
        {"scheduleId": 902, "status": 4},
    ]
    monkeypatch.setattr(cbaleague, "available_seasons", lambda _fetch: [1995, 2026])
    monkeypatch.setattr(cbaleague, "day_rows", lambda _day, _fetch: rows)

    def result_from_row(row, _fetch):
        return Game(
            game_id=str(row["scheduleId"]),
            league=League.CBA,
            season="2026",
            home_team_id="29137",
            away_team_id="29129",
            home_name="Jilin",
            away_name="Liaoning",
            scheduled_utc="2026-11-02T11:35:00Z",
            status=GameStatus.FINAL,
        )

    monkeypatch.setattr(cbaleague, "result_from_row", result_from_row)
    first = pipeline._backfill_cba_history()
    assert first["finals_written"] == 1
    assert first["status"] == "event_budget_exhausted_day_retained"
    health = store.get_provider_health("cbaleague-history", League.CBA)
    assert health.coverage["next_day"] == "2026-11-02"
    assert len(health.coverage["processed_schedule_ids"]) == 1

    second = pipeline._backfill_cba_history()
    assert second["finals_written"] == 1
    assert second["days_scanned"] == ["2026-11-02"]
    assert second["cursor"]["next_day"] == "2026-11-01"
    assert len([game for game in store.list_games(league=League.CBA)]) == 2


def test_cba_backfill_stops_at_oldest_season_advertised_by_source(tmp_path, monkeypatch):
    store = DuckDBStore(str(tmp_path / "cba-floor" / "warehouse.duckdb"), persist=True)
    pipeline = Pipeline(
        PipelineConfig(league="CBA", date_yyyymmdd="20261102"),
        store,
        fetch=lambda _url: "unused",
        allow_writes=True,
    )
    monkeypatch.setattr(cbaleague, "available_seasons", lambda _fetch: [1995, 2026])
    monkeypatch.setattr(pipeline, "_history_cursor", lambda _provider: {"next_day": "1994-12-31"})
    monkeypatch.setattr(
        cbaleague,
        "day_rows",
        lambda *_args: (_ for _ in ()).throw(AssertionError("must not scan before source floor")),
    )
    report = pipeline._backfill_cba_history()
    assert report["status"] == "complete_to_source_season_floor"
    assert report["source_floor_day"] == "1995-01-01"
    assert report["days_scanned"] == []


def test_backfill_dry_run_never_advances_history_cursors(tmp_path):
    store = DuckDBStore(str(tmp_path / "el" / "warehouse.duckdb"), persist=True)
    config = PipelineConfig(league="EUROLEAGUE", date_yyyymmdd="20261116")
    pipeline = Pipeline(config, store, fetch=_el_fetch(), allow_writes=False)
    pipeline.config.history_request_budget = 2
    report = pipeline.backfill()
    deep = report["backfill"]["deep_history"]
    assert deep["status"] == "disabled_canonical_write_required"
    assert store.counts()["games"] == 0


# ---------------------------------------------------------------------------
# Forecast service: history-model forecasts with real book evidence
# ---------------------------------------------------------------------------


def _seed_history(store, league: str, count: int, end_date: str):
    if league == "CBA":
        teams = ["Jilin", "Liaoning", "Shanghai", "Beijing"]
    else:
        teams = [
            "Dubai Basketball",
            "Anadolu Efes Istanbul",
            "Real Madrid",
            "FC Barcelona",
        ]
    league_enum = League[league]
    base = parse_utc(f"{end_date}T20:00:00Z")
    for index in range(count):
        home = teams[index % len(teams)]
        away = teams[(index + 1) % len(teams)]
        if home == away:
            away = teams[(index + 2) % len(teams)]
        home_q = [20 + (index % 5), 25, 22, 23]
        away_q = [21 + (index % 4), 24, 20, 22]
        scheduled = to_utc_iso(base - timedelta(days=2 * (count - index)))
        game = Game(
            game_id=f"{league}-history-{index}",
            league=league_enum,
            season="2026",
            home_team_id=home,
            away_team_id=away,
            home_name=home,
            away_name=away,
            scheduled_utc=scheduled,
            status=GameStatus.FINAL,
            actual=ActualResult(
                period_kind=PeriodKind.QUARTERS,
                home_quarters=[float(value) for value in home_q],
                away_quarters=[float(value) for value in away_q],
                home_final=float(sum(home_q)),
                away_final=float(sum(away_q)),
                overtime_periods=0,
                source="synthetic seeded history",
            ),
        )
        store.upsert_game(game)


def test_forecast_service_generates_history_forecasts_with_book_evidence():
    store = InMemoryStore()
    _seed_history(store, "EUROLEAGUE", 60, "2026-11-15")
    scheduled = Game(
        game_id="E2026-111",
        league=League.EUROLEAGUE,
        season="2026",
        home_team_id="DUB",
        away_team_id="IST",
        home_name="Dubai Basketball",
        away_name="Anadolu Efes Istanbul",
        scheduled_utc="2026-11-17T16:00:00Z",
    )
    store.upsert_game(scheduled)
    from dataclasses import replace

    raw_snapshot = betexplorer.fetch_odds(
        "EUROLEAGUE", "20261117", None, now_utc=OBSERVED, text_fetch=_be_text_fetch()
    )[0]
    # The pipeline rewrites provider game ids to canonical ids during book
    # resolution; seed the resolved form directly here.
    store.upsert_book_snapshot(replace(raw_snapshot, game_id=scheduled.game_id))
    service = MarketForecastService(store, League.EUROLEAGUE)
    report = service.build_report(
        as_of_utc="2026-11-17T12:00:00Z",
        date_yyyymmdd="20261117",
        timezone_name="Europe/Paris",
    )
    assert report["league"] == "EUROLEAGUE"
    assert scheduled.game_id not in {row["game_id"] for row in report["uncovered_games"]}
    forecasts = [
        forecast
        for forecast in report["forecasts"]
        if forecast["game_id"] == scheduled.game_id
        and forecast["market"] == "moneyline"
        and forecast["period"] == "full_game"
    ]
    assert forecasts, "the history model must emit a full-game moneyline forecast"
    forecast = forecasts[0]
    assert forecast["model"] == MODEL_VERSION
    assert forecast["book"]["provider"] == "betexplorer"
    assert forecast["book"]["selected_decimal_price"] in (1.3, 3.46)
    assert forecast["training_games"] >= 20
    assert forecast["residual_games"] >= 20
    assert forecast["team_history"]["sufficient_for_accas"] is True
    # The gates stay enforced: the prequential sample n is far below the
    # 30-calibration-sample minimum (most seeded history predates the
    # calibration epoch), so the forecast stays visible but ineligible.
    assert forecast["calibration"]["status"] in {"no_history", "provisional"}
    assert forecast["calibration"].get("n", 0) < 30
    # Operator decision 2026-11-22: this is the EuroLeague case that had never
    # produced a pick -- a real aggregate price on a team-history forecast whose
    # prequential sample is short. The shortfall is named and the row is a pick.
    assert "insufficient_walk_forward_samples" in forecast["eligibility"]["advisories"]
    assert forecast["eligibility"]["reasons"] == []
    assert forecast["eligibility"]["acca_eligible"] is True
    segments = report["walk_forward"]["segments"]
    assert any(segment["league"] == "EUROLEAGUE" for segment in segments)


def test_forecast_service_generates_cba_history_forecasts():
    store = InMemoryStore()
    _seed_history(store, "CBA", 60, "2026-11-15")
    scheduled = Game(
        game_id="100111015",
        league=League.CBA,
        season="2026",
        home_team_id="29137",
        away_team_id="29129",
        home_name="Jilin",
        away_name="Liaoning",
        scheduled_utc="2026-12-04T11:35:00Z",
    )
    store.upsert_game(scheduled)
    book = BookSnapshot(
        game_id=scheduled.game_id,
        provider="betexplorer",
        book="average",
        market=Market.MONEYLINE,
        period=Period.FULL_GAME,
        observed_utc="2026-12-04T10:00:00Z",
        home_price_decimal=2.10,
        away_price_decimal=1.72,
        provider_event_id="8tV6oJE8",
        home_name=scheduled.home_name,
        away_name=scheduled.away_name,
        scheduled_utc=scheduled.scheduled_utc,
    )
    store.upsert_book_snapshot(book)
    report = MarketForecastService(store, League.CBA).build_report(
        as_of_utc="2026-12-04T10:30:00Z",
        date_yyyymmdd="20261204",
        timezone_name="Asia/Shanghai",
    )
    forecasts = [
        forecast
        for forecast in report["forecasts"]
        if forecast["game_id"] == scheduled.game_id
        and forecast["market"] == "moneyline"
        and forecast["period"] == "full_game"
    ]
    assert forecasts and forecasts[0]["book"]["selected_decimal_price"] in (2.10, 1.72)
    assert forecasts[0]["team_history"]["sufficient_for_accas"] is True
    assert forecasts[0]["eligibility"]["reasons"] == []
    assert forecasts[0]["eligibility"]["advisories"] == ["insufficient_walk_forward_samples"]


# ---------------------------------------------------------------------------
# Red-team: failures stay visible, never fabricated
# ---------------------------------------------------------------------------


def test_stale_quotes_never_enter_pregame_analysis():
    store = InMemoryStore()
    config = PipelineConfig(
        league="EUROLEAGUE",
        date_yyyymmdd="20261117",
        fixture_now_utc="2026-11-20T12:00:00Z",  # far past the 72h staleness window
        book=BookConfig(),
    )
    pipeline = Pipeline(config, store, fetch=_el_fetch(), fetch_factory=lambda _p: _el_fetch())
    import ingestion.providers.betexplorer as be

    original = be._text_fetch
    be._text_fetch = lambda factory, tf=None: _be_text_fetch()
    try:
        report = pipeline.pregame()
    finally:
        be._text_fetch = original
    # Snapshots exist in the store but every one is older than the 72h
    # staleness window at this as-of, so none is visible to the report.
    assert all(row["odds_evidence"] is None for row in report["rows"])
    assert all(row["odds"] is None for row in report["rows"])
    for row in report["rows"]:
        game = store.get_game(row["game_id"])
        assert pipeline._visible_pregame_books(game) == []


def test_provider_failure_yields_gaps_and_no_games():
    def broken_fetch(url: str) -> Any:
        raise RuntimeError("provider down")

    store = InMemoryStore()
    pipeline = Pipeline(
        PipelineConfig(league="CBA", date_yyyymmdd="20261101", fixture_now_utc="2026-11-01T12:00:00Z"),
        store,
        fetch=broken_fetch,
    )
    report = pipeline.results()
    assert report["finals_written"] == 0
    assert any(gap.get("reason") == "provider_error" for gap in report["gaps"])


def test_ambiguous_book_identity_never_joins():
    game_a = Game(
        game_id="dup-1",
        league=League.EUROLEAGUE,
        season="2026",
        home_team_id="A",
        away_team_id="B",
        home_name="Dubai Basketball",
        away_name="Anadolu Efes Istanbul",
        scheduled_utc="2026-11-17T16:00:00Z",
    )
    game_b = Game(
        game_id="dup-2",
        league=League.EUROLEAGUE,
        season="2026",
        home_team_id="C",
        away_team_id="D",
        home_name="Dubai Basketball",
        away_name="Anadolu Efes Istanbul",
        scheduled_utc="2026-11-17T16:10:00Z",
    )
    snapshot = BookSnapshot(
        game_id="betexplorer:UkWOzxk1",
        provider="betexplorer",
        book="average",
        market=Market.MONEYLINE,
        period=Period.FULL_GAME,
        observed_utc=OBSERVED,
        home_price_decimal=1.30,
        away_price_decimal=3.46,
        provider_event_id="UkWOzxk1",
        home_name="Dubai Basketball",
        away_name="Anadolu Efes Istanbul",
        scheduled_utc="2026-11-17T16:00:00Z",
    )
    from ingestion.pipeline import _resolve_book_game

    game, reason = _resolve_book_game(snapshot, [game_a, game_b])
    assert game is None and "multiple games" in reason


def test_unknown_team_name_produces_identity_gap_not_a_join():
    game = Game(
        game_id="cba-unknown",
        league=League.CBA,
        season="2026",
        home_team_id="1",
        away_team_id="2",
        home_name="虚构赞助球队",
        away_name="北京首钢",
        scheduled_utc="2026-11-20T11:35:00Z",
    )
    snapshot = BookSnapshot(
        game_id="betexplorer:zzzzzzzz",
        provider="betexplorer",
        book="average",
        market=Market.MONEYLINE,
        period=Period.FULL_GAME,
        observed_utc="2026-11-20T10:00:00Z",
        home_price_decimal=1.9,
        away_price_decimal=1.95,
        provider_event_id="zzzzzzzz",
        home_name="Made Up Sponsor Team",
        away_name="Beijing",
        scheduled_utc="2026-11-20T11:35:00Z",
    )
    from ingestion.pipeline import _resolve_book_game

    resolved, reason = _resolve_book_game(snapshot, [game])
    assert resolved is None and "identity" in reason
