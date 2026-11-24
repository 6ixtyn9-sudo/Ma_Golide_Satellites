"""Japan B.League (B1) lane: server-rendered slate + keyless S3 snapshot.

Every body used here was served to an out-of-sandbox probe on 2026-11-23 (this
sandbox has no egress to bleague.jp or the S3 bucket) and is vendored under
``tests/fixtures/bleague/captures/``. The slate file is a capture *excerpt* —
one verbatim ``<li>`` block with a synthesised ``<ul>``/``</li>`` wrapper — and
the context file holds the verbatim ``Game`` object from snapshot 113.

Cases covered: the slate parse, the .NET ``/Date(ms+0900)/`` tip conversion, the
context's club identity, a reconciling quarter split, an overtime split that
must fall back to totals, a settled context without totals, a context answering
another key, a slate with no game block, a non-integer ``latestid``, the probe
receipt, the UTC-date filter across the JST boundary, and the gate's slate read.
"""

from __future__ import annotations

import json
from datetime import datetime, timezone
from typing import Any, Callable, Dict, List

import pytest

from ingestion.http import HttpError
from ingestion.league_registry import league_spec
from ingestion.models import GameStatus, League, PeriodKind
from ingestion.providers import bleague
from ingestion.providers import IdentityError
from ingestion.store import InMemoryStore

SLATE = "tests/fixtures/bleague/captures/schedule-2026-10-04-excerpt.html"
CONTEXT = "tests/fixtures/bleague/captures/context-506418-113.json"


def _slate() -> str:
    with open(SLATE, encoding="utf-8") as handle:
        return handle.read()


def _context_text() -> str:
    with open(CONTEXT, encoding="utf-8") as handle:
        return handle.read()


def _context() -> Dict[str, Any]:
    return json.loads(_context_text())


def _transport(context: Any | None = None, slate: str | None = None) -> Callable[[str], str]:
    body = _context_text() if context is None else json.dumps(context, ensure_ascii=False)
    page = _slate() if slate is None else slate

    def fetch(url: str) -> str:
        if "/latestid" in url:
            return "113"
        if url.endswith("/113.json"):
            return body
        if "bleague.jp/schedule/" in url:
            return page
        raise AssertionError(f"unexpected URL: {url}")

    return fetch


def test_registry_points_the_lane_at_the_site_and_the_snapshot():
    spec = league_spec(League.JAPAN_B_LEAGUE)
    assert (spec.schedule_source, spec.results_source, spec.gate_schedule_source) == (
        "bleague",
        "bleague",
        "bleague",
    )
    assert "mon=" in spec.source_detail


def test_slate_url_uses_the_honoured_parameter_names():
    # The host ignores ``month``; the captured working form is year/mon/day.
    url = bleague.schedule_url(datetime(2026, 10, 4).date())
    assert url.endswith("/schedule/?tab=1&year=2026&mon=10&day=4")


def test_slate_parse_reads_key_names_scores_status_and_tip():
    blocks = bleague.parse_slate(_slate())

    assert len(blocks) == 1
    block = blocks[0]
    assert block["schedule_key"] == "506418"
    assert block["home_name"] == "北海道"
    assert block["away_name"] == "仙台"
    assert (block["home_score"], block["away_score"]) == ("100", "89")
    assert block["status_text"] == "FINAL"
    assert block["tip_local"] == "14:05"


def test_dotnet_date_is_converted_with_its_own_offset():
    game = bleague._game_from_context(_context(), "506418", "2026-10-04")

    # /Date(1791090300000+0900)/ = 2026-10-04T05:05:00Z = 14:05 JST
    assert game.scheduled_utc == "2026-10-04T05:05:00Z"
    assert game.league is League.JAPAN_B_LEAGUE
    assert game.game_id == "BLEAGUE-506418"
    assert game.home_team_id == "702"
    assert game.away_team_id == "692"
    assert game.season == "2027"


def test_reconciling_quarters_are_written_with_their_finals():
    game = bleague._game_from_context(_context(), "506418", "2026-10-04")

    assert game.status is GameStatus.FINAL
    actual = game.actual
    assert actual.period_kind is PeriodKind.QUARTERS
    assert actual.home_quarters == [16.0, 26.0, 27.0, 31.0]
    assert actual.away_quarters == [23.0, 28.0, 17.0, 21.0]
    assert actual.home_final == 100.0
    assert actual.away_final == 89.0
    ok, _ = actual.reconciles()
    assert ok


def test_overtime_split_falls_back_to_totals_instead_of_a_bad_sum():
    body = _context()
    game = body["Game"]
    game["MaxPeriod"] = 5
    game["GameCurrentPeriod"] = 5
    game["HomeTeamScore"] = 108
    game["AwayTeamScore"] = 97

    parsed = bleague._game_from_context(body, "506418", "2026-10-04")

    assert parsed.status is GameStatus.FINAL
    assert parsed.actual.period_kind is PeriodKind.UNKNOWN
    assert parsed.actual.home_quarters == []
    assert parsed.actual.home_final == 108.0
    assert any("does not reconcile" in gap["detail"] for gap in bleague.last_gaps())


def test_settled_context_without_totals_is_refused_as_a_final():
    body = _context()
    del body["Game"]["HomeTeamScore"]

    parsed = bleague._game_from_context(body, "506418", "2026-10-04")

    assert parsed.status is not GameStatus.FINAL
    assert parsed.actual is None
    assert any("refused as a final" in gap["detail"] for gap in bleague.last_gaps())


def test_context_answering_another_key_is_refused():
    body = _context()
    body["Game"]["ScheduleKey"] = 506419

    with pytest.raises(IdentityError):
        bleague._game_from_context(body, "506418", "2026-10-04")


def test_discover_drops_a_game_outside_the_target_utc_date():
    # The capture tips 2026-10-04T05:05Z, so a 2026-10-05 request has no game
    # even though the same slate page is served for both days.
    games = bleague.discover_games("2026-10-05", _transport())

    assert games == []
    assert bleague.last_probe()["slate_games"] == 1
    assert bleague.last_probe()["games"] == 0


def test_discover_keeps_the_game_on_its_own_utc_date():
    games = bleague.discover_games("2026-10-04", _transport())

    assert [game.game_id for game in games] == ["BLEAGUE-506418"]


def test_slate_without_a_game_block_raises_instead_of_reading_as_empty():
    with pytest.raises(ValueError):
        bleague.discover_games("2026-10-04", _transport(slate="<html><body>no games</ul>"))


def test_non_integer_latestid_is_refused():
    def fetch(url: str) -> str:
        if "/latestid" in url:
            return "<html>error</html>"
        if "bleague.jp/schedule/" in url:
            return _slate()
        raise AssertionError(url)

    games = bleague.discover_games("2026-10-04", fetch)

    assert games == []
    assert any("snapshot id" in gap["detail"] for gap in bleague.last_gaps())


def test_probe_records_every_read_not_a_summary_claim():
    bleague.discover_games("2026-10-04", _transport())

    probe = bleague.last_probe()
    urls = [read["url"] for read in probe["reads"]]
    assert any("bleague.jp/schedule/" in url for url in urls)
    assert any(url.endswith("/latestid") for url in urls)
    assert any(url.endswith("/113.json") for url in urls)
    assert all(len(read["sha256"]) == 64 for read in probe["reads"])
    assert probe["requests"] == len(urls)


def test_gate_reads_the_slate_and_reports_the_utc_tip(monkeypatch):
    from ingestion import adaptive_gate

    seen: List[str] = []

    def fake_read(url: str, headers=None, **kwargs):  # noqa: ANN001, ANN003
        seen.append(url)
        return _slate()

    monkeypatch.setattr(adaptive_gate, "_read_limited_response", fake_read)
    config = adaptive_gate.LeagueGateConfig(
        label="Japan B.League",
        league="japan_b_league",
        league_key="JAPAN_B_LEAGUE",
        state_key="japan_b_league",
        acquisition_timezone="Asia/Tokyo",
        schedule_source="bleague",
    )
    slate = adaptive_gate._bleague_schedule(config, "2026-10-04")

    assert len(seen) == 2
    assert all("tab=1&year=" in url for url in seen)
    assert len(slate) == 1
    game = slate[0]
    assert game.game_id == "BLEAGUE-506418"
    assert game.status == "final"
    assert game.scheduled_utc == datetime(2026, 10, 4, 5, 5, tzinfo=timezone.utc)


def test_gate_fails_open_on_a_page_without_blocks(monkeypatch):
    from ingestion import adaptive_gate

    monkeypatch.setattr(
        adaptive_gate, "_read_limited_response", lambda url, **kw: "<html></html>"
    )
    config = adaptive_gate.LeagueGateConfig(
        label="Japan B.League",
        league="japan_b_league",
        league_key="JAPAN_B_LEAGUE",
        state_key="japan_b_league",
        acquisition_timezone="Asia/Tokyo",
        schedule_source="bleague",
    )

    with pytest.raises(adaptive_gate.GateUncertain):
        adaptive_gate._bleague_schedule(config, "2026-10-04")


def test_a_slate_read_for_another_day_counts_its_dropped_rows():
    """0 fixtures after a real read has a reason, and the reason is countable.

    The 16:37 tick answered a 2026-11-24 target with ``fixtures_found: 0`` for
    Japan and said nothing at all, because a row whose tip is outside the asked
    UTC day is dropped silently. That drop is correct; being unable to see it
    afterwards is not.
    """

    games = bleague.discover_games("2026-10-05", _transport())

    assert games == []
    probe = bleague.last_probe()
    assert probe["slate_games"] == 1
    assert probe["off_date_dropped"] == 1
    assert probe["requests"] >= 2  # both JST slate days, plus the context read


def test_zero_fixtures_on_a_read_slate_is_not_reported_as_a_void():
    from ingestion.models import CoverageState
    from ingestion.pipeline import Pipeline, PipelineConfig

    original = (bleague.discover_games, bleague.last_gaps, bleague.last_probe)
    bleague.discover_games = lambda *_a, **_k: []
    bleague.last_gaps = lambda: []
    bleague.last_probe = lambda: {"url": "https://www.bleague.jp/x", "requests": 1,
                                  "slate_games": 1, "off_date_dropped": 1, "bytes": 10}
    try:
        pipeline = Pipeline(
            PipelineConfig(league="JAPAN_B_LEAGUE", date_yyyymmdd="20261124"),
            InMemoryStore(),
        )
        games = pipeline._discover_games()
    finally:
        bleague.discover_games, bleague.last_gaps, bleague.last_probe = original

    assert games == []
    assert [g for g in pipeline.gaps if g.required_field == "official_schedule"] == []
    row = pipeline.store.get_provider_health("bleague", League.JAPAN_B_LEAGUE)
    assert row.state is CoverageState.MISSING
    assert row.coverage["off_date_dropped"] == 1


def test_each_slate_page_records_its_digest_and_how_many_keys_it_added():
    """Whether the host scopes the page by ``day`` is observed, not assumed.

    Two pages are fetched (the two JST days that can hold a tip). If the second
    adds no key, the probe says so with a digest instead of the contract
    docstring claiming a day scoping the host may not honour.
    """

    bleague.discover_games("2026-10-04", _transport())

    pages = bleague.last_probe()["pages"]
    assert [page["day"] for page in pages] == ["2026-10-04", "2026-10-05"]
    assert [page["new_keys"] for page in pages] == [1, 0]
    assert all(len(page["sha256"]) == 64 for page in pages)


def test_the_probe_says_whether_day_scopes_the_page(monkeypatch):
    """The 2026-11-23 settlement: `day` is echoed, not honoured.

    Four day=22/23/24/25 fetches returned the same ten game keys and differed
    by two bytes of echoed `day`, so a second page that adds no key is the
    expected shape on this host — and if it ever adds one, the probe says
    ``scoped`` instead of silently absorbing it.
    """

    calls = {"n": 0}
    real_parse = bleague.parse_slate

    def counting_parse(html: str):
        calls["n"] += 1
        return real_parse(html)

    monkeypatch.setattr(bleague, "parse_slate", counting_parse)

    bleague.discover_games("2026-10-04", _transport())
    assert bleague.last_probe()["day_param"] == "ignored"

    # Same page twice: the second adds nothing. Now make the second page carry
    # a key the first did not, which is what a genuinely day-scoped host looks
    # like.
    second = {"schedule_key": "506999"}
    state = {"n": 0}

    def two_pages(html: str):
        state["n"] += 1
        blocks = real_parse(html)
        return blocks if state["n"] == 1 else blocks + [second]

    monkeypatch.setattr(bleague, "parse_slate", two_pages)
    monkeypatch.setattr(
        bleague, "_fetch_context", lambda key, transport: _context()
    )
    monkeypatch.setattr(
        bleague,
        "_game_from_context",
        lambda payload, key, requested, *lane: (_ for _ in ()).throw(
            ValueError("off date")
        ),
    )

    bleague.discover_games("2026-10-04", _transport())
    assert bleague.last_probe()["day_param"] == "scoped"


def test_the_matcher_sees_the_english_club_name_not_the_japanese_slate_label():
    """Why a Japanese-name alias table is the wrong fix (captured 2026-11-24).

    The schedule page renders short Japanese labels (``福岡``, ``越谷``), but this
    lane does not read names from that page: it reads the S3 context, which
    carries ``HomeTeamNameE``/``HomeTeamNameJ``, and prefers the English field.
    A Latin board spelling therefore agrees by ordinary normalisation; only the
    Japanese field would need a hand-maintained pair per club, and it is never
    the candidate. ``HomeTeamNameJ`` stays as the fallback for a context that
    omits English, and that fallback is exactly what would fail here.
    """

    from ingestion.providers.betexplorer_direct import _team_names_agree

    game = bleague._game_from_context(_context(), "506418", "2026-10-04")

    assert game.home_name == "LEVANGA HOKKAIDO"  # HomeTeamNameE, per the capture
    assert _team_names_agree("LEVANGA HOKKAIDO", "Levanga Hokkaido", league="JAPAN_B_LEAGUE")[0] is True
    # The Japanese label is a different script: no token rule can join it.
    assert _team_names_agree("レバンガ北海道", "Levanga Hokkaido", league="JAPAN_B_LEAGUE")[0] is False


def _denied_transport(status: int = 403) -> Callable[[str], str]:
    """The schedule serves; every context pointer is refused, as S3 does."""

    page = _slate()

    def fetch(url: str) -> str:
        if "bleague.jp/schedule/" in url:
            return page
        raise HttpError(status, url, "<?xml version='1.0'?><Error><Code>AccessDenied</Code>")

    return fetch


def test_a_permanent_denial_is_a_reported_gap_not_a_dead_tick():
    """The 2026-11-24 scheduled tick died on exactly this: ``HTTP 403`` from
    ``/web_json/v2_genius_contexts/507403/latestid``, exit 1, no summary and no
    gap. A lane whose source refuses it must report that, not take the run with
    it -- every other lane's slate was lost with the process."""

    games = bleague.discover_games("2026-10-04", _denied_transport(403))

    assert games == []
    gaps = bleague.last_gaps()
    assert len(gaps) == 1, "one refusal is one gap, not one per key"
    detail = gaps[0]["detail"]
    assert "contexts refused" in detail and "HTTP 403" in detail
    assert "v2_genius_contexts" in detail
    # The receipt still shows the schedule read that succeeded.
    probe = bleague.last_probe()
    assert probe["reads"] and probe["contexts_denied"][0]["status"] == 403


def test_the_denial_probe_stops_after_two_keys():
    """Two keys is the smallest number that can tell a moved object from a
    closed tree; ten is a hammer."""

    page = _slate()
    # A second block, so discovery holds two keys.
    # The block holding the key, not the first <li> in the document.
    start = page.rfind("<li", 0, page.find("506418"))
    end = page.find("</li>", start) + len("</li>")
    two = page[:end] + page[start:end].replace("506418", "506419") + page[end:]
    calls: List[str] = []

    def fetch(url: str) -> str:
        calls.append(url)
        if "bleague.jp/schedule/" in url:
            return two
        raise HttpError(403, url, "AccessDenied")

    games = bleague.discover_games("2026-10-04", fetch)

    assert games == []
    denied = bleague.last_probe()["contexts_denied"]
    assert len(denied) == bleague.DENIAL_PROBES
    assert len({item["key"] for item in denied}) == bleague.DENIAL_PROBES


def test_a_transient_failure_still_raises():
    """A 500 is not the bucket saying no; it is the run saying it does not know."""

    with pytest.raises(HttpError):
        bleague.discover_games("2026-10-04", _denied_transport(500))


def test_the_one_lane_tick_reports_a_denial_instead_of_exiting_nonzero():
    """End of the chain: the pipeline turns the denial into coverage, not a crash."""

    from ingestion.pipeline import Pipeline, PipelineConfig

    config = PipelineConfig(
        league="JAPAN_B_ONE", date_yyyymmdd="20261125", fixture_now_utc="2026-11-24T16:00:00Z"
    )
    pipeline = Pipeline(
        config, InMemoryStore(), fetch=_denied_transport(403), allow_writes=False
    )

    games = pipeline._discover_games()  # must not raise: this is the dead tick

    assert games == []
    assert any(
        gap.provider == "bleague" and "contexts refused" in gap.detail
        for gap in pipeline.gaps
    )
