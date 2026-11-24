"""The script that puts a research section into a lane's dated cards.

Fenced-section handling, local-day slate selection, and the read chain
(fixtures page -> match page -> three fragments) are all exercised here. The read
chain runs on committed bytes through a dict-backed ``read_text``; no test opens a
socket.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from ingestion.models import (
    BookSnapshot,
    Game,
    GameStatus,
    League,
    Market,
    OperationalValidation,
    Period,
    PriceProvenance,
    SourcePermission,
)
from ingestion.models import StandingsRow as PublishedRow
from ingestion.providers import betexplorer_standings as bs
from scripts import research_standings as rs

FIXTURES = Path(__file__).resolve().parent / "fixtures" / "betexplorer"


def _fixture(name: str) -> str:
    return (FIXTURES / name).read_text(encoding="utf-8")


SECTION = ["", "line one", "line two"]


# -- the fence ---------------------------------------------------------------


def test_a_section_is_appended_to_an_existing_card():
    text = "PICKS CARD\n  a real row\n"

    out = rs.upsert_section(text, SECTION)

    assert out.startswith("PICKS CARD\n  a real row\n")
    assert out.count(rs.MARKER_START) == 1
    assert out.count(rs.MARKER_END) == 1
    assert "line one" in out


def test_the_section_is_replaced_in_place_not_repeated():
    once = rs.upsert_section("CARD\n", SECTION)
    twice = rs.upsert_section(once, ["", "replacement"])

    assert twice.count(rs.MARKER_START) == 1
    assert "line one" not in twice
    assert "replacement" in twice
    assert twice.startswith("CARD\n")


def test_text_after_the_fence_survives_a_rewrite():
    text = rs.upsert_section("CARD\n", SECTION) + "\nTRAILING NOTE\n"

    out = rs.upsert_section(text, ["", "new"])

    assert "TRAILING NOTE" in out
    assert out.index(rs.MARKER_END) < out.index("TRAILING NOTE")


def test_an_empty_card_becomes_just_the_fence():
    out = rs.upsert_section("", SECTION)

    assert out.startswith(rs.MARKER_START)
    assert out.rstrip("\n").endswith(rs.MARKER_END)


def test_the_fence_carries_no_marker_another_reader_owns():
    """``picks_*.txt`` is scanned for a compact marker and ticket cards for PENDING."""

    text = rs.upsert_section("CARD\n", SECTION)

    assert "PENDING" not in text.upper()


# -- writing -----------------------------------------------------------------


def test_both_cards_are_created_when_the_lane_had_none(tmp_path):
    written = rs.write_section(
        tmp_path, state_key="lithuania_nkl", target_date="2026-12-01", section=SECTION
    )

    assert set(written) == {"picks", "tickets"}
    picks = tmp_path / "lithuania_nkl" / "picks_2026-12-01.txt"
    tickets = tmp_path / "lithuania_nkl" / "auto_tickets_2026-12-01.txt"
    assert picks.is_file() and tickets.is_file()
    assert picks.read_text().count(rs.MARKER_START) == 1


def test_writing_twice_leaves_one_copy(tmp_path):
    for _ in range(2):
        rs.write_section(
            tmp_path, state_key="lithuania_nkl", target_date="2026-12-01", section=SECTION
        )

    text = (tmp_path / "lithuania_nkl" / "picks_2026-12-01.txt").read_text()

    assert text.count(rs.MARKER_START) == 1


def test_the_card_names_the_dated_file(tmp_path):
    path = rs.card_path(tmp_path, "lithuania_nkl", "tickets", "2026-12-01")

    assert path.name == "auto_tickets_2026-12-01.txt"


# -- the slate ---------------------------------------------------------------


def _game(
    game_id: str,
    *,
    scheduled_utc: str = "2026-12-01T17:00:00Z",
    status: GameStatus = GameStatus.SCHEDULED,
    home: str = "Telsiai",
    away: str = "Rytas 2",
) -> Game:
    return Game(
        game_id=game_id,
        league=League.LITHUANIA_NKL,
        season="2026",
        home_team_id="x",
        away_team_id="y",
        home_name=home,
        away_name=away,
        scheduled_utc=scheduled_utc,
        status=status,
    )


def test_only_games_on_the_local_day_are_scored():
    games = [
        _game("a", scheduled_utc="2026-12-01T17:00:00Z"),
        _game("b", scheduled_utc="2026-12-02T17:00:00Z"),
        _game("c", scheduled_utc="2026-12-01T17:00:00Z", status=GameStatus.FINAL),
    ]

    assert [game.game_id for game in rs.slate_games(games, "2026-12-01")] == ["a"]


def test_a_late_utc_tip_lands_on_the_next_local_card():
    """23:30Z is already the next day in a +2 lane, and the card is dated locally."""

    games = [_game("late", scheduled_utc="2026-12-01T23:30:00Z")]

    # UTC says 1 December; the +2 lane's own calendar says 2 December.
    assert [g.game_id for g in rs.slate_games(games, "2026-12-01")] == ["late"]
    assert [g.game_id for g in rs.slate_games(games, "2026-12-02", tz_offset_hours=2)] == [
        "late"
    ]
    assert rs.slate_games(games, "2026-12-01", tz_offset_hours=2) == []


def _published(team: str, *, position: int, played: int, wins: int, losses: int,
               pf: float, pa: float) -> PublishedRow:
    return PublishedRow(
        team_name=team,
        position=position,
        games_played=played,
        wins=wins,
        losses=losses,
        points_for=pf,
        points_against=pa,
        win_pct=round(wins / played, 4),
        source="betexplorer_published_table",
    )


TABLE = [
    _published("Telsiai", position=1, played=16, wins=15, losses=1, pf=1449.0, pa=1186.0),
    _published("Rytas 2", position=18, played=16, wins=1, losses=15, pf=1268.0, pa=1428.0),
]


def _snapshot(game_id: str) -> BookSnapshot:
    return BookSnapshot(
        game_id=game_id,
        provider="betexplorer",
        book="average",
        market=Market.MONEYLINE,
        period=Period.FULL_GAME,
        observed_utc="2026-11-25T18:00:00Z",
        home_price_decimal=1.15,
        away_price_decimal=6.50,
        price_provenance=PriceProvenance.AGGREGATE_CONSENSUS,
        source_permission=SourcePermission.UNRESOLVED,
        operational_validation=OperationalValidation.PROBATION,
        home_name="Telsiai",
        away_name="Rytas 2",
    )


def test_a_scored_slate_carries_rows_and_a_slip():
    section = rs.build_section(
        "LITHUANIA_NKL",
        [_game("a")],
        TABLE,
        {"a": [_snapshot("a")]},
        target_date="2026-12-01",
    )
    text = "\n".join(section)

    assert "1 fixture(s), 1 with a side, 1 priced" in text
    assert "SINGLE @1.150" in text
    assert "RESEARCH PAPER" in text


def test_a_slate_with_no_fixture_says_so_instead_of_vanishing():
    section = rs.build_section("LITHUANIA_NKL", [], TABLE, {}, target_date="2026-12-01")
    text = "\n".join(section)

    assert "0 fixture(s)" in text
    assert "honest zero" in text


# -- the read chain ----------------------------------------------------------


def _read_chain(calls: list[str]):
    widget = _fixture("nkl_match_stats_widget.html")
    overall = _fixture("nkl_standings_overall.html")
    fixtures = _fixture("el_fixtures_table.html")

    def read(url: str) -> str:
        calls.append(url)
        if "/standings/" in url:
            return overall
        if "/fixtures/" in url:
            return fixtures
        return widget

    return read


def test_a_lane_finds_its_match_page_from_its_fixtures_page():
    calls: list[str] = []

    rows, report = rs.fetch_published("EUROLEAGUE", _read_chain(calls), delay=0.0)

    assert len(rows) == 2
    assert report["match_url"].startswith("https://www.betexplorer.com/basketball/europe/euroleague/")
    # fixtures page + match page + three fragments
    assert report["requests"] == 4
    assert len(calls) == 5
    assert report["subs_read"] == ["away", "home", "overall"]


def test_a_given_match_page_skips_the_fixtures_read():
    calls: list[str] = []
    match = "https://www.betexplorer.com/basketball/lithuania/nkl/x/McCqxqs3/"

    rows, report = rs.fetch_published(
        "LITHUANIA_NKL", _read_chain(calls), match_url=match, delay=0.0
    )

    assert len(rows) == 2
    assert len(calls) == 4
    assert report["match_url"] == match


def test_a_fixtures_page_with_no_links_is_refused_by_name():
    with pytest.raises(bs.IdentityError) as caught:
        rs.fetch_published("LITHUANIA_NKL", lambda url: "<html></html>", delay=0.0)

    assert "no match links" in str(caught.value)


def test_the_read_chain_is_paced_between_requests():
    slept: list[float] = []
    reader = rs._read_text_factory(0.0)
    assert reader._delay == 0.0  # the factory respects the argument

    paced = bs.paced(lambda url: "x", delay=10.0, sleep=slept.append)
    paced("a")
    paced("b")

    assert slept == [10.0]


# -- the fixtures page as the slate ------------------------------------------

OBSERVED = "2026-11-23T06:12:00Z"


def test_the_fixtures_page_alone_yields_a_slate_and_its_prices():
    calls: list[str] = []

    def read(url: str) -> str:
        calls.append(url)
        return _fixture("el_fixtures_table.html")

    fixtures, snapshots, html, reads = rs.slate_from_page(
        "EUROLEAGUE", "2026-11-24", read, now_utc=OBSERVED
    )

    assert fixtures
    assert reads == 1 and len(calls) == 1
    first = fixtures[0]
    assert first.game_id.startswith("betexplorer:")
    assert first.home_name and first.away_name
    assert snapshots[first.game_id][0].home_price_decimal > 1.0
    assert "table-main" in html


def test_a_page_mode_run_never_reads_the_fixtures_page_twice():
    calls: list[str] = []
    widget = _fixture("nkl_match_stats_widget.html")
    overall = _fixture("nkl_standings_overall.html")
    fixtures = _fixture("el_fixtures_table.html")

    def read(url: str) -> str:
        calls.append(url)
        if "/standings/" in url:
            return overall
        if "/fixtures/" in url:
            return fixtures
        return widget

    games, snapshots, html, page_reads = rs.slate_from_page(
        "EUROLEAGUE", "2026-11-24", read, now_utc=OBSERVED
    )
    published, report = rs.fetch_published(
        "EUROLEAGUE", read, fixtures_html=html, delay=0.0
    )

    fixtures_reads = [url for url in calls if url.endswith("/fixtures/")]
    assert len(fixtures_reads) == 1  # the page served both the slate and the token
    assert len(published) == 2
    assert report["requests"] + page_reads == len(calls)


def test_an_empty_slate_from_a_real_page_still_writes_the_fence(tmp_path):
    """A lane with nothing to play today is an honest zero, not a missing file."""

    def read(url: str) -> str:
        return _fixture("el_fixtures_table.html") if "/fixtures/" in url else "<html></html>"

    games, snapshots, _html, _reads = rs.slate_from_page(
        "EUROLEAGUE", "2026-12-25", read, now_utc=OBSERVED
    )

    assert games == []
    section = rs.build_section(
        "EUROLEAGUE", games, TABLE, snapshots, target_date="2026-12-25"
    )
    written = rs.write_section(
        tmp_path, state_key="euroleague", target_date="2026-12-25", section=section
    )

    text = Path(written["picks"]).read_text()
    assert "0 fixture(s)" in text
    assert "honest zero" in text


# -- cross-process pacing ----------------------------------------------------


def _transport(tmp_path, *, reads=None, fail=False):
    calls: list[str] = []

    def read(url: str) -> str:
        calls.append(url)
        if fail:
            raise RuntimeError("HTTP 429")
        return "payload"

    slept: list[float] = []
    clock = {"now": 1000.0}
    transport = bs.cooldown(
        read,
        tmp_path / "stamp",
        delay=10.0,
        sleep=slept.append,
        now=lambda: clock["now"],
    )
    transport.calls_ref = calls  # type: ignore[attr-defined]
    return transport, slept, clock, calls


def test_a_second_process_waits_out_the_floor_the_first_one_used(tmp_path):
    transport, slept, clock, _calls = _transport(tmp_path)

    transport("a")  # no stamp yet: no wait
    assert slept == []

    clock["now"] += 3.0  # three seconds later, in another process
    transport("b")

    assert slept == [7.0]
    assert transport.waited == 7.0


def test_the_stamp_is_written_so_the_next_invocation_can_see_it(tmp_path):
    transport, _slept, clock, _calls = _transport(tmp_path)

    transport("a")

    assert float((tmp_path / "stamp").read_text()) == 1000.0


def test_a_failed_request_still_spends_the_host_tolerance(tmp_path):
    transport, _slept, clock, _calls = _transport(tmp_path, fail=True)

    with pytest.raises(RuntimeError):
        transport("a")

    assert float((tmp_path / "stamp").read_text()) == 1000.0


def test_a_zero_delay_disables_the_wait_and_the_stamp(tmp_path):
    transport, slept, clock, _calls = _transport(tmp_path)
    transport._delay = 0.0

    transport("a")
    transport("b")

    assert slept == []
    assert not (tmp_path / "stamp").exists()


# -- the failure block -------------------------------------------------------


def test_a_failed_read_replaces_a_stale_empty_slate_block(tmp_path):
    stale = rs.upsert_section(
        "",
        rs.build_section("LITHUANIA_NKL", [], TABLE, {}, target_date="2026-11-25"),
    )
    assert "honest zero" in stale

    section = rs.failure_section(
        "LITHUANIA_NKL", "2026-11-25", "HttpError: HTTP 429 for .../fixtures/"
    )
    out = rs.upsert_section(stale, section)

    assert "honest zero" not in out  # the claim that a read happened is gone
    assert "READ FAILED" in out
    assert "HTTP 429" in out
    assert out.count(rs.MARKER_START) == 1


def test_the_failure_block_says_why_the_old_block_was_not_kept():
    text = "\n".join(
        rs.failure_section("LITHUANIA_NKL", "2026-11-25", "IdentityError: no table")
    )

    assert "replaced, not kept" in text
    assert "claim a read that did not happen" in text
