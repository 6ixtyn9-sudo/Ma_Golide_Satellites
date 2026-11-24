"""The NKL results tick: the lane settles on the schedule, not only in a script.

``docs/NKL-SETTLEMENT-LIVE-RUN.md`` recorded three games settling when the reader
was called by hand. That proves the reader; it does not prove the lane, because
nothing wrote those results to the store and nothing would have run the next day.

These tests pin the wiring that makes ``results_source="nkl"`` a behaviour rather
than a label:

* the re-check picks stored pending NKL fixtures and settles them from the club
  and match pages, stamped with the match id as a provider id;
* a club page is read once per run and reused, so the request cost is two for the
  first game of a club and one for each after it;
* a page whose quarters do not reconcile with the stated final produces a gap and
  no game -- the fail-closed rule, on the path that actually writes;
* reads are paced to nkl.lt's ``Crawl-delay: 10``, applied between requests and
  never before the first.

No test touches the network: the transport is injected, and the crawl delay is
recorded rather than slept.
"""

from __future__ import annotations

import pytest

from ingestion.models import Game, GameStatus, League
from ingestion.pipeline import Pipeline, PipelineConfig
from ingestion.providers import nkl
from ingestion.store import InMemoryStore

from tests.test_nkl_match_page import _LISTING, _PAGE

#: The page as served with its final score block, which the transcription of the
#: quarters table alone does not carry.
PLAYED_PAGE = _PAGE.replace(
    "</div></div>", "</div></div><div class=nkl-match-score>83 : 88</div>", 1
)


def _listing(*, score: str = "83-88") -> str:
    return _LISTING.replace(
        '<td class="nkl-game-score loss"> 76-91</td>',
        f'<td class="nkl-game-score loss"> {score}</td>',
        1,
    )


def _game(
    *,
    game_id: str = "BEX-125749",
    tip_utc: str = "2026-10-01T16:00:00Z",  # 19:00 in Vilnius
    home: str = "Kretinga",
    away: str = "Neptunas 2",
    status: GameStatus = GameStatus.SCHEDULED,
) -> Game:
    return Game(
        game_id=game_id,
        league=League.LITHUANIA_NKL,
        season="2026",
        home_team_id="kretinga",
        away_team_id="neptunas-2",
        home_name=home,
        away_name=away,
        scheduled_utc=tip_utc,
        status=status,
        provider_ids={"betexplorer_fixtures": game_id},
    )


def _pipeline(store: InMemoryStore, page: str, *, date_yyyymmdd: str, now_utc: str):
    asked: list = []

    def fetch(url):
        asked.append(url)
        return page

    pipeline = Pipeline(
        PipelineConfig(
            league="LITHUANIA_NKL",
            date_yyyymmdd=date_yyyymmdd,
            fixture_now_utc=now_utc,
        ),
        store,
        fetch=fetch,
        allow_writes=False,
    )
    return pipeline, asked


@pytest.fixture(autouse=True)
def _no_real_sleep(monkeypatch):
    """Pacing is asserted by recording, never by waiting."""

    monkeypatch.setattr(nkl, "CRAWL_DELAY_SECONDS", 0.0)


def test_the_results_tick_settles_a_stored_pending_game():
    store = InMemoryStore()
    store.upsert_game(_game())
    pipeline, asked = _pipeline(
        store, PLAYED_PAGE, date_yyyymmdd="20261002", now_utc="2026-10-02T06:00:00Z"
    )
    # The club page listing is served for /komandos/ and the match page otherwise.
    pipeline.fetch = lambda url: _listing() if "/komandos/" in url else PLAYED_PAGE
    pipeline._injected_fetch = True

    finals = pipeline._recheck_pending_finals(set())

    assert [game.game_id for game in finals] == ["BEX-125749"]
    settled = finals[0]
    assert settled.status is GameStatus.FINAL
    assert settled.actual is not None
    assert settled.actual.reconciles() == (True, "")
    assert (settled.actual.home_final, settled.actual.away_final) == (83.0, 88.0)
    # Attribution: the federation's own match id is stamped beside the fixture id.
    assert settled.provider_ids["nkl"] == "125749"
    assert settled.provider_ids["betexplorer_fixtures"] == "BEX-125749"
    assert not [gap for gap in pipeline.gaps if gap.required_field == "nkl_result"]


def test_a_second_game_for_the_same_club_costs_one_more_request():
    """The club page lists a season: reading it once is the whole budget."""

    store = InMemoryStore()
    store.upsert_game(_game(game_id="BEX-125749", tip_utc="2026-10-01T16:00:00Z"))
    # The listing's second row is the fixture the other way round. find_row
    # matches a pair as a set, so the forecast's own club (Kretinga, at home here)
    # still routes to one club page -- which is exactly the cache being measured.
    store.upsert_game(
        _game(
            game_id="BEX-125760",
            tip_utc="2026-10-03T12:10:00Z",  # 2026-10-03 15:10 in Vilnius
            home="Kretinga",
            away="Vilkaviskis",
        )
    )
    listing = _listing()
    asked: list = []

    def fetch(url):
        asked.append(url)
        return listing if "/komandos/" in url else PLAYED_PAGE

    pipeline = Pipeline(
        PipelineConfig(
            league="LITHUANIA_NKL",
            date_yyyymmdd="20261004",
            fixture_now_utc="2026-10-04T06:00:00Z",
        ),
        store,
        fetch=fetch,
        allow_writes=False,
    )

    finals = pipeline._recheck_pending_finals(set())

    assert len(finals) == 2
    club_reads = [url for url in asked if "/komandos/" in url]
    match_reads = [url for url in asked if "/matches/" in url]
    assert len(club_reads) == 1, "one club page per club per run"
    assert len(match_reads) == 2


def test_a_game_the_page_has_not_played_is_a_gap_not_a_result():
    store = InMemoryStore()
    store.upsert_game(_game(tip_utc="2026-10-01T16:00:00Z"))
    pipeline, _ = _pipeline(
        store, PLAYED_PAGE, date_yyyymmdd="20261002", now_utc="2026-10-02T06:00:00Z"
    )
    # An unplayed row: no score cell digits, so the listing row carries no final.
    unplayed = _LISTING.replace(
        '<td class="nkl-game-score loss"> 76-91</td>', "<td class=nkl-game-score> - </td>"
    )
    pipeline.fetch = lambda url: unplayed if "/komandos/" in url else PLAYED_PAGE
    pipeline._injected_fetch = True

    assert pipeline._recheck_pending_finals(set()) == []
    gaps = [gap for gap in pipeline.gaps if gap.required_field == "nkl_result"]
    assert gaps and "no result yet" in gaps[0].detail


def test_a_page_that_does_not_reconcile_never_becomes_a_final():
    """The failure the whole route exists to catch: quarters that miss the final."""

    store = InMemoryStore()
    store.upsert_game(_game())
    # Stated final 90 : 88, quarters sum to 83 : 88 -- a dropped period, or an
    # overtime column this reader refuses to guess about.
    mismatched = _PAGE.replace(
        "</div></div>", "</div></div><div class=nkl-match-score>90 : 88</div>", 1
    )
    pipeline, _ = _pipeline(
        store, mismatched, date_yyyymmdd="20261002", now_utc="2026-10-02T06:00:00Z"
    )
    pipeline.fetch = lambda url: _listing() if "/komandos/" in url else mismatched
    pipeline._injected_fetch = True

    assert pipeline._recheck_pending_finals(set()) == []
    gaps = [gap for gap in pipeline.gaps if gap.required_field == "nkl_result"]
    assert gaps and ("reconcile" in gaps[0].detail or "periods" in gaps[0].detail)
    stored = store.list_games(league=League.LITHUANIA_NKL)[0]
    assert stored.actual is None and stored.status is GameStatus.SCHEDULED


def test_an_unmapped_club_is_named_in_the_gap_rather_than_guessed():
    store = InMemoryStore()
    store.upsert_game(_game(home="Kretinga", away="Nobody Athletic"))
    pipeline, _ = _pipeline(
        store, PLAYED_PAGE, date_yyyymmdd="20261002", now_utc="2026-10-02T06:00:00Z"
    )

    assert pipeline._recheck_pending_finals(set()) == []
    gaps = [gap for gap in pipeline.gaps if gap.required_field == "nkl_result"]
    assert gaps and "Nobody Athletic" in gaps[0].detail


def test_the_reader_paces_between_requests_and_never_before_the_first():
    slept: list = []
    reader = nkl.paced(lambda url: url, delay=10.0, sleep=slept.append)

    reader("https://nkl.lt/komandos/a/")
    assert slept == []
    reader("https://nkl.lt/matches/1/")
    reader("https://nkl.lt/matches/2/")

    assert slept == [10.0, 10.0]
    assert reader.calls == 3


def test_the_default_delay_is_the_hosts_own_number():
    """robots.txt says 10; the constant is where that is written down."""

    reader = nkl.paced(lambda url: url, sleep=lambda _seconds: None)
    assert reader._delay == nkl.CRAWL_DELAY_SECONDS
