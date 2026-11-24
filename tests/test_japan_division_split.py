"""Japan is two competitions, so it is two lanes -- one division each.

Run 45385819819 priced the top division (B.LEAGUE PREMIER, the schedule page's
``tab=1``) against BetExplorer's ``japan/b-league-one`` board, which is the
**second** division: no club name could ever agree, so the lane priced nothing
and still claimed full coverage. The two divisions were split into
``JAPAN_B_LEAGUE`` (PREMIER) and ``JAPAN_B_ONE`` on 2026-11-24.

The invariant these tests hold is narrow and load-bearing: a lane names one
division, the adapter reads that division's tab, the price source is the board
page captured for *that* division, and no lane inherits the other's slug.
"""

from __future__ import annotations

from typing import Any, Dict, List

import pytest

from ingestion.league_registry import parse_league, league_spec
from ingestion.models import League
from ingestion.providers import betexplorer, bleague

SLATE = "tests/fixtures/bleague/captures/schedule-2026-10-04-excerpt.html"
CONTEXT = "tests/fixtures/bleague/captures/context-506418-113.json"


def _transport(seen: List[str] | None = None):
    slate = open(SLATE, encoding="utf-8").read()
    context = open(CONTEXT, encoding="utf-8").read()

    def fetch(url: str) -> str:
        if seen is not None:
            seen.append(url)
        if "/latestid" in url:
            return "113"
        if url.endswith("/113.json"):
            return context
        if "bleague.jp/schedule/" in url:
            return slate
        raise AssertionError(f"unexpected URL: {url}")

    return fetch


def test_the_two_divisions_are_two_lanes_with_their_own_state():
    premier = league_spec(League.JAPAN_B_LEAGUE)
    one = league_spec(League.JAPAN_B_ONE)

    assert premier.league is not one.league
    assert premier.state_key != one.state_key
    assert (premier.state_key, one.state_key) == ("japan_b_league", "japan_b_one")
    for spec in (premier, one):
        assert (spec.schedule_source, spec.results_source, spec.gate_schedule_source) == (
            "bleague",
            "bleague",
            "bleague",
        )
        assert spec.timezone == "Asia/Tokyo"
    assert "PREMIER" in premier.source_detail
    assert "tab=1" in premier.source_detail
    assert "ONE" in one.source_detail
    assert "tab=2" in one.source_detail


def test_each_lane_owns_exactly_one_division_tab():
    assert bleague.tab_for(League.JAPAN_B_LEAGUE) == 1
    assert bleague.tab_for(League.JAPAN_B_ONE) == 2
    # The gate passes the registry's string, not the enum.
    assert bleague.tab_for("JAPAN_B_ONE") == 2
    with pytest.raises(ValueError, match="no B.League division tab"):
        bleague.tab_for("JAPAN_NBL")


def test_each_division_has_its_own_board_page():
    """Two divisions, two pages, and neither borrows the other's.

    The second division's page was matched game-for-game against the tab=2 slate.
    The top division's page was read on 2026-11-24 and served the heading
    "B.League Premier 2026/2027 Fixtures" with dated rows from 04.12. -- a
    third-party read, which is why the comment names its provenance instead of
    presenting it as this pipeline's own capture.
    """

    assert betexplorer.FIXTURES_URLS["JAPAN_B_ONE"].endswith("/japan/b-league-one/fixtures/")
    assert betexplorer.FIXTURES_URLS["JAPAN_B_LEAGUE"].endswith(
        "/japan/b-league-premier/fixtures/"
    )
    assert betexplorer.FIXTURES_URLS["JAPAN_B_ONE"] != betexplorer.FIXTURES_URLS["JAPAN_B_LEAGUE"]
    for league in (League.JAPAN_B_LEAGUE, League.JAPAN_B_ONE):
        spec = league_spec(league)
        assert [policy.provider for policy in spec.price_sources] == ["betexplorer"]
        # The direct day-board fallback remains on both, and is division-agnostic.
        assert spec.book_providers == ("betexplorer", "betexplorer_direct")


def test_the_b_one_lane_still_gets_the_aggregate_source_it_was_captured_with():
    from ingestion.models import PriceProvenance, SourcePermission

    policy = next(
        p for p in league_spec(League.JAPAN_B_ONE).price_sources if p.provider == "betexplorer"
    )
    assert policy.price_provenance is PriceProvenance.AGGREGATE_CONSENSUS
    assert policy.source_permission is SourcePermission.UNRESOLVED


def test_the_adapter_reads_the_tab_its_lane_owns():
    seen: List[str] = []
    games = bleague.discover_games("2026-10-04", _transport(seen), league=League.JAPAN_B_ONE)

    assert all("tab=2&year=" in url for url in seen if "bleague.jp/schedule/" in url)
    assert games and all(game.league is League.JAPAN_B_ONE for game in games)
    probe = bleague.last_probe()
    assert (probe["league"], probe["tab"]) == ("JAPAN_B_ONE", 2)


def test_a_lane_reading_another_divisions_tab_is_refused():
    with pytest.raises(ValueError, match="one lane is one division"):
        bleague.discover_games("2026-10-04", _transport(), tab=2, league=League.JAPAN_B_LEAGUE)


def test_a_result_re_read_keeps_the_lane_it_was_discovered_in():
    """Re-reading a B.ONE game's context must not relabel it PREMIER."""

    games = bleague.discover_games("2026-10-04", _transport(), league=League.JAPAN_B_ONE)
    assert games
    refreshed = bleague.fetch_game_result(games[0], _transport())
    assert refreshed.league is League.JAPAN_B_ONE
    assert refreshed.game_id == games[0].game_id


def test_the_gate_asks_each_lane_for_its_own_division(monkeypatch):
    from ingestion import adaptive_gate

    seen: List[str] = []
    slate = open(SLATE, encoding="utf-8").read()

    def fake_read(url: str, headers=None, **kwargs):  # noqa: ANN001, ANN003
        seen.append(url)
        return slate

    monkeypatch.setattr(adaptive_gate, "_read_limited_response", fake_read)
    config = next(c for c in adaptive_gate.LEAGUES if c.league_key == "JAPAN_B_ONE")
    adaptive_gate._bleague_schedule(config, "2026-10-04")

    assert seen and all("tab=2&year=" in url for url in seen)


def test_both_lanes_are_schedulable_and_distinct_in_the_gate_matrix():
    from ingestion import adaptive_gate

    by_key: Dict[str, Any] = {c.league_key: c for c in adaptive_gate.LEAGUES}
    assert by_key["JAPAN_B_LEAGUE"].state_key == "japan_b_league"
    assert by_key["JAPAN_B_ONE"].state_key == "japan_b_one"
    assert by_key["JAPAN_B_ONE"].schedule_source == "bleague"
    assert parse_league("JAPAN_B_ONE") is League.JAPAN_B_ONE
    # The established spelling still resolves to the PREMIER lane it always was.
    assert parse_league("Japan B.League") is League.JAPAN_B_LEAGUE


def test_the_two_lanes_never_share_a_price_page():
    """Distinct divisions must not resolve to one board page."""

    urls = list(betexplorer.FIXTURES_URLS.values())
    assert len(urls) == len(set(urls)), "two lanes claim the same fixtures page"
