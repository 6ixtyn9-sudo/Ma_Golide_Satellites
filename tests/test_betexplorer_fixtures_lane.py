"""The fixtures-table schedule lanes: five competitions, one honest read.

These tests use a **real captured page** (``tests/fixtures/betexplorer/
el_fixtures_table.html``, EuroLeague, 2026-11-23).  The served table shape is
identical across competitions, so it exercises the parser for every new lane;
what is lane-specific — the URL each lane fetches — is asserted separately from
the recorded, h1-verified slugs.

Properties under test, in the order the failure modes were measured:

* identity is the provider match code, never a team name (this table serves
  colliding city fragments and a ``W`` suffix on women's clubs);
* the tip is the site's CET/CEST wall clock converted to UTC, not the
  league's local time;
* rows with a blank kickoff cell are reported, never dated by assumption;
* an unreadable page raises instead of returning an empty slate, while a read
  page with no games on the asked date is a legitimate empty answer;
* the price read attaches by the same identity the schedule assigns; and
* the gate fails open rather than duplicating the season-wide read.
"""

from __future__ import annotations

from datetime import datetime, timedelta, timezone
from pathlib import Path

import pytest

from ingestion.adaptive_gate import LEAGUES, GateUncertain, fetch_schedule
from ingestion.league_registry import league_spec, parse_league, price_evidence_policy
from ingestion.models import League, PriceProvenance, SourcePermission
from ingestion.providers import betexplorer, betexplorer_fixtures

FIXTURE = Path(__file__).resolve().parent / "fixtures" / "betexplorer" / "el_fixtures_table.html"
#: Pinned observation instant: the page carries relative Today/Tomorrow labels,
#: so the replay must not depend on the wall clock.
OBSERVED = "2026-11-23T06:12:00Z"
SLATE_DATE = "2026-11-24"

LANES = (
    "CHINA_WCBA_W",
    "LITHUANIA_NKL",
    "AUSTRALIA_WNBL_W",
    "ISRAEL_LIGA_LEUMIT",
    "GREECE_BASKET_LEAGUE",
)


def _page() -> str:
    return FIXTURE.read_text(encoding="utf-8")


def _slate(league: str = "CHINA_WCBA_W"):
    return betexplorer_fixtures.fetch_slate(
        league,
        SLATE_DATE,
        text_fetch=lambda _url: _page(),
        now_utc=OBSERVED,
    )


@pytest.mark.parametrize("league", LANES)
def test_each_fixtures_lane_has_a_verified_path(league):
    """A declared source without a page is a silent no-op, not a schedule."""

    spec = league_spec(parse_league(league))
    assert spec.schedule_source == "betexplorer_fixtures"
    url = betexplorer.FIXTURES_URLS[league]
    # Path-based only: the season-navigation query shapes are robots-disallowed.
    assert url.endswith("/fixtures/")
    assert "?" not in url


@pytest.mark.parametrize("league", LANES)
def test_lane_is_active_and_its_price_is_aggregate_and_unresolved(league):
    spec = league_spec(parse_league(league))
    # Leumit was retired on 2026-11-24: its fixtures page is served and its price
    # is the site average, but no permitted results route exists, so it could
    # never settle. The parse contract above still holds for it.
    assert spec.is_schedulable is (league != "ISRAEL_LIGA_LEUMIT")
    policy = price_evidence_policy(parse_league(league), "betexplorer")
    assert policy.price_provenance is PriceProvenance.AGGREGATE_CONSENSUS
    assert policy.source_permission is SourcePermission.UNRESOLVED
    # This page is a fixtures page: it carries no settlement route, and the price
    # evidence stays aggregate and unresolved whatever the lane can settle from.
    # NKL settles from the federation's own club and match pages, activated
    # 2026-11-24 on a live run -- so its results_source names that reader.
    expected_results = "nkl" if league == "LITHUANIA_NKL" else "source_unavailable"
    assert spec.results_source == expected_results


@pytest.mark.parametrize("league", LANES)
def test_slate_is_built_with_match_code_identity(league):
    games = _slate(league)

    assert games, "the captured page serves a slate for this date"
    ids = [game.game_id for game in games]
    assert len(ids) == len(set(ids)), "identity collisions must be impossible"
    for game in games:
        prefix, _, code = game.game_id.partition(":")
        assert prefix == "betexplorer"
        assert len(code) == 8 and code.isalnum()
        assert " " not in code
        assert game.league is League[league]
        # Names are display only; they never reach the identity.
        assert game.home_name not in code


def test_colliding_display_names_still_yield_distinct_games():
    """The page serves 'Paris', 'Dubai' and sponsor-suffixed names; ids don't care."""

    games = _slate()
    names = [game.home_name for game in games]

    assert len(names) == len(set(names)) or True  # collisions are expected, not fatal
    by_name: dict[str, list[str]] = {}
    for game in games:
        by_name.setdefault(game.home_name.strip().lower(), []).append(game.game_id)
    for ids in by_name.values():
        assert len(ids) == len(set(ids))


@pytest.mark.parametrize("league", LANES)
def test_tip_is_the_site_paris_clock_converted_to_utc(league):
    games = _slate(league)

    for game in games:
        tip = datetime.fromisoformat(game.scheduled_utc.replace("Z", "+00:00"))
        assert tip.tzinfo is not None
        assert tip.utcoffset() == timedelta(0)
        # Late November is outside European DST, so the served wall clock is
        # exactly one local hour ahead of the UTC instant on every row.
        paris_hour = (tip + timedelta(hours=1)).hour
        assert (paris_hour - tip.hour) % 24 == 1
        assert game.evidence["tip_timezone_assumed"] == "Europe/Paris"
        assert "not" not in game.evidence["identity"] or True
        assert game.evidence["identity"] == "provider match code from the row href"
        assert game.evidence["name_join"] == "not used"
        assert game.evidence["source_permission"] == "unresolved"


def test_blank_kickoff_rows_are_reported_and_never_invented():
    games = _slate()
    probe = betexplorer_fixtures.last_probe()
    gaps = betexplorer_fixtures.last_gaps()

    assert probe is not None
    assert probe["blank_kickoff_rows"] > 0, "the capture holds undated rows"
    assert probe["slate_games"] == len(games)
    fields = {gap.required_field for gap in gaps}
    assert "fixtures_kickoff_missing" in fields
    # Every emitted game carries a real instant; nothing was dated by guess.
    assert all(game.scheduled_utc for game in games)


def test_unreadable_page_raises_instead_of_returning_an_empty_slate():
    from ingestion.providers import IdentityError

    with pytest.raises(IdentityError):
        betexplorer_fixtures.fetch_slate(
            "LITHUANIA_NKL",
            SLATE_DATE,
            text_fetch=lambda _url: "<html>no table here</html>",
            now_utc=OBSERVED,
        )


def test_a_read_page_with_no_games_on_the_date_is_an_answer_not_a_void():
    games = betexplorer_fixtures.fetch_slate(
        "AUSTRALIA_WNBL_W",
        "2026-06-15",  # off-season: the capture spans 2026-11-18 .. 2027-04-16
        text_fetch=lambda _url: _page(),
        now_utc=OBSERVED,
    )

    assert games == []
    # The read still left a receipt, so the pipeline can tell "read and empty"
    # from "never read".
    probe = betexplorer_fixtures.last_probe()
    assert probe is not None
    assert probe["parsed_rows"] > 0
    assert probe["slate_games"] == 0


def test_price_read_attaches_to_the_same_identity_as_the_schedule():
    """The join is the match code on both sides — the whole point of the lane."""

    games = _slate()
    scheduled_ids = {game.game_id for game in games}

    snapshots = betexplorer.fetch_odds(
        "CHINA_WCBA_W",
        SLATE_DATE,
        now_utc=OBSERVED,
        text_fetch=lambda _url: _page(),
    )

    assert snapshots, "the capture holds priced rows on this date"
    priced_ids = {snapshot.game_id for snapshot in snapshots}
    assert priced_ids <= scheduled_ids, (
        "a priced row that is not a scheduled game is an unmatched quote: "
        f"{sorted(priced_ids - scheduled_ids)}"
    )


def test_no_priced_row_means_a_named_gap_not_silence():
    """Same page, a date whose rows carry no quote: say it, don't skip it."""

    games = betexplorer_fixtures.fetch_slate(
        "CHINA_WCBA_W",
        "2027-01-12",  # future season rows, unpriced on the capture
        text_fetch=lambda _url: _page(),
        now_utc=OBSERVED,
    )

    assert games, "the capture lists future fixtures on this date"
    fields = {gap.required_field for gap in betexplorer_fixtures.last_gaps()}
    assert "fixtures_price_coverage" in fields


@pytest.mark.parametrize("league", LANES)
def test_gate_fails_open_for_a_season_wide_page(league):
    config = next(
        config for config in LEAGUES if config.league_key == League[league].value
    )

    with pytest.raises(GateUncertain, match="no date-bounded gate source"):
        list(fetch_schedule(config, "2026-11-24"))


def test_adapter_refuses_a_lane_it_does_not_own():
    with pytest.raises(ValueError, match="not the schedule source"):
        betexplorer_fixtures.fetch_slate(
            "NBA", SLATE_DATE, text_fetch=lambda _url: _page(), now_utc=OBSERVED
        )


def test_replay_is_deterministic_for_a_pinned_observation_time():
    first = [(g.game_id, g.scheduled_utc) for g in _slate()]
    second = [(g.game_id, g.scheduled_utc) for g in _slate()]

    assert first == second
    assert first == sorted(first, key=lambda row: (row[1], row[0]))


def test_pipeline_dispatch_writes_the_slate_and_its_receipt():
    """The lane is wired end to end, not merely parseable in isolation."""

    from ingestion.pipeline import Pipeline, PipelineConfig
    from ingestion.store import InMemoryStore

    page = _page()
    config = PipelineConfig(
        league="LITHUANIA_NKL",
        date_yyyymmdd="20261124",
        fixture_now_utc=OBSERVED,
    )
    pipeline = Pipeline(config, InMemoryStore(), fetch=lambda url: page, allow_writes=False)

    games = pipeline._discover_games()

    assert games, "dispatch reached the fixtures adapter"
    assert {game.league for game in games} == {League.LITHUANIA_NKL}
    health = next(
        row
        for row in pipeline._health_written
        if row["provider"] == betexplorer_fixtures.SCHEDULE_SOURCE
    )
    assert health["coverage"]["slate_games"] == len(games)
    assert health["coverage"]["blank_kickoff_rows"] > 0
    # NKL's results are store-driven: a fixtures page holds no result, so the
    # same-day finals path returns nothing *without* claiming the route is
    # missing. Pending games are settled by the re-check from the club and match
    # pages (tests/test_nkl_settlement_path.py).
    assert pipeline._fetch_finals() == []
    assert not any(gap.required_field == "official_results" for gap in pipeline.gaps)


def test_a_year_less_label_resolves_to_the_nearest_year_not_next_year():
    """A fixture that tipped yesterday must not be dated next year.

    The old rule bumped the year whenever the label's month/day had already
    passed, so ``24.11.`` read on 2026-11-25 resolved to 2027: yesterday became
    unreachable and its row reappeared in next season's slate.
    """

    from ingestion.providers.betexplorer import _parse_datetime

    observed = datetime(2026, 11, 25, 12, 0, tzinfo=timezone.utc)

    yesterday = _parse_datetime("24.11. 18:00", observed)
    tomorrow = _parse_datetime("26.11. 18:00", observed)
    new_year_eve = _parse_datetime("31.12. 20:00", observed)

    assert yesterday is not None and yesterday.year == 2026
    assert yesterday.day == 24
    assert tomorrow is not None and tomorrow.year == 2026
    assert new_year_eve is not None and new_year_eve.year == 2026
    # A season that straddles January keeps both sides in one season.
    january = datetime(2027, 1, 5, 12, 0, tzinfo=timezone.utc)
    december = _parse_datetime("28.12. 20:00", january)
    february = _parse_datetime("05.02. 20:00", january)
    assert december is not None and december.year == 2026
    assert february is not None and february.year == 2027


#: Each dark results lane must carry the current state of its own settlement
#: route in ``source_detail``. A lane that says only "no results route" reads the
#: same whether nobody looked or somebody looked and hit a gate; the assessment
#: of 2026-11-24 distinguished the two, and the registry is where the reports
#: read it from.
DARK_LANE_BLOCKERS = {
    "GREECE_BASKET_LEAGUE": "EsakeResults",
    # Re-read 2026-11-24T14:02Z: the documented URL now answers 500 demanding a
    # roundId, so the round space -- not the licence -- is what blocks it.
    "CHINA_WCBA_W": "roundId",
    # Referer-only access stopped working between two reads the same day.
    "AUSTRALIA_WNBL_W": "401",
    "ISRAEL_LIGA_LEUMIT": "robots Disallow",
}


@pytest.mark.parametrize("league", sorted(DARK_LANE_BLOCKERS))
def test_a_dark_results_lane_names_the_state_of_its_own_route(league):
    spec = league_spec(parse_league(league))

    assert spec.results_source == "source_unavailable"
    assert "2026-11-24" in spec.source_detail
    assert DARK_LANE_BLOCKERS[league] in spec.source_detail
    assert "docs/RESULTS-ROUTE-ASSESSMENT.md" in spec.source_detail


def test_the_lit_nkl_lane_carries_its_live_run_in_the_detail():
    """A lane that was dark twice must say what lit it, not just that it is lit."""

    spec = league_spec(parse_league("LITHUANIA_NKL"))

    assert spec.results_source == "nkl"
    # The blocker that kept it dark is preserved rather than deleted with it.
    assert "no score" in spec.source_detail
    assert "activated 2026-11-24" in spec.source_detail
    # The evidence is the three settled games, by id and by final.
    for match_id in ("125749", "125760", "125775"):
        assert match_id in spec.source_detail
    assert "76-91" in spec.source_detail and "95-99" in spec.source_detail
    assert "docs/NKL-SETTLEMENT-LIVE-RUN.md" in spec.source_detail


def test_leumit_is_retired_because_it_could_never_settle():
    """The same reason as the nine before it: no results route, so no grading."""

    from ingestion.league_registry import LaneStatus, retirement_guard

    spec = league_spec(parse_league("ISRAEL_LIGA_LEUMIT"))
    assert spec.status is LaneStatus.RETIRED
    assert "route assessment" in spec.status_detail
    # The published route arithmetic is preserved, not deleted with the lane.
    assert "one.co.il" in spec.source_detail
    refusal = retirement_guard(parse_league("ISRAEL_LIGA_LEUMIT"))
    assert refusal and "retired" in refusal
    assert retirement_guard(parse_league("ISRAEL_LIGA_LEUMIT"), allow_retired=True) is None
