"""BetExplorer price coverage: the slug map, the page parameter, and placeholders.

These tests exist because the failure mode is invisible: a lane whose fixtures
URL is missing from the map, or whose odds route is called with the wrong page
value, does not error — it returns no quotes and looks like a quiet market.

Verified against served bytes on 2026-11-23 (out-of-sandbox probe):
* every competition page below answered HTTP 200 and its ``<h1>`` matched;
* ``/match-odds/{id}/{page}/{tab}/bestOdds/?lang=en`` answers **only** with
  ``page=0``; word-form values returned 404 with an empty body;
* the fixtures table publishes odds only in the days before tip-off and uses
  ``1.00`` as a "no market" placeholder.
"""

from __future__ import annotations


from ingestion.league_registry import SPECS, league_spec
from ingestion.providers import betexplorer, betexplorer_direct
from ingestion.models import League


def test_every_lane_declaring_betexplorer_has_a_fixtures_url():
    """A declared provider without a page is a silent no-op, not a price source."""

    missing = [
        spec.league.value
        for spec in SPECS
        if any(policy.provider == "betexplorer" for policy in spec.price_sources)
        and spec.league.value not in betexplorer.FIXTURES_URLS
    ]
    assert missing == []


def test_the_verified_slugs_are_the_irregular_ones_served():
    # Spot-check the slugs that are easy to guess wrong.
    assert betexplorer.FIXTURES_URLS["ITALY_LBA"].endswith("/italy/lega-a/fixtures/")
    assert betexplorer.FIXTURES_URLS["TURKEY_BSL"].endswith("/turkey/super-lig/fixtures/")
    assert betexplorer.FIXTURES_URLS["JAPAN_B_ONE"].endswith("/japan/b-league-one/fixtures/")
    # Two divisions, two pages: the top one is not the second one by another name.
    assert betexplorer.FIXTURES_URLS["JAPAN_B_LEAGUE"].endswith(
        "/japan/b-league-premier/fixtures/"
    )
    assert betexplorer.FIXTURES_URLS["SOUTH_KOREA_KBL"].endswith("/south-korea/kbl/fixtures/")
    assert betexplorer.FIXTURES_URLS["ARGENTINA_LIGA_NACIONAL"].endswith("/argentina/liga-a/fixtures/")


def test_turkey_second_tier_is_not_the_top_flight():
    """``turkey/tbl`` serves 176 rows, but of the second tier — a real trap."""

    assert "/turkey/tbl/" not in betexplorer.FIXTURES_URLS["TURKEY_BSL"]


def test_bal_is_deliberately_unregistered_because_its_fixtures_table_is_empty():
    assert "BASKETBALL_AFRICA_LEAGUE" not in betexplorer.FIXTURES_URLS
    spec = league_spec(League.BASKETBALL_AFRICA_LEAGUE)
    assert all(policy.provider != "betexplorer" for policy in spec.price_sources)


def test_the_price_only_slugs_are_the_womens_and_eurocup_pages():
    """Two lanes take prices from BetExplorer but their slate from elsewhere.

    Verified 2026-11-24 against the served pages: the h1s answered
    "Euroleague Women 2026/2027" and "Eurocup 2026/2027". These are read for
    quotes only -- neither lane's schedule or results come from here.
    """

    assert betexplorer.FIXTURES_URLS["EUROLEAGUE_W"].endswith(
        "/europe/euroleague-women/fixtures/"
    )
    assert betexplorer.FIXTURES_URLS["EUROCUP"].endswith("/europe/eurocup/fixtures/")


def test_a_price_only_lane_keeps_its_own_schedule_and_results_source():
    """Registering a price source must not quietly move the slate."""

    for league, schedule in (
        (League.EUROLEAGUE_W, "fiba_elw"),
        (League.EUROCUP, "euroleague"),
    ):
        spec = league_spec(league)
        assert spec.schedule_source == schedule
        assert spec.results_source == schedule
        assert any(p.provider == "betexplorer" for p in spec.price_sources)


def test_a_womens_row_yields_an_aggregate_snapshot_on_the_womens_slug():
    """The new price-only route reads its own page and labels what it read.

    The row below is synthetic and only pins the wiring: which URL is fetched
    for ``EUROLEAGUE_W``, that the ``W`` suffix survives into the canonical
    names used for identity, and that the emitted price is
    ``aggregate_consensus`` rather than a named book. The live values seen on
    2026-11-24 (1.65/2.13 for Kibirkstis Vilnius W v Landes W) are recorded in
    ``docs/PROVIDERS.md``.
    """

    from ingestion.models import PriceProvenance

    fetched: list[str] = []

    def text_fetch(url: str) -> str:
        fetched.append(url)
        return (
            "<tr><td class=\"table-main__datetime\">25.11. 12:00</td>"
            "<td><a href=\"/basketball/europe/euroleague-women/"
            "kibirkstis-vilnius-landes/dEwy0DzS/\" class=\"in-match\">"
            "<span>Kibirkstis Vilnius W</span> - <span>Landes W</span></a></td>"
            "<td data-oid=\"a\" data-odd=\"1.65\"></td>"
            "<td data-oid=\"b\" data-odd=\"2.13\"></td></tr>"
        )

    snapshots = betexplorer.fetch_odds(
        "EUROLEAGUE_W",
        "20261125",
        now_utc="2026-11-24T18:39:00Z",
        text_fetch=text_fetch,
    )

    assert fetched == ["https://www.betexplorer.com/basketball/europe/euroleague-women/fixtures/"]
    assert len(snapshots) == 1
    snapshot = snapshots[0]
    assert snapshot.price_provenance is PriceProvenance.AGGREGATE_CONSENSUS
    assert (snapshot.home_price_decimal, snapshot.away_price_decimal) == (1.65, 2.13)
    # Names are stored through the alias table, which is what lets the snapshot
    # join FIBA's canonical game: the site's "Landes W" is FIBA's "Basket Landes".
    assert snapshot.home_name == "Kibirkstis Vilnius W"
    assert snapshot.away_name == "Basket Landes"


def test_verified_womens_alias_pairs_let_the_womens_price_join_happen():
    """Both sides read, then mapped -- never a similarity score.

    FIBA GDAP's ``officialName`` (captured 2026-11-18) is "Basket Landes" and
    "Casademont Zaragoza"; the site writes "Landes W" and "Zaragoza W". Without
    the alias rows those two prices could never join, whatever the route did.
    """

    from ingestion.identity import canonical_team_key

    for site_name, fiba_name in (
        ("Landes W", "Basket Landes"),
        ("Basket Landes W", "Basket Landes"),
        ("Zaragoza W", "Casademont Zaragoza"),
        ("Casademont Zaragoza W", "Casademont Zaragoza"),
    ):
        assert canonical_team_key(site_name) == canonical_team_key(fiba_name)

    # Pairs nobody has read both sides of stay unmatched on purpose: a guess
    # here would attach one club's price to another club's game.
    assert canonical_team_key("Mechelen W") != canonical_team_key("Kangoeroes Mechelen")
    assert canonical_team_key("Schio W") != canonical_team_key("Famila Schio")


def test_nba_g_league_stays_unregistered_because_the_host_has_no_market():
    """The slug exists; the market does not.

    ``/usa/nba-g-league/fixtures/`` serves "NBA G League 2026/2027" and lists the
    slate, but on 2026-11-24 every row carried an empty bookmaker count -- even
    games tipping that day -- and the match page answered "Unfortunately there
    isn't any bookmaker offering odds for this match at the moment". Registering
    it would produce a source that can only ever report "no price".
    """

    assert "NBA_G_LEAGUE" not in betexplorer.FIXTURES_URLS
    spec = league_spec(League.NBA_G_LEAGUE)
    assert all(policy.provider != "betexplorer" for policy in spec.price_sources)
    # The only price route it does declare is ESPN's relayed named-book one.
    assert [p.provider for p in spec.price_sources] == ["espn"]


def test_new_lanes_are_aggregate_consensus_prices():
    """The fixtures table publishes the site average, never a named book."""

    from ingestion.models import PriceProvenance, SourcePermission

    for league in (
        League.SPAIN_ACB,
        League.TURKEY_BSL,
        League.JAPAN_B_ONE,
        League.JAPAN_B_LEAGUE,
        League.BASKETBALL_CHAMPIONS_LEAGUE,
        League.EUROLEAGUE_W,
        League.EUROCUP,
    ):
        policy = [
            p for p in league_spec(league).price_sources if p.provider == "betexplorer"
        ][0]
        assert policy.price_provenance is PriceProvenance.AGGREGATE_CONSENSUS
        assert policy.source_permission is SourcePermission.UNRESOLVED


def test_odds_route_page_parameter_defaults_to_zero_and_never_a_word():
    """``page=1`` and word forms 404; forwarding them produced 'no prices'."""

    assert betexplorer_direct._comparison_page_param("") == "0"
    assert betexplorer_direct._comparison_page_param("<html>no callback here</html>") == "0"
    # A page-supplied word form is discarded rather than forwarded.
    page = "<a onclick=\"match_load_tabs('x', 'ha', 'y', 'z', '1x2')\">"
    assert betexplorer_direct._comparison_page_param(page) == "0"
    # A genuinely numeric page value is still honoured.
    numeric = "<a onclick=\"match_load_tabs('x', 'ha', 'y', 'z', '0')\">"
    assert betexplorer_direct._comparison_page_param(numeric) == "0"


def test_one_point_zero_zero_is_not_a_price():
    """``data-odd="1.00"`` is the site's placeholder for an absent market."""

    html = (
        "<tr><td class=\"table-main__datetime\">05.12. 18:00</td>"
        "<td><a href=\"/basketball/greece/basket-league/paok-promitheas/EFZxwyFq/\""
        " class=\"in-match\"><span>PAOK</span> - <span>Promitheas</span></a></td>"
        "<td data-oid=\"a\" data-odd=\"1.00\"></td>"
        "<td data-oid=\"b\" data-odd=\"9.80\"></td></tr>"
    )
    from datetime import datetime, timezone

    rows = betexplorer._rows(html, datetime(2026, 11, 23, tzinfo=timezone.utc))

    assert len(rows) == 1
    assert rows[0][5] == ()  # no prices, rather than a fabricated 1.00


def test_a_populated_row_still_yields_both_prices():
    html = (
        "<tr><td class=\"table-main__datetime\">05.12. 18:00</td>"
        "<td><a href=\"/basketball/greece/basket-league/paok-promitheas/EFZxwyFq/\""
        " class=\"in-match\"><span>PAOK</span> - <span>Promitheas</span></a></td>"
        "<td data-oid=\"a\" data-odd=\"1.62\"></td>"
        "<td data-oid=\"b\" data-odd=\"2.31\"></td></tr>"
    )
    from datetime import datetime, timezone

    rows = betexplorer._rows(html, datetime(2026, 11, 23, tzinfo=timezone.utc))

    assert rows[0][5] == ("1.62", "2.31")


def test_diacritics_are_folded_on_both_sides():
    """BetExplorer strips accents site-wide; league feeds keep them.

    Real pairs from the 2026-11-23 capture: ``Forca Lleida`` (site) vs
    ``Força Lleida`` (league), ``Breogan``/``Breogán``, ``Wurzburg``/``Würzburg``.
    """

    from ingestion.identity import normalize_team_name

    pairs = [
        ("Força Lleida", "Forca Lleida"),
        ("Breogán", "Breogan"),
        ("Würzburg", "Wurzburg"),
        ("PAOK", "PAOK"),
    ]
    for left, right in pairs:
        assert normalize_team_name(left) == normalize_team_name(right)


def test_folding_does_not_merge_different_clubs():
    from ingestion.identity import normalize_team_name

    assert normalize_team_name("Real Madrid") != normalize_team_name("Real Betis")
    assert normalize_team_name("Barcelona") != normalize_team_name("Barca")
