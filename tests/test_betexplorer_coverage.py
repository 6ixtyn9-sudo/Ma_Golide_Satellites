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
    assert betexplorer.FIXTURES_URLS["SOUTH_KOREA_KBL"].endswith("/south-korea/kbl/fixtures/")
    assert betexplorer.FIXTURES_URLS["ARGENTINA_LIGA_NACIONAL"].endswith("/argentina/liga-a/fixtures/")


def test_turkey_second_tier_is_not_the_top_flight():
    """``turkey/tbl`` serves 176 rows, but of the second tier — a real trap."""

    assert "/turkey/tbl/" not in betexplorer.FIXTURES_URLS["TURKEY_BSL"]


def test_bal_is_deliberately_unregistered_because_its_fixtures_table_is_empty():
    assert "BASKETBALL_AFRICA_LEAGUE" not in betexplorer.FIXTURES_URLS
    spec = league_spec(League.BASKETBALL_AFRICA_LEAGUE)
    assert all(policy.provider != "betexplorer" for policy in spec.price_sources)


def test_new_lanes_are_aggregate_and_research_only():
    """The fixtures table publishes the site average, never a named book."""

    from ingestion.models import PriceProvenance, SourcePermission

    for league in (
        League.SPAIN_ACB,
        League.TURKEY_BSL,
        League.JAPAN_B_ONE,
        League.BASKETBALL_CHAMPIONS_LEAGUE,
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
