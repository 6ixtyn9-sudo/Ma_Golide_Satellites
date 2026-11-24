"""The nkl.lt match page reader, against the markup the host actually serves.

The synthetic page below reproduces the two facts that break every other reader in
this repository: **unquoted attributes** and **newlines inside tags**.  It is
transcribed from the 2026-11-24 capture (200, 680,857 B, sha256 ``61a4fa2e…``,
robots ``Crawl-delay: 10``), not invented.
"""

from __future__ import annotations

from datetime import date, datetime

from zoneinfo import ZoneInfo

from ingestion.providers import nkl

#: Served shape: `class=nkl-quarters-table`, `<td\nclass=team-cell>`, unquoted href.
_PAGE = """
<div
class=nkl-top-info-bar><div
class=nkl-top-info-item>
<span
class=nkl-top-info-label>Rungtynių pradžia</span>
<span
class=nkl-top-info-val>2026-10-15, 19:00</span></div></div><div
class=nkl-card-box><table
class=nkl-quarters-table><thead><tr><th></th><th>Q1</th><th>Q2</th><th>Q3</th><th>Q4</th></tr></thead><tbody><tr><td
class=team-cell>
<a
href=https://nkl.lt/komandos/kretingos-kretinga-flexblow/ >
<img
src=http://nkl.lt/wp-content/uploads/2026/09/4.png alt>  <span>KRE</span>
</a></td><td>18</td><td>24</td><td>23</td><td>18</td></tr><tr><td
class=team-cell>
<a
href=https://nkl.lt/komandos/marijampoles-suduva/ >
<img
src=http://nkl.lt/wp-content/uploads/2026/09/17.png alt>  <span>SUD</span>
</a></td><td>31</td><td>21</td><td>15</td><td>21</td></tr></tbody></table></div>
"""


def test_it_reads_the_quarters_the_capture_showed():
    page = nkl.parse_match_page(_PAGE, url="https://nkl.lt/matches/125781/")

    assert page.home_quarters == [18.0, 24.0, 23.0, 18.0]
    assert page.away_quarters == [31.0, 21.0, 15.0, 21.0]
    assert page.quarter_sums() == (83.0, 88.0)
    assert (page.home_code, page.away_code) == ("KRE", "SUD")
    assert (page.home_slug, page.away_slug) == (
        "kretingos-kretinga-flexblow",
        "marijampoles-suduva",
    )
    assert page.played_on.isoformat() == "2026-10-15"
    assert page.tip_local.hour == 19


def test_a_sum_of_quarters_is_not_a_final():
    """83-88 is what the quarters add up to.  Comparing a sum to itself always
    reconciles and would hide a dropped overtime column, so the page must state a
    final before this record can settle anything."""

    page = nkl.parse_match_page(_PAGE, url="https://nkl.lt/matches/125781/")

    assert page.home_final is None
    assert page.to_actual_result() is None
    assert page.to_dict()["settleable"] is False
    assert any("no final score element" in reason for reason in page.reasons)


def test_a_stated_final_that_matches_makes_it_settleable():
    page = nkl.parse_match_page(
        _PAGE.replace(
            "</div></div>",
            "</div></div><div class=nkl-match-score>83 : 88</div>",
            1,
        ),
        url="https://nkl.lt/matches/125781/",
    )

    result = page.to_actual_result()

    assert result is not None
    assert result.reconciles() == (True, "")
    assert page.to_dict()["settleable"] is True


def test_a_fifth_period_column_is_refused_not_truncated():
    """Overtime grows the table by a column.  Truncating to Q1..Q4 would drop the
    extra period and produce a record that fails its own reconciliation."""

    page = nkl.parse_match_page(
        _PAGE.replace("<th>Q4</th>", "<th>Q4</th><th>OT1</th>"),
        url="https://nkl.lt/matches/125781/",
    )

    assert any("not Q1..Q4" in reason for reason in page.reasons)


def test_a_short_row_is_refused():
    page = nkl.parse_match_page(
        _PAGE.replace("<td>18</td><td>24</td><td>23</td><td>18</td>", "<td>18</td><td>24</td>"),
        url="https://nkl.lt/matches/125781/",
    )

    assert page.to_actual_result() is None
    assert any("four quarters per club" in reason for reason in page.reasons)


def test_an_empty_body_is_a_receipt_not_a_traceback():
    page = nkl.parse_match_page("", url="https://nkl.lt/matches/125781/")

    assert page.to_actual_result() is None
    assert "no nkl-quarters-table in the page" in page.reasons


def test_when_no_final_can_be_read_the_receipt_names_the_candidates():
    """The element stating the final is not yet known. Reporting the score-shaped
    pairs on the page is how the rule gets written from markup instead of guessed."""

    page = nkl.parse_match_page(_PAGE, url="https://nkl.lt/matches/125781/")
    reasons = " ".join(page.reasons)

    assert "no final score element" in reasons
    assert "score-shaped pairs" in reasons or "no 'Rungtynių" in reasons


def test_the_receipt_names_where_the_final_came_from():
    """A final that is really a summary row echoing the quarter sum would make
    reconciles() true by construction. The audit trail has to be in the receipt."""

    page = nkl.parse_match_page(
        _PAGE.replace(
            "</div></div>",
            "</div></div><div class=nkl-match-score><span>83</span> : <span>88</span></div>",
            1,
        ),
        url="https://nkl.lt/matches/125781/",
    )

    assert page.to_dict()["final_source"].startswith("class=nkl-match-score")
    assert "83" in page.to_dict()["final_source"]


def test_an_unplayed_fixture_is_not_a_parse_failure():
    """The schedule runs ahead of the scores. A page with no result table means
    the game has not been played, and saying 'parse failed' every night would
    bury a real markup change in that noise."""

    upcoming = _PAGE.replace(
        '<div\nclass=nkl-card-box><table\nclass=nkl-quarters-table>', "<div"
    )
    page = nkl.parse_match_page(upcoming)

    assert page.state == "no_result"
    assert page.played_on.isoformat() == "2026-10-15"  # the date cell still reads


def test_a_played_game_is_settled_and_a_five_column_one_is_incomplete():
    played = nkl.parse_match_page(
        _PAGE.replace("</div></div>", "</div></div><div class=nkl-match-score>83 : 88</div>", 1)
    )
    overtime = nkl.parse_match_page(
        _PAGE.replace("<th>Q4</th>", "<th>Q4</th><th>OT1</th>").replace(
            "</div></div>", "</div></div><div class=nkl-match-score>83 : 88</div>", 1
        )
    )

    assert played.state == "settled"
    assert overtime.state == "incomplete"
    assert overtime.to_actual_result() is None


#: Two rows of the club fixture table, transcribed from the 2026-11-24 capture:
#: `<tr class=nkl-match-row data-match-url=… >`, unquoted, date as `YYYY-MM-DD HH:MM:SS`.
_LISTING = """
<table class=nkl-matches-table-modern><tbody><tr class=nkl-match-row data-match-url=https://nkl.lt/matches/125749/ ><td style="color:#64748b; font-size:11px; white-space:nowrap;">2026-10-01 19:00:00</td><td><div style="display:flex; flex-direction:column; gap:3px;"><div class=nkl-game-team-cell> <img src=http://nkl.lt/wp-content/uploads/2026/09/4.png class=nkl-game-team-logo alt> <span>Kretingos Kretinga-FlexBlow</span></div><div class=nkl-game-team-cell> <img src=http://nkl.lt/wp-content/uploads/2026/09/18.png class=nkl-game-team-logo alt> <span>Klaipėdos Neptūnas-II</span></div></div></td><td class="nkl-game-score loss"> 76-91</td><td style="text-align:right; width: 40px;"> <a href=https://nkl.lt/matches/125749/ style="color:#075e32;">stat »</a></td></tr><tr class=nkl-match-row data-match-url=https://nkl.lt/matches/125760/ ><td style="color:#64748b; font-size:11px; white-space:nowrap;">2026-10-03 15:10:00</td><td><div style="display:flex; flex-direction:column; gap:3px;"><div class=nkl-game-team-cell> <img src=http://nkl.lt/wp-content/uploads/2026/09/19.png class=nkl-game-team-logo alt> <span>Vilkaviškio Perlas Go</span></div><div class=nkl-game-team-cell> <img src=http://nkl.lt/wp-content/uploads/2026/09/4.png class=nkl-game-team-logo alt> <span>Kretingos Kretinga-FlexBlow</span></div></div></td><td class="nkl-game-score win"> 84-65</td><td style="text-align:right; width: 40px;"> <a href=https://nkl.lt/matches/125760/ style="color:#075e32;">stat »</a></td></tr></tbody></table>
"""


def test_the_club_listing_gives_date_clubs_final_and_id():
    """One row is a joined view: it is what lets a forecast be found without
    opening a 680 KB match page first."""

    rows, counts = nkl.parse_club_page(_LISTING)

    assert counts["rows"] == 2 and counts["with_score"] == 2
    first = rows[0]
    assert first.match_id == "125749"
    assert first.url == "https://nkl.lt/matches/125749/"
    assert first.kickoff_local.date().isoformat() == "2026-10-01"
    assert (first.home_name, first.away_name) == (
        "Kretingos Kretinga-FlexBlow",
        "Klaipėdos Neptūnas-II",
    )
    assert (first.home_score, first.away_score) == (76.0, 91.0)


def test_an_unplayed_row_keeps_none_rather_than_repeating_the_last_score():
    """A fixture with no result cell must not inherit the previous row's numbers:
    zeros or a carry-over would settle a game that has not been played."""

    upcoming = _LISTING.replace(
        '<td class="nkl-game-score loss"> 76-91</td>',
        "<td class=nkl-game-score> - </td>",
    )

    rows, counts = nkl.parse_club_page(upcoming)

    assert counts["with_score"] == 1
    assert (rows[0].home_score, rows[0].away_score) == (None, None)
    assert (rows[1].home_score, rows[1].away_score) == (84.0, 65.0)


def test_every_nkl_club_maps_to_exactly_one_betexplorer_club():
    """Eighteen names on each side, reviewed on 2026-11-24. A duplicate target
    would settle one club's game with another's quarter scores."""

    targets = list(nkl.CLUB_NAMES.values())

    assert len(nkl.CLUB_NAMES) == 18
    assert len(set(targets)) == 18
    assert len({name.strip().lower() for name in nkl.CLUB_NAMES}) == 18


def test_the_two_olimpas_clubs_do_not_swap():
    """Palanga's club is called Olimpas and so is Plungė's. A similarity matcher
    swaps them; the reviewed table does not."""

    assert nkl.club_name("Palangos Palanga") == "BC Olimpas Palanga"
    assert nkl.club_name("Plungės Olimpas") == "Plunges Olimpas"


def test_find_row_matches_a_forecast_by_date_and_club_pair():
    rows, _ = nkl.parse_club_page(_LISTING)

    hit = nkl.find_row(
        rows,
        played_on=date(2026, 10, 1),
        home="Kretinga",
        away="Neptunas 2",
    )
    reversed_order = nkl.find_row(
        rows,
        played_on=date(2026, 10, 1),
        home="Neptunas 2",
        away="Kretinga",
    )

    assert hit is not None and hit.match_id == "125749"
    assert reversed_order is not None and reversed_order.match_id == "125749"


def test_find_row_refuses_a_wrong_date_or_an_unmapped_club():
    rows, _ = nkl.parse_club_page(_LISTING)

    assert nkl.find_row(rows, played_on=date(2026, 10, 2), home="Kretinga",
                        away="Neptunas 2") is None
    assert nkl.find_row(rows, played_on=date(2026, 10, 1), home="Unknown BC",
                        away="Neptunas 2") is None


def test_settle_forecast_reads_two_pages_and_produces_a_reconciling_record():
    played = _PAGE.replace(
        "</div></div>",
        "</div></div><div class=nkl-match-score>83 : 88</div>",
        1,
    )
    calls = []

    def read_text(url):
        calls.append(url)
        return _LISTING if "/komandos/" in url else played

    settlement = nkl.settle_forecast(
        read_text, played_on=date(2026, 10, 1), home="Kretinga", away="Neptunas 2"
    )

    assert settlement.requests_made == 2
    assert calls == [
        "https://nkl.lt/komandos/kretingos-kretinga-flexblow/",
        "https://nkl.lt/matches/125749/",
    ]
    result = settlement.result
    assert result is not None
    assert result.reconciles() == (True, "")
    assert (result.home_final, result.away_final) == (83.0, 88.0)


def test_a_cached_club_page_costs_one_request_per_extra_game():
    played = _PAGE.replace(
        "</div></div>", "</div></div><div class=nkl-match-score>83 : 88</div>", 1
    )
    cache = {}
    calls = []

    def read_text(url):
        calls.append(url)
        return _LISTING if "/komandos/" in url else played

    for _ in range(2):
        nkl.settle_forecast(
            read_text, played_on=date(2026, 10, 1), home="Kretinga", away="Neptunas 2",
            cache=cache,
        )

    assert calls.count("https://nkl.lt/komandos/kretingos-kretinga-flexblow/") == 1
    assert len(calls) == 3


def test_an_unmapped_club_makes_no_request_at_all():
    calls = []

    settlement = nkl.settle_forecast(
        lambda url: calls.append(url), played_on=date(2026, 10, 1),
        home="Some New Club", away="Kretinga",
    )

    assert calls == []
    assert settlement.requests_made == 0
    assert "no NKL mapping" in settlement.reasons[0]


def test_a_game_without_a_result_is_not_settled():
    upcoming = _LISTING.replace(
        '<td class="nkl-game-score loss"> 76-91</td>', "<td class=nkl-game-score> - </td>"
    )
    calls = []

    def read_text(url):
        calls.append(url)
        return upcoming

    settlement = nkl.settle_forecast(
        read_text, played_on=date(2026, 10, 1), home="Kretinga", away="Neptunas 2"
    )

    assert settlement.result is None
    assert settlement.requests_made == 1  # club page only; no match page fetched
    assert "no result yet" in settlement.reasons[0]


def test_quarters_that_disagree_with_the_final_are_refused_not_written():
    """This is what an overtime column that the reader did not see looks like: the
    page states a final the four quarters cannot reach."""

    page = _PAGE.replace(
        "</div></div>", "</div></div><div class=nkl-match-score>83 : 92</div>", 1
    )

    def read_text(url):
        return _LISTING if "/komandos/" in url else page

    settlement = nkl.settle_forecast(
        read_text, played_on=date(2026, 10, 1), home="Kretinga", away="Neptunas 2"
    )

    assert settlement.result is None
    assert "do not reconcile" in settlement.reasons[0]


def test_a_miss_says_which_kind_of_miss_it_is():
    """'no row' was the sentence that cost a live run: it is the same words for an
    empty page, a page of other dates, and a name-spelling mismatch."""

    rows, _ = nkl.parse_club_page(_LISTING)

    empty = nkl.explain_miss([], played_on=date(2026, 10, 1), home="Kretinga",
                             away="Neptunas 2")
    wrong_day = nkl.explain_miss(rows, played_on=date(2026, 10, 5), home="Kretinga",
                                 away="Neptunas 2")
    right_day = nkl.explain_miss(rows, played_on=date(2026, 10, 1), home="Someone",
                                 away="Neptunas 2")

    assert "no fixture rows at all" in empty[0]
    assert "no fixture row on 2026-10-05" in wrong_day[0]
    assert "rows between 2026-10-01 and 2026-10-03" in wrong_day[0]
    # The names are printed as the page spells them, so a mapping gap is visible.
    assert "Kretingos Kretinga-FlexBlow v Klaipėdos Neptūnas-II" in right_day[0]


def test_a_miss_caused_by_an_unmapped_name_is_named_as_such():
    rows = [
        nkl.ListingRow(
            match_id="1",
            url="https://nkl.lt/matches/1/",
            kickoff_local=datetime(2026, 10, 1, 19, 0, tzinfo=ZoneInfo("Europe/Vilnius")),
            home_name="Kretingos Kretinga-FlexBlow",
            away_name="KlaipÄ—dos NeptÅ«nas-II",  # a latin-1 decode of UTF-8
            home_score=76.0,
            away_score=91.0,
        )
    ]

    reasons = nkl.explain_miss(rows, played_on=date(2026, 10, 1), home="Kretinga",
                               away="Neptunas 2")

    assert "no BetExplorer mapping for" in reasons[0]
    assert "KlaipÄ—dos" in reasons[0]
