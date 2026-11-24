"""The nkl.lt match page reader, against the markup the host actually serves.

The synthetic page below reproduces the two facts that break every other reader in
this repository: **unquoted attributes** and **newlines inside tags**.  It is
transcribed from the 2026-11-24 capture (200, 680,857 B, sha256 ``61a4fa2e…``,
robots ``Crawl-delay: 10``), not invented.
"""

from __future__ import annotations

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
