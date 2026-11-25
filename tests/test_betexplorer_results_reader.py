"""The aggregator results reader: what it reads, and what it refuses to read.

The fixtures here are **synthetic**, assembled from the only markup on this host
that has been captured in full -- the fixtures row the existing parser already
handles (`in-match` link, two `<span>` names, a `table-main__datetime` cell) --
plus a score cell.  They prove the reader's rules; they do not prove the served
results page's markup.  That is what the capture tool is for: the first real page
replaces this file.
"""

from __future__ import annotations

from datetime import date

import pytest

from ingestion.providers import betexplorer_results as reader


def _row(
    code: str = "AbCd1234",
    home: str = "Perth",
    away: str = "Sydney",
    score: str = "89:84",
    when: str = "",
) -> str:
    """One results row, in the shape the 2026-11-24 capture actually served.

    ``data-test`` comes *before* ``href`` and both clubs sit inside the one
    ``in-match`` anchor; the total is a second, unclassed link in the next cell.
    """

    date_cell = f'<td class="table-main__datetime">{when}</td>' if when else ""
    return (
        "<tr>"
        + date_cell
        + '<td class="h-text-left">'
        f'<a data-test="11015187" href="/basketball/australia/wnbl-women/perth-sydney/{code}/"'
        ' class="in-match">'
        f"<span><strong>{home}</strong></span> - <span>{away}</span></a></td>"
        '<td class="h-text-center">'
        f'<a href="/basketball/australia/wnbl-women/perth-sydney/{code}/">{score}</a></td>'
        '<td class="table-main__odds colored" data-odd="1.05">1.05</td>'
        "</tr>"
    )


def _page(*rows: str) -> str:
    return '<table class="table-main"><tbody>' + "".join(rows) + "</tbody></table>"


def test_a_finished_row_yields_code_clubs_totals_and_the_date():
    page = reader.parse_results(_page(_row(when="22.11.2026 04:00")))

    assert page.counts["rows"] == 1
    assert page.counts["read"] == 1
    assert len(page.rows) == 1

    row = page.rows[0]
    assert row.match_id == "AbCd1234"
    assert (row.home, row.away) == ("Perth", "Sydney")
    assert (row.home_score, row.away_score) == (89, 84)
    assert row.played_on == date(2026, 11, 22)
    assert row.source == reader.SOURCE


def test_a_kickoff_clock_is_not_a_score():
    """`20:00` meets the colon shape and fails the only test that matters."""

    page = reader.parse_results(_page(_row(score="", when="22.11.2026 20:00")))

    assert page.counts["read"] == 0
    assert page.counts["score_not_one"] == 1


def test_a_forfeit_is_left_unread_rather_than_invented():
    """20:0 is a real final and an exact complement of a clock; we skip it."""

    page = reader.parse_results(_page(_row(score="20:0")))

    assert page.counts["read"] == 0


def test_two_score_shaped_pairs_are_an_ambiguity_not_a_choice():
    page = reader.parse_results(_page(_row(score="89:84 (55:44)")))

    assert page.counts["read"] == 0
    assert page.counts["score_not_one"] == 1


def test_a_span_outside_the_match_anchor_is_not_a_club():
    """The page wraps round labels, tabs and odds in spans too; only the anchor's
    content is a club list."""

    page = reader.parse_results(
        _page(_row().replace("<tr>", "<tr><td><span>Round 8</span></td>", 1))
    )

    assert page.counts["read"] == 1
    assert (page.rows[0].home, page.rows[0].away) == ("Perth", "Sydney")


def test_three_spans_inside_the_anchor_are_counted_not_guessed():
    three = _row().replace("- <span>Sydney</span>", "- <span>Refs</span> <span>Sydney</span>")

    page = reader.parse_results(_page(three))

    assert page.counts["read"] == 0
    assert page.counts["names_not_two"] == 1


def test_a_row_without_a_match_code_is_counted_not_dropped():
    no_code = _row().replace('class="in-match"', 'class="not-a-match"')

    page = reader.parse_results(_page(no_code))

    assert page.counts["rows"] == 1
    assert page.counts["without_code"] == 1


def test_an_undated_result_reads_but_is_not_dated():
    page = reader.parse_results(_page(_row(when="22.11.")))

    assert page.counts["read"] == 1
    assert page.counts["undated"] == 1
    assert page.rows[0].played_on is None


def test_the_same_match_twice_is_read_once():
    page = reader.parse_results(_page(_row(), _row()))

    assert page.counts["read"] == 1
    assert page.counts["duplicate_match_id"] == 1


def test_an_empty_body_raises_rather_than_reading_as_no_finals():
    with pytest.raises(ValueError):
        reader.parse_results("  ")


# --- the match page: the split the page states beside the total ---------------
#
# Observed on the served page on 2026-11-25 (NKL, code McCqxqs3):
#
#     24.11.2026 - 12:00
#     93:107
#     (15:25, 30:28, 17:27, 31:27)
#
# The wrapper elements were not captured, so these fixtures are written at the
# text level on purpose: the reader must work from what the page *says*, not from
# a tag name that was guessed at.  The numbers and the shape are the observed
# ones.


def _match_page(final: str = "93:107", split: str = "(15:25, 30:28, 17:27, 31:27)") -> str:
    return (
        "<html><body><div class='wrap'>"
        "<h3>Jurbarkas</h3><p class='detail'>24.11.2026 - 12:00</p>"
        f"<p class='score'>{final}</p>"
        f"<p class='score-detail'>{split}</p>"
        "<p>Head-to-head</p>"
        "<table><tr><td>Jurbarkas</td><td>92:86</td><td>22.01.2026</td></tr></table>"
        "</div></body></html>"
    )


def test_the_match_page_split_reconciles_with_its_own_total():
    periods = reader.parse_match_page(_match_page())

    assert periods is not None
    assert (periods.home_final, periods.away_final) == (93, 107)
    assert periods.home_quarters == (15.0, 30.0, 17.0, 31.0)
    assert periods.away_quarters == (25.0, 28.0, 27.0, 27.0)
    assert periods.overtime == ()
    assert periods.matches


def test_a_split_that_does_not_add_up_is_refused():
    # 15+30+17+31 = 93, so a stated total of 94 is not this split's total.
    assert reader.parse_match_page(_match_page(final="94:107")) is None


def test_an_overtime_period_is_carried_not_dropped():
    periods = reader.parse_match_page(_match_page("100:107", "(15:25, 30:28, 17:27, 31:27, 7:0)"))

    assert periods is not None
    assert len(periods.regulation) == 4
    assert periods.home_overtime == 7.0 and periods.away_overtime == 0.0
    assert periods.matches


def test_a_page_with_no_split_yields_nothing():
    assert reader.parse_match_page(_match_page(split="")) is None


def test_two_splits_on_one_page_are_an_ambiguity_not_a_choice():
    page = _match_page() + _match_page()
    assert reader.parse_match_page(page) is None


def test_a_split_with_too_few_periods_is_refused():
    assert reader.parse_match_page(_match_page("62:81", "(31:40, 31:41)")) is None


def test_a_split_without_a_total_beside_it_is_refused():
    page = "<html><body><p>(15:25, 30:28, 17:27, 31:27)</p></body></html>"
    assert reader.parse_match_page(page) is None
