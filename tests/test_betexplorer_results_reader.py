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
    when: str = "22.11.2026 04:00",
) -> str:
    return (
        "<tr>"
        f'<td class="table-main__datetime">{when}</td>'
        f'<td><a href="/basketball/australia/wnbl-women/perth-sydney/{code}/" class="in-match">'
        f"<span>{home}</span></a></td>"
        f'<td><a href="/basketball/australia/wnbl-women/perth-sydney/{code}/" class="in-match">'
        f"<span><strong>{away}</strong></span></a></td>"
        f'<td class="table-main__result">{score}</td>'
        "</tr>"
    )


def _page(*rows: str) -> str:
    return '<table class="table-main"><tbody>' + "".join(rows) + "</tbody></table>"


def test_a_finished_row_yields_code_clubs_totals_and_the_date():
    page = reader.parse_results(_page(_row()))

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


def test_two_clubs_are_required_and_a_third_span_is_counted():
    page = reader.parse_results(
        _page(_row().replace("<tr>", "<tr><td><span>Round 8</span></td>", 1))
    )

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
