"""The ESAKE reader: what it reads, and what it refuses to read.

The fixture is **derived, not captured**. It is assembled from the 2026-11-24
probe's verbatim snippets — the `esake-program-game-final-score` block holding
`<span>100&nbsp;-&nbsp;64</span>`, the `EsakegameView?idgame=1439C9D6` identity,
and the «21 Νοε / 15:00» day-and-tip cell — plus deliberately plain surrounding
markup. It proves the parser's rules; it does **not** prove the live page's
container classes. That is the point of shipping the reader with a probe script:
the first real fetch replaces this file.

What is asserted, in the order the guesses would cost the most:

* a finished row yields identity, both clubs, both totals, a date and a tip;
* a row with no `final-score` block is *not* a final — it may be unplayed, and a
  parser that invents 0-0 for it would invent a settlement;
* three club-shaped spans are an ambiguity to count, never a coin flip;
* an unreadable month leaves the date unknown, and no date means no final;
* a page with no row container, or with no game identity, raises: neither may
  read as "no games in this round".
"""

from __future__ import annotations

from datetime import date
from pathlib import Path

import pytest

from ingestion.providers import IdentityError, esake

FIXTURE = Path(__file__).parent / "fixtures" / "esake" / "results_excerpt_derived.html"


def _page() -> str:
    return FIXTURE.read_text(encoding="utf-8")


def _parse(html: str | None = None, *, season_start_year: int | None = 2026):
    return esake.parse_results(_page() if html is None else html, season_start_year=season_start_year)


def test_a_finished_row_carries_identity_both_clubs_both_totals_and_a_date():
    page = _parse()

    assert page.counts["rows"] == 2
    assert len(page.finals) == 1
    final = page.finals[0]
    assert final.idgame == "1439C9D6"
    assert final.home == "ΚΟΛΟΣΣΟΣ H HOTELS COLLECTION"
    assert final.away == "ΠΡΟΜΗΘΕΑΣ"
    assert (final.home_score, final.away_score) == (100, 64)
    assert final.played_on == date(2026, 11, 21)
    assert final.tip_local == "15:00"
    assert final.url == "https://www.esake.gr/el/action/EsakegameView?idgame=1439C9D6&mode=2"
    # Attribution travels with the row: the terms that permit copying ask for it.
    assert final.source.startswith("Ε.Σ.Α.Κ.Ε.")


def test_a_row_without_a_final_score_block_is_not_a_final():
    page = _parse()

    unplayed = next(row for row in page.rows if row.idgame == "AAAABBBB")
    assert unplayed.has_result is False
    assert unplayed.is_final is False
    assert (unplayed.home, unplayed.away) == ("ΑΡΗΣ", "ΠΑΟΚ")
    assert page.counts["without_final"] == 1
    assert {"idgame": "AAAABBBB", "reason": "no final-score block in the row"} in page.incomplete


def test_a_bare_score_outside_the_final_block_is_not_evidence():
    # A date cell or a bracket label can wear the same shape; only the block counts.
    html = (
        '<div class="esake-program-game row">'
        "<span>ΑΡΗΣ</span><span>100 - 64</span><span>ΠΑΟΚ</span>"
        '<a href="/el/action/EsakegameView?idgame=CCCCDDDD&amp;mode=2">x</a>'
        "</div>"
    )
    page = _parse(html)

    assert page.finals == ()
    assert page.counts["without_final"] == 1


def test_three_club_shaped_spans_are_counted_not_guessed():
    html = (
        '<div class="esake-program-game row">'
        "<span>ΑΡΗΣ</span><span>ΠΑΟΚ</span><span>ΟΛΥΜΠΙΑΚΟΣ</span>"
        '<div class="esake-program-game-final-score"><span>80&nbsp;-&nbsp;70</span></div>'
        '<a href="/el/action/EsakegameView?idgame=EEEEFFFF&amp;mode=2">x</a>'
        "</div>"
    )
    page = _parse(html)

    assert page.finals == (), "an ambiguous row must not be promoted to a final"
    assert page.counts["names_ambiguous"] == 1
    assert page.rows[0].home is None and page.rows[0].away is None


def test_an_unreadable_month_leaves_the_date_unknown_so_nothing_settles():
    html = (
        '<div class="esake-program-game row">'
        '<div class="when">21 Ξεφ</div><span>ΑΡΗΣ</span><span>ΠΑΟΚ</span>'
        '<div class="esake-program-game-final-score"><span>80&nbsp;-&nbsp;70</span></div>'
        '<a href="/el/action/EsakegameView?idgame=11112222&amp;mode=2">x</a>'
        "</div>"
    )
    page = _parse(html)

    assert page.rows[0].has_result is True
    assert page.rows[0].played_on is None
    assert page.finals == ()
    assert page.counts["unresolved_date"] == 1


def test_a_when_cell_we_cannot_classify_makes_the_row_ambiguous_not_a_final():
    """A dated span we cannot read is not silently promoted to a club name.

    The cost of guessing here would be two clubs swapped on a settled game, so
    the row is counted as ambiguous and never reaches ``finals``.
    """

    html = (
        '<div class="esake-program-game row">'
        "<span>21 Ξεφ</span><span>ΑΡΗΣ</span><span>ΠΑΟΚ</span>"
        '<div class="esake-program-game-final-score"><span>80&nbsp;-&nbsp;70</span></div>'
        '<a href="/el/action/EsakegameView?idgame=55556666&amp;mode=2">x</a>'
        "</div>"
    )
    page = _parse(html)

    assert page.counts["names_ambiguous"] == 1
    assert page.finals == ()


def test_without_a_season_year_a_result_cannot_be_dated():
    page = _parse(season_start_year=None)

    assert page.finals == ()
    assert page.counts["unresolved_date"] == 1


def test_a_page_with_no_row_container_raises_instead_of_reading_as_an_empty_round():
    html = '<a href="/el/action/EsakegameView?idgame=1439C9D6&amp;mode=2">x</a>'

    with pytest.raises(IdentityError) as raised:
        _parse(html)
    assert "row container" in str(raised.value)


def test_a_page_with_no_game_identity_raises():
    html = '<div class="esake-program-game row"><span>ΑΡΗΣ</span></div>'

    with pytest.raises(IdentityError) as raised:
        _parse(html)
    assert "no EsakegameView game identity" in str(raised.value)


def test_an_empty_body_raises_rather_than_answering():
    with pytest.raises(IdentityError):
        _parse("   ")


def test_a_duplicate_identity_is_counted_once():
    row = (
        '<div class="esake-program-game row">'
        "<span>21 Νοε</span><span>ΑΡΗΣ</span><span>ΠΑΟΚ</span>"
        '<div class="esake-program-game-final-score"><span>80&nbsp;-&nbsp;70</span></div>'
        '<a href="/el/action/EsakegameView?idgame=33334444&amp;mode=2">x</a>'
        "</div>"
    )
    page = _parse(row + row)

    assert page.counts["duplicate_idgame"] == 1
    assert len(page.finals) == 1


@pytest.mark.parametrize(
    "label, month",
    [
        ("Νοε", 11),  # the one label the capture actually shows
        ("ΝΟΕ", 11),
        ("Νοέ", 11),
        ("Ιουν", 6),
        ("Ιουλ", 7),
        ("Μαϊ", 5),
        ("Ξεφ", None),
        ("", None),
    ],
)
def test_month_labels_resolve_or_refuse(label, month):
    assert esake.month_from_label(label) == month


def test_the_matchday_url_keeps_the_parameters_the_probe_validated():
    assert esake.results_url(8) == (
        "https://www.esake.gr/el/action/EsakeResults"
        "?idchampionship=184645B9&idseason=00000001&idteam=&series=8"
    )
    assert "series=" not in esake.results_url()
    assert esake.game_url("1439c9d6").endswith("idgame=1439C9D6&mode=2")


def test_the_checksum_names_the_bytes_that_were_read():
    body = _page()

    assert esake.checksum(body) == esake.checksum(body)
    assert esake.checksum(body) != esake.checksum(body + " ")
    assert len(esake.checksum(body)) == 64


def test_the_page_reports_counts_so_a_drifted_markup_is_visible():
    payload = _parse().to_dict()

    assert payload["counts"]["containers"] == 2
    assert payload["counts"]["with_final"] == 1
    assert payload["finals"][0]["played_on"] == "2026-11-21"
    assert any(item["idgame"] == "AAAABBBB" for item in payload["incomplete"])
