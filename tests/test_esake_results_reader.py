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


# --- the served page, as captured -------------------------------------------
#
# The fixtures below are *captured*, not derived: `results_row_with_onclick_
# literal.html` reproduces one row of the 2026-11-24 series-8 page, and
# `results_excerpt_captured.html` is the verbatim bytes the operator printed
# around the first game anchor, truncated where that printout stopped.  Both
# carry the toggle handler that the reader used to mistake for a row start.

CAPTURED = Path(__file__).parent / "fixtures" / "esake" / "results_excerpt_captured.html"
CAPTURED_ROW = Path(__file__).parent / "fixtures" / "esake" / "results_row_with_onclick_literal.html"


def test_the_toggle_handler_is_not_a_row_start():
    """The bug that wrote zero finals: `class="` is required, `.closest(` is not.

    The handler names the container twice in one attribute, so splitting on the
    bare token cut every row at its own toggle -- fourteen containers for seven
    games, none of them holding both the identity and the score.
    """

    page = _parse(CAPTURED_ROW.read_text(encoding="utf-8"))

    assert page.counts["containers"] == 1
    assert page.counts["rows"] == 1
    assert page.counts["no_identity"] == 0
    assert len(page.finals) == 1

    row = page.finals[0]
    assert row.idgame == "3C7521D5"
    assert (row.home, row.away) == ("ΑΡΗΣ", "ΠΑΟΚ")
    assert (row.home_score, row.away_score) == (88, 79)
    assert row.played_on == date(2026, 11, 21)
    assert row.tip_local == "19:00"


def test_the_hidden_per_quarter_row_does_not_make_the_clubs_ambiguous():
    """Four repeats of two names is still two clubs, not an ambiguity."""

    page = _parse(CAPTURED_ROW.read_text(encoding="utf-8"))

    assert page.counts["names_ambiguous"] == 0
    assert "ΑΡΗΣ" in CAPTURED_ROW.read_text(encoding="utf-8")


def test_the_captured_fragment_yields_identity_where_the_page_once_yielded_none():
    """Verbatim bytes, truncated mid-row: the identity reads, the total does not.

    That asymmetry is the honest answer for a truncated printout, and it is the
    opposite of what the reader used to say about the whole page (seven containers
    with no identity at all).
    """

    page = _parse(CAPTURED.read_text(encoding="utf-8"))

    assert page.counts["containers"] == 1
    assert page.counts["rows"] == 1
    assert page.counts["no_identity"] == 0
    assert page.rows[0].idgame == "3C7521D5"
    assert page.counts["with_final"] == 0


def test_a_date_cell_that_breaks_inside_the_month_name_still_reads():
    """The served page wraps «Νοεμβρίου» across lines; a date is still a date.

    Captured raw bytes carry newlines and indentation between «Νοε» and
    «μβρίου», so the whitespace-tolerant pass is what turns 7 finals that could
    not be dated into 7 finals that can.
    """

    html = (
        '<div class="esake-program-game">'
        "<h5>08η Αγωνιστική</h5>"
        "<div>Σαβ 21 Νοε\n                μβρίου 19:00 μμ</div>"
        '<div class="esake-program-game-final-score row equal">'
        '<div class="col-lg-5 text-right"><span>ΑΡΗΣ</span></div>'
        '<div class="col-lg-2 text-center"><span>84&nbsp;-&nbsp;87</span></div>'
        '<div class="col-lg-5 text-left"><span>ΑΕΚ</span></div>'
        "</div>"
        '<a href="/el/action/EsakegameView?idgame=3C7521D5&mode=2">p</a>'
        "</div>"
    )

    page = _parse(html)

    assert page.counts["unresolved_date"] == 0
    assert len(page.finals) == 1
    assert page.finals[0].played_on == date(2026, 11, 21)


def test_a_club_cell_that_nests_a_span_keeps_the_whole_name():
    """`<span>(.*?)</span>` stops at the inner close tag and trims the club."""

    html = (
        '<div class="esake-program-game">'
        '<div>Σαβ 21 Νοεμβρίου 19:00</div>'
        '<div class="esake-program-game-final-score row equal">'
        '<div class="col-lg-5 text-right"><span>ΚΟΛΟΣΣΟΣ <span class="x">H</span> '
        "HOTELS</span></div>"
        '<div class="col-lg-2 text-center"><span>99&nbsp;-&nbsp;77</span></div>'
        '<div class="col-lg-5 text-left"><span>ΑΡΗΣ</span></div>'
        "</div>"
        '<a href="/el/action/EsakegameView?idgame=95F2FDAA&mode=2">p</a>'
        "</div>"
    )

    page = _parse(html)

    assert page.counts["names_ambiguous"] == 0
    assert page.finals[0].home == "ΚΟΛΟΣΣΟΣ H HOTELS"


def test_the_date_comes_from_the_clock_cell_not_from_anywhere_on_the_row():
    """A day in another cell is not this game's day.

    The row carries venue and TV cells, and the page carries plenty of other
    dates.  Reading the clock cell first is what keeps «12 Μαρτίου» from some
    heading out of a November fixture.
    """

    html = (
        '<div class="esake-program-game">'
        "<h5>12 Μαρτίου - deadline</h5>"
        '<div class="esake-program-game-info one-col"><img src="/skn/clock.svg" class="mr-xs">'
        "Σαβ 21 Νοε - 15:00</div>"
        '<div class="esake-program-game-final-score row equal">'
        '<div class="col-lg-5 text-right"><span>ΑΡΗΣ</span></div>'
        '<div class="col-lg-2 text-center"><span>84&nbsp;-&nbsp;87</span></div>'
        '<div class="col-lg-5 text-left"><span>ΑΕΚ</span></div>'
        "</div>"
        '<a href="/el/action/EsakegameView?idgame=3C7521D5&mode=2">p</a>'
        "</div>"
    )

    page = _parse(html)

    assert page.finals[0].played_on == date(2026, 11, 21)


def test_the_probe_reports_the_cell_and_the_row_side_by_side():
    """`date_probe` exists because a reader saying None has two causes, and one
    call should tell "never received" from "did not match"."""

    html = (
        '<div class="esake-program-game">'
        '<div class="esake-program-game-info one-col">Σαβ 21 Νοε - 15:00</div>'
        "</div>"
    )

    probe = esake.date_probe(html)

    assert probe["cell_text"].startswith("Σαβ 21 Νοε")
    assert probe["cell_matches"] == [["21", "Νοε"]]
    assert probe["cell_result"] == (21, 11)
    assert probe["result"] == (21, 11)
    assert "\u039d\u03bf\u03b5" in probe["text_head"]


def test_a_latin1_decoded_capture_hides_the_date_and_repairing_it_restores_it():
    """The whole Greek saga in one test.

    A UTF-8 page decoded as ISO-8859-1 -- requests' default when a host declares
    no charset -- turns «Νοε» into two latin-1 characters, one of them the control
    code U+009D. Nothing about the row changes except its characters, so the
    totals still read, the names still read, and only the date disappears: the
    month table has no entry for «Î\\x9dÎ¿Îµ». Re-encoding to latin-1 and decoding
    as UTF-8 is the whole repair, and it is done at the boundary rather than
    guessed at inside the parser.
    """

    captured = (
        '<div class="esake-program-game">'
        '<div class="esake-program-game-info one-col">Σαβ 21 Νοε - 15:00</div>'
        '<div class="esake-program-game-final-score row equal">'
        '<div class="col-lg-5 text-right"><span>ΑΡΗΣ</span></div>'
        '<div class="col-lg-2 text-center"><span>84&nbsp;-&nbsp;87</span></div>'
        '<div class="col-lg-5 text-left"><span>ΑΕΚ</span></div>'
        "</div>"
        '<a href="/el/action/EsakegameView?idgame=3C7521D5&mode=2">p</a>'
        "</div>"
    )
    mojibake = captured.encode("utf-8").decode("latin-1")

    broken = _parse(mojibake)

    assert broken.counts["rows"] == 1
    assert broken.counts["with_final"] == 0
    assert broken.counts["unresolved_date"] == 1
    assert broken.rows[0].home_score == 84  # the totals survive the mojibake

    repaired = _parse(mojibake.encode("latin-1").decode("utf-8"))

    assert repaired.counts["with_final"] == 1
    assert repaired.finals[0].played_on == date(2026, 11, 21)
    assert (repaired.finals[0].home, repaired.finals[0].away) == ("ΑΡΗΣ", "ΑΕΚ")
