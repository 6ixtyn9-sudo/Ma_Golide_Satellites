"""The published-table reader: arithmetic, refusals, pacing, and no token reuse.

Every test runs on committed bytes. The two fixtures under
``tests/fixtures/betexplorer/`` state their own provenance in a header comment --
which characters are served and which are a marked continuation -- because the
capture that proved this route reached the sandbox as a printed window rather than
a saved file. Nothing here asserts a number that was not on the page.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from ingestion.models import StandingsRow
from ingestion.providers import betexplorer_standings as bs
from ingestion.providers import IdentityError

FIXTURES = Path(__file__).resolve().parent / "fixtures" / "betexplorer"

TOKEN = "2BG7ii0m"
MATCH_ID = "McCqxqs3"
MATCH_URL = (
    "https://www.betexplorer.com/basketball/lithuania/nkl/jurbarkas-vytis-vdu/McCqxqs3/"
)
OVERALL_URL = (
    "https://www.betexplorer.com/basketball/lithuania/nkl/standings/"
    f"?table=table&table_sub=overall&ts={TOKEN}&dcheck=0&as-ajax=1&l=en&event_context={MATCH_ID}"
)


def _fixture(name: str) -> str:
    return (FIXTURES / name).read_text(encoding="utf-8")


def _widget() -> str:
    return _fixture("nkl_match_stats_widget.html")


def _overall() -> str:
    return _fixture("nkl_standings_overall.html")


def _fragment(
    rows: str,
    *,
    sub: str = "overall",
    head: str | None = None,
) -> str:
    """Build a fragment with the served head, for the cases the capture lacks."""

    columns = head or (
        '<th class="rank col_rank" title="Rank">#</th>'
        '<th class="participant_name col_name" title="Team">Team</th>'
        '<th class="matches" data-type="matches" title="Matches Played">MP</th>'
        '<th class="wins" data-type="wins" title="Wins">W</th>'
        '<th class="losses" data-type="losses" title="Losses">L</th>'
        '<th class="goals" data-type="goals" title="Total Points">TP</th>'
        '<th class="form" data-type="form" title="Last 5 matches">Form</th>'
    )
    return (
        f'<table id="table-type-1" class="stats-table stats-main table-1"><thead>'
        f'<tr class="main">{columns}</tr></thead><tbody>{rows}</tbody></table>'
    )


def _row(
    *,
    name: str,
    rank: int = 1,
    played: int | None = 16,
    wins: int | None = 15,
    losses: int | None = 1,
    goals: str | None = "1449:1186",
    club_id: str = "IHmmS3rt",
    title: str | None = "Promotion - NKL (Winners stage: )",
) -> str:
    title_attr = f' title="{title}"' if title else ""
    cells = [
        f'<td class="rank col_rank no q1"{title_attr}>{rank}.</td>',
        f'<td class="participant_name col_participant_name col_name">'
        f'<span class="team_name_span">'
        f"<a onclick=\"javascript:getUrlByWinType('/basketball/team/x/{club_id}/');\">"
        f"{name}</a></span></td>",
    ]
    for cls, value in (
        ("matches_played", played),
        ("wins", wins),
        ("losses", losses),
    ):
        cells.append(
            f'<td class="{cls} col_{cls}">{"" if value is None else value}</td>'
        )
    cells.append(f'<td class="goals col_goals">{"" if goals is None else goals}</td>')
    cells.append('<td class="form col_form"><div class="matches-6"></div></td>')
    return (
        f'<tr class="odd glib-participant-{club_id}" data-def-order="0">'
        + "".join(cells)
        + "</tr>"
    )


# -- parsing the served bytes ------------------------------------------------


def test_parses_the_served_row_with_published_arithmetic():
    table = bs.parse_standings_table(_overall())

    telsiai = table.rows[0]
    assert isinstance(telsiai, StandingsRow)
    assert telsiai.team_name == "Telsiai"
    assert telsiai.position == 1
    assert (telsiai.games_played, telsiai.wins, telsiai.losses) == (16, 15, 1)
    assert (telsiai.points_for, telsiai.points_against) == (1449.0, 1186.0)
    # W / MP, rounded, not recalled.
    assert telsiai.win_pct == round(15 / 16, 4)
    assert telsiai.source == bs.SOURCE


def test_stage_label_is_carried_rather_than_flattened():
    table = bs.parse_standings_table(_overall())

    assert table.stage == "Promotion - NKL (Winners stage: )"
    assert table.rows[0].section == "Promotion - NKL (Winners stage: )"
    assert table.rows[1].section == "Relegation - NKL (Places 9-16)"
    assert table.report["stage"] == "Promotion - NKL (Winners stage: )"


def test_club_id_rides_alongside_the_display_name():
    table = bs.parse_standings_table(_overall())

    assert table.team_ids == {"Telsiai": "IHmmS3rt", "Synthetic Club": "SYNTH01"}


def test_form_glyph_is_not_turned_into_a_streak():
    """Only one form glyph class was ever seen, so no sign may be claimed."""

    table = bs.parse_standings_table(_overall())

    assert table.rows[0].streak is None
    assert table.rows[0].last10_pct is None


def test_two_rows_parse_independently():
    table = bs.parse_standings_table(_overall())

    assert [row.team_name for row in table.rows] == ["Telsiai", "Synthetic Club"]
    assert table.rows[1].win_pct == round(9 / 16, 4)
    assert table.report["rows_parsed"] == 2
    assert table.report["rows_skipped"] == 0


def test_columns_are_found_by_class_not_by_position():
    head = (
        '<th class="rank col_rank" title="Rank">#</th>'
        '<th class="participant_name col_name" title="Team">Team</th>'
        '<th class="matches" title="Matches Played">MP</th>'
        '<th class="wins" title="Wins">W</th>'
        '<th class="losses" title="Losses">L</th>'
        '<th class="goals" title="Total Points">TP</th>'
        '<th class="form" title="Last 5 matches">Form</th>'
    )
    # No Pts column: the capture had one, another competition may not, and the
    # parser must not have learned TP's position from it.
    table = bs.parse_standings_table(_fragment(_row(name="Telsiai"), head=head))

    assert table.rows[0].points_for == 1449.0
    assert table.report["has_points_column"] is False


def test_points_column_is_reported_when_present():
    assert bs.parse_standings_table(_overall()).report["has_points_column"] is True


# -- refusals ----------------------------------------------------------------


def test_a_fixtures_table_is_not_read_as_standings():
    fixtures = _fixture("nkl_standings_overall.html").replace(
        'class="stats-table stats-main table-1"',
        'class="table-main table-main--leaguefixtures"',
    )
    with pytest.raises(IdentityError):
        bs.parse_standings_table(fixtures)


def test_an_empty_body_names_its_own_size():
    with pytest.raises(IdentityError) as caught:
        bs.parse_standings_table("")

    assert "0 bytes" in str(caught.value)
    assert "no standings table" in str(caught.value)


def test_a_head_without_win_columns_is_refused():
    head = (
        '<th class="rank col_rank" title="Rank">#</th>'
        '<th class="participant_name col_name" title="Team">Team</th>'
        '<th class="matches" title="MP">MP</th>'
    )
    with pytest.raises(IdentityError) as caught:
        bs.parse_standings_table(_fragment(_row(name="Telsiai"), head=head))

    assert "missing one of" in str(caught.value)


def test_a_row_missing_core_cells_is_skipped_and_counted():
    rows = _row(name="Telsiai") + _row(name="No Games", played=None, wins=None, losses=None)
    table = bs.parse_standings_table(_fragment(rows))

    assert [row.team_name for row in table.rows] == ["Telsiai"]
    assert table.report["rows_skipped"] == 1
    assert table.report["skip_reasons"] == {"missing_core_cells": 1}


def test_a_self_contradicting_record_is_skipped_not_repaired():
    rows = _row(name="Telsiai") + _row(name="Impossible", wins=14, losses=1)
    table = bs.parse_standings_table(_fragment(rows))

    assert [row.team_name for row in table.rows] == ["Telsiai"]
    assert table.report["skip_reasons"] == {"inconsistent_record": 1}


def test_a_table_that_parses_to_nothing_is_an_error_not_an_empty_list():
    with pytest.raises(IdentityError) as caught:
        bs.parse_standings_table(_fragment(_row(name="Impossible", wins=14, losses=1)))

    assert "parsed to zero rows" in str(caught.value)


def test_unknown_league_is_rejected_before_any_request():
    with pytest.raises(ValueError):
        bs.competition_prefix("NOT_A_LEAGUE")


# -- finding the token on the match page -------------------------------------


def test_the_match_page_publishes_all_three_sub_tables():
    links = bs.find_standings_links(_widget())

    assert set(links) == {"overall", "home", "away"}
    assert links["overall"].startswith("https://www.betexplorer.com/basketball/lithuania/nkl/")
    assert f"ts={TOKEN}" in links["home"]
    assert f"event_context={MATCH_ID}" in links["away"]


def test_the_form_tab_is_a_link_but_never_a_sub_table():
    widget = _widget()
    assert "table=form" in widget  # the page does publish one
    assert all("table=form" not in url for url in bs.find_standings_links(widget).values())


def test_a_page_without_the_widget_reports_what_it_did_see():
    with pytest.raises(IdentityError) as caught:
        bs.fetch_standings(
            "LITHUANIA_NKL",
            lambda url: "<html><body>fixtures only</body></html>",
            match_url=MATCH_URL,
            delay=0.0,
        )

    assert "published no standings table links" in str(caught.value)
    assert "tabs seen: none" in str(caught.value)


# -- the whole read ----------------------------------------------------------


def _reader(pages: dict[str, str], calls: list[str] | None = None):
    def read(url: str) -> str:
        if calls is not None:
            calls.append(url)
        for key, page in pages.items():
            if key in url:
                return page
        raise AssertionError(f"unexpected read: {url}")

    return read


def test_one_read_per_sub_table_plus_the_match_page():
    home = _fragment(_row(name="Telsiai", played=9, wins=8, losses=1), sub="home")
    away = _fragment(_row(name="Telsiai", played=7, wins=7, losses=0), sub="away")
    calls: list[str] = []
    reader = _reader(
        {"table_sub=home": home, "table_sub=away": away, "table_sub=overall": _overall(),
         MATCH_ID: _widget()},
        calls,
    )

    result = bs.fetch_standings("LITHUANIA_NKL", reader, match_url=MATCH_URL, delay=0.0)

    assert result.report["requests"] == 4
    assert len(calls) == 4
    assert result.report["subs_read"] == ["away", "home", "overall"]
    assert result.report["token_seen"] is True
    assert result.league == "LITHUANIA_NKL"


def test_home_and_away_records_come_from_the_split_tables():
    home = _fragment(_row(name="Telsiai", played=9, wins=8, losses=1), sub="home")
    away = _fragment(_row(name="Telsiai", played=7, wins=7, losses=0), sub="away")
    reader = _reader(
        {"table_sub=home": home, "table_sub=away": away, "table_sub=overall": _overall(),
         MATCH_ID: _widget()}
    )

    rows = bs.fetch_standings(
        "LITHUANIA_NKL", reader, match_url=MATCH_URL, delay=0.0
    ).rows

    assert rows[0].home_record == "8-1"
    assert rows[0].away_record == "7-0"
    # A club only in the overall table keeps its records unset, not "0-0".
    assert rows[1].home_record is None
    assert rows[1].away_record is None


def test_a_split_that_does_not_reconstruct_the_total_is_not_attached():
    """Two fragments can straddle a game; the split must then add up or be dropped."""

    home = _fragment(_row(name="Telsiai", played=9, wins=8, losses=1), sub="home")
    away = _fragment(_row(name="Telsiai", played=6, wins=6, losses=0), sub="away")
    reader = _reader(
        {"table_sub=home": home, "table_sub=away": away, "table_sub=overall": _overall(),
         MATCH_ID: _widget()}
    )

    result = bs.fetch_standings("LITHUANIA_NKL", reader, match_url=MATCH_URL, delay=0.0)

    # 8-1 at home plus 6-0 away is 14-1, not the 15-1 the overall row states.
    assert result.rows[0].home_record is None
    assert result.rows[0].away_record is None
    assert result.report["split_mismatch"] == ["Telsiai"]


def test_a_split_that_reconstructs_the_total_is_attached_and_reported_clean():
    home = _fragment(_row(name="Telsiai", played=9, wins=8, losses=1), sub="home")
    away = _fragment(_row(name="Telsiai", played=7, wins=7, losses=0), sub="away")
    reader = _reader(
        {"table_sub=home": home, "table_sub=away": away, "table_sub=overall": _overall(),
         MATCH_ID: _widget()}
    )

    result = bs.fetch_standings("LITHUANIA_NKL", reader, match_url=MATCH_URL, delay=0.0)

    assert result.report["split_mismatch"] == []
    assert result.report["subs_missing"] == []


def test_the_overall_table_is_required():
    reader = _reader(
        {
            "table_sub=home": _fragment(_row(name="Telsiai"), sub="home"),
            MATCH_ID: _widget(),
        }
    )
    with pytest.raises(IdentityError) as caught:
        bs.fetch_standings(
            "LITHUANIA_NKL", reader, match_url=MATCH_URL, delay=0.0, subs=("home",)
        )

    assert "overall table was not requested" in str(caught.value)


def test_two_tokens_on_one_page_are_refused():
    widget = _widget().replace(
        "table_sub=home&ts=2BG7ii0m", "table_sub=home&ts=DIFFERENT"
    )
    with pytest.raises(IdentityError) as caught:
        bs.fetch_standings(
            "LITHUANIA_NKL", lambda url: widget, match_url=MATCH_URL, delay=0.0
        )

    assert "distinct standings tokens" in str(caught.value)


def test_the_token_is_never_carried_between_runs():
    """A run reads the page it was given; nothing is memoised across calls."""

    def run(token: str, calls: list[str]) -> None:
        widget = _widget().replace(f"ts={TOKEN}", f"ts={token}")
        reader = _reader(
            {"table_sub=overall": _overall(), MATCH_ID: widget}, calls
        )
        bs.fetch_standings("LITHUANIA_NKL", reader, match_url=MATCH_URL, delay=0.0,
                           subs=("overall",))

    first: list[str] = []
    second: list[str] = []
    run("FIRST123", first)
    run("SECOND99", second)

    assert any("ts=FIRST123" in url for url in first)
    assert any("ts=SECOND99" in url for url in second)
    assert not any("FIRST123" in url for url in second)


# -- pacing ------------------------------------------------------------------


def test_the_delay_is_between_requests_never_before_the_first():
    slept: list[float] = []
    reader = bs.paced(lambda url: "x", delay=10.0, sleep=slept.append)

    reader("a")
    assert slept == []

    reader("b")
    reader("c")
    assert slept == [10.0, 10.0]
    assert reader.calls == 3


def test_the_published_floor_is_ten_seconds():
    assert bs.CRAWL_DELAY_SECONDS == 10.0
    assert bs.paced(lambda url: "x", sleep=lambda _: None)._delay == 10.0


def test_fetch_standings_uses_the_paced_reader():
    slept: list[float] = []
    pages = {
        "table_sub=overall": _overall(),
        MATCH_ID: _widget(),
    }
    reader = _reader(pages)

    result = bs.fetch_standings(
        "LITHUANIA_NKL",
        reader,
        match_url=MATCH_URL,
        delay=10.0,
        sleep=slept.append,
        subs=("overall",),
    )

    assert result.report["requests"] == 2
    assert slept == [10.0]
