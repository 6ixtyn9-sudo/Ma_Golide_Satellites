"""The date+tip join: the one identity question that needs no club names.

The federation names clubs in Greek and this lane's canonical games carry the
board's Latin names, so a name join would need one hand-made transliteration per
club. The clock is better evidence than a name: a game's own tip, to the minute,
in Europe/Athens, either identifies it uniquely among the finals published for
that date or it does not. When it does not, the join must say so.

What is asserted here:

* a UTC tip converts to the Athens wall clock the federation publishes, across
  both Greek offsets (EET +02:00, EEST +03:00);
* a unique minute attaches;
* two finals at the same minute are **refused** -- never resolved by picking one;
* a same-date final at a different minute is a different answer from "no date at
  all", because the two failures need different fixes;
* an already-settled game is not re-joined.
"""

from __future__ import annotations

from datetime import date

import pytest

from ingestion.providers.esake import ResultRow
from scripts.probe_esake_results import _athens_tip, _join


def _final(idgame: str, played_on: date, tip: str | None, score=(80, 70)) -> ResultRow:
    return ResultRow(
        idgame=idgame,
        home="ΑΡΗΣ",
        away="ΠΑΟΚ",
        home_score=score[0],
        away_score=score[1],
        played_on=played_on,
        tip_local=tip,
        url=f"https://www.esake.gr/el/action/EsakegameView?idgame={idgame}&mode=2",
    )


def _game(game_id: str, scheduled_utc: str, *, settled: bool = False) -> dict:
    return {
        "game_id": game_id,
        "scheduled_utc": scheduled_utc,
        "home": "Aris",
        "away": "PAOK",
        "settled": settled,
    }


@pytest.mark.parametrize(
    "utc, expected",
    [
        ("2026-11-21T13:00:00Z", ("2026-11-21", "15:00")),  # EET, +02:00
        ("2026-09-15T12:00:00Z", ("2026-09-15", "15:00")),  # EEST, +03:00
        ("2026-11-21T21:30:00Z", ("2026-11-21", "23:30")),
        ("2026-11-21T22:30:00Z", ("2026-11-22", "00:30")),  # crosses the date line
    ],
)
def test_a_utc_tip_becomes_the_athens_wall_clock_the_federation_publishes(utc, expected):
    assert _athens_tip(utc) == expected


def test_an_unparseable_tip_is_absent_rather_than_guessed():
    assert _athens_tip("") is None
    assert _athens_tip("not-a-time") is None


def test_a_unique_minute_attaches():
    report = _join(
        [_game("betexplorer:abc", "2026-11-21T13:00:00Z")],
        [_final("1439C9D6", date(2026, 11, 21), "15:00")],
    )

    assert report["attachable"] == 1
    assert report["ambiguous"] == 0
    assert report["examples"][0]["idgame"] == "1439C9D6"
    assert report["examples"][0]["athens_tip"] == "2026-11-21 15:00"


def test_two_finals_at_the_same_minute_are_refused_not_picked():
    report = _join(
        [_game("betexplorer:abc", "2026-11-21T13:00:00Z")],
        [
            _final("AAAAAAAA", date(2026, 11, 21), "15:00"),
            _final("BBBBBBBB", date(2026, 11, 21), "15:00"),
        ],
    )

    assert report["attachable"] == 0
    assert report["ambiguous"] == 1
    assert report["examples"][0]["note"].startswith("refused")
    assert report["examples"][0]["candidates"] == ["AAAAAAAA", "BBBBBBBB"]


def test_a_same_date_game_at_another_minute_is_a_different_answer():
    report = _join(
        [_game("betexplorer:abc", "2026-11-21T15:30:00Z")],  # 17:30 Athens
        [_final("1439C9D6", date(2026, 11, 21), "15:00")],
    )

    assert report["attachable"] == 0
    assert report["same_date_no_tip"] == 1
    assert report["unmatched"] == 0


def test_a_date_the_federation_does_not_publish_is_unmatched():
    report = _join(
        [_game("betexplorer:abc", "2026-11-30T13:00:00Z")],
        [_final("1439C9D6", date(2026, 11, 21), "15:00")],
    )

    assert report["unmatched"] == 1


def test_a_settled_game_is_not_joined_again():
    report = _join(
        [_game("betexplorer:abc", "2026-11-21T13:00:00Z", settled=True)],
        [_final("1439C9D6", date(2026, 11, 21), "15:00")],
    )

    assert report["attachable"] == 0
    assert report["stored_unsettled"] == 0


def test_a_final_without_a_tip_is_counted_apart():
    report = _join(
        [_game("betexplorer:abc", "2026-11-21T13:00:00Z")],
        [_final("1439C9D6", date(2026, 11, 21), None)],
    )

    assert report["finals_without_tip"] == 1
    assert report["attachable"] == 0
    assert report["unmatched"] == 1
