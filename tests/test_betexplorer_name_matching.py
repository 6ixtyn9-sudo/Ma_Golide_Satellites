"""BetExplorer board-name matching, checked against served board spellings.

Every ``(canonical, board)`` pair below uses a board name **quoted verbatim from a
competition page captured on 2026-11-23** (out-of-sandbox probe) against the name
the corresponding Ma Golide adapter produces. These are the real joins the price
pipeline depends on: a name miss is not an error, it is a price that silently
does not attach.

The two hard cases this file exists for:

* ``Karşıyaka`` — BetExplorer publishes ASCII, and the Turkish dotless ``ı`` has
  no combining mark, so NFKD alone leaves ``karsıyaka`` and the row never joins.
* ``Élan Chalon`` / ``Chalon/Saone`` — structurally different names with no shared
  token subset in either direction; only an explicit, league-scoped pair can join
  them, and that pair is a standing invitation to re-verify each season.
"""

from __future__ import annotations

import pytest

from ingestion.providers.betexplorer_direct import (
    _LEAGUE_NAME_PAIRS,
    _team_names_agree,
)

BOARD_DATE = "2026-12-05"

#: (league, canonical name from our adapter, board spelling as served)
REAL_PAIRS = [
    ("SPAIN_ACB", "Surne Bilbao Basket", "Bilbao"),
    ("SPAIN_ACB", "Lenovo Tenerife", "Tenerife"),
    ("SPAIN_ACB", "Força Lleida", "Forca Lleida"),
    ("SPAIN_ACB", "Saski Baskonia", "Baskonia"),
    ("SPAIN_ACB", "UCAM Murcia", "Murcia"),
    ("ITALY_LBA", "Umana Reyer Venezia", "Venezia"),
    ("ITALY_LBA", "Bertram Derthona Tortona", "Tortona"),
    ("TURKEY_BSL", "Karşıyaka", "Karsiyaka"),
    ("TURKEY_BSL", "Esenler Erokspor", "Esenler Erokspor"),
    ("FRANCE_LNB_PRO_A", "Élan Chalon", "Chalon/Saone"),
    ("FRANCE_LNB_PRO_A", "Gravelines-Dunkerque", "Gravelines-Dunkerque"),
    ("GERMANY_BBL", "Skyliners Frankfurt", "Frankfurt"),
    ("GERMANY_BBL", "Würzburg Baskets", "Wurzburg"),
    ("GREECE_BASKET_LEAGUE", "Promitheas Patras", "Promitheas"),
    ("GREECE_BASKET_LEAGUE", "PAOK", "PAOK"),
    ("ISRAEL_PREMIER_LEAGUE", "Hapoel Tel Aviv", "Hapoel Tel Aviv"),
    ("JAPAN_B_LEAGUE", "Levanga Hokkaido", "Hokkaido"),
    ("JAPAN_B_LEAGUE", "Koshigaya Alphas", "Koshigaya Alphas"),
    ("SOUTH_KOREA_KBL", "Goyang Sono Skygunners", "Goyang Sono"),
    ("SOUTH_KOREA_KBL", "Seoul SK Knights", "Seoul Knights"),
    ("BRAZIL_NBB", "Caxias do Sul", "Caxias do Sul"),
    ("ARGENTINA_LIGA_NACIONAL", "Ferro Carril Oeste", "Ferro"),
    ("ARGENTINA_LIGA_NACIONAL", "Racing de Chivilcoy", "Racing de Chivilcoy"),
    ("BASKETBALL_CHAMPIONS_LEAGUE", "BC Rytas", "Rytas"),
    ("BASKETBALL_CHAMPIONS_LEAGUE", "Sabah BC", "Sabah"),
    ("BASKETBALL_AFRICA_LEAGUE", "RSSB Tigers", "RSSB Tigers"),
    ("BASKETBALL_AFRICA_LEAGUE", "Petro Atletico", "Petro Atletico"),
]


@pytest.mark.parametrize("league,canonical,board", REAL_PAIRS)
def test_served_board_names_join_their_canonical_club(league, canonical, board):
    ok, method = _team_names_agree(
        canonical, board, league=league, board_date=BOARD_DATE
    )
    assert ok, f"{canonical!r} did not join board name {board!r} in {league} ({method!r})"


@pytest.mark.parametrize("league,canonical,board", REAL_PAIRS)
def test_agreement_is_symmetric(league, canonical, board):
    forward, _ = _team_names_agree(canonical, board, league=league, board_date=BOARD_DATE)
    backward, _ = _team_names_agree(board, canonical, league=league, board_date=BOARD_DATE)
    assert forward == backward is True


def test_turkish_dotless_i_folds_to_ascii():
    """``ı`` has no combining mark, so NFKD alone is not enough."""

    from ingestion.identity import normalize_team_name

    assert normalize_team_name("Karşıyaka") == normalize_team_name("Karsiyaka")
    assert normalize_team_name("Şafak") == normalize_team_name("Safak")


def test_structural_pairs_are_league_scoped_not_global():
    """A global alias would let ``Chalon`` match in a league that never uses it."""

    assert ("FRANCE_LNB_PRO_A", "elanchalon", "chalonsaone") in _LEAGUE_NAME_PAIRS
    ok, method = _team_names_agree(
        "Élan Chalon", "Chalon/Saone", league="SPAIN_ACB", board_date=BOARD_DATE
    )
    assert ok is False
    ok, method = _team_names_agree(
        "Élan Chalon", "Chalon/Saone", league="FRANCE_LNB_PRO_A", board_date=BOARD_DATE
    )
    assert (ok, method) == (True, "league_alias")


def test_unrelated_clubs_do_not_join():
    for league, left, right in [
        ("SPAIN_ACB", "Real Madrid", "Barcelona"),
        ("ITALY_LBA", "Virtus Bologna", "Olimpia Milano"),
        ("TURKEY_BSL", "Anadolu Efes", "Fenerbahce"),
        ("SOUTH_KOREA_KBL", "Seoul SK Knights", "Busan KCC Egis"),
    ]:
        ok, _ = _team_names_agree(left, right, league=league, board_date=BOARD_DATE)
        assert ok is False, f"{left!r} joined {right!r} in {league}"


def test_numeric_squad_markers_are_not_folded_away():
    """Under-21 sides must not join the senior club of the same name."""

    ok, _ = _team_names_agree("Fenerbahce U21", "Fenerbahce", league="TURKEY_BSL", board_date=BOARD_DATE)
    assert ok is False
