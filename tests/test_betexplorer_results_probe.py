"""The results capture tool: bounded, verbatim, and honest about not parsing.

There are no committed match-page bodies yet — that is precisely why the
parser does not exist.  These tests therefore use a **synthetic** page and say
so: they prove the tool's contract (bounded requests, verbatim bytes with a
digest, no invented result), not that any markup has been recognised.
"""

from __future__ import annotations

from datetime import datetime, timezone
import hashlib
import json
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))

from probe_betexplorer_results import (  # noqa: E402
    MAX_CAPTURE,
    MIN_INTERVAL_SECONDS,
    capture,
    finished_candidates,
    index_signals,
    match_url,
    results_url,
    shape_signals,
)

NOW = datetime(2026, 11, 25, 12, 0, tzinfo=timezone.utc)

#: Synthetic, minimal, and labelled as such: the fixtures row shape is the one
#: the committed capture uses, the match page is a stand-in that exists only to
#: exercise the capture contract.
FIXTURES = (
    '<table class="table-main">'
    "<tr><th>Round</th></tr>"
    '<tr><td class="table-main__datetime">24.11. 18:00</td>'
    '<td class="h-text-left"><a href="/basketball/greece/basket-league/'
    'olympiacos-panathinaikos/AbCd1234/" class="in-match">'
    "<span>Olympiacos</span> - <span>Panathinaikos</span></a></td>"
    '<td class="table-main__odds" data-odd="1.85"></td>'
    '<td class="table-main__odds" data-odd="1.95"></td></tr>'
    '<tr><td class="table-main__datetime">26.11. 18:00</td>'
    '<td class="h-text-left"><a href="/basketball/greece/basket-league/'
    'aris-paok/ZzYy9876/" class="in-match">'
    "<span>Aris</span> - <span>PAOK</span></a></td></tr>"
    "</table>"
)

MATCH_PAGE = (
    "<html><body><table class='table-main'>"
    "<tr><td>1</td><td>24</td><td>22</td></tr>"
    "<tr><td>2</td><td>19</td><td>21</td></tr>"
    "<tr><td>F</td><td>89</td><td>75</td></tr>"
    "</table></body></html>"
)


def test_match_url_is_a_path_on_the_verified_slug():
    url = match_url("GREECE_BASKET_LEAGUE", "AbCd1234")

    assert url.startswith("https://www.betexplorer.com/basketball/greece/")
    assert url.endswith("/AbCd1234/")
    assert "?" not in url, "query shapes are robots-disallowed"
    assert "fixtures" not in url


def test_only_games_far_enough_past_tip_are_candidates():
    rows = finished_candidates("GREECE_BASKET_LEAGUE", FIXTURES, now=NOW, limit=5)

    # The 24.11 game is final by NOW; the 26.11 game is in the future.
    assert [row["event_id"] for row in rows] == ["AbCd1234"]


def test_capture_saves_verbatim_bytes_and_a_digest(tmp_path: Path):
    urls: list[str] = []

    def read_text(url: str) -> str:
        urls.append(url)
        return FIXTURES if url.endswith("/fixtures/") else MATCH_PAGE

    sleeps: list[float] = []
    payload = capture(
        "GREECE_BASKET_LEAGUE",
        max_games=4,
        date=None,
        out_dir=tmp_path / "probe",
        now=NOW,
        read_text=read_text,
        sleep=sleeps.append,
    )

    assert payload["captured"] == 1
    saved = Path(payload["games"][0]["path"])
    assert saved.read_text(encoding="utf-8") == MATCH_PAGE
    assert payload["games"][0]["sha256"] == hashlib.sha256(MATCH_PAGE.encode()).hexdigest()
    assert payload["games"][0]["bytes"] == len(MATCH_PAGE.encode())
    # One fixtures read plus one match page: no speculative extra requests.
    assert len(urls) == 2
    # The index on disk is the artifact a later parser is written against.
    index = json.loads((tmp_path / "probe" / "index.json").read_text())
    assert index["schema_version"] == "magolide.results_probe.v1"
    assert index["games"][0]["status"] == "captured"


def test_capture_produces_no_result_of_any_kind(tmp_path: Path):
    payload = capture(
        "GREECE_BASKET_LEAGUE",
        max_games=2,
        date=None,
        out_dir=tmp_path / "probe",
        now=NOW,
        read_text=lambda url: FIXTURES if url.endswith("/fixtures/") else MATCH_PAGE,
        sleep=lambda _s: None,
    )

    for game in payload["games"]:
        assert game["status"] in {"captured", "fetch_failed"}
        # No score field of any kind: the tool cannot invent one, so it must
        # not carry a slot that looks like it might have.
        assert not {"home_score", "away_score", "home_final", "away_final", "actual"} & set(game)
    assert payload["note"].startswith("Bodies are served pages saved verbatim")
    assert "reconciles()" in payload["note"]


def test_request_budget_is_hard_capped():
    assert capture.__defaults__ is None  # explicit kwargs only
    assert MAX_CAPTURE <= 12
    assert MIN_INTERVAL_SECONDS >= 2.0


def test_budget_is_respected_across_many_finished_games(tmp_path: Path):
    rows = "".join(
        f'<tr><td class="table-main__datetime">2{d}.11. 18:00</td>'
        f'<td class="h-text-left"><a href="/basketball/greece/basket-league/'
        f'home{d}-away{d}/Ev{d:06d}/" class="in-match">'
        f"<span>Home{d}</span> - <span>Away{d}</span></a></td></tr>"
        for d in range(1, 10)
    )
    page = f'<table class="table-main">{rows}</table>'
    calls: list[str] = []

    def read_text(url: str) -> str:
        calls.append(url)
        return page if url.endswith("/fixtures/") else MATCH_PAGE

    payload = capture(
        "GREECE_BASKET_LEAGUE",
        max_games=3,
        date=None,
        out_dir=tmp_path / "probe",
        now=NOW,
        read_text=read_text,
        sleep=lambda _s: None,
    )

    assert payload["captured"] == 3
    assert len([url for url in calls if not url.endswith("/fixtures/")]) == 3


def test_a_failed_capture_is_a_receipt_not_a_crash(tmp_path: Path):
    def read_text(url: str) -> str:
        if url.endswith("/fixtures/"):
            return FIXTURES
        raise RuntimeError("HTTP 429")

    payload = capture(
        "GREECE_BASKET_LEAGUE",
        max_games=2,
        date=None,
        out_dir=tmp_path / "probe",
        now=NOW,
        read_text=read_text,
        sleep=lambda _s: None,
    )

    assert payload["captured"] == 0
    assert payload["games"][0]["status"] == "fetch_failed"
    assert "429" in payload["games"][0]["detail"]


def test_shape_signals_describe_markup_without_interpreting_it():
    signals = shape_signals(MATCH_PAGE)

    assert signals["chars"] == len(MATCH_PAGE)
    assert signals["td_cells"] > 0
    assert signals["integer_cells"] >= 6
    assert signals["first_integer_cells"]
    # Counts only: no key may claim a score.
    assert not {"home_score", "away_score", "final"} & set(signals)


def test_the_results_index_url_is_the_same_slug_without_a_query():
    """Path-only, like every other page this tool reads: robots allows paths."""

    from ingestion.providers.betexplorer import FIXTURES_URLS

    url = results_url("AUSTRALIA_WNBL_W")
    assert url == FIXTURES_URLS["AUSTRALIA_WNBL_W"].replace("fixtures/", "results/")
    assert url.endswith("/australia/wnbl-women/results/")
    assert "?" not in url


def test_the_index_signals_count_what_the_page_offers_without_parsing_a_score():
    body = (
        '<table class="table-main">'
        '<tr><td><a href="/basketball/australia/wnbl-women/perth-sydney/AbCd1234/">'
        "Perth - Sydney</a></td><td>89:84</td></tr>"
        '<tr><td><a href="/basketball/australia/wnbl-women/melbourne-adelaide/EfGh5678/">'
        "Melbourne - Adelaide</a></td><td>77 : 80</td></tr>"
        "</table>"
    )

    signals = index_signals(body)

    assert signals["has_table"] is True
    assert signals["match_codes"] == 2
    assert signals["distinct_match_codes"] == 2
    assert signals["colon_scores"] == 2
    assert signals["first_codes"] == ["AbCd1234", "EfGh5678"]
    assert signals["has_no_matches_hint"] is False


def test_an_empty_results_page_says_so_rather_than_looking_healthy():
    signals = index_signals('<table class="table-main"><p>No matches found</p></table>')

    assert signals["match_codes"] == 0
    assert signals["has_no_matches_hint"] is True
