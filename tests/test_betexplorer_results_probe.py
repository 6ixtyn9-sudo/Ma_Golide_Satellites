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

import probe_betexplorer_results as probe  # noqa: E402
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


def test_the_first_match_window_is_the_row_not_the_page():
    """A count says the page has 34 codes; the window says what a row looks like."""

    filler = "<td>padding</td>" * 200
    body = (
        '<table class="table-main"><tr><td>header</td></tr>'
        + filler
        + '<tr><td><a href="/basketball/australia/wnbl-women/perth-sydney/AbCd1234/">'
        "Perth - Sydney</a></td><td>89:84</td></tr>"
        + filler
        + "</table>"
    )

    window = probe.first_match_window(body)

    assert "AbCd1234" in window
    assert "89:84" in window
    assert len(window) <= probe.WINDOW_CHARS
    assert len(window) < len(body)
    # The row is the point of the file: it sits near the start, not past the
    # middle, because a capture that ends before the row teaches nothing.
    assert window.index("AbCd1234") < 500


def test_a_page_with_no_match_code_has_no_window_rather_than_a_guess():
    assert probe.first_match_window('<table class="table-main"><p>No matches found</p></table>') == ""


def test_read_replays_a_captured_page_without_a_request(tmp_path):
    """`--read` is the replay: the saved bytes, no network, honest counts."""

    league = "AUSTRALIA_WNBL_W"
    saved = tmp_path / league
    saved.mkdir(parents=True)
    (saved / "results-index.html").write_text(
        '<table class="table-main"><tr>'
        '<td class="h-text-left">'
        '<a data-test="11015187" href="/basketball/australia/wnbl-women/perth-sydney/AbCd1234/"'
        ' class="in-match"><span><strong>Perth</strong></span> - <span>Sydney</span></a></td>'
        '<td class="h-text-center">'
        '<a href="/basketball/australia/wnbl-women/perth-sydney/AbCd1234/">89:84</a></td>'
        "</tr></table>",
        encoding="utf-8",
    )

    assert probe.main(["--league", league, "--read", "--out", str(tmp_path)]) == 0


def test_the_index_codes_come_back_distinct_and_in_order():
    body = (
        '<a href="/basketball/australia/wnbl-women/a-b/AbCd1234/">x</a>'
        '<a href="/basketball/australia/wnbl-women/a-b/AbCd1234/">x</a>'
        '<a href="/basketball/australia/wnbl-women/a-b/EfGh5678/">x</a>'
    )

    assert probe.index_codes(body) == ["AbCd1234", "EfGh5678"]


def test_match_pages_are_fetched_bounded_spaced_and_receipted(tmp_path):
    """One request per league answers "can this host settle?" -- no more."""

    fetched = []
    slept = []

    def read_text(url):
        fetched.append(url)
        return (
            "<table><tr><td>Q1</td><td>22</td><td>18</td></tr>"
            "<tr><td>Q2</td><td>19</td><td>21</td></tr></table>"
        )

    payload = probe.capture_match_pages(
        "AUSTRALIA_WNBL_W",
        matches=[
            {"code": "AbCd1234", "url": "/basketball/australia/wnbl-women/a-b/AbCd1234/"},
            {"code": "EfGh5678", "url": "/basketball/australia/wnbl-women/c-d/EfGh5678/"},
            {"code": "IjKl9012", "url": "/basketball/australia/wnbl-women/e-f/IjKl9012/"},
        ],
        out_dir=tmp_path,
        read_text=read_text,
        sleep=slept.append,
        limit=2,
    )

    assert len(fetched) == 2
    assert slept == [probe.MATCH_INTERVAL_SECONDS]  # one gap between two requests
    assert payload["requested"] == 2
    assert [page["status"] for page in payload["pages"]] == ["captured", "captured"]
    assert payload["pages"][0]["period_markers"] == 2
    assert payload["pages"][0]["period_examples"] == ["Q1", "Q2"]
    assert (tmp_path / "match_AbCd1234.html").exists()
    assert (tmp_path / "match-pages.json").exists()


def test_a_failed_match_fetch_is_a_receipt_not_a_crash(tmp_path):
    def read_text(url):
        raise RuntimeError("429 from the host")

    payload = probe.capture_match_pages(
        "AUSTRALIA_WNBL_W",
        matches=[{"code": "AbCd1234", "url": "/basketball/australia/wnbl-women/a-b/AbCd1234/"}],
        out_dir=tmp_path,
        read_text=read_text,
        limit=1,
    )

    assert payload["pages"][0]["status"] == "fetch_failed"
    assert "429" in payload["pages"][0]["detail"]


def test_a_nav_tab_is_not_a_match_code():
    """`fixtures` is eight lowercase letters and matched the code pattern.

    The first `--match 1` run opened `/…/nkl/fixtures/` and reported a captured
    page for `event_id: fixtures` -- a success that taught nothing.
    """

    body = (
        '<a href="/basketball/lithuania/nkl/results/" class="list-tabs__item__in">Results</a>'
        '<a href="/basketball/lithuania/nkl/fixtures/" class="list-tabs__item__in">Fixtures</a>'
        '<a href="/basketball/lithuania/nkl/olimpas-palanga-omega-tauras/IeUABOGO/">x</a>'
    )

    assert probe.index_codes(body) == ["IeUABOGO"]
    assert probe.index_matches(body) == [
        {"code": "IeUABOGO", "url": "/basketball/lithuania/nkl/olimpas-palanga-omega-tauras/IeUABOGO/"}
    ]


def test_a_relative_href_from_the_index_is_made_absolute():
    """`No scheme supplied` is a tool fault wearing the look of a host refusal.

    The first `--match 1` run on the corrected code handed the client
    `/basketball/...` verbatim and got `HTTP 0 ... Invalid URL`. The receipt said
    so, which is the only reason it took one run to see.
    """

    assert probe.absolute_url("/basketball/australia/wnbl-women/a-b/AbCd1234/") == (
        "https://www.betexplorer.com/basketball/australia/wnbl-women/a-b/AbCd1234/"
    )
    assert probe.absolute_url("https://example.com/x") == "https://example.com/x"


def test_match_fetches_use_absolute_urls(tmp_path):
    seen = []

    def read_text(url):
        seen.append(url)
        return "<table><tr><td>Q1</td></tr></table>"

    probe.capture_match_pages(
        "AUSTRALIA_WNBL_W",
        matches=[{"code": "EaZEfXoI", "url": "/basketball/australia/wnbl-women/a-b/EaZEfXoI/"}],
        out_dir=tmp_path,
        read_text=read_text,
        limit=1,
    )

    assert seen == [
        "https://www.betexplorer.com/basketball/australia/wnbl-women/a-b/EaZEfXoI/"
    ]


def test_a_429_stops_the_run_rather_than_arguing_with_the_host(tmp_path):
    """One landed page, one throttle, and no third request.

    Three match pages a few seconds apart drew a 429 on 2026-11-24. The host was
    naming its own limit; the receipt keeps it and the run ends.
    """

    fetched = []

    def read_text(url):
        fetched.append(url)
        if len(fetched) == 2:
            raise RuntimeError("HTTP 429 for %s: " % url)
        return "<html><body>ok</body></html>"

    payload = probe.capture_match_pages(
        "AUSTRALIA_WNBL_W",
        matches=[
            {"code": "AbCd1234", "url": "/basketball/australia/wnbl-women/a-b/AbCd1234/"},
            {"code": "EfGh5678", "url": "/basketball/australia/wnbl-women/c-d/EfGh5678/"},
            {"code": "IjKl9012", "url": "/basketball/australia/wnbl-women/e-f/IjKl9012/"},
        ],
        out_dir=tmp_path,
        read_text=read_text,
        sleep=lambda _: None,
        limit=3,
    )

    assert len(fetched) == 2  # the third was never asked for
    assert payload["stopped_on_429"] is True
    assert [page["status"] for page in payload["pages"]] == ["captured", "fetch_failed"]


def test_an_age_gate_is_saved_as_evidence_but_never_counted_as_a_match(tmp_path):
    """The failure mode this catches: 110 KB that reports as a captured page.

    The 2026-11-24 match "captures" were `<h1 id="ageTitle">Verificação de idade
    obrigatória</h1>` — an 18+ modal, not a scoreboard. Size said yes; the page
    said nothing.
    """

    gate = (
        '<div id="ageModal" class="age__modal">'
        '<h1 id="ageTitle">Verificação de idade obrigatória</h1></div>'
    )

    payload = probe.capture_match_pages(
        "LITHUANIA_NKL",
        matches=[{"code": "nRzm7OsT", "url": "/basketball/lithuania/nkl/a-b/nRzm7OsT/"}],
        out_dir=tmp_path,
        read_text=lambda url: gate,
        limit=1,
    )

    page = payload["pages"][0]
    assert page["status"] == "gated"
    assert page["gate"].startswith("age gate")
    assert payload["gated"] == 1
    assert (tmp_path / "match_nRzm7OsT.html").exists()  # still evidence


def test_a_real_body_is_not_mistaken_for_a_gate():
    assert probe.gate_reason("<table><tr><td>Q1</td></tr></table>") is None


def test_the_cooldown_holds_off_between_runs_not_only_inside_one(tmp_path):
    """Three runs a few seconds apart are a burst; the stamp is on disk so the
    next run sees what the last one did."""

    stamp = tmp_path / ".host-cooldown"
    slept = []

    first = probe.HostCooldown(stamp, seconds=30)
    assert first.wait(now=1000.0, sleep=slept.append) == 0.0
    assert slept == []

    second = probe.HostCooldown(stamp, seconds=30)
    assert second.wait(now=1005.0, sleep=slept.append) == 25.0
    assert slept == [25.0]

    # The second run acted at 1030, so the third may not act before 1060: asking
    # at 1050 owes ten seconds. The gap between actions stays 30, which is the
    # property that matters -- a wait cannot be banked and spent early.
    third = probe.HostCooldown(stamp, seconds=30)
    assert third.wait(now=1050.0, sleep=slept.append) == 10.0
    assert slept == [25.0, 10.0]


def test_a_capture_reports_the_wait_it_served(tmp_path):
    stamp = tmp_path / ".host-cooldown"
    probe.HostCooldown(stamp, seconds=30).wait(now=probe.time.time(), sleep=lambda _: None)

    payload = probe.capture_match_pages(
        "AUSTRALIA_WNBL_W",
        matches=[{"code": "AbCd1234", "url": "/basketball/australia/wnbl-women/a-b/AbCd1234/"}],
        out_dir=tmp_path,
        read_text=lambda url: "<html>ok</html>",
        sleep=lambda _: None,
        limit=1,
        cooldown=probe.HostCooldown(stamp, seconds=30),
    )

    assert 29.0 <= payload["cooldown_seconds_waited"] <= 30.0
    assert payload["age_gate_authorisation"].startswith("Operator authorisation")


def test_gate_hints_read_what_the_accept_button_writes():
    """Passing a gate means reproducing what it records, not guessing its name."""

    body = (
        "<script>"
        "document.getElementById('ageYes').onclick=function(){"
        "document.cookie='be_age_verified=1; path=/; max-age=31536000';"
        "localStorage.setItem('ageAccepted', 'true');"
        "};</script>"
    )

    hints = probe.gate_hints(body)

    assert "be_age_verified=1; path=/; max-age=31536000" in hints
    assert "ageAccepted" in hints
    # A class name is not a way in. It is reported, but never as a hint.
    assert "age__modal" not in hints


_GATED_COOKIE_PAGE = (
    '<div class="age__modal"><h1 id="ageTitle">Verificação de idade obrigatória</h1>'
    '<button onclick="document.cookie=\'age_verified=1; path=/; max-age=31536000\';">'
    "Entrar</button></div>"
)

_GATED_STORAGE_PAGE = (
    '<div class="age__modal"><h1 id="ageTitle">Verificação de idade obrigatória</h1>'
    "<button onclick=\"localStorage.setItem('ageVerified','true');\">Entrar</button></div>"
)

_MATCH_PAGE = "<table><tr><td>Q1</td><td>88</td><td>Q2</td><td>91</td></tr></table>"


def test_the_accept_button_writes_a_cookie_we_can_replay(tmp_path):
    """Passing the gate means sending back what it recorded, not guessing names."""

    assert probe.gate_cookies(_GATED_COOKIE_PAGE) == [("age_verified", "1")]

    seen = []

    def gate_read_text(url, cookie):
        seen.append(cookie)
        return _MATCH_PAGE

    payload = probe.capture_match_pages(
        "LITHUANIA_NKL",
        matches=[{"code": "nRzm7OsT", "url": "/basketball/lithuania/nkl/a-b/nRzm7OsT/"}],
        out_dir=tmp_path,
        read_text=lambda url: _GATED_COOKIE_PAGE,
        limit=1,
        gate_accept=True,
        gate_read_text=gate_read_text,
    )

    page = payload["pages"][0]
    assert seen == ["age_verified=1"]
    assert page["status"] == "captured"
    assert page["gate_passed_by"] == ["age_verified"]
    assert payload["gated"] == 0
    assert payload["requests_made"] == 2  # one gate, one replay: never a loop
    assert (tmp_path / "match_nRzm7OsT.html").read_text() == _MATCH_PAGE


def test_a_storage_only_gate_says_it_needs_a_browser(tmp_path):
    """localStorage cannot travel in a header. Claiming otherwise would be a lie
    that only shows up as `gated: 0` with an empty page."""

    calls = []

    def gate_read_text(url, cookie):  # pragma: no cover - must not be reached
        calls.append(cookie)
        return _MATCH_PAGE

    payload = probe.capture_match_pages(
        "CHINA_WCBA_W",
        matches=[{"code": "tOWc8ZhA", "url": "/basketball/china/wcba-women/a-b/tOWc8ZhA/"}],
        out_dir=tmp_path,
        read_text=lambda url: _GATED_STORAGE_PAGE,
        limit=1,
        gate_accept=True,
        gate_read_text=gate_read_text,
    )

    page = payload["pages"][0]
    assert calls == []
    assert payload["requests_made"] == 1
    assert page["status"] == "gated"
    assert "needs a browser" in page["gate"]
    assert "ageVerified" in page["gate"]


def test_accepting_the_gate_is_off_unless_asked_for(tmp_path):
    """The authority travels with the capture; the act of using it does not
    happen by accident."""

    payload = probe.capture_match_pages(
        "LITHUANIA_NKL",
        matches=[{"code": "nRzm7OsT", "url": "/basketball/lithuania/nkl/a-b/nRzm7OsT/"}],
        out_dir=tmp_path,
        read_text=lambda url: _GATED_COOKIE_PAGE,
        limit=1,
    )

    assert payload["gate_accept"] is False
    assert payload["requests_made"] == 1
    assert payload["pages"][0]["status"] == "gated"


def test_the_gate_hints_read_needs_no_league(tmp_path, capsys):
    """Reading a saved page is offline work; demanding a lane for it made the one
    command that answers the gate question fail with a usage error."""

    page = tmp_path / "gate.html"
    page.write_text(_GATED_COOKIE_PAGE, encoding="utf-8")

    assert probe.main(["--gate-hints", str(page)]) == 0

    payload = json.loads(capsys.readouterr().out)
    assert payload["gate"].startswith("age gate")
    assert payload["hints"] == ["age_verified=1; path=/; max-age=31536000"]
    assert "age__modal" in payload["mentions"]  # context, reported as such
    assert payload["replayable"] is True
    assert payload["age_gate_authorisation"].startswith("Operator authorisation")


#: The shape the real page turned out to have (2026-11-24): a modal, and far from
#: it a language switcher that clears ``lang`` plus a bare ``mode`` write. Both are
#: cookie writes. Neither is consent -- replaying them returned the same 110 KB gate.
_GATE_WITH_UNRELATED_WRITES = (
    '<div class="age__modal"><h1 id="ageTitle">Verificação de idade obrigatória</h1>'
    '<button id="ageYes">Sim</button><button id="ageNo">Não</button></div>'
    + "<p>banner</p>" * 400
    + "<script>document.cookie=\"lang=; expires=Thu, 01 Jan 1970 00:00:00 GMT; path=/\";"
    'document.cookie="mode";</script>'
)


def test_a_cookie_write_far_from_the_gate_is_not_a_hint():
    writes = probe.gate_cookie_writes(_GATE_WITH_UNRELATED_WRITES)

    assert [write["name"] for write in writes] == ["lang", "mode"]
    assert all(write["near_gate"] is False for write in writes)
    assert probe.gate_cookies(_GATE_WITH_UNRELATED_WRITES) == []


def test_the_cli_calls_a_language_reset_what_it_is(tmp_path, capsys):
    page = tmp_path / "gate.html"
    page.write_text(_GATE_WITH_UNRELATED_WRITES, encoding="utf-8")

    assert probe.main(["--gate-hints", str(page)]) == 1  # nothing replayable, and says so

    payload = json.loads(capsys.readouterr().out)
    assert payload["hints"] == []
    assert payload["replayable"] is False
    assert payload["unrelated_writes"][0]["raw"].startswith("lang=; expires=Thu, 01 Jan 1970")
    assert payload["unrelated_writes"][0]["distance"] > probe.GATE_PROXIMITY_CHARS


def test_accepting_sends_nothing_when_the_only_writes_are_unrelated(tmp_path):
    """One request, and a reason a human can check -- not a header full of noise
    that comes back gated and reads like the host refusing."""

    def must_not_be_called(url, cookie):  # pragma: no cover - must not be reached
        raise AssertionError(f"replayed {cookie!r} for a write the gate never made")

    payload = probe.capture_match_pages(
        "LITHUANIA_NKL",
        matches=[{"code": "nRzm7OsT", "url": "/basketball/lithuania/nkl/a-b/nRzm7OsT/"}],
        out_dir=tmp_path,
        read_text=lambda url: _GATE_WITH_UNRELATED_WRITES,
        limit=1,
        gate_accept=True,
        gate_read_text=must_not_be_called,
    )

    page = payload["pages"][0]
    assert payload["requests_made"] == 1
    assert page["status"] == "gated"
    assert "no cookie written near the accept control" in page["gate"]
    assert "lang" in page["gate"]


def test_the_gate_script_read_shows_the_handler_and_the_bundles(tmp_path, capsys):
    page = tmp_path / "gate.html"
    page.write_text(
        '<div class="age__modal"><h1 id="ageTitle">t</h1>'
        '<button id="ageYes" onclick="localStorage.setItem(\'ageOk\',\'1\');loc();">Sim</button></div>'
        '<script src="/static/js/consent.bundle.js"></script>',
        encoding="utf-8",
    )

    assert probe.main(["--gate-script", str(page)]) == 0

    payload = json.loads(capsys.readouterr().out)
    assert payload["external_scripts"] == ["/static/js/consent.bundle.js"]
    assert payload["inline_script_blocks"] == 0
    assert payload["accept_handlers"][0]["control"] == "ageYes"
    assert "ageOk" in payload["accept_handlers"][0]["excerpt"]
    assert payload["accept_handlers"][0]["looks_like_code"] is True


def test_the_handler_is_found_even_when_the_button_carries_none(tmp_path, capsys):
    """The 2026-11-24 page: `<button id="ageYes">` with no onclick, banners after
    it, and the behaviour attached far below. Reading only the first occurrence
    of the id showed ads and declared nothing."""

    page = tmp_path / "gate.html"
    page.write_text(
        '<div class="age__modal"><h1 id="ageTitle">18+</h1>'
        '<div class="age__buttons"><button id="ageYes">Sim</button>'
        '<button id="ageNo">Não</button></div></div>'
        + '<div class="banner"><iframe src="/ad"></iframe></div>' * 20
        + "<script>var el=document.getElementById('ageYes');"
        "el.addEventListener('click',function(){document.cookie='be18=1;path=/';});</script>",
        encoding="utf-8",
    )

    assert probe.main(["--gate-script", str(page)]) == 0

    payload = json.loads(capsys.readouterr().out)
    windows = [item for item in payload["accept_handlers"] if item["control"] == "ageYes"]
    assert len(windows) == 2  # the markup, and the listener
    assert windows[0]["looks_like_code"] is False
    assert windows[1]["looks_like_code"] is True
    assert "addEventListener" in windows[1]["excerpt"]
    # And the block that owns the behaviour is named, so it can be read next.
    assert payload["inline_blocks"][0]["mentions_controls"] == ["ageYes"]
    assert payload["inline_blocks"][0]["writes"] is True


def test_the_gate_read_reports_whether_a_match_page_is_under_the_modal(tmp_path, capsys):
    page = tmp_path / "gate.html"
    page.write_text(
        '<div class="age__modal"><h1 id="ageTitle">18+</h1>'
        '<button id="ageYes">Sim</button></div>',
        encoding="utf-8",
    )

    assert probe.main(["--gate-script", str(page)]) == 0

    shape = json.loads(capsys.readouterr().out)["shape"]
    assert shape["td_cells"] == 0 and shape["period_markers"] == 0


def test_every_receipt_names_the_build_that_made_it(tmp_path, capsys):
    """Three sessions were spent reading output from a checkout one merge behind.
    A receipt that cannot name its own build cannot be told apart from a host
    answering badly."""

    page = tmp_path / "gate.html"
    page.write_text(_GATED_COOKIE_PAGE, encoding="utf-8")

    probe.main(["--gate-script", str(page)])
    payload = json.loads(capsys.readouterr().out)

    assert payload["tool"]["script"] == "probe_betexplorer_results.py"
    assert set(payload["tool"]) == {"script", "git_commit", "modified"}
    # Whatever git says, it says something -- None is "I could not tell", which
    # is a different statement from a commit hash and is allowed to appear.
    assert payload["tool"]["git_commit"] is None or isinstance(payload["tool"]["git_commit"], str)


def test_the_build_stamp_survives_a_missing_git(monkeypatch):
    def no_git(*args, **kwargs):
        raise OSError("git not found")

    monkeypatch.setattr(probe.subprocess, "run", no_git)

    build = probe.probe_build()

    assert build == {"script": "probe_betexplorer_results.py", "git_commit": None, "modified": None}


def test_a_writing_inline_block_is_shown_because_the_id_search_cannot_see_it(tmp_path, capsys):
    """The real page attaches nothing by id: `ageYes` occurs once, as markup, and
    two small blocks write cookies. A handler that selects by class is invisible
    to an id search, so the code is printed instead of summarised."""

    page = tmp_path / "gate.html"
    page.write_text(
        '<div class="age__modal"><h1 id="ageTitle">18+</h1><button id="ageYes">Sim</button></div>'
        "<script>var m=document.cookie.match(/be18=1/);"
        "document.querySelector('.age__btn--yes').onclick=function(){"
        "document.cookie='be18=1; path=/';location.reload();};</script>",
        encoding="utf-8",
    )

    probe.main(["--gate-script", str(page)])
    payload = json.loads(capsys.readouterr().out)

    block = payload["inline_blocks"][0]
    assert block["mentions_controls"] == []  # by id, nothing -- and that is the point
    assert block["writes"] is True
    assert "querySelector('.age__btn--yes')" in block["excerpt"]
    assert "be18=1" in block["excerpt"]


def test_a_supplied_cookie_is_sent_and_named(tmp_path):
    """The end of the road for this route: a key read from the page's own bundle.
    The receipt names it, so a pass never reads as something the tool worked out."""

    seen = []

    def read_text(url, cookie):
        seen.append(cookie)
        return _MATCH_PAGE

    payload = probe.capture_match_pages(
        "LITHUANIA_NKL",
        matches=[{"code": "nRzm7OsT", "url": "/basketball/lithuania/nkl/a-b/nRzm7OsT/"}],
        out_dir=tmp_path,
        read_text=read_text,
        limit=1,
        cookies=[("be18", "1")],
    )

    assert seen == ["be18=1"]
    assert payload["cookies_sent"] == ["be18"]
    assert payload["requests_made"] == 1
    assert payload["pages"][0]["status"] == "captured"


def test_the_bundle_read_names_the_key_it_finds(tmp_path):
    """The page holds the modal and nothing that acts on it, so the behaviour is
    in a bundle or nowhere. This is the reader for that bundle."""

    js = (
        "function n(){var e=document.getElementById('ageYes');"
        "document.querySelector('.age__btn--yes').onclick=function(){"
        "document.cookie='be18=1; path=/; max-age=31536000';location.reload();}}"
    )

    report = probe.gate_bundle_report("https://example.test/all.min.js", out_dir=tmp_path,
                                     read_text=lambda url: js)

    assert report["status"] == "captured"
    assert report["tokens_found"] == ["age__btn--yes", "ageYes"]
    assert report["cookie_writes"][0]["excerpt"].startswith("document.cookie='be18=1")
    assert Path(report["path"]).read_text() == js
    assert report["sha256"] == hashlib.sha256(js.encode()).hexdigest()


def test_a_bundle_that_cannot_be_fetched_is_a_receipt_not_a_crash(tmp_path):
    def boom(url):
        raise RuntimeError("HTTP 429 for https://example.test/all.min.js")

    report = probe.gate_bundle_report("https://example.test/all.min.js", out_dir=tmp_path,
                                     read_text=boom)

    assert report["status"] == "fetch_failed"
    assert "429" in report["detail"]
    assert "windows" not in report


class _Client:
    def __init__(self):
        self.calls = []

    def get_text(self, url, headers=None):
        self.calls.append((url, headers))
        return _MATCH_PAGE


def test_a_cookie_is_sent_as_a_header_and_not_as_the_headers_dict():
    """The 2026-11-24 receipt: `fetch_failed: dictionary update sequence element #0
    has length 1` beside `gated: 0`. The request never left; the cookie string had
    been handed to get_text as `headers`."""

    client = _Client()
    read = probe.cookie_reader(client)

    assert read("https://example.test/x", "be18=1") == _MATCH_PAGE
    assert client.calls == [("https://example.test/x", {"Cookie": "be18=1"})]


def test_the_bundle_read_finds_a_key_named_through_the_pages_own_helper():
    """The gate calls set_cookie('...'), not document.cookie=. A reader that stops
    at the raw assignment reports one write and misses the key."""

    js = (
        "function set_cookie(e,t,a,s,n,i){document.cookie=e+\"=\"+escape(t)}"
        "function init_age_wall(){var s=document.getElementById('ageYes');"
        "s.onclick=function(){set_cookie('be18','1',365,'/');location.reload()}}"
    )

    report = probe.bundle_map(js)

    assert report["gate_keys"] == ["be18"]
    assert report["set_cookie_calls"][0]["value"] == "'1'"
    assert report["age_wall_region"].startswith("init_age_wall")
    # The raw assignment inside the helper is still reported, for contrast.
    assert report["cookie_writes"]


def test_a_failed_request_is_not_reported_as_a_passed_gate(tmp_path):
    """gated: 0 with nothing fetched reads as success. The summary says otherwise."""

    def boom(url, cookie):
        raise ValueError("dictionary update sequence element #0 has length 1; 2 is required")

    payload = probe.capture_match_pages(
        "LITHUANIA_NKL",
        matches=[{"code": "nRzm7OsT", "url": "/basketball/lithuania/nkl/a-b/nRzm7OsT/"}],
        out_dir=tmp_path,
        read_text=boom,
        limit=1,
        cookies=[("be18", "1")],
    )

    assert payload["gated"] == 0
    assert payload["summary"] == "0 captured, 0 gated, 1 failed of 1 requested; 1 request(s) never returned a body"


def test_a_page_that_ships_the_modal_and_a_table_is_captured_not_gated(tmp_path):
    """BetExplorer puts the age modal in every page's markup and hides it
    client-side. Marker presence alone is not a refusal; the bundle says so:
    init_age_wall() checks `data-age-wall-enabled` and `currentGeoLocation`."""

    body = (
        '<div class="age__modal"><h1 id="ageTitle">18+</h1>'
        '<button id="ageYes">Sim</button></div>'
        "<table><tr><td>Q1</td><td>88</td><td>91</td><td>Q2</td><td>22</td><td>19</td></tr></table>"
    )

    payload = probe.capture_match_pages(
        "LITHUANIA_NKL",
        matches=[{"code": "nRzm7OsT", "url": "/basketball/lithuania/nkl/a-b/nRzm7OsT/"}],
        out_dir=tmp_path,
        read_text=lambda url: body,
        limit=1,
    )

    page = payload["pages"][0]
    assert page["status"] == "captured"
    assert page["gate_markup"].startswith("age gate")
    assert page["td_cells"] == 6
    assert payload["gated"] == 0


def test_the_gates_own_cookies_do_not_replace_the_callers(tmp_path):
    """`cookies` holds the header for every request in the loop. Reassigning it
    with the gate's pairs would swap the operator's key for a `lang=` reset on
    the next page -- the same overwrite, one loop later."""

    gated = _GATED_COOKIE_PAGE + "<p>banner</p>"
    seen = []

    def read_text(url, cookie):
        seen.append(cookie)
        return gated

    def gate_read_text(url, cookie):
        return gated  # still gated: the replay path sets `cookies` in the old code

    probe.capture_match_pages(
        "LITHUANIA_NKL",
        matches=[
            {"code": "nRzm7OsT", "url": "/basketball/lithuania/nkl/a-b/nRzm7OsT/"},
            {"code": "AbCd1234", "url": "/basketball/lithuania/nkl/c-d/AbCd1234/"},
        ],
        out_dir=tmp_path,
        read_text=read_text,
        sleep=lambda _: None,
        limit=2,
        cookies=[("age_verified", "true")],
        gate_accept=True,
        gate_read_text=gate_read_text,
    )

    assert seen
    assert all(cookie == "age_verified=true" for cookie in seen)
