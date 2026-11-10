"""Predixsport adapter tests (fixture-based, no network).

The committed fixtures under ``tests/fixtures/predixsport/`` are SYNTHETIC
DOCUMENTATION_ONLY payloads authored from the documented schema (see that
directory's README); no authenticated live call has ever been made — the
2026-11-06 unauthenticated sweep got ``Missing Authorization header``, which
is recorded as evidence that auth is required.  HTTP is exercised through
``http.HttpClient`` over a fake session so the real retry/status/error
handling runs without a socket.  Helpers carry the unique ``px_`` prefix so
they cannot collide with other agents' conftest helpers.
"""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from ingestion import http
from ingestion.models import CoverageState
from ingestion.providers import predixsport as px

FIXTURES = Path(__file__).parent / "fixtures" / "predixsport"
FAKE_KEY = "px-dummy-test-key-not-a-real-credential"


class PxFakeResponse:
    """Minimal stand-in for requests.Response."""

    def __init__(self, status_code: int = 200, text: str = "{}", headers=None):
        self.status_code = status_code
        self.text = text
        self.headers = headers or {}


class PxFakeSession:
    """Minimal stand-in for requests.Session, recording every call."""

    def __init__(self, responses):
        self.responses = list(responses)
        self.calls = []  # list of (url, headers dict)
        self.providers_seen = []  # provider names handed to fetch_factory

    def get(self, url, headers=None, timeout=None):
        self.calls.append((url, dict(headers or {})))
        if not self.responses:
            raise AssertionError(f"unexpected extra request: {url}")
        return self.responses.pop(0)


def px_factory(responses):
    """fetch_factory(provider) -> FetchFn over an authenticated HttpClient.

    Mirrors the prod wiring: the factory receives the provider name, then the
    returned FetchFn performs the GET with ``px._headers()`` baked in.
    """
    session = PxFakeSession(responses)

    def factory(provider: str):
        session.providers_seen.append(provider)
        client = http.HttpClient(session=session, sleep=lambda _s: None)
        return lambda url: client.get_json(url, headers=px._headers())

    return factory, session


def px_fixture_factory(name: str):
    """Pure-fixture fetch_factory serving one JSON file (no HTTP layer)."""

    def factory(provider: str):
        assert provider == px.PROVIDER
        with open(FIXTURES / name) as f:
            data = json.load(f)

        def _fetch(url: str):
            return data

        return _fetch

    return factory


def px_load(name: str) -> dict:
    with open(FIXTURES / name) as f:
        return json.load(f)


def px_json(payload, status: int = 200) -> PxFakeResponse:
    return PxFakeResponse(status_code=status, text=json.dumps(payload))


PX_URLS = {
    "upcoming?days": "upcoming_nba.json",
    "px-nba-401909287": "detail_match.json",
    "px-nba-disagreement": "detail_disagreement.json",
}


# ---------------------------------------------------------------------------
# Disabled state (no key) — required test 1
# ---------------------------------------------------------------------------


def test_disabled_without_key(monkeypatch):
    monkeypatch.delenv(px.ENV_KEY, raising=False)
    assert px.is_enabled() is False

    # The factory must never even be built when disabled.
    def _unused_factory(provider):
        raise AssertionError("fetch_factory must not be called while disabled")

    assert px.fetch_upcoming("nba", 7, _unused_factory) == []
    assert px.fetch_match_detail("px-nba-1", "nba", _unused_factory) is None


def test_enabled_requires_nonempty_key(monkeypatch):
    monkeypatch.delenv(px.ENV_KEY, raising=False)
    monkeypatch.setenv(px.ENV_KEY, "   ")
    assert px.is_enabled() is False  # whitespace-only counts as missing
    monkeypatch.setenv(px.ENV_KEY, FAKE_KEY)
    assert px.is_enabled() is True


# ---------------------------------------------------------------------------
# Auth header — required test 2
# ---------------------------------------------------------------------------


def test_auth_header_built_from_env(monkeypatch):
    monkeypatch.delenv(px.ENV_KEY, raising=False)
    # Direct call without a key raises the module's UnauthorizedError.
    with pytest.raises(px.UnauthorizedError):
        px._headers()
    # With a (dummy) key the Bearer header is built from the environment.
    monkeypatch.setenv(px.ENV_KEY, FAKE_KEY)
    assert px._headers() == {"Authorization": f"Bearer {FAKE_KEY}"}
    # And the header really reaches the wire through the prod-style factory.
    factory, session = px_factory([px_json(px_load("upcoming_nba.json"))])
    px.fetch_upcoming("nba", 7, factory)
    assert len(session.calls) == 1
    url, headers = session.calls[0]
    assert url == f"{px.BASE_URL}/v1/predictions/nba/upcoming?days=7"
    assert headers["Authorization"] == f"Bearer {FAKE_KEY}"
    assert session.providers_seen == [px.PROVIDER]


# ---------------------------------------------------------------------------
# Match detail — required tests 3 and 4
# ---------------------------------------------------------------------------


def test_detail_derives_scores_with_home_minus_away_sign(monkeypatch):
    monkeypatch.setenv(px.ENV_KEY, FAKE_KEY)
    fixture = px_load("detail_match.json")
    total = fixture["predicted_total_points"]  # 228.4
    diff = fixture["predicted_point_diff"]  # -6.2 (HOME minus AWAY)
    assert (total, diff) == (228.4, -6.2)
    # Python-arithmetic check of the documented derivation.
    expected_home = (total + diff) / 2  # (228.4 - 6.2) / 2 = 111.1
    expected_away = (total - diff) / 2  # (228.4 + 6.2) / 2 = 117.3
    assert expected_home == pytest.approx(111.1)
    assert expected_away == pytest.approx(117.3)

    factory, session = px_factory([px_json(fixture)])
    prediction = px.fetch_match_detail("px-nba-401909287", "nba", factory)
    assert prediction is not None
    # The negative diff flips the score edge to the away side.
    assert prediction.expected_home_score == pytest.approx(expected_home)
    assert prediction.expected_home_score == pytest.approx(111.1)
    assert prediction.expected_away_score == pytest.approx(expected_away)
    assert prediction.expected_away_score == pytest.approx(117.3)
    assert prediction.expected_total == pytest.approx(228.4)
    assert prediction.expected_margin == pytest.approx(-6.2)
    assert session.calls[0][0] == f"{px.BASE_URL}/v1/predictions/nba/px-nba-401909287"


def test_pred_from_probability_not_margin(monkeypatch):
    """Independent-model disagreement must be preserved: the probability
    model favours AWAY (0.58) while the margin model favours HOME (+6.2).
    Pred computes 2 from probability; the margin keeps its positive sign."""
    monkeypatch.setenv(px.ENV_KEY, FAKE_KEY)
    factory, _session = px_factory([px_json(px_load("detail_disagreement.json"))])
    prediction = px.fetch_match_detail("px-nba-disagreement", "nba", factory)

    assert prediction is not None
    assert prediction.prob_home == pytest.approx(0.42)
    assert prediction.prob_away == pytest.approx(0.58)
    assert prediction.prob_away > prediction.prob_home
    assert prediction.prob_away > prediction.prob_home
    # The margin/scores side keeps the HOME edge — not "repaired" to agree.
    assert prediction.expected_margin == pytest.approx(6.2)
    assert prediction.expected_margin > 0
    assert prediction.expected_home_score == pytest.approx(117.3)
    assert prediction.expected_away_score == pytest.approx(111.1)
    assert prediction.expected_home_score > prediction.expected_away_score


# ---------------------------------------------------------------------------
# Upcoming list — required test 5
# ---------------------------------------------------------------------------


def test_missing_total_leaves_scores_none(monkeypatch):
    """List-only records carry probabilities (+ margin) but no totals: the
    prediction keeps scores None instead of fabricating them."""
    monkeypatch.setenv(px.ENV_KEY, FAKE_KEY)
    factory, _session = px_factory([px_json(px_load("upcoming_nba.json"))])
    gaps = []
    predictions = px.fetch_upcoming("nba", 7, factory, gaps=gaps)
    assert len(predictions) == 2
    assert len(gaps) == 1  # the malformed third entry is skipped with a gap
    assert gaps[0].reason is CoverageState.MISSING

    first = predictions[0]
    assert first.game_id == "px-nba-401909287"
    assert first.provider == "predixsport"
    assert first.produced_by == "predixsport-nba-earlyseason"  # preserved
    assert first.prob_home == pytest.approx(0.713845)
    assert first.prob_away == pytest.approx(0.286155)
    assert first.expected_margin == pytest.approx(7.636)  # home minus away
    assert first.expected_home_score is None
    assert first.expected_away_score is None
    assert first.expected_total is None

    second = predictions[1]
    assert second.expected_margin == pytest.approx(-9.575)
    assert second.expected_home_score is None


# ---------------------------------------------------------------------------
# Defensive parsing / errors
# ---------------------------------------------------------------------------


def test_list_rejects_out_of_range_probabilities_with_gap(monkeypatch):
    monkeypatch.setenv(px.ENV_KEY, FAKE_KEY)
    payload = {
        "matches": [
            {
                "match_id": "px-percent-scale",
                "probabilities": {"home": 71.4, "away": 28.6},  # 0-100, not 0-1
            }
        ]
    }
    factory, _session = px_factory([px_json(payload)])
    gaps = []
    predictions = px.fetch_upcoming("nba", 7, factory, gaps=gaps)
    assert predictions == []
    assert len(gaps) == 1
    assert gaps[0].reason is CoverageState.MISSING


def test_detail_total_without_diff_leaves_scores_none(monkeypatch):
    """Total without a diff: no derivation is possible for the split."""
    monkeypatch.setenv(px.ENV_KEY, FAKE_KEY)
    payload = {
        "match_id": "px-nba-total-only",
        "probabilities": {"home": 0.5, "away": 0.5},
        "predicted_total_points": 228.4,
    }
    factory, _session = px_factory([px_json(payload)])
    prediction = px.fetch_match_detail("px-nba-total-only", "nba", factory)
    assert prediction.expected_total == pytest.approx(228.4)
    assert prediction.expected_home_score is None
    assert prediction.expected_away_score is None
    assert prediction.expected_margin is None


@pytest.mark.parametrize("status", [401, 404])
def test_unknown_or_unauthorized_raises_http_error(monkeypatch, status):
    monkeypatch.setenv(px.ENV_KEY, FAKE_KEY)
    factory, _session = px_factory(
        [PxFakeResponse(status_code=status, text="Missing Authorization header")]
    )
    with pytest.raises(http.HttpError) as excinfo:
        px.fetch_match_detail("px-unknown", "nba", factory)
    assert excinfo.value.status == status
    # The key must never leak into surfaced error text.
    assert FAKE_KEY not in str(excinfo.value)


def test_invalid_league_is_rejected(monkeypatch):
    monkeypatch.setenv(px.ENV_KEY, FAKE_KEY)
    factory, session = px_factory([])
    with pytest.raises(ValueError):
        px.fetch_upcoming("nfl", 7, factory)
    with pytest.raises(ValueError):
        px.fetch_match_detail("px-1", "football", factory)
    assert session.calls == []  # nothing was called


# ---------------------------------------------------------------------------
# Smoke
# ---------------------------------------------------------------------------


def test_smoke_disabled_reports_documentation_only_without_calling(monkeypatch):
    monkeypatch.delenv(px.ENV_KEY, raising=False)
    factory, session = px_factory([])
    report = px.smoke("nba", factory)
    assert report["enabled"] is False
    assert report["evidence"] == "DOCUMENTATION_ONLY"
    assert session.calls == []


def test_smoke_enabled_makes_list_call_and_reports_counts(monkeypatch):
    monkeypatch.setenv(px.ENV_KEY, FAKE_KEY)
    factory, session = px_factory([px_json(px_load("upcoming_nba.json"))])
    report = px.smoke("nba", factory)
    assert report["enabled"] is True
    assert report["upcoming_count"] == 2
    assert report["gaps"] == 1
    # Never claims LIVE coverage without an observed authenticated response.
    assert report["evidence"] == "DOCUMENTATION_ONLY"
    assert "Missing Authorization header" in report["evidence_note"]
    assert len(session.calls) == 1
    assert session.calls[0][1]["Authorization"] == f"Bearer {FAKE_KEY}"


def test_smoke_enabled_surfaces_redacted_error(monkeypatch):
    monkeypatch.setenv(px.ENV_KEY, FAKE_KEY)
    factory, _session = px_factory([PxFakeResponse(status_code=401, text="denied")])
    report = px.smoke("euroleague", factory)
    assert report["enabled"] is True
    assert report["upcoming_count"] == 0
    assert "HttpError" in report["error"]
    assert FAKE_KEY not in report["error"]
