"""Shared test helpers: fixture loading, a URL-matching fake fetch, and a
network guard.

No ordinary test performs network I/O. Providers receive ``fetch`` callables
that serve small committed JSON fixtures from ``tests/fixtures/``. An autouse
fixture blocks outbound sockets; only tests marked ``@pytest.mark.live`` (the
explicitly invoked live smoke module) may reach the network. The helpers are
exposed both as plain functions and as pytest fixtures so test modules never
need to import ``conftest`` directly (import-mode safe).
"""

from __future__ import annotations

import json
import socket
from pathlib import Path
from typing import Any, Dict, Union

import pytest

FIXTURES_DIR = Path(__file__).resolve().parent / "fixtures"


def _block_socket(*args, **kwargs):
    raise RuntimeError(
        "network access attempted in an ordinary test; use committed fixtures "
        "or mark the test with @pytest.mark.live"
    )


@pytest.fixture(autouse=True)
def no_network(request, monkeypatch):
    """Block outbound sockets unless the test is explicitly marked live."""
    if request.node.get_closest_marker("live"):
        yield
        return
    monkeypatch.setattr(socket, "socket", _block_socket, raising=True)
    monkeypatch.setattr(socket, "create_connection", _block_socket, raising=True)
    yield


def load_json_fixture(*parts: str) -> Any:
    """Load a committed JSON fixture relative to ``tests/fixtures/``."""
    path = FIXTURES_DIR.joinpath(*parts)
    if not path.is_file():
        raise FileNotFoundError(f"fixture not found: {path}")
    with open(path, encoding="utf-8") as handle:
        return json.load(handle)


def make_fixture_fetch(mapping: Dict[str, Union[str, Path, Any]]):
    """Build a ``fetch`` callable that matches URLs to fixtures by substring.

    The LONGEST matching mapping key wins, so overlapping keys stay
    unambiguous (e.g. ``events/<id>`` also appears inside the predictor
    URL, but the longer ``competitions/<id>/predictor`` key takes
    precedence there). Mapping values are fixture paths relative to
    ``tests/fixtures/`` or already-loaded payloads (dict/list). Unmatched
    URLs raise ``KeyError`` — callers (and ``providers.espn.smoke``) treat
    that as missing data.
    """

    def _fetch(url: str) -> Any:
        matches = [key for key in mapping if key in url]
        if not matches:
            raise KeyError(f"no fixture mapped for URL: {url}")
        value = mapping[max(matches, key=len)]
        if isinstance(value, (str, Path)):
            return load_json_fixture(str(value))
        return value

    return _fetch


@pytest.fixture
def load_fixture():
    """``load_json_fixture`` as a pytest fixture (import-mode safe)."""
    return load_json_fixture


@pytest.fixture
def fixture_fetch():
    """``make_fixture_fetch`` factory as a pytest fixture."""
    return make_fixture_fetch
