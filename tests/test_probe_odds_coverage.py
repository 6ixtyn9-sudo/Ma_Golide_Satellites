"""The coverage probe must classify, never guess.

The synthetic payload below mirrors the *documented* catalogue shape (a list of
objects carrying ``key``/``group``/``title``/``active``/``has_outrights``). It
exists to test our classification and redaction, not to assert the live
contract: the live shape is whatever the provider returns, and the probe saves
the raw body so a disagreement is visible as evidence rather than a traceback.
"""

from __future__ import annotations

import importlib.util
from pathlib import Path

from ingestion.models import League
from ingestion.providers.theoddsapi import SPORT_KEYS

ROOT = Path(__file__).resolve().parent.parent


def _load():
    spec = importlib.util.spec_from_file_location(
        "probe_odds_api_coverage", ROOT / "scripts" / "probe_odds_api_coverage.py"
    )
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


probe = _load()

CATALOGUE = [
    {"key": "basketball_nba", "group": "Basketball", "title": "NBA", "active": True,
     "has_outrights": False},
    {"key": "basketball_spain_acb", "group": "Basketball", "title": "Liga ACB",
     "active": True, "has_outrights": False},
    {"key": "basketball_something_else", "group": "Basketball", "title": "Other",
     "active": False, "has_outrights": False},
    {"key": "soccer_epl", "group": "Soccer", "title": "EPL", "active": True,
     "has_outrights": False},
    {"key": "", "group": "Basketball", "title": "no key", "active": True},
    "not-a-row",
]


def test_the_key_is_redacted_from_every_recorded_url():
    url = probe.catalogue_url("SECRET-VALUE")
    assert "SECRET-VALUE" in url  # it has to be sent
    assert "SECRET-VALUE" not in probe.redact_url(url)
    assert "<redacted>" in probe.redact_url(url)


def test_only_basketball_rows_are_kept_and_malformed_rows_are_skipped():
    keys = [row["key"] for row in probe.basketball_keys(CATALOGUE)]
    assert keys == ["basketball_nba", "basketball_something_else", "basketball_spain_acb"]
    assert "soccer_epl" not in keys


def test_a_non_list_catalogue_yields_nothing_rather_than_a_guess():
    assert probe.basketball_keys({"error_code": 401}) == []
    assert probe.basketball_keys(None) == []


def test_classification_separates_mapped_unmapped_and_unrelated():
    leagues = [League.NBA, League.SPAIN_ACB]

    assert probe.classify("basketball_nba", SPORT_KEYS, leagues) == ("already mapped", "NBA")
    # Listed by the provider, absent from the registry's admitted map.
    assert probe.classify("basketball_spain_acb", SPORT_KEYS, leagues) == (
        "listed, unmapped",
        "SPAIN_ACB",
    )
    assert probe.classify("basketball_croatia_premijer", SPORT_KEYS, leagues) == (
        "listed, no league",
        None,
    )


def test_an_empty_catalogue_is_not_reported_as_no_coverage():
    """Silence from the provider must not read as "the league is not carried"."""

    # The script's own contract: no body -> exit 1 and an explicit warning.
    source = (ROOT / "scripts" / "probe_odds_api_coverage.py").read_text(encoding="utf-8")
    assert "Do not read this as 'no coverage'" in source
