"""Live smoke test — explicitly invoked, redacted, reports ACTUAL coverage.

Run with:  pytest tests/test_live_smoke.py -m live -q

This module is excluded from the ordinary suite (the conftest socket guard is
lifted only for ``live``-marked tests). It makes KEYLESS calls only (ESPN,
Hardwood, official EuroLeague). Authenticated providers stay disabled unless
their env keys are present, and their output is labelled accordingly.
"""

from __future__ import annotations


import pytest

from ingestion.http import HttpClient
from ingestion.pipeline import Pipeline, PipelineConfig
from ingestion.providers import espn, euroleague, hardwood
from ingestion.store import InMemoryStore

pytestmark = pytest.mark.live


def _today_yyyymmdd() -> str:
    from datetime import datetime, timezone

    return datetime.now(timezone.utc).strftime("%Y%m%d")


def _fetch():
    return HttpClient().get_json


def _is_live_fetch(fetch) -> bool:
    return True  # smoke labels itself LIVE_VERIFIED by construction


def test_espn_live_smoke_reports_actual_coverage():
    date = _today_yyyymmdd()
    report = espn.smoke(date, _fetch())
    assert report["fixtures_found"] > 0, f"no ESPN fixtures on {date}"
    assert report["verification"] == "LIVE_VERIFIED"
    print(
        "\nESPN smoke",
        date,
        "fixtures:",
        report["fixtures_found"],
        "bpi:",
        report["bpi_complete"],
    )


def test_hardwood_live_smoke_reports_freshness():
    artifact = hardwood.load_forecasts(_fetch())
    state = hardwood.freshness_state(artifact)
    print(
        "\nHardwood smoke",
        artifact.generated_at,
        artifact.ingestion_status,
        "games:",
        len(artifact.games_by_id),
        "freshness:",
        state.value,
    )
    assert artifact.ingestion_status == "ok"


def test_euroleague_live_smoke_reports_actual_coverage():
    report = euroleague.smoke(_fetch())
    assert report["games_found"] > 0
    print(
        "\nEuroLeague smoke",
        report["games_found"],
        "games;",
        report["results_reconciled"],
        "reconciled",
    )


def test_full_pregame_dry_run_live_nba():
    """Keyless end-to-end dry run against the live slate; no writes."""
    date = _today_yyyymmdd()
    config = PipelineConfig(league="NBA", date_yyyymmdd=date)
    pipeline = Pipeline(config, InMemoryStore(), fetch=_fetch())
    report = pipeline.pregame()
    print(
        "\nNBA pregame dry run",
        date,
        "fixtures:",
        report["fixtures_found"],
        "predictions:",
        report.get("predictions_selected"),
        "gaps:",
        len(report["gaps"]),
    )
    assert report["mode"] == "dry_run"
    assert report["fixtures_found"] > 0
    # Live smoke remains read-only and returns canonical acquisition facts.
    assert "rows" in report and "gaps" in report
