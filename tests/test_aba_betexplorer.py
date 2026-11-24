"""ABA BetExplorer aggregate-price research lane.

The ABA route is deliberately independent of the canonical ABA League
schedule/result adapter. These tests exercise only the aggregate-price adapter
shape against a hermetic representative table, then prove that the pipeline
stamps research-only evidence onto a canonical ABA fixture.
"""

from __future__ import annotations

from forecasting.types import BookEvidence
from ingestion.league_registry import league_spec, price_evidence_policy
from ingestion.models import (
    Game,
    GameStatus,
    League,
    OperationalValidation,
    PriceProvenance,
    SourcePermission,
)
from ingestion.pipeline import Pipeline, PipelineConfig
from ingestion.providers import aba_liga, betexplorer
from ingestion.store import InMemoryStore


OBSERVED_UTC = "2026-11-20T12:00:00Z"
TARGET_DATE = "20261120"
MATCH_ID = "AbCdEf12"

# Contract-shaped table for the generic, already-live-verified BetExplorer
# moneyline parser. It is a hermetic parser fixture, not an ABA schedule/result
# source capture or a claim that the route supplies named-book evidence.
ABA_FIXTURES_HTML = f'''<table class="table-main table-main--leaguefixtures">
<tr><td class="table-main__datetime">20.11.2026 20:00</td>
<td class="h-text-left"><a href="/basketball/europe/aba-league/partizan-buducnost/{MATCH_ID}/" class="in-match"><span>Partizan</span> - <span>Buducnost</span></a></td>
<td class="table-main__odds"><button data-odd="1.55"></button></td>
<td class="table-main__odds"><button data-odd="2.50"></button></td></tr>
</table>'''


def _canonical_aba_game() -> Game:
    return Game(
        game_id="manual-aba-partizan-buducnost",
        league=League.ABA_LEAGUE,
        season="2026-27",
        home_team_id="partizan",
        away_team_id="buducnost",
        home_name="Partizan",
        away_name="Buducnost",
        scheduled_utc="2026-11-20T19:00:00Z",
        status=GameStatus.SCHEDULED,
        provider_ids={"betexplorer": MATCH_ID},
    )


def test_aba_registry_uses_official_canonical_source_but_keeps_research_prices_separate():
    spec = league_spec(League.ABA_LEAGUE)
    assert (
        spec.schedule_source,
        spec.results_source,
        spec.gate_schedule_source,
    ) == ("aba_liga", "aba_liga", "aba_liga")
    assert spec.book_providers == ("betexplorer", "betexplorer_direct")

    policy = price_evidence_policy(League.ABA_LEAGUE, "betexplorer")
    assert policy.price_provenance is PriceProvenance.AGGREGATE_CONSENSUS
    assert policy.source_permission is SourcePermission.UNRESOLVED
    assert policy.operational_validation is OperationalValidation.PROBATION
    assert "never a named-book ticket" in policy.detail


def test_aba_betexplorer_route_parses_only_prematch_aggregate_moneyline():
    snapshots = betexplorer.fetch_odds(
        "ABA_LEAGUE",
        TARGET_DATE,
        now_utc=OBSERVED_UTC,
        text_fetch=lambda _url: ABA_FIXTURES_HTML,
    )

    assert betexplorer.FIXTURES_URLS["ABA_LEAGUE"].endswith("/aba-league/fixtures/")
    assert betexplorer.evidence_for("ABA_LEAGUE") == betexplorer.ABA_EVIDENCE
    assert len(snapshots) == 1
    snapshot = snapshots[0]
    assert snapshot.provider_event_id == MATCH_ID
    assert snapshot.scheduled_utc == "2026-11-20T19:00:00Z"
    assert snapshot.home_name == "Partizan Mozzart Bet Belgrade"
    assert snapshot.away_name == "Buducnost"
    assert (snapshot.home_price_decimal, snapshot.away_price_decimal) == (1.55, 2.50)
    assert snapshot.price_provenance is PriceProvenance.AGGREGATE_CONSENSUS
    assert snapshot.source_permission is SourcePermission.UNRESOLVED
    assert snapshot.operational_validation is OperationalValidation.PROBATION


def test_aba_pipeline_maps_and_restamps_aggregate_evidence(monkeypatch):
    store = InMemoryStore()
    game = _canonical_aba_game()
    store.upsert_game(game)
    monkeypatch.setattr(
        betexplorer,
        "_text_fetch",
        lambda _factory, _text_fetch=None: (lambda _url: ABA_FIXTURES_HTML),
    )
    monkeypatch.setattr(aba_liga, "discover_games", lambda *_args, **_kwargs: [game])
    monkeypatch.setattr(aba_liga, "last_gaps", lambda: [])

    report = Pipeline(
        PipelineConfig(
            league="ABA_LEAGUE",
            date_yyyymmdd=TARGET_DATE,
            fixture_now_utc=OBSERVED_UTC,
        ),
        store,
        fetch=lambda _url: {},
    ).pregame()

    snapshots = store.get_book_snapshots(game.game_id)
    assert len(snapshots) == 1
    snapshot = snapshots[0]
    assert snapshot.game_id == game.game_id
    assert snapshot.provider_event_id == MATCH_ID
    assert snapshot.price_provenance is PriceProvenance.AGGREGATE_CONSENSUS
    assert snapshot.source_permission is SourcePermission.UNRESOLVED
    assert snapshot.operational_validation is OperationalValidation.PROBATION
    # Operator decision 2026-11-22: an aggregate consensus quote may price a
    # paper leg. What it may never do is pass itself off as a bookmaker's own
    # quote, so the provenance and the label are asserted either side of it.
    evidence = BookEvidence.from_snapshot(snapshot, frozen=False)
    assert evidence.ticket_price_eligible
    assert evidence.price_provenance is PriceProvenance.AGGREGATE_CONSENSUS
    assert "NOT A NAMED-BOOK QUOTE" in evidence.price_provenance_label

    assert report["book_provider_status"]["betexplorer"]["status"] == "fetched"
    assert report["book_provider_status"]["betexplorer"]["snapshots_mapped"] == 1
    assert not any(gap["required_field"] == "official_schedule" for gap in report["gaps"])
