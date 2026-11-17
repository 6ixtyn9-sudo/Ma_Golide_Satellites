"""Headless acquisition orchestration: smoke, pregame, freeze, results, coverage.

Provider adapters return canonical records. This module coordinates those
records and a Store; it has no spreadsheet, web-application, or external
read-model dependency. Every command is dry-run unless the caller explicitly
enables canonical writes.
"""

from __future__ import annotations

from datetime import datetime, timedelta
import logging
from dataclasses import dataclass, field, replace
from typing import Any, Callable, Dict, List, Optional, Sequence
from zoneinfo import ZoneInfo

from .acquisition import AcquisitionClient
from .books import (
    BookConfig,
    newest_snapshot,
    period_coverage,
    select_moneyline_snapshots,
    select_total_snapshots,
)
from .http import HttpClient
from .identity import (
    canonical_team_key,
    now_utc_iso,
    parse_utc,
    times_within,
    to_utc_iso,
)
from .models import (
    BookSnapshot,
    CoverageGap,
    CoverageReport,
    CoverageState,
    Market,
    Game,
    GameStatus,
    League,
    Period,
    Prediction,
    ProviderHealth,
    payload_hash,
)
from .predictions import (
    DEFAULT_PROVIDER_ORDER,
    SelectionResult,
    freeze_game_predictions,
    odds_for_prediction,
    resolve_frozen_bundle,
    select_prediction,
)
from .leagues import native_periods
from .providers import betexplorer, cbaleague, espn, euroleague, hardwood, isportsapi, sportradar

logger = logging.getLogger(__name__)

DEFAULT_TZ = {
    "NBA": "America/New_York",
    "EUROLEAGUE": "Europe/Paris",
    "CBA": "Asia/Shanghai",
    "WNBA": "America/New_York",
    "NCAA_M": "America/New_York",
    "NCAA_W": "America/New_York",
}
# CBA acquisition now follows the official league data API's Beijing-local
# calendar day (schedules/day is keyed by Beijing date); its display and
# ledger grouping use Asia/Shanghai as before.
ACQUISITION_TZ = {**DEFAULT_TZ, "CBA": "Asia/Shanghai"}
BOOK_PROVIDERS = ("therundown", "sportsgameodds", "theoddsapi")
# EuroLeague and CBA bookmaker prices come from the keyless BetExplorer
# fixture tables (representative average moneyline odds, LIVE_VERIFIED
# 2026-11-16). CBA keeps its paid iSportsAPI adapter as an additional
# source when a key exists; it is never assumed.
BOOK_PROVIDERS_BY_LEAGUE = {
    "CBA": ("isportsapi", "betexplorer"),
    "EUROLEAGUE": ("betexplorer",),
    "NCAA_M": ("theoddsapi",),
    "NCAA_W": ("theoddsapi",),
    "WNBA": ("theoddsapi",),
}

ESPN_LEAGUES = {League.NBA, League.WNBA, League.NCAA_M, League.NCAA_W}


def _book_providers(league: str) -> Sequence[str]:
    # NBA/EuroLeague retain the existing overrideable provider tuple.
    return BOOK_PROVIDERS_BY_LEAGUE.get(league.upper(), BOOK_PROVIDERS)


BOOK_STALENESS_HOURS = 72.0


@dataclass
class PipelineConfig:
    league: str
    date_yyyymmdd: str
    tz_name: Optional[str] = None
    book: BookConfig = field(default_factory=BookConfig)
    freshness_hours: float = 12.0
    provider_order: Sequence[str] = DEFAULT_PROVIDER_ORDER
    # Bounded acquisition plan; these are code/config defaults, never workflow
    # dispatch inputs. Backfill acquires official results, not invented training.
    history_days: int = 21
    history_request_budget: int = 21
    deep_history_event_budget: int = 50
    deep_history_start_season: Optional[int] = None
    context_request_budget: int = 40
    http_request_budget: int = 1200
    # Used only by committed offline fixtures so replay evidence is stable.
    fixture_now_utc: Optional[str] = None

    @property
    def league_enum(self) -> League:
        return League[self.league.upper()]

    @property
    def timezone(self) -> str:
        return self.tz_name or DEFAULT_TZ[self.league.upper()]


class Pipeline:
    """One configured acquisition run against one canonical store."""

    def __init__(
        self,
        config: PipelineConfig,
        store,
        fetch: Optional[Callable[[str], Any]] = None,
        fetch_factory: Optional[Callable[[str], Callable[[str], Any]]] = None,
        allow_writes: bool = False,
    ) -> None:
        self.config = config
        self.store = store
        self.fetch = fetch or HttpClient().get_json
        cache_dir = None
        store_path = getattr(store, "path", None)
        if fetch is None and store_path and store_path != ":memory:":
            # ESPN is keyless: its durable cache contains no credentialed URLs.
            # Other providers continue to use their existing bounded HttpClient.
            from pathlib import Path

            cache_dir = str(Path(store_path).parent / "http-cache" / "espn")
        self.acquisition = AcquisitionClient(
            self.fetch,
            cache_dir=cache_dir,
            request_budget=config.http_request_budget,
            now_fn=self._now_iso,
        )
        self.espn_fetch = self.acquisition
        self.fetch_factory = fetch_factory or _default_fetch_factory
        self.allow_writes = allow_writes
        self.gaps: List[CoverageGap] = []
        self.quota: Dict[str, Any] = {}
        self.context_report: Dict[str, Any] = {}

    def _now_iso(self) -> str:
        if self.config.fixture_now_utc:
            return to_utc_iso(parse_utc(self.config.fixture_now_utc))
        return now_utc_iso()

    def _reset_stage(self) -> None:
        """Keep gaps and quota scoped to one command, not cumulative duplicates."""

        self.gaps = []
        self.quota = {}
        self.context_report = {}

    def _standings_smoke(self) -> Dict[str, Any]:
        """Measure standings endpoint coverage without persisting/model use."""
        league = self.config.league_enum
        if league not in {League.NBA, League.EUROLEAGUE}:
            return {
                "status": "not_supported_by_current_adapter",
                "league": league.value,
                "role": "evidence_coverage_only",
            }
        captured: Dict[str, Any] = {}
        observed = self._now_iso()

        standings_round = None
        source_fetch = self.espn_fetch if league is League.NBA else self.fetch

        def capture(url: str):
            payload = source_fetch(url)
            captured["payload"] = payload
            return payload

        try:
            if league is League.NBA:
                rows = espn.fetch_standings(capture)
                status = "ok" if rows else "empty"
                provider = "espn"
            else:
                year = int(self.config.date_yyyymmdd[:4])
                if int(self.config.date_yyyymmdd[4:6]) < 7:
                    year -= 1
                provider_report: Dict[str, str] = {}
                rows = euroleague.fetch_standings(
                    capture,
                    seasoncode=f"E{year}",
                    report=provider_report,
                )
                status = provider_report.get("status", "unavailable")
                standings_round = provider_report.get("round")
                provider = "euroleague"
            if league is League.NBA:
                standings_round = None
            row_values = [row.__dict__ if hasattr(row, "__dict__") else row for row in rows]
            return {
                "status": status,
                "provider": provider,
                "round": standings_round,
                "observed_utc": observed,
                "teams_seen": len(rows),
                "payload_hash": payload_hash(captured.get("payload")) if "payload" in captured else None,
                "normalized_rows_hash": payload_hash(row_values),
                "role": "evidence_coverage_only_not_a_model_input",
            }
        except Exception as exc:
            return {
                "status": "provider_error",
                "provider": "espn" if league is League.NBA else "euroleague",
                "observed_utc": observed,
                "error_type": type(exc).__name__,
                "role": "evidence_coverage_only_not_a_model_input",
            }

    # -- commands ------------------------------------------------------------

    def smoke(self) -> Dict[str, Any]:
        """Report observed provider coverage without writing canonical state."""

        self._reset_stage()
        league = self.config.league.upper()
        if league in {"NBA", "WNBA", "NCAA_M", "NCAA_W"}:
            report = espn.smoke(self.config.date_yyyymmdd, self.espn_fetch, self.config.league_enum)
            if league == "NBA":
                report["hardwood"] = self._hardwood_smoke()
            report["standings"] = self._standings_smoke()
            report["bookmakers"] = self._bookmakers_smoke()
            return report
        if league == "CBA":
            report = cbaleague.schedule_smoke(
                self.config.date_yyyymmdd, self.fetch, self.config.league_enum
            )
            report["isportsapi"] = isportsapi.schedule_smoke(
                self.config.date_yyyymmdd, self.fetch, self.config.league_enum
            )
            try:
                report["sportradar"] = sportradar.smoke(self.config.date_yyyymmdd, self.fetch)
            except Exception as exc:
                report["sportradar"] = {"status": "PROVIDER_ERROR", "detail": str(exc)[:200]}
            report["bookmakers"] = self._bookmakers_smoke()
            return report
        if league == "EUROLEAGUE":
            report = euroleague.smoke(self.fetch)
            report["standings"] = self._standings_smoke()
            report["predixsport"] = self._predixsport_smoke()
            report["bookmakers"] = self._bookmakers_smoke()
            return report
        raise ValueError(f"unsupported league for smoke: {league}")

    def pregame(self) -> Dict[str, Any]:
        """Acquire fixtures, coherent predictions, and observed book evidence."""

        self._reset_stage()
        games = self._discover_games()
        games = self._acquire_context(games)
        selections = self._acquire_predictions(games)
        self._acquire_book_lines(games)

        rows = []
        for game in games:
            selection = selections[game.game_id]
            pred = selection.selected
            snapshots = self._visible_pregame_books(game)
            quarter_lines, full_game_line, book_evidence, book_gaps = self._book_fields(
                game, snapshots
            )
            self.gaps.extend(book_gaps)
            odds, odds_evidence = self._odds_evidence_for(snapshots, pred)
            rows.append(
                {
                    "game_id": game.game_id,
                    "matchup": f"{game.away_name} @ {game.home_name}",
                    "tip_utc": game.scheduled_utc,
                    "prediction_provider": selection.provider,
                    "prediction_reason": selection.reason,
                    "prediction": _prediction_facts(pred),
                    "prediction_frozen": bool(
                        pred and self.store.is_prediction_frozen(game.game_id, pred.provider)
                    ),
                    "odds": odds,
                    "odds_evidence": odds_evidence,
                    "full_game_line": full_game_line,
                    "quarters": (
                        quarter_lines if game.league not in {League.NCAA_M, League.NCAA_W} else {}
                    ),
                    "halves": (
                        quarter_lines if game.league in {League.NCAA_M, League.NCAA_W} else {}
                    ),
                    "native_periods": quarter_lines,
                    "book_evidence": book_evidence,
                }
            )

        return self._report(
            mode="write" if self.allow_writes else "dry_run",
            fixtures_found=len(games),
            extra={
                "rows": rows,
                "predictions_selected": _count_by(
                    [s.provider for s in selections.values() if s.provider]
                ),
                "espn_context": self.context_report,
            },
        )

    def freeze(self) -> Dict[str, Any]:
        """Freeze the latest valid pre-tip predictions and book observations."""

        self._reset_stage()
        tip_games = self._games_with_tip_passed()
        frozen: Dict[str, Dict[str, str]] = {}
        books_frozen = 0
        for game in tip_games:
            frozen[game.game_id] = freeze_game_predictions(self.store, game, game.scheduled_utc)
            books_frozen += self._freeze_books(game)
        return self._report(
            mode="write" if self.allow_writes else "dry_run",
            fixtures_found=len(tip_games),
            extra={
                "frozen_predictions": frozen,
                "book_snapshots_frozen": books_frozen,
            },
        )

    def results(self) -> Dict[str, Any]:
        """Acquire reconciled official finals and join frozen pre-tip evidence."""

        self._reset_stage()
        try:
            finals = self._fetch_finals()
        except Exception as exc:
            # A provider outage must not fabricate an empty-but-successful
            # slate; it becomes an explicit provider_error gap and no finals
            # are written (same semantics as the backfill paths).
            logger.warning("result fetch failed: %s", exc)
            self.gaps.append(
                CoverageGap(
                    game_id=None,
                    required_field="official_results",
                    reason=CoverageState.PROVIDER_ERROR,
                    provider=self.config.league.lower(),
                    detail=str(exc)[:200],
                    observed_utc=self._now_iso(),
                )
            )
            finals = []
        rows = []
        used_frozen = 0
        quarantined = []
        for game in finals:
            if game.actual is None:
                quarantined.append({"game_id": game.game_id, "reason": "no_official_result"})
                continue
            ok, why = game.actual.reconciles()
            if not ok:
                quarantined.append({"game_id": game.game_id, "reason": why})
                continue

            self.store.upsert_game(game)
            freeze_game_predictions(self.store, game, game.scheduled_utc)
            self._freeze_books(game)
            pred, was_frozen = resolve_frozen_bundle(
                self.store,
                game.game_id,
                game.scheduled_utc,
                tuple(self.config.provider_order),
            )
            if was_frozen:
                used_frozen += 1
            if pred is None:
                self.gaps.append(
                    CoverageGap(
                        game_id=game.game_id,
                        required_field="pregame_prediction",
                        reason=CoverageState.MISSING,
                        detail="missing_pregame_prediction",
                        observed_utc=self._now_iso(),
                    )
                )
            snapshots = self.store.get_book_snapshots(game.game_id)
            odds, odds_evidence = self._frozen_odds_evidence(snapshots, pred, game)
            rows.append(
                {
                    "game_id": game.game_id,
                    "tip_utc": game.scheduled_utc,
                    "home_final": game.actual.home_final,
                    "away_final": game.actual.away_final,
                    "result_source": game.actual.source,
                    "prediction": _prediction_facts(pred),
                    "prediction_frozen": was_frozen,
                    "odds": odds,
                    "odds_evidence": odds_evidence,
                }
            )

        return self._report(
            mode="write" if self.allow_writes else "dry_run",
            fixtures_found=len(finals),
            extra={
                "rows": rows,
                "finals_written": len(rows),
                "frozen_predictions_used": used_frozen,
                "quarantined": quarantined,
            },
        )

    def coverage(self) -> Dict[str, Any]:
        """Summarize actual stored coverage without network access."""

        self._reset_stage()
        games = self.store.list_games(league=self.config.league_enum)
        any_book = BookConfig()
        with_preds = 0
        with_fg_line = 0
        native_labels = [period.value for period in native_periods(self.config.league_enum)]
        per_period: Dict[str, int] = {label: 0 for label in native_labels}
        finals = 0
        for game in games:
            if self.store.get_predictions(game.game_id):
                with_preds += 1
            snapshots = self.store.get_book_snapshots(game.game_id)
            if select_total_snapshots(snapshots, any_book):
                with_fg_line += 1
            counts = period_coverage(snapshots, any_book, native_periods(self.config.league_enum))
            for label in per_period:
                per_period[label] += counts.get(label, 0)
            if (
                game.status is GameStatus.FINAL
                and game.actual is not None
                and game.actual.reconciles()[0]
            ):
                finals += 1
        return self._report(
            mode="store",
            fixtures_found=len(games),
            extra={
                "games_in_store": len(games),
                "games_with_prediction": with_preds,
                "games_with_full_game_line": with_fg_line,
                "full_game_line_coverage": {
                    "with_line": with_fg_line,
                    "total_games": len(games),
                },
                "quarter_lines": (per_period if native_labels == ["Q1", "Q2", "Q3", "Q4"] else {}),
                "quarter_line_coverage": (
                    per_period if native_labels == ["Q1", "Q2", "Q3", "Q4"] else {}
                ),
                "half_lines": per_period if native_labels == ["H1", "H2"] else {},
                "native_period_lines": per_period,
                "native_period_line_coverage": per_period,
                "finals_reconciled": finals,
                "gaps": [],
            },
        )

    # -- acquisition helpers -------------------------------------------------

    def _discover_games(self) -> List[Game]:
        league = self.config.league_enum
        espn_books: Dict[str, List[BookSnapshot]] = {}
        if league in ESPN_LEAGUES:
            games, espn_books = espn.discover_games_with_evidence(
                self.config.date_yyyymmdd,
                self.espn_fetch,
                league,
                observed_utc=self._now_iso(),
            )
        elif league is League.CBA:
            # Official keyless league data (cbaleague data-server). The paid
            # iSportsAPI adapter remains book-only; one discovery source per
            # league prevents duplicate canonical games.
            games = cbaleague.discover_games(self.config.date_yyyymmdd, self.fetch, league)
            self.gaps.extend(cbaleague.last_gaps())
        elif league is League.EUROLEAGUE:
            games = euroleague.discover_games(_iso(self.config.date_yyyymmdd), self.fetch)
        else:
            raise ValueError(f"unsupported league for discovery: {league.value}")
        for game in games:
            self.store.upsert_game(game)
        for game_id, snapshots in espn_books.items():
            for snapshot in snapshots:
                try:
                    self.store.upsert_book_snapshot(snapshot)
                except Exception as exc:
                    logger.warning("ESPN scoreboard book store failed: %s", exc)
                    self.gaps.append(
                        CoverageGap(
                            game_id=game_id,
                            required_field="espn_scoreboard_book",
                            reason=CoverageState.PROVIDER_ERROR,
                            provider="espn",
                            detail=str(exc)[:200],
                            observed_utc=self._now_iso(),
                        )
                    )
        # Bounded pre-tip odds catch-up: if any upcoming game within 36h has no book snapshot,
        # re-fetch the scoreboard once per league letting book_snapshots_from_event do normal parsing.
        # Cost is one extra HTTP request per league at most, within http_request_budget.
        try:
            self._espn_odds_catchup(games)
        except Exception as exc:
            logger.warning("ESPN odds catch-up failed: %s", exc)
        return games

    def _espn_odds_catchup(self, games: List[Game]) -> None:
        """Re-fetch scoreboard once if any near-tip game lacks book coverage.

        Catches per-market gaps (e.g. moneyline present but total missing) and
        benefits from open→close promotion. DraftKings lines post near tip
        (verified 1/6 next-day NBA games had odds at 11:49Z while 8/8 same-day
        had odds). One extra fetch per league per run is bounded and keeps
        every snapshot timestamped.
        """
        if self.config.league_enum not in ESPN_LEAGUES:
            return
        now = parse_utc(self._now_iso())
        need = []
        for game in games:
            if game.status is not GameStatus.SCHEDULED:
                continue
            tip = parse_utc(game.scheduled_utc)
            if tip <= now or tip - now > timedelta(hours=36):
                continue
            snaps = self.store.get_book_snapshots(game.game_id)
            # Free-tier fix: re-fetch if any core market is missing, not only when zero snapshots.
            has_moneyline = any(s.market is Market.MONEYLINE for s in snaps)
            has_total = any(s.market is Market.TOTAL and s.line is not None for s in snaps)
            has_spread = any(s.market is Market.SPREAD and s.line is not None for s in snaps)
            if not (has_moneyline and has_total and has_spread):
                need.append(game.game_id)
        if not need:
            return
        # One extra scoreboard fetch per league per run (bounded, uses AcquisitionClient budget/cooldowns)
        try:
            _, catchup_books = espn.discover_games_with_evidence(
                self.config.date_yyyymmdd,
                self.espn_fetch,
                self.config.league_enum,
                observed_utc=self._now_iso(),
            )
        except Exception:
            return
        hits = 0
        for gid in need:
            snaps = catchup_books.get(gid) or []
            for snap in snaps:
                try:
                    self.store.upsert_book_snapshot(snap)
                    hits += 1
                except Exception:
                    pass
        # Telemetry for monitoring: visible in candidate_funnel via gaps? Use context_report
        self.context_report["odds_catchup_attempts"] = 1
        self.context_report["odds_catchup_hits"] = hits
        self.context_report["odds_catchup_games_without_books"] = len(need)

    def _acquire_context(self, games: List[Game]) -> List[Game]:
        """Enrich upcoming ESPN games with bounded summary/form/injury evidence."""

        if self.config.league_enum not in ESPN_LEAGUES:
            return games
        now = parse_utc(self._now_iso())
        eligible = [
            index
            for index, game in enumerate(games)
            if game.status is GameStatus.SCHEDULED and parse_utc(game.scheduled_utc) > now
        ]
        eligible.sort(key=lambda index: parse_utc(games[index].scheduled_utc))
        attempted = 0
        enriched = 0
        for index in eligible[: self.config.context_request_budget]:
            game = games[index]
            attempted += 1
            try:
                games[index] = espn.fetch_game_context(
                    game.provider_ids.get("espn", game.game_id),
                    self.espn_fetch,
                    game,
                    observed_utc=self._now_iso(),
                )
                self.store.upsert_game(games[index])
                enriched += 1
            except Exception as exc:
                logger.warning("ESPN summary context failed for %s: %s", game.game_id, exc)
                self.gaps.append(
                    CoverageGap(
                        game_id=game.game_id,
                        required_field="espn_summary_context",
                        reason=CoverageState.PROVIDER_ERROR,
                        provider="espn",
                        detail=str(exc)[:200],
                        observed_utc=self._now_iso(),
                    )
                )
        skipped = max(0, len(eligible) - attempted)
        if skipped:
            self.gaps.append(
                CoverageGap(
                    game_id=None,
                    required_field="espn_summary_context_budget",
                    reason=CoverageState.UNSUPPORTED,
                    provider="espn",
                    detail=(
                        f"{skipped} upcoming games exceeded the per-run summary "
                        f"budget of {self.config.context_request_budget}"
                    ),
                    observed_utc=self._now_iso(),
                )
            )
        injury_report: Dict[str, Any] = {"status": "not_applicable"}
        if self.config.league_enum is League.NBA and eligible:
            injury_observed = self._now_iso()
            try:
                snapshot = espn.fetch_nba_injuries(
                    self.espn_fetch,
                    observed_utc=injury_observed,
                )
                covered_teams = set(snapshot["teams"])
                attached_games = 0
                for game in games:
                    if (
                        game.status is not GameStatus.SCHEDULED
                        or parse_utc(game.scheduled_utc) <= parse_utc(injury_observed)
                    ):
                        continue
                    team_ids = {str(game.home_team_id), str(game.away_team_id)}
                    espn_evidence = dict(game.evidence.get("espn") or {})
                    espn_evidence["league_injuries"] = {
                        "schema_version": snapshot["schema_version"],
                        "observed_utc": snapshot["observed_utc"],
                        "source_modified_utc": snapshot["source_modified_utc"],
                        "payload_hash": snapshot["payload_hash"],
                        "teams": {
                            team_id: snapshot["teams"].get(team_id, [])
                            for team_id in team_ids
                        },
                        "team_coverage_complete": team_ids <= covered_teams,
                    }
                    game.evidence = {**game.evidence, "espn": espn_evidence}
                    self.store.upsert_game(game)
                    attached_games += 1
                injury_report = {
                    "status": "ok",
                    "observed_utc": snapshot["observed_utc"],
                    "source_modified_utc": snapshot["source_modified_utc"],
                    "teams_seen": snapshot["teams_seen"],
                    "injury_entries_seen": snapshot["injury_entries_seen"],
                    "dropped_missing_athlete_id": snapshot["dropped_missing_athlete_id"],
                    "dropped_non_object_entries": snapshot["dropped_non_object_entries"],
                    "games_attached": attached_games,
                }
            except Exception as exc:
                injury_report = {"status": "provider_error", "error_type": type(exc).__name__}
                self.gaps.append(
                    CoverageGap(
                        game_id=None,
                        required_field="espn_nba_injuries",
                        reason=CoverageState.PROVIDER_ERROR,
                        provider="espn",
                        detail=f"NBA injury snapshot unavailable: {type(exc).__name__}",
                        observed_utc=injury_observed,
                    )
                )
        self.context_report = {
            "eligible_upcoming_games": len(eligible),
            "attempted": attempted,
            "enriched": enriched,
            "budget": self.config.context_request_budget,
            "budget_skipped": skipped,
            "nba_injuries": injury_report,
        }
        return games

    def _acquire_predictions(self, games: List[Game]) -> Dict[str, SelectionResult]:
        selections: Dict[str, SelectionResult] = {}
        now = self._now_iso()
        artifact = None
        if self.config.league.upper() == "NBA":
            try:
                artifact = hardwood.load_forecasts(self.fetch)
            except Exception as exc:
                logger.warning("hardwood artifact unavailable: %s", exc)
                self.gaps.append(
                    CoverageGap(
                        game_id=None,
                        required_field="hardwood_artifact",
                        reason=CoverageState.PROVIDER_ERROR,
                        provider="hardwood",
                        detail=str(exc)[:200],
                        observed_utc=now,
                    )
                )

        cba_probabilities: Dict[str, Prediction] = {}
        if self.config.league_enum is League.CBA:
            if sportradar.is_enabled():
                try:
                    cba_probabilities = sportradar.fetch_cba_probabilities(
                        self.config.date_yyyymmdd,
                        games,
                        self.fetch,
                        now_utc=self.config.fixture_now_utc,
                    )
                except Exception as exc:
                    logger.warning("Sportradar CBA probabilities unavailable: %s", exc)
                    self.gaps.append(
                        CoverageGap(
                            game_id=None,
                            required_field="sportradar_probabilities",
                            reason=CoverageState.PROVIDER_ERROR,
                            provider="sportradar",
                            detail=str(exc)[:200],
                            observed_utc=now,
                        )
                    )
            else:
                self.gaps.append(
                    CoverageGap(
                        game_id=None,
                        required_field="sportradar_probabilities",
                        reason=CoverageState.UNAUTHORIZED,
                        provider="sportradar",
                        detail="missing SPORTRADAR_API_KEY",
                        observed_utc=now,
                    )
                )

        for game in games:
            # Stored pre-tip observations compete with fresh observations. A
            # transient provider failure must not blank a still-fresh bundle.
            candidates: List[Prediction] = [
                pred for pred in self.store.get_predictions(game.game_id) if pred.pre_tip
            ]
            league = self.config.league_enum
            if league in ESPN_LEAGUES:
                try:
                    candidates.append(
                        espn.fetch_bpi(
                            game.provider_ids.get("espn", game.game_id),
                            self.espn_fetch,
                            game,
                            now_utc=self.config.fixture_now_utc,
                        )
                    )
                except Exception as exc:
                    self.gaps.append(
                        CoverageGap(
                            game_id=game.game_id,
                            required_field="espn_bpi",
                            reason=CoverageState.PROVIDER_ERROR,
                            provider="espn",
                            detail=str(exc)[:200],
                            observed_utc=now,
                        )
                    )
                if league is League.NBA and artifact is not None:
                    candidates.extend(_hardwood_candidates(game, artifact, self.fetch, self.gaps))
            elif league is League.CBA:
                prediction = cba_probabilities.get(game.game_id)
                if prediction is not None:
                    candidates.append(prediction)
            elif league is League.EUROLEAGUE:
                candidates.extend(_predixsport_candidates(game, self.fetch_factory))

            selection = select_prediction(
                candidates,
                game,
                now,
                freshness_hours=self.config.freshness_hours,
                provider_order=tuple(self.config.provider_order),
            )
            tip = parse_utc(game.scheduled_utc)
            for pred in [selection.selected, *selection.shadows]:
                if pred is not None:
                    pred.pre_tip = parse_utc(pred.observed_utc) < tip
            for shadow in selection.shadows:
                try:
                    self.store.upsert_prediction(shadow, frozen=False)
                except Exception as exc:
                    logger.warning("shadow store failed: %s", exc)
            if selection.selected is not None:
                try:
                    self.store.upsert_prediction(selection.selected, frozen=False)
                except Exception as exc:
                    logger.warning("prediction store failed: %s", exc)
            selections[game.game_id] = selection
            self.gaps.extend(selection.gaps)
        return selections

    def _book_provider_options(self, provider: str) -> Dict[str, Any]:
        """Pass deterministic acquisition context to providers when available."""
        if provider == "theoddsapi":
            return {"tz_name": self.config.timezone}
        if provider == "betexplorer" and self.config.fixture_now_utc:
            # Recorded BetExplorer pages include relative Today/Tomorrow labels.
            # Pin fixture replays to their configured observation time instead
            # of the wall clock so results remain stable across calendar days.
            return {"now_utc": self.config.fixture_now_utc}
        return {}

    def _acquire_book_lines(self, games: List[Game]) -> Dict[str, List[BookSnapshot]]:
        raw_snapshots: List[BookSnapshot] = []
        for provider in _book_providers(self.config.league):
            module = _book_module(provider)
            if module is None or not module.is_enabled():
                continue
            try:
                options = self._book_provider_options(provider)
                raw_snapshots.extend(
                    module.fetch_odds(
                        self.config.league.upper(),
                        self.config.date_yyyymmdd,
                        self.fetch_factory,
                        **options,
                    )
                )
                self.quota[provider] = module.last_quota() if hasattr(module, "last_quota") else {}
            except Exception as exc:
                logger.warning("%s fetch failed: %s", provider, exc)
                self.gaps.append(
                    CoverageGap(
                        game_id=None,
                        required_field="book_lines",
                        reason=CoverageState.PROVIDER_ERROR,
                        provider=provider,
                        detail=str(exc)[:200],
                        observed_utc=self._now_iso(),
                    )
                )

        mapped: Dict[str, List[BookSnapshot]] = {}
        for raw in raw_snapshots:
            game, reason = _resolve_book_game(raw, games)
            if game is None:
                self.gaps.append(
                    CoverageGap(
                        game_id=str(raw.game_id),
                        required_field="book_identity",
                        reason=CoverageState.MISSING,
                        provider=raw.provider,
                        detail=reason,
                        observed_utc=self._now_iso(),
                    )
                )
                continue
            provider_event_id = str(raw.provider_event_id or raw.game_id)
            known_event_id = game.provider_ids.get(raw.provider)
            if known_event_id and str(known_event_id) != provider_event_id:
                self.gaps.append(
                    CoverageGap(
                        game_id=game.game_id,
                        required_field="book_identity",
                        reason=CoverageState.MISSING,
                        provider=raw.provider,
                        detail=(
                            f"provider event {provider_event_id} contradicts "
                            f"existing mapping {known_event_id}"
                        ),
                        observed_utc=self._now_iso(),
                    )
                )
                continue
            if not known_event_id:
                game.provider_ids[raw.provider] = provider_event_id
                self.store.upsert_game(game)
            snapshot = replace(raw, game_id=game.game_id, provider_event_id=provider_event_id)
            per_game = mapped.setdefault(game.game_id, [])
            if snapshot not in per_game:
                per_game.append(snapshot)
            try:
                self.store.upsert_book_snapshot(snapshot)
            except Exception as exc:
                logger.warning("book store failed: %s", exc)
        return mapped

    def _visible_pregame_books(self, game: Game) -> List[BookSnapshot]:
        """Return fresh, pre-tip canonical observations visible to this report."""

        now = parse_utc(self._now_iso())
        tip = parse_utc(game.scheduled_utc)
        cutoff = min(now, tip)
        visible = []
        for snapshot in self.store.get_book_snapshots(game.game_id):
            observed = parse_utc(snapshot.observed_utc)
            if observed > cutoff or observed >= tip:
                continue
            if now - observed > timedelta(hours=BOOK_STALENESS_HOURS):
                continue
            visible.append(snapshot)
        return visible

    def _book_fields(self, game: Game, snapshots: List[BookSnapshot]):
        quarters: Dict[str, Optional[float]] = {}
        evidence: Dict[str, Optional[Dict[str, Any]]] = {}
        gaps: List[CoverageGap] = []
        periods = (
            *((period.value, period) for period in native_periods(game.league)),
            ("full_game", Period.FULL_GAME),
        )
        full_game_line = None
        for label, period in periods:
            chosen = newest_snapshot(select_total_snapshots(snapshots, self.config.book, period))
            value = float(chosen.line) if chosen is not None else None
            evidence[label] = _book_evidence(chosen)
            if period is Period.FULL_GAME:
                full_game_line = value
            else:
                quarters[label] = value
            if chosen is None:
                gaps.append(
                    CoverageGap(
                        game_id=game.game_id,
                        required_field=f"book_{period.value.lower()}_line",
                        reason=(CoverageState.MISSING if snapshots else CoverageState.UNSUPPORTED),
                        provider=self.config.book.provider,
                        detail=(
                            f"no fresh pre-match {period.value} total line "
                            "for the configured source"
                        ),
                        observed_utc=self._now_iso(),
                    )
                )
        return quarters, full_game_line, evidence, gaps

    def _odds_evidence_for(
        self, snapshots: List[BookSnapshot], pred: Optional[Prediction]
    ) -> tuple[Optional[float], Optional[Dict[str, Any]]]:
        candidates = []
        for snapshot in select_moneyline_snapshots(snapshots, self.config.book):
            price = odds_for_prediction(snapshot, pred)
            if price is not None:
                candidates.append((snapshot, price))
        if not candidates:
            return None, None
        snapshot, price = max(candidates, key=lambda item: parse_utc(item[0].observed_utc))
        return price, _book_evidence(snapshot)

    def _frozen_odds_evidence(
        self,
        snapshots: List[BookSnapshot],
        pred: Optional[Prediction],
        game: Game,
    ) -> tuple[Optional[float], Optional[Dict[str, Any]]]:
        tip = parse_utc(game.scheduled_utc)
        candidates = [
            snapshot
            for snapshot in select_moneyline_snapshots(snapshots, self.config.book)
            if parse_utc(snapshot.observed_utc) < tip
        ]
        frozen = [
            snapshot
            for snapshot in candidates
            if self.store.is_book_snapshot_frozen(game.game_id, snapshot)
        ]
        return self._odds_evidence_for(frozen or candidates, pred)

    def _freeze_books(self, game: Game) -> int:
        from .store import FrozenWriteError

        frozen = 0
        by_key: Dict[tuple, BookSnapshot] = {}
        for snapshot in self.store.get_book_snapshots(game.game_id):
            if not snapshot.pre_match or parse_utc(snapshot.observed_utc) >= parse_utc(
                game.scheduled_utc
            ):
                continue
            key = (
                snapshot.provider,
                snapshot.book,
                snapshot.market.value,
                snapshot.period.value,
            )
            previous = by_key.get(key)
            if previous is None or parse_utc(snapshot.observed_utc) > parse_utc(
                previous.observed_utc
            ):
                by_key[key] = snapshot
        for _key, snapshot in sorted(by_key.items(), key=lambda item: str(item[0])):
            if self.store.is_book_snapshot_frozen(game.game_id, snapshot):
                continue
            try:
                self.store.upsert_book_snapshot(snapshot, frozen=True)
                frozen += 1
            except FrozenWriteError:
                pass
        return frozen

    def _games_with_tip_passed(self) -> List[Game]:
        now = parse_utc(self._now_iso())
        return [
            game
            for game in self.store.list_games(league=self.config.league_enum)
            if parse_utc(game.scheduled_utc) <= now
        ]

    def _default_history_start_season(self) -> int:
        target = datetime.strptime(self.config.date_yyyymmdd, "%Y%m%d").date()
        current = target.year
        if self.config.league_enum is not League.WNBA and target.month >= 10:
            current += 1
        return current - 1

    def _store_history_cursor(
        self, season: int, event_index: int, index_year: int, event_count: int
    ) -> None:
        now = self._now_iso()
        self.store.upsert_provider_health(
            ProviderHealth(
                provider="espn-history",
                league=self.config.league_enum,
                last_attempt_utc=now,
                last_success_utc=now,
                state=CoverageState.OK,
                detail="durable deep-history acquisition cursor",
                coverage={
                    "next_season": int(season),
                    "next_event_index": int(event_index),
                    "last_index_year": int(index_year),
                    "last_index_event_count": int(event_count),
                    "floor_season": espn.HISTORY_FLOORS[self.config.league_enum],
                },
            )
        )

    def _backfill_deep_history(self) -> Dict[str, Any]:
        """Ingest a bounded batch from the core season/event manifest."""

        budget = max(0, int(self.config.deep_history_event_budget))
        if not budget:
            return {"status": "disabled", "event_budget": 0, "ledger_dates": []}
        base = {
            "event_budget": budget,
            "ledger_dates": [],
            "events_acquired": 0,
            "final_scores_written": 0,
            "already_persisted": 0,
            "coverage": {},
            "learning_or_calibration_claimed": False,
        }
        if not (self.allow_writes and getattr(self.store, "persist", False)):
            return {
                **base,
                "status": "disabled_canonical_write_required",
                "detail": "dry runs cannot advance a cursor past unpersisted history",
            }

        health = self.store.get_provider_health("espn-history", self.config.league_enum)
        previous = health.coverage if health else {}
        default_start = self._default_history_start_season()
        if self.config.deep_history_start_season is not None:
            season = int(self.config.deep_history_start_season)
            event_index = 0
        else:
            season = int(previous.get("next_season") or default_start)
            event_index = int(previous.get("next_event_index") or 0)
        season = min(season, default_start)
        floor = espn.HISTORY_FLOORS[self.config.league_enum]
        base["floor_season"] = floor
        base["start_cursor"] = {"season": season, "event_index": event_index}
        if season < floor:
            return {
                **base,
                "status": "complete_to_verified_index_floor",
                "cursor": {"season": season, "event_index": event_index},
            }

        def index_fetch(url: str):
            return self.acquisition.get_json(url, use_cache=True, ttl_seconds=7 * 86400)

        try:
            manifest = espn.season_event_index(season, index_fetch, self.config.league_enum)
        except Exception as exc:
            self.gaps.append(
                CoverageGap(
                    game_id=None,
                    required_field="deep_history_season_index",
                    reason=CoverageState.PROVIDER_ERROR,
                    provider="espn",
                    detail=f"{season}: {str(exc)[:180]}",
                    observed_utc=self._now_iso(),
                )
            )
            return {
                **base,
                "status": "provider_error",
                "cursor": {"season": season, "event_index": event_index},
            }

        event_ids = manifest["event_ids"]
        base["index"] = {
            "season": season,
            "provider_count": manifest["count"],
            "resolved_event_ids": len(event_ids),
            "completeness": manifest["completeness"],
        }
        if not event_ids:
            next_season = season - 1
            self._store_history_cursor(next_season, 0, season, 0)
            return {
                **base,
                "status": "index_empty",
                "cursor": {"season": next_season, "event_index": 0},
            }

        acquired = 0
        final_scores = 0
        already = 0
        coverage_counts: Dict[str, int] = {}
        ledger_dates = set()
        processed = 0
        status = "acquired"
        while event_index < len(event_ids) and processed < budget:
            event_id = event_ids[event_index]
            existing = self.store.get_game(event_id)
            existing_espn = (existing.evidence.get("espn") or {}) if existing else {}
            existing_history = existing_espn.get("history") or {}
            period_parser_current = (
                self.config.league_enum is not League.NCAA_W
                or existing_history.get("result_period_parser_version")
                == espn.RESULT_PERIOD_PARSER_VERSION
            )
            if (
                existing
                and existing_history
                and existing_espn.get("summary")
                and period_parser_current
            ):
                already += 1
                event_index += 1
                processed += 1
                self._store_history_cursor(season, event_index, season, len(event_ids))
                continue
            try:
                game = espn.fetch_historical_event(
                    event_id,
                    self.espn_fetch,
                    self.config.league_enum,
                    observed_utc=self._now_iso(),
                    index_year=season,
                )
            except Exception as exc:
                logger.warning("ESPN deep history failed for %s: %s", event_id, exc)
                self.gaps.append(
                    CoverageGap(
                        game_id=event_id,
                        required_field="deep_history_event",
                        reason=CoverageState.PROVIDER_ERROR,
                        provider="espn",
                        detail=str(exc)[:200],
                        observed_utc=self._now_iso(),
                    )
                )
                status = "provider_error_cursor_retained"
                break

            self.store.upsert_game(game)
            acquired += 1
            processed += 1
            event_index += 1
            coverage = game.evidence["espn"]["history"]["coverage"]
            coverage_counts[coverage] = coverage_counts.get(coverage, 0) + 1
            if game.actual and game.actual.home_final is not None:
                final_scores += 1
            local_date = (
                parse_utc(game.scheduled_utc)
                .astimezone(ZoneInfo(self.config.timezone))
                .strftime("%Y%m%d")
            )
            ledger_dates.add(local_date)
            if coverage == "final_score_only":
                self.gaps.append(
                    CoverageGap(
                        game_id=event_id,
                        required_field="native_periods",
                        reason=CoverageState.MISSING,
                        provider="espn",
                        detail=game.evidence["espn"]["history"]["period_note"],
                        observed_utc=self._now_iso(),
                    )
                )
            elif coverage in {"unreconciled_periods", "non_final_provider_status"}:
                self.gaps.append(
                    CoverageGap(
                        game_id=event_id,
                        required_field="historical_result_integrity",
                        reason=CoverageState.PROVIDER_ERROR,
                        provider="espn",
                        detail=game.evidence["espn"]["history"]["period_note"],
                        observed_utc=self._now_iso(),
                    )
                )
            self._store_history_cursor(season, event_index, season, len(event_ids))

        if event_index >= len(event_ids):
            season -= 1
            event_index = 0
            self._store_history_cursor(season, 0, season + 1, len(event_ids))
        return {
            **base,
            "status": status,
            "events_acquired": acquired,
            "final_scores_written": final_scores,
            "already_persisted": already,
            "coverage": coverage_counts,
            "ledger_dates": sorted(ledger_dates),
            "cursor": {"season": season, "event_index": event_index},
        }

    def _record_history_cursor(self, provider: str, coverage: Dict[str, Any]) -> None:
        """Persist a durable non-ESPN history cursor in provider_health."""
        now = self._now_iso()
        self.store.upsert_provider_health(
            ProviderHealth(
                provider=provider,
                league=self.config.league_enum,
                last_attempt_utc=now,
                last_success_utc=now,
                state=CoverageState.OK,
                detail="durable deep-history acquisition cursor",
                coverage=coverage,
            )
        )

    def _history_cursor(self, provider: str) -> Dict[str, Any]:
        health = self.store.get_provider_health(provider, self.config.league_enum)
        return dict(health.coverage or {}) if health else {}

    def _ledger_date_for(self, game: Game) -> str:
        return (
            parse_utc(game.scheduled_utc)
            .astimezone(ZoneInfo(self.config.timezone))
            .strftime("%Y%m%d")
        )

    def _backfill_euroleague_history(self) -> Dict[str, Any]:
        """Ingest whole seasons of reconcilable finals, one request per season.

        The official season schedule carries each played game's official
        score and per-quarter partials (verified 2026-11-16), so one request
        covers a full season of reconcilable history. Seasons walk backwards
        under ``history_request_budget``; the cursor advances only on
        canonical writes, mirroring the ESPN deep-history contract. There is
        deliberately no league-wide game cap: whole seasons enter the store.
        """
        budget = max(0, int(self.config.history_request_budget))
        base: Dict[str, Any] = {
            "status": "acquired",
            "seasons": [],
            "finals_written": 0,
            "quarantined_rows": 0,
            "ledger_dates": [],
            "learning_or_calibration_claimed": False,
        }
        if not budget:
            return {**base, "status": "disabled"}
        if not (self.allow_writes and getattr(self.store, "persist", False)):
            return {
                **base,
                "status": "disabled_canonical_write_required",
                "detail": "dry runs cannot advance a cursor past unpersisted history",
            }

        cursor = self._history_cursor("euroleague-history")
        season = int(cursor.get("next_season") or self._default_history_start_season())
        season = min(season, self._default_history_start_season())
        ledger_dates: set = set()
        while budget > 0 and season > 1900:
            seasoncode = f"E{season}"
            failures: List[Dict[str, str]] = []
            try:
                results = euroleague.season_results(seasoncode, self.fetch, failures)
            except Exception as exc:
                self.gaps.append(
                    CoverageGap(
                        game_id=None,
                        required_field="euroleague_season_schedule",
                        reason=CoverageState.PROVIDER_ERROR,
                        provider="euroleague",
                        detail=f"{seasoncode}: {str(exc)[:180]}",
                        observed_utc=self._now_iso(),
                    )
                )
                base["status"] = "provider_error_cursor_retained"
                base["cursor"] = {"next_season": season}
                return base
            written = 0
            for game in results:
                self.store.upsert_game(game)
                written += 1
                ledger_dates.add(self._ledger_date_for(game))
            for failure in failures:
                self.gaps.append(
                    CoverageGap(
                        game_id=failure.get("game_id"),
                        required_field="historical_result_integrity",
                        reason=CoverageState.PROVIDER_ERROR,
                        provider="euroleague",
                        detail=str(failure.get("error", ""))[:200],
                        observed_utc=self._now_iso(),
                    )
                )
            base["seasons"].append(
                {
                    "season": seasoncode,
                    "finals_written": written,
                    "quarantined_rows": len(failures),
                }
            )
            base["finals_written"] += written
            base["quarantined_rows"] += len(failures)
            budget -= 1
            season -= 1
            self._record_history_cursor("euroleague-history", {"next_season": season})
        base["cursor"] = {"next_season": season}
        base["ledger_dates"] = sorted(ledger_dates)
        return base

    def _backfill_cba_history(self) -> Dict[str, Any]:
        """Walk official CBA days backwards, ingesting reconciled finals.

        Each finished game costs one period-score POST request, bounded by
        ``deep_history_event_budget``; day scans are bounded separately by
        ``history_request_budget``. The source's season list sets the oldest
        day this cursor will scan. A day with an unprocessed final or a result
        failure is retained for retry; completed days are idempotently written
        before the cursor moves backwards.
        """
        event_budget = max(0, int(self.config.deep_history_event_budget))
        day_budget = max(0, int(self.config.history_request_budget))
        base: Dict[str, Any] = {
            "status": "acquired",
            "days_scanned": [],
            "finals_written": 0,
            "ledger_dates": [],
            "learning_or_calibration_claimed": False,
        }
        if not event_budget or not day_budget:
            return {**base, "status": "disabled"}
        if not (self.allow_writes and getattr(self.store, "persist", False)):
            return {
                **base,
                "status": "disabled_canonical_write_required",
                "detail": "dry runs cannot advance a cursor past unpersisted history",
            }

        try:
            seasons = cbaleague.available_seasons(self.fetch)
        except Exception as exc:
            self.gaps.append(
                CoverageGap(
                    game_id=None,
                    required_field="cba_season_manifest",
                    reason=CoverageState.PROVIDER_ERROR,
                    provider="cbaleague",
                    detail=str(exc)[:200],
                    observed_utc=self._now_iso(),
                )
            )
            base.update(status="provider_error_cursor_retained", finals_written=0)
            return base
        floor = datetime(min(seasons), 1, 1).date()
        base["source_floor_day"] = floor.isoformat()
        cursor = self._history_cursor("cbaleague-history")
        target = datetime.strptime(self.config.date_yyyymmdd, "%Y%m%d").date()
        day = datetime.strptime(cursor["next_day"], "%Y-%m-%d").date() if cursor.get("next_day") else target
        processed_ids = set(cursor.get("processed_schedule_ids") or [])
        if day < floor:
            base.update(
                status="complete_to_source_season_floor",
                cursor={"next_day": day.isoformat()},
            )
            return base

        scanned = 0
        events = 0
        status = "acquired"
        while scanned < day_budget and events < event_budget and day >= floor:
            day_iso = day.isoformat()
            try:
                rows = cbaleague.day_rows(day_iso, self.fetch)
            except Exception as exc:
                self.gaps.append(
                    CoverageGap(
                        game_id=None,
                        required_field="cba_schedule_day",
                        reason=CoverageState.PROVIDER_ERROR,
                        provider="cbaleague",
                        detail=f"{day_iso}: {str(exc)[:180]}",
                        observed_utc=self._now_iso(),
                    )
                )
                status = "provider_error_cursor_retained"
                break
            scanned += 1
            base["days_scanned"].append(day_iso)
            event_budget_exhausted = False
            result_failure = False
            for row in rows:
                schedule_id = str(row.get("scheduleId") or "")
                if row.get("status") != 4 or schedule_id in processed_ids:
                    continue
                if events >= event_budget:
                    event_budget_exhausted = True
                    break
                try:
                    game = cbaleague.result_from_row(row, self.fetch)
                except Exception as exc:
                    self.gaps.append(
                        CoverageGap(
                            game_id=str(row.get("scheduleId") or "") or None,
                            required_field="historical_result_integrity",
                            reason=CoverageState.PROVIDER_ERROR,
                            provider="cbaleague",
                            detail=str(exc)[:200],
                            observed_utc=self._now_iso(),
                        )
                    )
                    result_failure = True
                    break
                self.store.upsert_game(game)
                events += 1
                processed_ids.add(schedule_id)
                base["ledger_dates"].append(self._ledger_date_for(game))

            if result_failure:
                status = "provider_error_cursor_retained"
                break
            if event_budget_exhausted:
                status = "event_budget_exhausted_day_retained"
                break

            day -= timedelta(days=1)
            processed_ids = set()
            if day < floor:
                status = "complete_to_source_season_floor"
                break

        if status != "provider_error_cursor_retained" and day >= floor:
            self._record_history_cursor(
                "cbaleague-history",
                {
                    "next_day": day.isoformat(),
                    "processed_schedule_ids": sorted(processed_ids),
                },
            )
        elif status != "provider_error_cursor_retained":

            # Keep a terminal cursor just below the discovered source floor so
            # later runs can stop without walking dates the API never serves.
            self._record_history_cursor(
                "cbaleague-history",
                {
                    "next_day": (floor - timedelta(days=1)).isoformat(),
                    "processed_schedule_ids": [],
                },
            )
        base["finals_written"] = events
        base["cursor"] = {"next_day": day.isoformat()}
        base["status"] = status
        base["ledger_dates"] = sorted(set(base["ledger_dates"]))
        return base

    def backfill(self) -> Dict[str, Any]:
        """Acquire recent scoreboards plus a bounded durable deep-history batch.

        Historical evidence never implies training, calibration, or acca
        eligibility. Only write-enabled runs advance the deep event cursor.
        """

        self._reset_stage()
        configured_days = max(0, int(self.config.history_days))
        budget = max(0, int(self.config.history_request_budget))
        days = min(configured_days, budget)
        espn_enabled = self.config.league_enum in ESPN_LEAGUES
        total_found = 0
        total_written = 0
        date_reports = []
        ledger_dates = set()
        if espn_enabled and configured_days > budget:
            self.gaps.append(
                CoverageGap(
                    game_id=None,
                    required_field="historical_backfill_budget",
                    reason=CoverageState.UNSUPPORTED,
                    provider="espn",
                    detail=(
                        f"requested {configured_days} days; bounded to {budget} "
                        "scoreboard requests in this run"
                    ),
                    observed_utc=self._now_iso(),
                )
            )

        target = datetime.strptime(self.config.date_yyyymmdd, "%Y%m%d").date()
        if espn_enabled and days:
            for offset in range(1, days + 1):
                day = target - timedelta(days=offset)
                date = day.strftime("%Y%m%d")

                def cached_fetch(url: str):
                    return self.acquisition.get_json(url, use_cache=True, ttl_seconds=86400)

                found = 0
                written = 0
                reconciled = 0
                try:
                    finals = espn.fetch_recent_results(
                        date,
                        cached_fetch,
                        self.config.league_enum,
                        observed_utc=lambda: self.acquisition.last_observed_utc or self._now_iso(),
                    )
                except Exception as exc:
                    logger.warning("ESPN historical backfill failed for %s: %s", date, exc)
                    self.gaps.append(
                        CoverageGap(
                            game_id=None,
                            required_field="historical_results",
                            reason=CoverageState.PROVIDER_ERROR,
                            provider="espn",
                            detail=f"{date}: {str(exc)[:180]}",
                            observed_utc=self._now_iso(),
                        )
                    )
                    date_reports.append(
                        {"date": date, "status": "provider_error", "finals_found": 0}
                    )
                    continue
                for game in finals:
                    found += 1
                    total_found += 1
                    if game.actual is None:
                        continue
                    ok, reason = game.actual.reconciles()
                    if not ok:
                        self.gaps.append(
                            CoverageGap(
                                game_id=game.game_id,
                                required_field=(
                                    "native_periods"
                                    if game.actual.period_kind.value == "unknown"
                                    else "historical_reconciliation"
                                ),
                                reason=(
                                    CoverageState.MISSING
                                    if game.actual.period_kind.value == "unknown"
                                    else CoverageState.PROVIDER_ERROR
                                ),
                                provider="espn",
                                detail=reason,
                                observed_utc=self._now_iso(),
                            )
                        )
                    else:
                        reconciled += 1
                    self.store.upsert_game(game)
                    written += 1
                    total_written += 1
                if written:
                    ledger_dates.add(date)
                date_reports.append(
                    {
                        "date": date,
                        "status": "acquired",
                        "finals_found": found,
                        "finals_reconciled": reconciled,
                        "canonical_persistence": (
                            "duckdb_csv" if self.allow_writes else "dry_run_not_persisted"
                        ),
                    }
                )

        if espn_enabled:
            deep = self._backfill_deep_history()
        elif self.config.league_enum is League.EUROLEAGUE:
            deep = self._backfill_euroleague_history()
        elif self.config.league_enum is League.CBA:
            deep = self._backfill_cba_history()
        else:
            deep = {"status": "unsupported_league", "ledger_dates": []}
        ledger_dates.update(deep.get("ledger_dates") or [])
        status = (
            "disabled"
            if not days and deep["status"] in {"disabled", "unsupported_league"}
            else "acquired"
        )
        return self._report(
            mode="write" if self.allow_writes else "dry_run",
            fixtures_found=total_found + int(deep.get("events_acquired") or 0),
            extra={
                "finals_written": total_written + int(deep.get("final_scores_written") or 0),
                "backfill": {
                    "status": status,
                    "requested_days": configured_days,
                    "request_budget": budget,
                    "dates": date_reports,
                    "ledger_dates": sorted(ledger_dates),
                    "deep_history": deep,
                    "learning_or_calibration_claimed": False,
                },
            },
        )

    def _fetch_finals(self) -> List[Game]:
        league = self.config.league_enum
        if league in ESPN_LEAGUES:
            return espn.fetch_recent_results(
                self.config.date_yyyymmdd,
                self.espn_fetch,
                league,
                observed_utc=self._now_iso(),
            )
        if league is League.CBA:
            finals = cbaleague.fetch_recent_results(
                self.config.date_yyyymmdd, self.fetch, league
            )
            self.gaps.extend(cbaleague.last_gaps())
            return finals
        if league is League.EUROLEAGUE:
            return self._euroleague_finals(self._discover_games())
        raise ValueError(f"unsupported league for results: {league.value}")

    def _euroleague_finals(self, games: List[Game]) -> List[Game]:
        finals = []
        for game in games:
            if game.status is GameStatus.FINAL:
                try:
                    finals.append(euroleague.fetch_game_result(game, self.fetch))
                except Exception as exc:
                    logger.warning("euroleague result failed for %s: %s", game.game_id, exc)
                    self.gaps.append(
                        CoverageGap(
                            game_id=game.game_id,
                            required_field="actual_result",
                            reason=CoverageState.PROVIDER_ERROR,
                            provider="euroleague",
                            detail=str(exc)[:200],
                            observed_utc=self._now_iso(),
                        )
                    )
        return finals

    # -- smoke helpers -------------------------------------------------------

    def _hardwood_smoke(self) -> Dict[str, Any]:
        try:
            artifact = hardwood.load_forecasts(self.fetch)
            state = hardwood.freshness_state(artifact)
            fetch_target = getattr(self.fetch, "__self__", self.fetch)
            return {
                "verification": (
                    "LIVE_VERIFIED" if isinstance(fetch_target, HttpClient) else "FIXTURE_VERIFIED"
                ),
                "generated_at": artifact.generated_at,
                "ingestion_status": artifact.ingestion_status,
                "freshness_state": state.value,
                "games_in_artifact": len(artifact.games_by_id),
            }
        except Exception as exc:
            return {"verification": "PROVIDER_ERROR", "detail": str(exc)[:200]}

    def _predixsport_smoke(self) -> Dict[str, Any]:
        from .providers import predixsport

        if not predixsport.is_enabled():
            return {
                "verification": "UNAUTHORIZED",
                "enabled": False,
                "note": "PREDIXSPORT_API_KEY missing; provider disabled",
            }
        try:
            return predixsport.smoke(self.config.league.lower(), self.fetch_factory)
        except Exception as exc:
            return {"verification": "PROVIDER_ERROR", "detail": str(exc)[:200]}

    def _bookmakers_smoke(self) -> Dict[str, Any]:
        output = {}
        for provider in _book_providers(self.config.league):
            module = _book_module(provider)
            if module is None:
                continue
            if not module.is_enabled():
                output[provider] = {
                    "enabled": False,
                    "verification": "UNAUTHORIZED",
                    "note": "API key missing; provider disabled",
                }
                continue
            try:
                options = self._book_provider_options(provider)
                output[provider] = module.smoke(
                    self.config.league.upper(), self.fetch_factory, **options
                )
            except Exception as exc:
                output[provider] = {
                    "enabled": True,
                    "verification": "PROVIDER_ERROR",
                    "detail": str(exc)[:200],
                }
        return output

    # -- report --------------------------------------------------------------

    def _report(self, mode: str, fixtures_found: int, extra: Dict[str, Any]) -> Dict[str, Any]:
        report = CoverageReport(
            league=self.config.league.upper(),
            generated_utc=self._now_iso(),
            mode=mode,
            fixtures_found=fixtures_found,
            finals_written=extra.get("finals_written", 0),
            frozen_predictions_used=extra.get("frozen_predictions_used", 0),
            predictions_selected=extra.get("predictions_selected", {}),
            gaps=[gap.__dict__.copy() for gap in self.gaps],
            quota=self.quota,
        )
        report_dict = report.to_dict()
        store_kind = {
            "DuckDBStore": "duckdb",
            "InMemoryStore": "in_memory_ephemeral",
        }.get(type(self.store).__name__, type(self.store).__name__)
        report_dict["acquisition"] = self.acquisition.telemetry()
        report_dict["persistence"] = {
            "kind": store_kind,
            "write_enabled": bool(getattr(self.store, "persist", False)),
            "unattended_state_configured": store_kind == "duckdb",
            "system_of_record": ("duckdb_csv" if store_kind == "duckdb" else "ephemeral"),
        }
        report_dict.update(extra)
        return report_dict


# ---------------------------------------------------------------------------
# Provider and identity helpers
# ---------------------------------------------------------------------------


def _prediction_facts(pred: Optional[Prediction]) -> Optional[Dict[str, Any]]:
    if pred is None:
        return None
    return {
        "provider": pred.provider,
        "observed_utc": pred.observed_utc,
        "source_modified_utc": pred.source_modified_utc,
        "source_revision": pred.source_revision,
        "model": pred.model,
        "model_version": pred.model_version,
        "produced_by": pred.produced_by,
        "pre_tip": pred.pre_tip,
        "prob_home": pred.prob_home,
        "prob_away": pred.prob_away,
        "expected_home_score": pred.expected_home_score,
        "expected_away_score": pred.expected_away_score,
        "expected_total": pred.expected_total,
        "expected_margin": pred.expected_margin,
        "auxiliary_metrics": pred.auxiliary_metrics,
    }


def _book_evidence(snapshot: Optional[BookSnapshot]) -> Optional[Dict[str, Any]]:
    if snapshot is None:
        return None
    return {
        "provider": snapshot.provider,
        "book": snapshot.book,
        "observed_utc": snapshot.observed_utc,
        "source_modified_utc": snapshot.source_modified_utc,
        "payload_hash": snapshot.payload_hash,
    }


def _book_identity_mismatch(snapshot: BookSnapshot, game: Game) -> str:
    if snapshot.home_name and canonical_team_key(snapshot.home_name) != canonical_team_key(
        game.home_name
    ):
        return "home-team metadata contradicts the mapped game"
    if snapshot.away_name and canonical_team_key(snapshot.away_name) != canonical_team_key(
        game.away_name
    ):
        return "away-team metadata contradicts the mapped game"
    if snapshot.scheduled_utc:
        try:
            if not times_within(snapshot.scheduled_utc, game.scheduled_utc):
                return "start-time metadata contradicts the mapped game"
        except (TypeError, ValueError):
            return "invalid start-time metadata"
    return ""


def _resolve_book_game(snapshot: BookSnapshot, games: Sequence[Game]) -> tuple[Optional[Game], str]:
    """Resolve a snapshot to one game without trusting an ID coincidence."""

    raw_id = str(snapshot.provider_event_id or snapshot.game_id)
    known = [
        game for game in games if str(game.provider_ids.get(snapshot.provider) or "") == raw_id
    ]
    if snapshot.provider_event_id is None:
        known.extend(game for game in games if game.game_id == raw_id and game not in known)
    if len(known) > 1:
        return None, "provider/canonical ID maps to multiple games"
    if known:
        mismatch = _book_identity_mismatch(snapshot, known[0])
        return (None, mismatch) if mismatch else (known[0], "")

    if not (snapshot.home_name and snapshot.away_name and snapshot.scheduled_utc):
        return None, "unmapped provider event lacks home/away/start identity metadata"
    candidates = []
    for game in games:
        if _book_identity_mismatch(snapshot, game):
            continue
        if canonical_team_key(snapshot.home_name) != canonical_team_key(game.home_name):
            continue
        if canonical_team_key(snapshot.away_name) != canonical_team_key(game.away_name):
            continue
        candidates.append(game)
    if len(candidates) == 1:
        return candidates[0], ""
    if len(candidates) > 1:
        return None, "ambiguous home/away/start match across multiple games"
    return None, "no game matched the provider event's home/away/start identity"


def _default_fetch_factory(provider: str):
    """Build an authenticated fetch without exposing credentials in URLs."""

    import os

    if provider == "predixsport":
        key = os.environ.get("PREDIXSPORT_API_KEY", "")
        return (
            HttpClient(default_headers={"Authorization": f"Bearer {key}"}).get_json
            if key
            else HttpClient().get_json
        )
    if provider == "therundown":
        from .providers import therundown

        return (
            HttpClient(default_headers=therundown.auth_headers()).get_json
            if therundown.is_enabled()
            else HttpClient().get_json
        )
    if provider == "sportsgameodds":
        key = os.environ.get("SPORTSGAMEODDS_API_KEY", "")
        return (
            HttpClient(default_headers={"x-api-key": key}).get_json
            if key
            else HttpClient().get_json
        )
    return HttpClient().get_json


def _book_module(provider: str):
    from .providers import isportsapi, sportsgameodds, theoddsapi, therundown

    return {
        "betexplorer": betexplorer,
        "therundown": therundown,
        "sportsgameodds": sportsgameodds,
        "theoddsapi": theoddsapi,
        "isportsapi": isportsapi,
    }[provider]


def _hardwood_candidates(game: Game, artifact, fetch, gaps: List[CoverageGap]) -> List[Prediction]:
    output: List[Prediction] = []
    local_gaps: List[CoverageGap] = []
    prediction = hardwood.fetch_prediction(game, fetch, gaps=local_gaps)
    if prediction is not None:
        output.append(prediction)
    gaps.extend(local_gaps)
    return output


def _predixsport_candidates(game: Game, fetch_factory) -> List[Prediction]:
    from .providers import predixsport

    if not predixsport.is_enabled():
        return []
    try:
        predictions = predixsport.fetch_upcoming(
            "nba" if game.league is League.NBA else "euroleague", 7, fetch_factory
        )
    except Exception as exc:
        logger.warning("predixsport upcoming failed: %s", exc)
        return []
    for prediction in predictions:
        if prediction.game_id in {
            game.provider_ids.get("predixsport"),
            game.game_id,
        }:
            return [prediction]
    return []


def _iso(date_yyyymmdd: str) -> str:
    return f"{date_yyyymmdd[:4]}-{date_yyyymmdd[4:6]}-{date_yyyymmdd[6:8]}"


def _count_by(values: Sequence[str]) -> Dict[str, int]:
    output: Dict[str, int] = {}
    for value in values:
        output[value] = output.get(value, 0) + 1
    return output
