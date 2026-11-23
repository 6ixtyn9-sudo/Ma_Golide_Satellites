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
from .http import HttpClient, HttpError
from .identity import (
    canonical_team_key,
    now_utc_iso,
    parse_utc,
    team_names_agree,
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
    PriceProvenance,
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
from .league_registry import SPECS, league_spec, parse_league, price_evidence_policy
from .leagues import native_periods
from .providers import (
    aba_liga,
    acb,
    betexplorer,
    betexplorer_direct,
    cbaleague,
    espn,
    euroleague,
    fiba_euroleague_w,
    hardwood,
    israel_ibl,
    lba,
)

logger = logging.getLogger(__name__)

# Games in these states will never produce an official final to wait for.
_CLOSED_STATUSES = frozenset({GameStatus.POSTPONED, GameStatus.CANCELLED, GameStatus.ABANDONED})

# All public configuration is derived from the central registry.  This keeps
# workflow/CLI/warehouse identities aligned when a new basketball competition
# is added and prevents a second hand-maintained timezone matrix.
DEFAULT_TZ = {spec.league.value: spec.timezone for spec in SPECS}
ACQUISITION_TZ = {spec.league.value: spec.acquisition_timezone for spec in SPECS}
BOOK_PROVIDERS = ("theoddsapi", "betexplorer_direct")
# Explicit per-league tuples are retained as a public compatibility surface and
# include the operator-approved BetExplorer direct fallback.  A source-limited
# league still sends no request unless its canonical schedule produced a future
# candidate fixture for the fallback to match.
BOOK_PROVIDERS_BY_LEAGUE = {spec.league.value: spec.book_providers for spec in SPECS}

# Registry declares which leagues have a validated ESPN schedule/result route.
# ESPN's deeper core-event-index history and BPI are independent contracts with
# narrower allow-lists in the adapter; a new scoreboard lane must not inherit
# either one by association.
ESPN_LEAGUES = {spec.league for spec in SPECS if spec.schedule_source == "espn"}
ESPN_DEEP_HISTORY_LEAGUES = frozenset(espn.HISTORY_FLOORS)
ESPN_BPI_LEAGUES = espn.BPI_LEAGUES

# Most ESPN lanes can lean on a separately validated deep-history index. NBA G
# League cannot: only its date-bounded scoreboard/result contract is approved.
# A cold lane therefore needs a wider *bounded calendar* bootstrap before the
# empirical model can accumulate both residual and team-history evidence. This
# is intentionally an acquisition plan, not a ForecastConfig override: it does
# not loosen a prediction, calibration, price, or paper-ticket gate.
DEFAULT_RECENT_HISTORY_DAYS = 21
NBA_G_LEAGUE_BOOTSTRAP_HISTORY_DAYS = 56


def default_history_backfill_plan(
    league: League,
    *,
    history_days: Optional[int] = None,
    history_request_budget: Optional[int] = None,
) -> Dict[str, int | str]:
    """Return the bounded scoreboard-history plan for one operator run.

    The wider NBA G League plan uses the existing ESPN ``nba-development``
    date endpoint only. It neither probes ESPN's unvalidated core event index
    nor asserts complete historical coverage. Explicit CLI/operator values
    always win; when only ``history_days`` is supplied, its request budget
    follows that value so an intended override is not silently clipped.
    """

    default_days = (
        NBA_G_LEAGUE_BOOTSTRAP_HISTORY_DAYS
        if league is League.NBA_G_LEAGUE
        else DEFAULT_RECENT_HISTORY_DAYS
    )
    if history_days is None and history_request_budget is None:
        return {
            "history_days": default_days,
            "history_request_budget": default_days,
            "policy": (
                "nba_g_league_date_bounded_empirical_history_bootstrap"
                if league is League.NBA_G_LEAGUE
                else "default_recent_scoreboard_backfill"
            ),
        }
    resolved_days = default_days if history_days is None else int(history_days)
    resolved_budget = (
        resolved_days if history_request_budget is None else int(history_request_budget)
    )
    return {
        "history_days": resolved_days,
        "history_request_budget": resolved_budget,
        "policy": "operator_override",
    }


def _book_providers(league: str) -> Sequence[str]:
    parsed = parse_league(league)
    # Preserve the existing overrideable NBA global for fixture/regression
    # harnesses and local operator experiments.  The global includes the same
    # direct fallback advertised by every registry lane; other lanes remain
    # explicitly registry-scoped.
    if parsed is League.NBA:
        return BOOK_PROVIDERS
    return BOOK_PROVIDERS_BY_LEAGUE.get(parsed.value, ())


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
    # Scheduled result/draft work must not implicitly spend a named-price
    # provider credit. The workflow enables this only with a budget permit at
    # the dated ticket-acquisition cutoff.
    acquire_book_lines: bool = True
    # Used only by committed offline fixtures so replay evidence is stable.
    fixture_now_utc: Optional[str] = None

    @property
    def league_enum(self) -> League:
        return parse_league(self.league)

    @property
    def timezone(self) -> str:
        return self.tz_name or league_spec(self.league_enum).timezone


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
        # Fixture tests provide a one-argument fetcher.  Live FIBA acquisition
        # needs its public APIM header and RSC text transport, so it owns a
        # small official client only when no fixture fetcher was supplied.
        self._fiba_client = (
            None
            if fetch is not None or config.league_enum is not League.EUROLEAGUE_W
            else fiba_euroleague_w.FibaClient()
        )
        self._injected_fetch = fetch is not None
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
        # Per-stage source telemetry. This distinguishes a true no-market
        # response from a provider that was intentionally unavailable because
        # its optional credential was not configured.
        self.book_provider_status: Dict[str, Dict[str, Any]] = {}
        self.prediction_provider_status: Dict[str, Dict[str, Any]] = {}

    def _now_iso(self) -> str:
        if self.config.fixture_now_utc:
            return to_utc_iso(parse_utc(self.config.fixture_now_utc))
        return now_utc_iso()

    def _aba_text_fetch(self):
        """Use an injected fixture transport, else the adapter's text client.

        The ordinary pipeline fetch is JSON-oriented. ABA League publishes
        server-rendered HTML, so passing the injected transport only preserves
        hermetic tests while live runs let the adapter select ``get_text``.
        """

        return self.fetch if self._injected_fetch else None

    def _reset_stage(self) -> None:
        """Keep gaps and quota scoped to one command, not cumulative duplicates."""

        self.gaps = []
        self.quota = {}
        self.context_report = {}
        self.book_provider_status = {}
        self.prediction_provider_status = {}

    def _stamp_book_evidence(self, snapshot: BookSnapshot) -> BookSnapshot:
        """Apply the registry evidence contract to every acquired price quote.

        Provider adapters may parse a named bookmaker correctly, but must not
        decide whether their source is permitted or operationally validated.
        The registry is the sole promotion point. Unknown providers remain
        visible as named observations while failing the ticket-price gate.
        """

        policy = price_evidence_policy(self.config.league_enum, snapshot.provider)
        return replace(
            snapshot,
            price_provenance=policy.price_provenance,
            source_permission=policy.source_permission,
            operational_validation=policy.operational_validation,
        )

    @staticmethod
    def _book_policy_status(provider: str, snapshot: BookSnapshot) -> Dict[str, str]:
        """Expose applied evidence axes without treating them as a price quote."""

        return {
            "provider": provider,
            "price_provenance": snapshot.price_provenance.value,
            "source_permission": snapshot.source_permission.value,
            "operational_validation": snapshot.operational_validation.value,
        }

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
        league = self.config.league_enum.value
        if self.config.league_enum in ESPN_LEAGUES:
            report = espn.smoke(self.config.date_yyyymmdd, self.espn_fetch, self.config.league_enum)
            if league == "NBA":
                report["hardwood"] = self._hardwood_smoke()
                # Same one-call probe the EuroLeague lane exposes, so the league
                # that actually has slates can answer "is Predixsport enabled
                # and returning rows?" without waiting for a scheduled run.
                report["predixsport"] = self._predixsport_smoke()
            if self.config.league_enum not in ESPN_BPI_LEAGUES:
                # Schedule/result support has been independently checked, but
                # no external pre-tip model is represented until that contract
                # is separately validated. The empirical model remains an
                # eligible forecast surface only after its own history gate.
                report["predictions"] = {
                    "status": "no_external_prediction_provider_validated",
                    "policy": "empirical_history_only_after_reconciled_history_threshold",
                }
            report["standings"] = self._standings_smoke()
            report["bookmakers"] = self._bookmakers_smoke()
            return report
        if league == "ABA_LEAGUE":
            report = aba_liga.smoke(text_fetch=self._aba_text_fetch())
            report["bookmakers"] = self._bookmakers_smoke()
            return report
        if league == "CBA":
            report = cbaleague.schedule_smoke(
                self.config.date_yyyymmdd, self.fetch, self.config.league_enum
            )
            report["bookmakers"] = self._bookmakers_smoke()
            return report
        if league == "ITALY_LBA":
            try:
                report = lba.smoke(self.fetch, self._lba_season_start_year())
            except Exception as exc:
                report = {"status": "provider_error", "provider": "lba",
                          "error_type": type(exc).__name__}
            report["bookmakers"] = self._bookmakers_smoke()
            return report
        if league == "ISRAEL_PREMIER_LEAGUE":
            try:
                year = self._israel_season_end_year()
                report = israel_ibl.smoke(
                    self.fetch if self._injected_fetch else None, year
                )
            except Exception as exc:
                report = {"status": "provider_error", "provider": "israel_ibl",
                          "error_type": type(exc).__name__}
            report["bookmakers"] = self._bookmakers_smoke()
            return report
        if league == "SPAIN_ACB":
            try:
                year = self._acb_season_end_year()
                report = acb.smoke(self.fetch, year)
            except Exception as exc:
                report = {"status": "provider_error", "provider": "acb",
                          "error_type": type(exc).__name__}
            report["bookmakers"] = self._bookmakers_smoke()
            return report
        if league in ("EUROLEAGUE", "EUROCUP"):
            # EuroCup reuses the validated EuroLeague Basketball adapter with
            # competition code U; a failed probe is reported, never raised.
            try:
                report = euroleague.smoke(self.fetch, self._euro_seasoncode())
            except Exception as exc:
                report = {
                    "status": "provider_error",
                    "provider": "euroleague",
                    "season_code": self._euro_seasoncode(),
                    "error_type": type(exc).__name__,
                }
            if league == "EUROLEAGUE":
                report["standings"] = self._standings_smoke()
                report["predixsport"] = self._predixsport_smoke()
            report["bookmakers"] = self._bookmakers_smoke()
            return report
        if league == "EUROLEAGUE_W":
            report = fiba_euroleague_w.smoke(
                self._fiba_schedule_fetch,
                verification="LIVE_VERIFIED" if self._fiba_client is not None else "FIXTURE_VERIFIED",
            )
            report["predictions"] = {
                "status": "not_available_from_official_fiba_fixture_result_detail_path",
                "policy": "empirical_history_only_until_a_separately_validated_pre_tip_provider_exists",
            }
            report["bookmakers"] = self._bookmakers_smoke()
            return report
        spec = league_spec(self.config.league_enum)
        return {
            "league": spec.league.value,
            "status": "source_unavailable",
            "schedule": {
                "provider": spec.schedule_source,
                "state": CoverageState.UNSUPPORTED.value,
                "detail": spec.source_detail,
            },
            "results": {
                "provider": spec.results_source,
                "state": CoverageState.UNSUPPORTED.value,
                "detail": spec.source_detail,
            },
            "bookmakers": self._bookmakers_smoke(),
            "policy": "canonical stored games/history remain forecastable; no schedule or result is fabricated",
        }

    def pregame(self) -> Dict[str, Any]:
        """Acquire fixtures, coherent predictions, and observed book evidence."""

        self._reset_stage()
        games = self._discover_games()
        games = self._acquire_context(games)
        selections = self._acquire_predictions(games)
        if self.config.acquire_book_lines:
            self._acquire_book_lines(games)
        else:
            # A no-paid-provider window must not suppress the keyless direct
            # fallback.  The workflow's existing flag protects The Odds API
            # credits, while this bounded source can still fill exact named
            # full-game moneyline, total, and spread gaps in its live read-only
            # rehearsal.
            self._acquire_book_lines(games, direct_fallback_only=True)
            self.book_provider_status["policy"] = {
                "enabled": False,
                "status": "paid_sources_deferred_direct_fallback_attempted",
                "detail": (
                    "paid named-book acquisition was deferred outside the ticket window; "
                    "the bounded direct BetExplorer fallback was still attempted for games "
                    "without a fresh named full-game moneyline"
                ),
            }
            self.gaps.append(
                CoverageGap(
                    game_id=None,
                    required_field="paid_named_book_price",
                    reason=CoverageState.NOT_APPLICABLE,
                    provider="price_budget",
                    detail="paid provider acquisition deferred outside the dated ticket-acquisition window",
                    observed_utc=self._now_iso(),
                )
            )

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
                "prediction_provider_status": self.prediction_provider_status,
                "book_provider_status": self.book_provider_status,
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
                    provider=self.config.league_enum.value.lower(),
                    detail=str(exc)[:200],
                    observed_utc=self._now_iso(),
                )
            )
            finals = []
        rechecked = self._recheck_pending_finals({game.game_id for game in finals})
        finals += rechecked
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

        written = {row["game_id"] for row in rows}
        return self._report(
            mode="write" if self.allow_writes else "dry_run",
            fixtures_found=len(finals),
            extra={
                "rows": rows,
                "finals_written": len(rows),
                "frozen_predictions_used": used_frozen,
                "quarantined": quarantined,
                # Earlier league-local days that gained late finals this run.
                "rechecked_dates": sorted(
                    {_iso(self._ledger_date_for(g)) for g in rechecked if g.game_id in written}
                ),
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

    def _fiba_schedule_fetch(self, url: str) -> Any:
        """Fetch official FIBA GDAP JSON, preserving fixture injection."""

        if self._fiba_client is not None:
            payload = self._fiba_client.get_json(url)
            self.quota["fiba"] = self._fiba_client.last_quota()
            return payload
        return self.fetch(url)

    def _fiba_detail_fetch(self, url: str) -> Dict[str, Any]:
        """Fetch one official FIBA game page detail, preserving fixture injection."""

        if self._fiba_client is not None:
            detail = self._fiba_client.game_detail(url)
            self.quota["fiba"] = self._fiba_client.last_quota()
            return detail
        return self.fetch(url)

    def _unsupported_source_gap(self, required_field: str, *, provider: str, detail: str) -> None:
        """Surface missing compliant source capability without inventing data."""

        self.gaps.append(
            CoverageGap(
                game_id=None,
                required_field=required_field,
                reason=CoverageState.UNSUPPORTED,
                provider=provider,
                detail=detail,
                observed_utc=self._now_iso(),
            )
        )

    def _stored_slate_games(self) -> List[Game]:
        """Return only canonical games already supplied by a validated import.

        This enables the normal prediction/freeze/forecast lifecycle for a
        newly registered league while its automatic official adapter remains
        unavailable.  It never turns an unavailable source into an empty
        success: callers emit a CoverageGap separately.
        """

        wanted = self.config.date_yyyymmdd
        timezone = ZoneInfo(self.config.timezone)
        return [
            game
            for game in self.store.list_games(league=self.config.league_enum)
            if parse_utc(game.scheduled_utc).astimezone(timezone).strftime("%Y%m%d") == wanted
        ]

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
        elif league is League.ABA_LEAGUE:
            games = aba_liga.discover_games(
                self.config.date_yyyymmdd, text_fetch=self._aba_text_fetch()
            )
            self.gaps.extend(aba_liga.last_gaps())
        elif league is League.CBA:
            # Public league data from the CBA data-server remains the sole
            # discovery source, preventing duplicate canonical games.
            games = cbaleague.discover_games(self.config.date_yyyymmdd, self.fetch, league)
            self.gaps.extend(cbaleague.last_gaps())
        elif league is League.ITALY_LBA:
            games = lba.discover_games(
                _iso(self.config.date_yyyymmdd), self.fetch, self._lba_season_start_year()
            )
            for gap in lba.last_gaps():
                self.gaps.append(
                    CoverageGap(
                        game_id=(f"LBA-{gap['game_id']}" if gap.get("game_id") else None),
                        required_field="schedule_identity",
                        reason=CoverageState.MISSING,
                        provider="lba",
                        detail=str(gap.get("detail", ""))[:200],
                        observed_utc=self._now_iso(),
                    )
                )
        elif league is League.ISRAEL_PREMIER_LEAGUE:
            games = israel_ibl.discover_games(
                _iso(self.config.date_yyyymmdd),
                self.fetch if self._injected_fetch else None,
                self._israel_season_end_year(),
            )
            for gap in israel_ibl.last_gaps():
                self.gaps.append(
                    CoverageGap(
                        game_id=(f"IBL-{gap['game_id']}" if gap.get("game_id") else None),
                        required_field="schedule_identity",
                        reason=CoverageState.MISSING,
                        provider="israel_ibl",
                        detail=str(gap.get("detail", ""))[:200],
                        observed_utc=self._now_iso(),
                    )
                )
        elif league is League.SPAIN_ACB:
            games = acb.discover_games(
                _iso(self.config.date_yyyymmdd), self.fetch, self._acb_season_end_year()
            )
            for gap in acb.last_gaps():
                self.gaps.append(
                    CoverageGap(
                        game_id=(f"ACB-{gap['game_id']}" if gap.get("game_id") else None),
                        required_field="schedule_identity",
                        reason=CoverageState.MISSING,
                        provider="acb",
                        detail=str(gap.get("detail", ""))[:200],
                        observed_utc=self._now_iso(),
                    )
                )
        elif league in (League.EUROLEAGUE, League.EUROCUP):
            games = euroleague.discover_games(
                _iso(self.config.date_yyyymmdd), self.fetch, self._euro_seasoncode()
            )
        elif league is League.EUROLEAGUE_W:
            games = fiba_euroleague_w.discover_games(
                _iso(self.config.date_yyyymmdd), self._fiba_schedule_fetch
            )
            # Future FIBA bracket placeholders are intentionally skipped as
            # non-canonical fixtures, but their bounded coverage telemetry is
            # part of this stage's report rather than a hidden empty-slate.
            self.gaps.extend(fiba_euroleague_w.last_gaps())
        else:
            spec = league_spec(league)
            self._unsupported_source_gap(
                "official_schedule",
                provider=spec.schedule_source,
                detail=(
                    spec.source_detail
                    or "No compliant automated schedule adapter is registered for this league."
                ),
            )
            games = self._stored_slate_games()
        for game in games:
            self.store.upsert_game(game)
        if league in ESPN_LEAGUES:
            scoreboard_snapshots = [snapshot for snapshots in espn_books.values() for snapshot in snapshots]
            pre_match_games = {
                snapshot.game_id for snapshot in scoreboard_snapshots if snapshot.pre_match
            }
            self.book_provider_status["espn"] = {
                "enabled": True,
                "status": (
                    "observed_pre_match_markets"
                    if pre_match_games
                    else "no_pre_match_markets_in_scoreboard"
                ),
                "snapshots_observed": len(scoreboard_snapshots),
                "pre_match_snapshots": sum(
                    snapshot.pre_match for snapshot in scoreboard_snapshots
                ),
                "games_with_pre_match_snapshots": len(pre_match_games),
            }
            if scoreboard_snapshots:
                stamped = self._stamp_book_evidence(scoreboard_snapshots[0])
                self.book_provider_status["espn"]["evidence_policy"] = self._book_policy_status(
                    "espn", stamped
                )
        for game_id, snapshots in espn_books.items():
            for snapshot in snapshots:
                snapshot = self._stamp_book_evidence(snapshot)
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
                    self.store.upsert_book_snapshot(self._stamp_book_evidence(snap))
                    hits += 1
                except Exception:
                    pass
        # Telemetry for monitoring: visible in candidate_funnel via gaps? Use context_report
        self.context_report["odds_catchup_attempts"] = 1
        self.context_report["odds_catchup_hits"] = hits
        self.context_report["odds_catchup_games_without_books"] = len(need)
        if "espn" in self.book_provider_status:
            self.book_provider_status["espn"]["catchup_snapshots_stored"] = hits
            self.book_provider_status["espn"]["catchup_games_requested"] = len(need)

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
        if self.config.league_enum is League.NBA:
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

        # Predixsport publishes a league list, not a canonical-game keyed
        # endpoint. Fetch that list once per relevant league run, then apply a
        # strict ID-or-home/away/UTC-tip join. Calling it once per game both
        # wastes the configured provider allowance and could create different
        # source observations inside the same slate.
        predix_by_game = self._predixsport_candidates_for_games(games)

        for game in games:
            # Stored pre-tip observations compete with fresh observations. A
            # transient provider failure must not blank a still-fresh bundle.
            candidates: List[Prediction] = [
                pred for pred in self.store.get_predictions(game.game_id) if pred.pre_tip
            ]
            league = self.config.league_enum
            if league in ESPN_BPI_LEAGUES:
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

            # NBA must not bypass this independent provider just because ESPN
            # BPI is also available. Provider priority continues to make the
            # validated complete ESPN/Hardwood bundle the selected forecast;
            # a matching Predixsport observation is retained as a coherent
            # shadow for prospective diagnostics.
            candidates.extend(predix_by_game.get(game.game_id, ()))

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

    def _predixsport_candidates_for_games(
        self, games: List[Game]
    ) -> Dict[str, List[Prediction]]:
        """Fetch and strictly join one Predixsport upcoming list for a slate.

        The adapter's source ID is preserved in each returned prediction, but
        it cannot be used as a canonical game ID until it matches an existing
        recorded provider ID or a unique home/away/UTC-tip identity. Ambiguous
        or incomplete source rows become visible coverage gaps; they are never
        guessed into a fixture.
        """

        output: Dict[str, List[Prediction]] = {game.game_id: [] for game in games}
        league = self.config.league_enum
        if not games or league not in {League.NBA, League.EUROLEAGUE}:
            return output

        from .providers import predixsport

        if not predixsport.is_enabled():
            self.prediction_provider_status["predixsport"] = {
                "enabled": False,
                "status": "disabled_missing_credential",
                "records_returned": 0,
                "records_matched": 0,
            }
            return output

        # Quota window: the free tier is 200 SHARED requests per calendar
        # month, so an upstream call is only made when the previous one is
        # older than UPCOMING_WINDOW_HOURS. The window is recorded for every
        # attempt, including dry runs and failures, because the request has
        # already been spent upstream either way. Stored pre-tip observations
        # stay in the candidate pool, so a skipped window loses no evidence.
        now_iso = self._now_iso()
        window = self._predixsport_window()
        next_call = window.get("next_call_utc")
        if next_call:
            try:
                due = parse_utc(next_call) <= parse_utc(now_iso)
            except ValueError:
                due = True  # unreadable telemetry must not wedge the lane
            if not due:
                self.prediction_provider_status["predixsport"] = {
                    "enabled": True,
                    "status": "skipped_quota_window",
                    "records_returned": 0,
                    "records_matched": 0,
                    "window_hours": predixsport.UPCOMING_WINDOW_HOURS,
                    "last_call_utc": window.get("last_call_utc"),
                    "next_call_utc": next_call,
                }
                return output

        source_gaps: List[CoverageGap] = []
        # Response shape (key names and row count only, never values): the
        # documented schema has never been seen live, so "the feed is empty"
        # and "the rows arrived under a key we do not read" both otherwise
        # surface as a bare records_returned: 0.
        response_shape: Dict[str, Any] = {}
        try:
            predictions = predixsport.fetch_upcoming(
                "nba" if league is League.NBA else "euroleague",
                predixsport.UPCOMING_DAYS,
                self.fetch_factory,
                gaps=source_gaps,
                diagnostics=response_shape,
            )
        except Exception as exc:
            # A failed request still consumed an upstream call, so the window
            # is recorded before the error is reported.
            self._record_predixsport_window(now_iso, ok=False)
            self.prediction_provider_status["predixsport"] = {
                "enabled": True,
                "status": "provider_error",
                "records_returned": 0,
                "records_matched": 0,
                "detail": f"{type(exc).__name__}: {str(exc)[:160]}",
            }
            self.gaps.append(
                CoverageGap(
                    game_id=None,
                    required_field="predixsport_upcoming_predictions",
                    reason=CoverageState.PROVIDER_ERROR,
                    provider="predixsport",
                    detail=f"{type(exc).__name__}: {str(exc)[:180]}",
                    observed_utc=self._now_iso(),
                )
            )
            return output

        self.gaps.extend(source_gaps)
        matched = 0
        unmatched = 0
        for prediction in predictions:
            game, reason = _match_predixsport_prediction(prediction, games)
            if game is None:
                unmatched += 1
                self.gaps.append(
                    CoverageGap(
                        game_id=None,
                        required_field="predixsport_prediction_identity",
                        reason=CoverageState.MISSING,
                        provider="predixsport",
                        detail=reason,
                        observed_utc=self._now_iso(),
                    )
                )
                continue
            output[game.game_id].append(replace(prediction, game_id=game.game_id))
            matched += 1

        self._record_predixsport_window(now_iso, ok=True)
        self.prediction_provider_status["predixsport"] = {
            "enabled": True,
            "status": "fetched",
            "records_returned": len(predictions),
            "records_matched": matched,
            "records_unmatched": unmatched,
            "records_rejected": len(source_gaps),
            "response": response_shape,
        }
        return output

    def _betexplorer_direct_fallback_games(self, games: Sequence[Game]) -> List[Game]:
        """Return games missing an exact named full-game core market.

        The direct source can provide its displayed moneyline, Over/Under, and
        Asian Handicap rows.  It is therefore a real fallback for any absent
        core full-game market, not merely a moneyline rescue.  An explicitly
        selected provider other than the direct fallback remains an operator
        filter and is never overridden.
        """

        selected_provider = (self.config.book.provider or "").strip().lower()
        if selected_provider and selected_provider not in {"none", "betexplorer_direct"}:
            return []
        if selected_provider == "betexplorer_direct":
            return list(games)

        def has_core_market(snapshot: BookSnapshot, market: Market) -> bool:
            if (
                snapshot.market is not market
                or snapshot.period is not Period.FULL_GAME
                or snapshot.price_provenance is not PriceProvenance.NAMED_BOOK
                or not self.config.book.matches_source(snapshot)
            ):
                return False
            if market is Market.MONEYLINE:
                return snapshot.home_price_decimal is not None and snapshot.away_price_decimal is not None
            if market is Market.TOTAL:
                return (
                    snapshot.line is not None
                    and snapshot.over_price_decimal is not None
                    and snapshot.under_price_decimal is not None
                )
            return (
                snapshot.line is not None
                and snapshot.home_price_decimal is not None
                and snapshot.away_price_decimal is not None
            )

        candidates: List[Game] = []
        for game in games:
            visible = self._visible_pregame_books(game)
            if not all(
                any(has_core_market(snapshot, market) for snapshot in visible)
                for market in (Market.MONEYLINE, Market.TOTAL, Market.SPREAD)
            ):
                candidates.append(game)
        return candidates

    def _book_provider_options(
        self, provider: str, games: Optional[Sequence[Game]] = None
    ) -> Dict[str, Any]:
        """Pass deterministic, bounded acquisition context to book providers."""
        if provider == "theoddsapi":
            return {"tz_name": self.config.timezone}
        if provider == "betexplorer" and self.config.fixture_now_utc:
            # Recorded fixture-table pages include relative Today/Tomorrow labels.
            # Pin fixture replays to their configured observation time instead
            # of the wall clock so results remain stable across calendar days.
            return {"now_utc": self.config.fixture_now_utc}
        if provider == "betexplorer_direct":
            # The direct fallback receives only canonical scheduled fixtures;
            # it fetches only the distinct Paris-local date boards represented
            # by those games, then comparison rows for strict ordered matches.
            options: Dict[str, Any] = {"now_utc": self.config.fixture_now_utc or self._now_iso()}
            if games is not None:
                options["games"] = list(games)
            return options
        return {}

    def _acquire_book_lines(
        self, games: List[Game], *, direct_fallback_only: bool = False
    ) -> Dict[str, List[BookSnapshot]]:
        """Acquire and reconcile configured book sources with explicit health states.

        Missing optional credentials, an empty provider response, an identity
        rejection, and a successful mapped response are materially different
        causes of "no matched odds".  Preserve that distinction in the
        pregame report rather than making an intentionally-disabled source
        look like an empty slate.
        """

        raw_snapshots: List[BookSnapshot] = []
        providers = (
            ("betexplorer_direct",) if direct_fallback_only else _book_providers(self.config.league)
        )
        if not providers:
            self.book_provider_status["policy"] = {
                "enabled": False,
                "status": "not_configured_for_league",
                "detail": (
                    "no validated named-book source is registered; forecasts remain "
                    "unpriced and ticket-ineligible rather than borrowing another league's market"
                ),
                "price_provenance": "no_price",
                "source_permission": "unresolved",
                "operational_validation": "unavailable",
            }
            self._unsupported_source_gap(
                "named_book_price",
                provider="price_registry",
                detail="No named-book price source is registered; operate as an explicit unpriced forecast lane.",
            )
        for provider in providers:
            module = _book_module(provider)
            if module is None:
                self.book_provider_status[provider] = {
                    "enabled": False,
                    "status": "adapter_unavailable",
                }
                continue
            if not module.is_enabled():
                self.book_provider_status[provider] = {
                    "enabled": False,
                    "status": "disabled",
                    "detail": "optional provider credential is not configured",
                }
                continue
            try:
                provider_games: Sequence[Game] = games
                if provider == "betexplorer_direct":
                    provider_games = self._betexplorer_direct_fallback_games(games)
                    if not provider_games:
                        self.book_provider_status[provider] = {
                            "enabled": True,
                            "status": "fallback_not_needed",
                            "snapshots_fetched": 0,
                            "snapshots_mapped": 0,
                            "identity_rejections": 0,
                            "detail": "a fresh named full-game moneyline already exists for every slate game",
                        }
                        continue
                options = self._book_provider_options(provider, provider_games)
                provider_snapshots = module.fetch_odds(
                    self.config.league_enum.value,
                    self.config.date_yyyymmdd,
                    self.fetch_factory,
                    **options,
                )
                raw_snapshots.extend(provider_snapshots)
                self.quota[provider] = module.last_quota() if hasattr(module, "last_quota") else {}
                provider_gaps = module.last_gaps() if hasattr(module, "last_gaps") else []
                self.gaps.extend(provider_gaps)
                permit_deferred = any(
                    gap.required_field == "provider_credit_permit" for gap in provider_gaps
                )
                self.book_provider_status[provider] = {
                    "enabled": True,
                    "status": (
                        "budget_permit_rejected"
                        if permit_deferred
                        else "fetched" if provider_snapshots else "fetched_no_snapshots"
                    ),
                    "snapshots_fetched": len(provider_snapshots),
                    "snapshots_mapped": 0,
                    "identity_rejections": 0,
                }
                if hasattr(module, "last_receipt"):
                    # The direct BetExplorer fallback publishes its bounded
                    # source URLs/digests and identity census here.  Keep that
                    # receipt with the report rather than hiding a no-price
                    # result behind a generic empty-provider label.
                    self.book_provider_status[provider]["receipt"] = module.last_receipt()
            except Exception as exc:
                logger.warning("%s fetch failed: %s", provider, exc)
                self.book_provider_status[provider] = {
                    "enabled": True,
                    "status": "provider_error",
                    "error_type": type(exc).__name__,
                }
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
                status = self.book_provider_status.get(raw.provider)
                if status is not None:
                    status["identity_rejections"] = int(status["identity_rejections"]) + 1
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
                status = self.book_provider_status.get(raw.provider)
                if status is not None:
                    status["identity_rejections"] = int(status["identity_rejections"]) + 1
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
            snapshot = self._stamp_book_evidence(
                replace(raw, game_id=game.game_id, provider_event_id=provider_event_id)
            )
            per_game = mapped.setdefault(game.game_id, [])
            if snapshot not in per_game:
                per_game.append(snapshot)
                status = self.book_provider_status.get(raw.provider)
                if status is not None:
                    status["snapshots_mapped"] = int(status["snapshots_mapped"]) + 1
            try:
                self.store.upsert_book_snapshot(snapshot)
            except Exception as exc:
                logger.warning("book store failed: %s", exc)

        for status in self.book_provider_status.values():
            if status.get("status") != "fetched":
                continue
            if not status.get("snapshots_mapped"):
                status["status"] = "identity_unmatched"
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
                snapshot.line if snapshot.market in {Market.TOTAL, Market.SPREAD} else None,
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

    def _lba_season_start_year(self) -> int:
        """LBA season FIRST year for this run's target date (2026 == 2026/27)."""
        from datetime import datetime as _dt

        return lba.season_start_year(_dt.strptime(self.config.date_yyyymmdd, "%Y%m%d"))

    def _israel_season_end_year(self) -> int:
        """Season END year for the Israel board route (``cYear``)."""
        from datetime import datetime as _dt

        return israel_ibl.season_end_year(_dt.strptime(self.config.date_yyyymmdd, "%Y%m%d"))

    def _acb_season_end_year(self) -> int:
        """ACB season END year for this run's target date (2027 == 2026-27)."""
        from datetime import datetime as _dt

        target = _dt.strptime(self.config.date_yyyymmdd, "%Y%m%d")
        return acb.season_end_year(target)

    def _euro_seasoncode(self) -> str:
        """Season code for the EuroLeague Basketball lanes (E… / U…).

        The provider identifies the competition by the season-code letter, so a
        EuroCup run must never inherit the EuroLeague code. The season start
        year follows the same July boundary used by the standings probe.
        """
        year = int(self.config.date_yyyymmdd[:4])
        if int(self.config.date_yyyymmdd[4:6]) < 7:
            year -= 1
        competition = (
            euroleague.COMPETITION_EUROCUP
            if self.config.league_enum is League.EUROCUP
            else euroleague.COMPETITION_EUROLEAGUE
        )
        return euroleague.season_code(competition, year)

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
        permanent_skips = 0
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
                if isinstance(exc, HttpError) and exc.status == 404:
                    # A 404 on one indexed event is permanent: the manifest
                    # entry has no summary page, so retaining the cursor would
                    # wedge every later run on the same event. Record the gap
                    # and advance like an already-persisted event.
                    logger.warning(
                        "ESPN deep history permanent 404 for %s: %s", event_id, exc
                    )
                    self.gaps.append(
                        CoverageGap(
                            game_id=event_id,
                            required_field="deep_history_event",
                            reason=CoverageState.MISSING,
                            provider="espn",
                            detail=f"permanent HTTP 404: {str(exc)[:190]}",
                            observed_utc=self._now_iso(),
                        )
                    )
                    permanent_skips += 1
                    processed += 1
                    event_index += 1
                    self._store_history_cursor(
                        season, event_index, season, len(event_ids)
                    )
                    continue
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
        if permanent_skips and status == "acquired":
            status = "acquired_with_permanent_skips"
        return {
            **base,
            "status": status,
            "events_acquired": acquired,
            "final_scores_written": final_scores,
            "already_persisted": already,
            "permanent_skips": permanent_skips,
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

    def _predixsport_window(self) -> Dict[str, Any]:
        """Read the upstream quota window for this league (quota telemetry)."""
        health = self.store.get_provider_health("predixsport", self.config.league_enum)
        return dict(health.coverage or {}) if health else {}

    def _record_predixsport_window(self, attempted_utc: str, ok: bool) -> None:
        """Record that an upstream call was spent, and when the next is due.

        This is quota telemetry, not an acquisition cursor: it is written even
        on a dry run or a failed request, because the shared 200/month
        allowance has already been consumed by the attempt.
        """
        from datetime import timedelta

        from .providers import predixsport

        try:
            attempted = parse_utc(attempted_utc)
        except ValueError:  # pragma: no cover - defensive, observed_utc is ours
            return
        next_call = attempted + timedelta(hours=predixsport.UPCOMING_WINDOW_HOURS)
        previous = self._predixsport_window()
        coverage = {
            "last_call_utc": to_utc_iso(attempted),
            "next_call_utc": to_utc_iso(next_call),
            "window_hours": predixsport.UPCOMING_WINDOW_HOURS,
            "calls_recorded": int(previous.get("calls_recorded") or 0) + 1,
            "last_call_ok": bool(ok),
        }
        try:
            self.store.upsert_provider_health(
                ProviderHealth(
                    provider="predixsport",
                    league=self.config.league_enum,
                    last_attempt_utc=coverage["last_call_utc"],
                    last_success_utc=coverage["last_call_utc"] if ok else None,
                    state=CoverageState.OK if ok else CoverageState.PROVIDER_ERROR,
                    detail=(
                        f"predixsport upstream quota window "
                        f"({predixsport.UPCOMING_WINDOW_HOURS}h); not an acquisition cursor"
                    ),
                    coverage=coverage,
                )
            )
        except Exception as exc:  # pragma: no cover - telemetry must never wedge a run
            logger.warning("predixsport quota window not recorded: %s", exc)

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
        competition = (
            euroleague.COMPETITION_EUROCUP
            if self.config.league_enum is League.EUROCUP
            else euroleague.COMPETITION_EUROLEAGUE
        )
        ledger_dates: set = set()
        while budget > 0 and season > 1900:
            seasoncode = euroleague.season_code(competition, season)
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

    def _backfill_aba_liga_history(self) -> Dict[str, Any]:
        """Backfill ABA finals from the externally referenced route contract.

        The calendar is the complete season index, while each final's detail
        page supplies scoped team IDs and reconciled periods. The cursor
        advances only through games actually written; an invalid detail blocks
        the next game number instead of silently jumping a corrupted result.
        """

        event_budget = max(0, int(self.config.deep_history_event_budget))
        season_budget = max(0, int(self.config.history_request_budget))
        base: Dict[str, Any] = {
            "status": "acquired",
            "event_budget": event_budget,
            "season_request_budget": season_budget,
            "seasons": [],
            "events_acquired": 0,
            "final_scores_written": 0,
            "administrative_disposition_count": 0,
            # A visible calendar placeholder is not a quarantine. Keep the
            # provider-error total separate so a high count cannot be read as
            # a parser/reconciliation failure in an otherwise clean batch.
            "not_applicable_rows": 0,
            "quarantined_rows": 0,
            "ledger_dates": [],
            "excluded_seasons": [],
            # Receipts are durable historical-completeness attestations. A
            # receipt-complete older season is reported when skipped so the
            # cursor never silently burns an event budget replaying it.
            "receipt_complete_seasons_skipped": [],
            "learning_or_calibration_claimed": False,
        }
        if not event_budget or not season_budget:
            return {**base, "status": "disabled"}
        if not (self.allow_writes and getattr(self.store, "persist", False)):
            return {
                **base,
                "status": "disabled_canonical_write_required",
                "detail": "dry runs cannot advance an ABA history cursor",
            }

        cursor = self._history_cursor("aba-liga-history")
        raw_receipts = cursor.get("season_receipts") or {}
        if not isinstance(raw_receipts, dict) or any(
            not isinstance(key, str) or not isinstance(value, dict)
            for key, value in raw_receipts.items()
        ):
            self.gaps.append(
                CoverageGap(
                    game_id=None,
                    required_field="aba_liga_season_receipt",
                    reason=CoverageState.PROVIDER_ERROR,
                    provider=aba_liga.PROVIDER,
                    detail="stored ABA history receipts have an invalid shape; cursor retained",
                    observed_utc=self._now_iso(),
                )
            )
            return {
                **base,
                "status": "provider_error_cursor_retained",
                "cursor": cursor,
            }
        # Coverage receipts are compact count attestations, not raw source
        # pages. They travel with the cursor so a later bounded batch cannot
        # erase evidence that an older season was fully reconciled.
        season_receipts: Dict[str, Dict[str, Any]] = {
            key: dict(value) for key, value in raw_receipts.items()
        }
        target_day = datetime.strptime(self.config.date_yyyymmdd, "%Y%m%d").date()
        current = aba_liga.season_end_year(target_day)
        season = min(int(cursor.get("next_season") or current), current)
        after_game_no = int(cursor.get("last_game_no") or 0)
        raw_administrative_count = cursor.get("administrative_disposition_count", 0)
        if (
            isinstance(raw_administrative_count, bool)
            or not isinstance(raw_administrative_count, int)
            or raw_administrative_count < 0
        ):
            self.gaps.append(
                CoverageGap(
                    game_id=None,
                    required_field="aba_liga_administrative_disposition",
                    reason=CoverageState.PROVIDER_ERROR,
                    provider=aba_liga.PROVIDER,
                    detail="stored administrative-disposition count is invalid; cursor retained",
                    observed_utc=self._now_iso(),
                )
            )
            return {
                **base,
                "status": "provider_error_cursor_retained",
                "cursor": cursor,
            }
        administrative_disposition_count = raw_administrative_count
        last_completed_season = cursor.get("last_completed_season")
        last_completed_final_count = cursor.get("last_completed_final_count")
        remaining_events = event_budget
        remaining_seasons = season_budget
        ledger_dates: set = set()
        base["season_receipts"] = []

        def cursor_coverage(next_season: int, last_game_no: int) -> Dict[str, Any]:
            coverage: Dict[str, Any] = {
                "next_season": next_season,
                "last_game_no": last_game_no,
                "floor_season": aba_liga.FIRST_SEASON_END_YEAR,
            }
            if administrative_disposition_count:
                coverage["administrative_disposition_count"] = administrative_disposition_count
            if last_completed_season is not None:
                coverage["last_completed_season"] = int(last_completed_season)
                coverage["last_completed_final_count"] = int(last_completed_final_count or 0)
            if season_receipts:
                coverage["season_receipts"] = season_receipts
            return coverage

        def stored_reconciled_final_count(covered_season: int) -> int:
            return sum(
                game.status is GameStatus.FINAL
                and game.actual is not None
                and game.actual.reconciles()[0]
                and int((game.evidence.get(aba_liga.PROVIDER) or {}).get("season_end_year") or 0)
                == covered_season
                for game in self.store.list_games(league=League.ABA_LEAGUE)
            )

        def record_season_receipt(
            covered_season: int,
            calendar_final_count: int,
            reconciled_final_count: int,
            administrative_count: int,
            adapter_gaps: List[CoverageGap],
        ) -> None:
            """Record a count-only receipt after a clean complete calendar scan.

            A calendar terminal row is either a period-reconciled played final
            or a separately verified official administrative disposition.
            Unpinned schedule placeholders are separate visible telemetry and
            never enter either denominator.
            """

            key = str(covered_season)
            receipt = {
                "season_end_year": covered_season,
                "calendar_final_count": calendar_final_count,
                "reconciled_final_count": reconciled_final_count,
                "administrative_disposition_count": administrative_count,
                "not_applicable_rows": sum(
                    gap.reason is CoverageState.NOT_APPLICABLE
                    and gap.required_field != "official_administrative_disposition"
                    for gap in adapter_gaps
                ),
                "blocking_gap_count": 0,
                "status": "calendar_terminal_count_reconciled",
                "source": aba_liga.PROVIDER,
                "evidence": aba_liga.EVIDENCE,
            }
            previous = season_receipts.get(key)
            if previous is not None and all(
                previous.get(field) == value for field, value in receipt.items()
            ):
                receipt["completed_at_utc"] = previous.get("completed_at_utc")
            else:
                receipt["completed_at_utc"] = self._now_iso()
            season_receipts[key] = receipt
            base["season_receipts"].append(dict(receipt))

        def completed_receipt(covered_season: int) -> Optional[Dict[str, Any]]:
            """Return a trustworthy immutable receipt for a prior season.

            Cursor coverage is operational state and can outlive code/schema
            changes, so merely finding a dictionary under ``season_receipts``
            is not enough to skip source acquisition. Only the exact
            zero-blocking-gap attestation emitted above is eligible. Anything
            malformed or stale falls through to the normal fail-closed replay.
            """

            receipt = season_receipts.get(str(covered_season))
            required_fields = {
                "season_end_year",
                "calendar_final_count",
                "reconciled_final_count",
                "administrative_disposition_count",
                "not_applicable_rows",
                "blocking_gap_count",
                "status",
                "source",
                "evidence",
                "completed_at_utc",
            }
            if not isinstance(receipt, dict) or set(receipt) != required_fields:
                return None
            counts = (
                receipt.get("calendar_final_count"),
                receipt.get("reconciled_final_count"),
                receipt.get("administrative_disposition_count"),
                receipt.get("not_applicable_rows"),
                receipt.get("blocking_gap_count"),
            )
            if (
                receipt.get("season_end_year") != covered_season
                or any(
                    isinstance(value, bool) or not isinstance(value, int) or value < 0
                    for value in counts
                )
                or receipt["calendar_final_count"] <= 0
                or receipt["calendar_final_count"]
                != receipt["reconciled_final_count"] + receipt["administrative_disposition_count"]
                or receipt["blocking_gap_count"] != 0
                or receipt.get("status") != "calendar_terminal_count_reconciled"
                or receipt.get("source") != aba_liga.PROVIDER
                or receipt.get("evidence") != aba_liga.EVIDENCE
                or not isinstance(receipt.get("completed_at_utc"), str)
                or not receipt["completed_at_utc"]
            ):
                return None
            return dict(receipt)

        # Once the current season was fully scanned, the monotonic history
        # cursor moves to older seasons. Re-read only its lightweight calendar
        # index before doing so: a late-posted final below the old cursor must
        # trigger a bounded replay, never become a silent permanent hole.
        if int(last_completed_season or 0) == current:
            try:
                indexed_count = aba_liga.calendar_final_count(
                    current, text_fetch=self._aba_text_fetch()
                )
                self.gaps.extend(aba_liga.last_gaps())
            except Exception as exc:
                self.gaps.append(
                    CoverageGap(
                        game_id=None,
                        required_field="aba_liga_season_completeness",
                        reason=CoverageState.PROVIDER_ERROR,
                        provider=aba_liga.PROVIDER,
                        detail=f"{current}: {str(exc)[:180]}",
                        observed_utc=self._now_iso(),
                    )
                )
                return {
                    **base,
                    "status": "provider_error_cursor_retained",
                    "cursor": cursor,
                }
            stored_count = stored_reconciled_final_count(current)
            current_receipt = season_receipts.get(str(current), {})
            stored_administrative_count = current_receipt.get(
                "administrative_disposition_count", 0
            )
            if (
                isinstance(stored_administrative_count, bool)
                or not isinstance(stored_administrative_count, int)
                or stored_administrative_count < 0
            ):
                self.gaps.append(
                    CoverageGap(
                        game_id=None,
                        required_field="aba_liga_season_completeness",
                        reason=CoverageState.PROVIDER_ERROR,
                        provider=aba_liga.PROVIDER,
                        detail=f"season {current}: stored administrative-disposition count is invalid",
                        observed_utc=self._now_iso(),
                    )
                )
                return {
                    **base,
                    "status": "provider_error_cursor_retained",
                    "cursor": cursor,
                }
            stored_terminal_count = stored_count + stored_administrative_count
            if indexed_count != stored_terminal_count:
                self.gaps.append(
                    CoverageGap(
                        game_id=None,
                        required_field="aba_liga_season_completeness",
                        reason=CoverageState.MISSING,
                        provider=aba_liga.PROVIDER,
                        detail=(
                            f"season {current}: calendar has {indexed_count} terminal row(s), "
                            f"store has {stored_count} reconciled final(s) and "
                            f"{stored_administrative_count} administrative disposition(s); "
                            "replaying from game 0"
                        ),
                        observed_utc=self._now_iso(),
                    )
                )
                base["completeness_repair"] = {
                    "season_end_year": current,
                    "calendar_final_count": indexed_count,
                    "stored_reconciled_final_count": stored_count,
                }
                if stored_administrative_count:
                    base["completeness_repair"]["stored_administrative_disposition_count"] = (
                        stored_administrative_count
                    )
                season, after_game_no = current, 0
                administrative_disposition_count = 0
                # Do not repeatedly reset a bounded repair to zero. Subsequent
                # successful batches retain the normal current-season cursor
                # until it reaches a fresh complete receipt below.
                last_completed_season = None
                last_completed_final_count = None
        while (
            remaining_events > 0
            and remaining_seasons > 0
            and season >= aba_liga.FIRST_SEASON_END_YEAR
        ):
            # Historical receipts are immutable count attestations. Do not
            # spend a later batch's detail-event budget replaying a season
            # that was already fully reconciled, even if an older batch left
            # a stale partial game cursor below it. The active season remains
            # deliberately exempt: the calendar recheck above must continue
            # to detect finals posted late below its cursor.
            receipt = completed_receipt(season) if season < current else None
            if receipt is not None:
                stored_count = stored_reconciled_final_count(season)
                expected_terminal_count = receipt["calendar_final_count"]
                if (
                    stored_count + receipt["administrative_disposition_count"]
                    == expected_terminal_count
                ):
                    skipped = {
                        "season_end_year": season,
                        "status": "receipt_complete_skipped",
                        "calendar_final_count": expected_terminal_count,
                        "reconciled_final_count": receipt["reconciled_final_count"],
                        "administrative_disposition_count": receipt[
                            "administrative_disposition_count"
                        ],
                        "not_applicable_rows": receipt["not_applicable_rows"],
                    }
                    base["receipt_complete_seasons_skipped"].append(skipped)
                    base["seasons"].append(skipped)
                    season -= 1
                    after_game_no = 0
                    administrative_disposition_count = 0
                    # This is no source request, so do not consume the season
                    # request budget. Use the available batch to continue at
                    # the next incomplete historical season instead.
                    self._record_history_cursor(
                        "aba-liga-history", cursor_coverage(season, after_game_no)
                    )
                    continue
            exclusion_detail = aba_liga.HISTORY_EXCLUDED_SEASONS.get(season)
            if exclusion_detail is not None:
                exclusion = {
                    "season_end_year": season,
                    "status": "excluded_by_operator_policy",
                    "detail": exclusion_detail,
                }
                self.gaps.append(
                    CoverageGap(
                        game_id=None,
                        required_field="aba_liga_history_policy",
                        reason=CoverageState.NOT_APPLICABLE,
                        provider=aba_liga.PROVIDER,
                        detail=f"season {season}: {exclusion_detail}",
                        observed_utc=self._now_iso(),
                    )
                )
                base["excluded_seasons"].append(exclusion)
                base["seasons"].append(exclusion)
                season -= 1
                after_game_no = 0
                administrative_disposition_count = 0
                remaining_seasons -= 1
                self._record_history_cursor(
                    "aba-liga-history", cursor_coverage(season, after_game_no)
                )
                continue
            try:
                games, attempted, exhausted, calendar_final_count = aba_liga.season_finals(
                    season,
                    start_after_game_no=after_game_no,
                    limit=remaining_events,
                    text_fetch=self._aba_text_fetch(),
                )
            except Exception as exc:
                self.gaps.append(
                    CoverageGap(
                        game_id=None,
                        required_field="aba_liga_season_calendar",
                        reason=CoverageState.PROVIDER_ERROR,
                        provider=aba_liga.PROVIDER,
                        detail=f"{season}: {str(exc)[:180]}",
                        observed_utc=self._now_iso(),
                    )
                )
                base["status"] = "provider_error_cursor_retained"
                break
            adapter_gaps = aba_liga.last_gaps()
            administrative_dispositions = aba_liga.last_administrative_dispositions()
            administrative_disposition_count += len(administrative_dispositions)
            base["administrative_disposition_count"] += len(administrative_dispositions)
            self.gaps.extend(adapter_gaps)
            # ``NOT_APPLICABLE`` makes an intentional, visible statement
            # about a placeholder or an administrative result; it is never a
            # quarantined provider failure. Administrative dispositions have
            # their own total above, so keep them out of generic
            # placeholder telemetry as well.
            blocking_gaps = [
                gap for gap in adapter_gaps if gap.reason is not CoverageState.NOT_APPLICABLE
            ]
            not_applicable_rows = sum(
                gap.reason is CoverageState.NOT_APPLICABLE
                and gap.required_field != "official_administrative_disposition"
                for gap in adapter_gaps
            )
            for game in games:
                self.store.upsert_game(game)
                base["events_acquired"] += 1
                base["final_scores_written"] += 1
                remaining_events -= 1
                ledger_dates.add(self._ledger_date_for(game))
            base["seasons"].append(
                {
                    "season_end_year": season,
                    "attempted_game_numbers": attempted,
                    "calendar_final_count": calendar_final_count,
                    "finals_written": len(games),
                    "administrative_disposition_count": len(administrative_dispositions),
                    "not_applicable_rows": not_applicable_rows,
                    "exhausted": exhausted,
                    "quarantined_rows": len(blocking_gaps),
                }
            )
            base["not_applicable_rows"] += not_applicable_rows
            base["quarantined_rows"] += len(blocking_gaps)
            if blocking_gaps:
                # The adapter stops on the first bad detail. Retain the last
                # fully resolved terminal row (played or administrative) so a
                # malformed row remains the next retry, never a silent hole.
                resolved_game_no = aba_liga.last_history_progress_game_no()
                if resolved_game_no > after_game_no:
                    after_game_no = resolved_game_no
                    self._record_history_cursor(
                        "aba-liga-history", cursor_coverage(season, after_game_no)
                    )
                base["status"] = (
                    "partial_provider_error"
                    if games or administrative_dispositions
                    else "provider_error_cursor_retained"
                )
                break
            if exhausted:
                if calendar_final_count == 0:
                    # An all-placeholder index has no canonical result
                    # evidence. It is not a historical-completeness receipt,
                    # even though an empty store would trivially satisfy a
                    # zero-to-zero count comparison.
                    self.gaps.append(
                        CoverageGap(
                            game_id=None,
                            required_field="aba_liga_season_receipt",
                            reason=CoverageState.MISSING,
                            provider=aba_liga.PROVIDER,
                            detail=(
                                f"season {season}: calendar has no final rows; "
                                "receipt withheld and replay retained at game 0"
                            ),
                            observed_utc=self._now_iso(),
                        )
                    )
                    self._record_history_cursor("aba-liga-history", cursor_coverage(season, 0))
                    base["status"] = "provider_error_cursor_retained"
                    break
                reconciled_final_count = stored_reconciled_final_count(season)
                terminal_count = reconciled_final_count + administrative_disposition_count
                if terminal_count != calendar_final_count:
                    # A complete index traversal with an unclassified terminal
                    # row is not a completed season. Rewind it rather than
                    # hiding a persistence hole or a missing adjudication.
                    self.gaps.append(
                        CoverageGap(
                            game_id=None,
                            required_field="aba_liga_season_receipt",
                            reason=CoverageState.MISSING,
                            provider=aba_liga.PROVIDER,
                            detail=(
                                f"season {season}: calendar has {calendar_final_count} terminal row(s), "
                                f"store has {reconciled_final_count} reconciled final(s) and "
                                f"{administrative_disposition_count} administrative disposition(s); "
                                "receipt withheld and replay retained at game 0"
                            ),
                            observed_utc=self._now_iso(),
                        )
                    )
                    self._record_history_cursor("aba-liga-history", cursor_coverage(season, 0))
                    base["status"] = "provider_error_cursor_retained"
                    break
                record_season_receipt(
                    season,
                    calendar_final_count,
                    reconciled_final_count,
                    administrative_disposition_count,
                    adapter_gaps,
                )
                if season == current:
                    last_completed_season = season
                    last_completed_final_count = calendar_final_count
                season -= 1
                after_game_no = 0
                administrative_disposition_count = 0
                remaining_seasons -= 1
                self._record_history_cursor(
                    "aba-liga-history", cursor_coverage(season, after_game_no)
                )
                continue
            # A bounded successful batch resumes after the last resolved
            # terminal row, including an administrative disposition that never
            # becomes a canonical Game.
            resolved_game_no = aba_liga.last_history_progress_game_no()
            if resolved_game_no > after_game_no:
                after_game_no = resolved_game_no
                self._record_history_cursor(
                    "aba-liga-history", cursor_coverage(season, after_game_no)
                )
            break
        base["cursor"] = self._history_cursor("aba-liga-history")
        base["ledger_dates"] = sorted(ledger_dates)
        if season < aba_liga.FIRST_SEASON_END_YEAR and base["status"] == "acquired":
            base["status"] = (
                "complete_to_policy_floor_with_exclusions"
                if aba_liga.HISTORY_EXCLUDED_SEASONS
                else "complete_to_contract_floor"
            )
        return base

    def _backfill_fiba_euroleague_w_history(self) -> Dict[str, Any]:
        """Ingest bounded, reconciled FIBA ELW finals from the current GDAP edition.

        GDAP's competition schedule is one official request; individual game
        pages are then bounded by ``deep_history_event_budget`` because period
        scores must come from usable official game detail, never inferred from
        a final score alone.
        """

        budget = max(0, int(self.config.deep_history_event_budget))
        base: Dict[str, Any] = {
            "event_budget": budget,
            "events_acquired": 0,
            "final_scores_written": 0,
            "ledger_dates": [],
            "learning_or_calibration_claimed": False,
        }
        if not budget:
            return {**base, "status": "disabled"}
        if not (self.allow_writes and getattr(self.store, "persist", False)):
            return {
                **base,
                "status": "disabled_canonical_write_required",
                "detail": "dry runs do not claim a durable FIBA history backfill",
            }
        try:
            games = fiba_euroleague_w.discover_games(None, self._fiba_schedule_fetch)
        except Exception as exc:
            self.gaps.append(
                CoverageGap(
                    game_id=None,
                    required_field="fiba_elw_schedule",
                    reason=CoverageState.PROVIDER_ERROR,
                    provider="fiba",
                    detail=str(exc)[:200],
                    observed_utc=self._now_iso(),
                )
            )
            return {**base, "status": "provider_error"}
        self.gaps.extend(fiba_euroleague_w.last_gaps())
        now = parse_utc(self._now_iso())
        existing = {
            game.game_id: game
            for game in self.store.list_games(league=League.EUROLEAGUE_W)
        }

        def schedule_score_changed(stored: Game, fresh: Game) -> bool:
            """Whether GDAP now disagrees with stored schedule-index provenance."""

            fresh_scores = (
                fresh.provider_ids.get("fiba_elw_schedule_team_a_score"),
                fresh.provider_ids.get("fiba_elw_schedule_team_b_score"),
            )
            stored_scores = (
                stored.provider_ids.get("fiba_elw_schedule_team_a_score"),
                stored.provider_ids.get("fiba_elw_schedule_team_b_score"),
            )
            return fresh_scores != (None, None) and fresh_scores != stored_scores

        missing_finals: List[Game] = []
        corrected_finals: List[Game] = []
        for game in games:
            if game.status is not GameStatus.FINAL or parse_utc(game.scheduled_utc) >= now:
                continue
            stored = existing.get(game.game_id)
            if stored is None or stored.actual is None or not stored.actual.reconciles()[0]:
                missing_finals.append(game)
            elif schedule_score_changed(stored, game):
                # A current GDAP score correction invalidates the old detail
                # provenance. Re-fetch and reconcile it before retaining the
                # canonical final; do not silently count it as already done.
                corrected_finals.append(game)
        candidates = {game.game_id: game for game in [*missing_finals, *corrected_finals]}
        # A durable circular cursor prevents one permanently malformed game
        # detail from consuming the first slot forever once a season grows past
        # the per-run budget. It still wraps, so transient detail failures are
        # retried on a later write-enabled run.
        ordered_candidates = sorted(candidates.values(), key=lambda game: (game.scheduled_utc, game.game_id))
        cursor = self._history_cursor("fiba-elw-history")
        after = (str(cursor.get("scheduled_utc") or ""), str(cursor.get("game_id") or ""))
        remaining = [
            game
            for game in ordered_candidates
            if (game.scheduled_utc, game.game_id) > after
        ]
        final_rows = (remaining or ordered_candidates)[:budget]
        candidate_ids = set(candidates)
        base["already_reconciled"] = sum(
            game.status is GameStatus.FINAL
            and parse_utc(game.scheduled_utc) < now
            and game.game_id not in candidate_ids
            for game in games
        )
        base["score_corrections_detected"] = len(corrected_finals)
        base["score_corrections_written"] = 0
        failures = 0
        written_rows: List[Game] = []
        corrected_ids = {game.game_id for game in corrected_finals}
        for game in final_rows:
            try:
                result = fiba_euroleague_w.fetch_game_result(game, self._fiba_detail_fetch)
            except Exception as exc:
                failures += 1
                self.gaps.append(
                    CoverageGap(
                        game_id=game.game_id,
                        required_field="historical_fiba_game_detail",
                        reason=CoverageState.PROVIDER_ERROR,
                        provider="fiba",
                        detail=str(exc)[:200],
                        observed_utc=self._now_iso(),
                    )
                )
                continue
            self.store.upsert_game(result)
            written_rows.append(game)
            base["events_acquired"] += 1
            base["final_scores_written"] += 1
            base["score_corrections_written"] += int(game.game_id in corrected_ids)
            base["ledger_dates"].append(self._ledger_date_for(result))
        # Cursor advancement means a canonical detail was actually written.
        # In particular an all-failed batch keeps the old cursor and retries
        # on the next write-enabled run instead of declaring false progress.
        if written_rows:
            last = written_rows[-1]
            self._record_history_cursor(
                "fiba-elw-history",
                {"scheduled_utc": last.scheduled_utc, "game_id": last.game_id},
            )
        status = (
            "acquired"
            if not failures
            else "provider_error_cursor_retained"
            if not written_rows
            else "partial_provider_error"
        )
        return {
            **base,
            "status": status,
            "quarantined_rows": failures,
            "cursor": self._history_cursor("fiba-elw-history"),
            "ledger_dates": sorted(set(base["ledger_dates"])),
        }

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
        if self.config.league_enum not in ESPN_LEAGUES | {
            League.ABA_LEAGUE,
            League.EUROLEAGUE,
            League.EUROCUP,
            League.EUROLEAGUE_W,
            League.CBA,
        }:
            spec = league_spec(self.config.league_enum)
            self._unsupported_source_gap(
                "official_history_results",
                provider=spec.results_source,
                detail=(
                    spec.source_detail
                    or "No compliant automated historical-result adapter is registered for this league."
                ),
            )
            return self._report(
                mode="write" if self.allow_writes else "dry_run",
                fixtures_found=0,
                extra={
                    "finals_written": 0,
                    "backfill": {
                        "status": "source_unavailable",
                        "requested_days": max(0, int(self.config.history_days)),
                        "request_budget": max(0, int(self.config.history_request_budget)),
                        "dates": [],
                        "ledger_dates": [],
                        "deep_history": {"status": "source_unavailable", "ledger_dates": []},
                        "learning_or_calibration_claimed": False,
                    },
                },
            )
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
                    # Only fully-final scoreboards are cached: a late game that
                    # was live at fetch time must be re-read on the next run.
                    return self.acquisition.get_json(
                        url,
                        use_cache=True,
                        ttl_seconds=86400,
                        cacheable=espn.scoreboard_is_settled,
                    )

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

        if self.config.league_enum in ESPN_DEEP_HISTORY_LEAGUES:
            deep = self._backfill_deep_history()
        elif espn_enabled:
            # A date-bounded scoreboard/result route is enough for the bounded
            # recent-result history pass above. It is not evidence that this
            # league's core event index has been validated for durable deep
            # history, so do not advance an unreviewed cursor.
            deep = {
                "status": "not_validated_for_league",
                "ledger_dates": [],
                "learning_or_calibration_claimed": False,
            }
        elif self.config.league_enum is League.ABA_LEAGUE:
            deep = self._backfill_aba_liga_history()
        elif self.config.league_enum in (League.EUROLEAGUE, League.EUROCUP):
            deep = self._backfill_euroleague_history()
        elif self.config.league_enum is League.EUROLEAGUE_W:
            deep = self._backfill_fiba_euroleague_w_history()
        elif self.config.league_enum is League.CBA:
            deep = self._backfill_cba_history()
        else:
            deep = {"status": "unsupported_league", "ledger_dates": []}
        ledger_dates.update(deep.get("ledger_dates") or [])
        status = (
            "disabled"
            if not days
            and deep["status"]
            in {"disabled", "unsupported_league", "not_validated_for_league"}
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
        if league is League.ABA_LEAGUE:
            finals = aba_liga.fetch_recent_results(
                self.config.date_yyyymmdd, text_fetch=self._aba_text_fetch()
            )
            self.gaps.extend(aba_liga.last_gaps())
            return finals
        if league is League.CBA:
            finals = cbaleague.fetch_recent_results(
                self.config.date_yyyymmdd, self.fetch, league
            )
            self.gaps.extend(cbaleague.last_gaps())
            return finals
        if league is League.ITALY_LBA:
            return self._lba_finals(self._discover_games())
        if league is League.ISRAEL_PREMIER_LEAGUE:
            return self._israel_ibl_finals(self._discover_games())
        if league is League.SPAIN_ACB:
            return self._acb_finals(self._discover_games())
        if league in (League.EUROLEAGUE, League.EUROCUP):
            return self._euroleague_finals(self._discover_games())
        if league is League.EUROLEAGUE_W:
            return self._fiba_euroleague_w_finals(self._discover_games())
        spec = league_spec(league)
        self._unsupported_source_gap(
            "official_results",
            provider=spec.results_source,
            detail=(
                spec.source_detail
                or "No compliant automated official-result adapter is registered for this league."
            ),
        )
        return []

    def _lba_finals(self, games: List[Game]) -> List[Game]:
        """Refresh LBA finals from per-match play-by-play (periods authoritative).

        Play-by-play carries no team names, so identity stays as discovered from
        the calendar; periods derive from cumulative period ends and must
        reconcile, or the game stays a recorded gap.
        """
        finals: List[Game] = []
        for game in games:
            if game.status is not GameStatus.FINAL:
                continue
            try:
                refreshed = lba.fetch_game_result(game, self.fetch)
            except Exception as exc:
                logger.warning("lba result failed for %s: %s", game.game_id, exc)
                self.gaps.append(
                    CoverageGap(
                        game_id=game.game_id,
                        required_field="actual_result",
                        reason=CoverageState.PROVIDER_ERROR,
                        provider="lba",
                        detail=str(exc)[:200],
                        observed_utc=self._now_iso(),
                    )
                )
                continue
            if refreshed.actual is None:
                continue
            ok, reason = refreshed.actual.reconciles()
            if not ok:
                self.gaps.append(
                    CoverageGap(
                        game_id=refreshed.game_id,
                        required_field="historical_reconciliation",
                        reason=CoverageState.MISSING,
                        provider="lba",
                        detail=reason[:200],
                        observed_utc=self._now_iso(),
                    )
                )
                continue
            finals.append(refreshed)
        return finals

    def _israel_ibl_finals(self, games: List[Game]) -> List[Game]:
        """Refresh Israel finals from the board, keeping failures as gaps."""
        finals: List[Game] = []
        fetch = self.fetch if self._injected_fetch else None
        for game in games:
            if game.status is not GameStatus.FINAL:
                continue
            try:
                refreshed = israel_ibl.fetch_game_result(game, fetch)
            except Exception as exc:
                logger.warning("israel_ibl result failed for %s: %s", game.game_id, exc)
                self.gaps.append(
                    CoverageGap(
                        game_id=game.game_id,
                        required_field="actual_result",
                        reason=CoverageState.PROVIDER_ERROR,
                        provider="israel_ibl",
                        detail=str(exc)[:200],
                        observed_utc=self._now_iso(),
                    )
                )
                continue
            if refreshed.actual is None:
                continue
            ok, reason = refreshed.actual.reconciles()
            if not ok:
                self.gaps.append(
                    CoverageGap(
                        game_id=refreshed.game_id,
                        required_field="historical_reconciliation",
                        reason=CoverageState.MISSING,
                        provider="israel_ibl",
                        detail=reason[:200],
                        observed_utc=self._now_iso(),
                    )
                )
                continue
            finals.append(refreshed)
        return finals

    def _acb_finals(self, games: List[Game]) -> List[Game]:
        """Re-fetch ACB finals from the per-game header (periods authoritative).

        Discovery already resolves identity and periods from the header; the
        results stage refreshes only games the schedule marks finished so a
        late correction is never lost, and a failure stays a gap.
        """
        finals: List[Game] = []
        for game in games:
            if game.status is not GameStatus.FINAL:
                continue
            try:
                refreshed = acb.fetch_game_result(game, self.fetch)
            except Exception as exc:
                logger.warning("acb result failed for %s: %s", game.game_id, exc)
                self.gaps.append(
                    CoverageGap(
                        game_id=game.game_id,
                        required_field="actual_result",
                        reason=CoverageState.PROVIDER_ERROR,
                        provider="acb",
                        detail=str(exc)[:200],
                        observed_utc=self._now_iso(),
                    )
                )
                continue
            if refreshed.actual is None:
                continue
            ok, reason = refreshed.actual.reconciles()
            if not ok:
                self.gaps.append(
                    CoverageGap(
                        game_id=refreshed.game_id,
                        required_field="historical_reconciliation",
                        reason=CoverageState.MISSING,
                        provider="acb",
                        detail=reason[:200],
                        observed_utc=self._now_iso(),
                    )
                )
                continue
            finals.append(refreshed)
        return finals

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

    def _fiba_euroleague_w_finals(self, games: List[Game]) -> List[Game]:
        """Fetch period-complete official FIBA finals; leave unusable detail pending."""

        finals = []
        for game in games:
            if game.status is not GameStatus.FINAL:
                continue
            try:
                finals.append(fiba_euroleague_w.fetch_game_result(game, self._fiba_detail_fetch))
            except Exception as exc:
                logger.warning("FIBA EuroLeague Women result failed for %s: %s", game.game_id, exc)
                self.gaps.append(
                    CoverageGap(
                        game_id=game.game_id,
                        required_field="official_fiba_game_detail",
                        reason=CoverageState.PROVIDER_ERROR,
                        provider="fiba",
                        detail=str(exc)[:200],
                        observed_utc=self._now_iso(),
                    )
                )
        return finals

    def _recheck_pending_finals(self, already: set) -> List[Game]:
        """Finals for earlier-day games that tipped but never got a result.

        Result fetches are keyed by the run's league-local date, so a game
        still live (or not yet officially closed) at the last run of its day
        would otherwise stay pending forever. Every league re-checks only the
        earlier days in the ``history_days`` window that still hold such a
        game; days with nothing pending cost no requests.
        """
        days = max(0, int(self.config.history_days))
        if not days:
            return []
        target = datetime.strptime(self.config.date_yyyymmdd, "%Y%m%d").date()
        window_start_utc = f"{target - timedelta(days=days + 1)}T00:00:00Z"

        def is_pending(game: Game) -> bool:
            return (
                game.game_id not in already
                and game.status not in _CLOSED_STATUSES
                and (game.status is not GameStatus.FINAL or game.actual is None)
                and self._ledger_date_for(game) < self.config.date_yyyymmdd
            )

        pending = {
            game.game_id: game
            for game in self.store.list_games(
                league=self.config.league_enum,
                on_or_after_utc=window_start_utc,
                before_utc=self._now_iso(),
            )
            if is_pending(game)
        }
        # Do not silently make an old pending final disappear merely because
        # it aged out of the bounded network recheck window. It is deliberately
        # not retried forever here (that would defeat the budget); surface one
        # explicit result-unresolvable marker for operator remediation.
        expired = [
            game
            for game in self.store.list_games(
                league=self.config.league_enum,
                before_utc=window_start_utc,
            )
            if is_pending(game)
        ]
        if expired:
            oldest = min(expired, key=lambda game: game.scheduled_utc)
            self.gaps.append(
                CoverageGap(
                    game_id=None,
                    required_field="pending_final_expired",
                    reason=CoverageState.MISSING,
                    provider=self.config.league_enum.value.lower(),
                    detail=(
                        f"result_unresolvable: {len(expired)} pending final(s) aged beyond "
                        f"the {days}-day recheck window; oldest={oldest.game_id}"
                    ),
                    observed_utc=self._now_iso(),
                )
            )
        if not pending:
            return []
        league = self.config.league_enum
        if league not in ESPN_LEAGUES | {
            League.ABA_LEAGUE,
            League.EUROLEAGUE,
            League.EUROLEAGUE_W,
            League.CBA,
        }:
            spec = league_spec(league)
            self._unsupported_source_gap(
                "pending_results_recheck",
                provider=spec.results_source,
                detail=(
                    spec.source_detail
                    or "No compliant result adapter is registered for pending-final rechecks."
                ),
            )
            return []
        if league in {League.EUROLEAGUE, League.EUROLEAGUE_W}:
            # Each official European source exposes a competition-wide schedule;
            # a single request covers all pending FIBA ELW games as well.
            batches = [None]
        else:
            batches = sorted({self._ledger_date_for(g) for g in pending.values()})
        finals: List[Game] = []
        for date in batches:
            try:
                if league is League.EUROLEAGUE:
                    day = self._euroleague_finals(
                        [g for g in euroleague.discover_games(None, self.fetch) if g.game_id in pending]
                    )
                elif league is League.EUROLEAGUE_W:
                    season_games = fiba_euroleague_w.discover_games(
                        None, self._fiba_schedule_fetch
                    )
                    self.gaps.extend(fiba_euroleague_w.last_gaps())
                    day = self._fiba_euroleague_w_finals(
                        [g for g in season_games if g.game_id in pending]
                    )
                elif league in ESPN_LEAGUES:
                    day = espn.fetch_recent_results(
                        date, self.espn_fetch, league, observed_utc=self._now_iso()
                    )
                elif league is League.ABA_LEAGUE:
                    day = aba_liga.fetch_recent_results(
                        date, text_fetch=self._aba_text_fetch()
                    )
                    self.gaps.extend(aba_liga.last_gaps())
                else:
                    day = cbaleague.fetch_recent_results(date, self.fetch, league)
                    self.gaps.extend(cbaleague.last_gaps())
            except Exception as exc:  # one bad day never blocks the others
                logger.warning("pending result re-check failed for %s: %s", date, exc)
                self.gaps.append(
                    CoverageGap(
                        game_id=None,
                        required_field="pending_results_recheck",
                        reason=CoverageState.PROVIDER_ERROR,
                        provider=league.value.lower(),
                        detail=f"{date or 'season'}: {str(exc)[:180]}",
                        observed_utc=self._now_iso(),
                    )
                )
                continue
            for fetched in day:
                stored = pending.get(fetched.game_id)
                if stored is None:
                    continue
                if team_names_agree(stored.home_name, fetched.away_name) and (
                    team_names_agree(stored.away_name, fetched.home_name)
                ):
                    # A neutral-site or re-designated fixture can come back with
                    # home and away exchanged. The frozen ticket is settled by
                    # team identity, so this is not fatal — but the operator has
                    # to see that the provider moved the sides.
                    self.gaps.append(
                        CoverageGap(
                            game_id=fetched.game_id,
                            required_field="result_side_identity",
                            reason=CoverageState.MISSING,
                            provider=league.value.lower(),
                            detail=(
                                "result reports home/away the other way round versus the "
                                "frozen fixture; frozen legs settle on their own team identity"
                            ),
                            observed_utc=self._now_iso(),
                        )
                    )
                finals.append(fetched)
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
            return predixsport.smoke(self.config.league_enum.value.lower(), self.fetch_factory)
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
                    self.config.league_enum.value, self.fetch_factory, **options
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
            league=self.config.league_enum.value,
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
    return HttpClient().get_json


def _book_module(provider: str):
    from .providers import theoddsapi

    return {
        "betexplorer": betexplorer,
        "betexplorer_direct": betexplorer_direct,
        "theoddsapi": theoddsapi,
    }.get(provider)


def _hardwood_candidates(game: Game, artifact, fetch, gaps: List[CoverageGap]) -> List[Prediction]:
    output: List[Prediction] = []
    local_gaps: List[CoverageGap] = []
    prediction = hardwood.fetch_prediction(game, fetch, gaps=local_gaps)
    if prediction is not None:
        output.append(prediction)
    gaps.extend(local_gaps)
    return output


def _match_predixsport_prediction(
    prediction: Prediction, games: Sequence[Game]
) -> tuple[Optional[Game], str]:
    """Resolve one provider prediction to exactly one canonical fixture.

    A known Predixsport event ID is sufficient because it was previously
    bound to the canonical game map. Otherwise require both ordered teams and
    a UTC tip within the existing strict cross-provider tolerance. The helper
    deliberately has no fuzzy-name, date-only, or array-position fallback.
    """

    source = prediction.auxiliary_metrics or {}
    source_match_id = str(source.get("source_match_id") or prediction.game_id).strip()
    direct = [
        game
        for game in games
        if source_match_id
        and str(game.provider_ids.get("predixsport") or "").strip() == source_match_id
    ]
    if len(direct) == 1:
        return direct[0], ""
    if len(direct) > 1:
        return None, f"Predixsport ID {source_match_id!r} maps to multiple canonical games"

    home = source.get("source_home_name")
    away = source.get("source_away_name")
    tip = source.get("source_tipoff_utc")
    if not all(isinstance(value, str) and value.strip() for value in (home, away, tip)):
        return (
            None,
            f"Predixsport prediction {source_match_id!r} lacks provider ID mapping and home/away/UTC-tip identity",
        )

    matches: List[Game] = []
    for game in games:
        try:
            same_teams = (
                canonical_team_key(str(home)) == canonical_team_key(game.home_name)
                and canonical_team_key(str(away)) == canonical_team_key(game.away_name)
            )
            if same_teams and times_within(str(tip), game.scheduled_utc):
                matches.append(game)
        except (TypeError, ValueError):
            return (
                None,
                f"Predixsport prediction {source_match_id!r} has an invalid UTC tip {tip!r}",
            )

    if len(matches) == 1:
        return matches[0], ""
    if not matches:
        return (
            None,
            f"Predixsport prediction {source_match_id!r} has no exact home/away/UTC-tip match",
        )
    return (
        None,
        f"Predixsport prediction {source_match_id!r} matches multiple canonical games",
    )


def _iso(date_yyyymmdd: str) -> str:
    return f"{date_yyyymmdd[:4]}-{date_yyyymmdd[4:6]}-{date_yyyymmdd[6:8]}"


def _count_by(values: Sequence[str]) -> Dict[str, int]:
    output: Dict[str, int] = {}
    for value in values:
        output[value] = output.get(value, 0) + 1
    return output
