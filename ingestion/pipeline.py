"""Headless acquisition orchestration: smoke, pregame, freeze, results, coverage.

Provider adapters return canonical records. This module coordinates those
records and a Store; it has no spreadsheet, web-application, or external
read-model dependency. Every command is dry-run unless the caller explicitly
enables canonical writes.
"""

from __future__ import annotations

from datetime import timedelta
import logging
from dataclasses import dataclass, field, replace
from typing import Any, Callable, Dict, List, Optional, Sequence

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
    Game,
    GameStatus,
    League,
    Period,
    Prediction,
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
from .providers import espn, euroleague, hardwood, isportsapi, sportradar

logger = logging.getLogger(__name__)

DEFAULT_TZ = {
    "NBA": "America/New_York",
    "EUROLEAGUE": "Europe/Paris",
    "CBA": "Asia/Shanghai",
    "NCAA_M": "America/New_York",
    "NCAA_W": "America/New_York",
}
# iSportsAPI schedule dates are documented as GMT+0 calendar windows; the
# other providers use their competition-local slate date.
ACQUISITION_TZ = {**DEFAULT_TZ, "CBA": "UTC"}
BOOK_PROVIDERS = ("therundown", "sportsgameodds", "theoddsapi")
BOOK_PROVIDERS_BY_LEAGUE = {
    "CBA": ("isportsapi",),
    "NCAA_M": ("theoddsapi",),
    "NCAA_W": ("theoddsapi",),
}


def _book_providers(league: str) -> Sequence[str]:
    # NBA/EuroLeague retain the existing overrideable provider tuple.
    return BOOK_PROVIDERS_BY_LEAGUE.get(league.upper(), BOOK_PROVIDERS)
BOOK_STALENESS_HOURS = 24.0


@dataclass
class PipelineConfig:
    league: str
    date_yyyymmdd: str
    tz_name: Optional[str] = None
    book: BookConfig = field(default_factory=BookConfig)
    freshness_hours: float = 12.0
    provider_order: Sequence[str] = DEFAULT_PROVIDER_ORDER
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
        self.fetch_factory = fetch_factory or _default_fetch_factory
        self.allow_writes = allow_writes
        self.gaps: List[CoverageGap] = []
        self.quota: Dict[str, Any] = {}

    def _now_iso(self) -> str:
        if self.config.fixture_now_utc:
            return to_utc_iso(parse_utc(self.config.fixture_now_utc))
        return now_utc_iso()

    def _reset_stage(self) -> None:
        """Keep gaps and quota scoped to one command, not cumulative duplicates."""

        self.gaps = []
        self.quota = {}

    # -- commands ------------------------------------------------------------

    def smoke(self) -> Dict[str, Any]:
        """Report observed provider coverage without writing canonical state."""

        self._reset_stage()
        league = self.config.league.upper()
        if league in {"NBA", "NCAA_M", "NCAA_W"}:
            report = espn.smoke(
                self.config.date_yyyymmdd, self.fetch, self.config.league_enum
            )
            if league == "NBA":
                report["hardwood"] = self._hardwood_smoke()
            report["bookmakers"] = self._bookmakers_smoke()
            return report
        if league == "CBA":
            report = isportsapi.schedule_smoke(
                self.config.date_yyyymmdd, self.fetch, self.config.league_enum
            )
            try:
                report["sportradar"] = sportradar.smoke(
                    self.config.date_yyyymmdd, self.fetch
                )
            except Exception as exc:
                report["sportradar"] = {
                    "status": "PROVIDER_ERROR", "detail": str(exc)[:200]
                }
            report["bookmakers"] = self._bookmakers_smoke()
            return report
        if league == "EUROLEAGUE":
            report = euroleague.smoke(self.fetch)
            report["predixsport"] = self._predixsport_smoke()
            report["bookmakers"] = self._bookmakers_smoke()
            return report
        raise ValueError(f"unsupported league for smoke: {league}")

    def pregame(self) -> Dict[str, Any]:
        """Acquire fixtures, coherent predictions, and observed book evidence."""

        self._reset_stage()
        games = self._discover_games()
        selections = self._acquire_predictions(games)
        self._acquire_book_lines(games)

        rows = []
        for game in games:
            selection = selections[game.game_id]
            pred = selection.selected
            snapshots = self._visible_pregame_books(game)
            quarter_lines, full_game_line, book_evidence, book_gaps = (
                self._book_fields(game, snapshots)
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
                        pred
                        and self.store.is_prediction_frozen(
                            game.game_id, pred.provider
                        )
                    ),
                    "odds": odds,
                    "odds_evidence": odds_evidence,
                    "full_game_line": full_game_line,
                    "quarters": (
                        quarter_lines
                        if game.league not in {League.NCAA_M, League.NCAA_W}
                        else {}
                    ),
                    "halves": (
                        quarter_lines
                        if game.league in {League.NCAA_M, League.NCAA_W}
                        else {}
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
            },
        )

    def freeze(self) -> Dict[str, Any]:
        """Freeze the latest valid pre-tip predictions and book observations."""

        self._reset_stage()
        tip_games = self._games_with_tip_passed()
        frozen: Dict[str, Dict[str, str]] = {}
        books_frozen = 0
        for game in tip_games:
            frozen[game.game_id] = freeze_game_predictions(
                self.store, game, game.scheduled_utc
            )
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
        finals = self._fetch_finals()
        rows = []
        used_frozen = 0
        quarantined = []
        for game in finals:
            if game.actual is None:
                quarantined.append(
                    {"game_id": game.game_id, "reason": "no_official_result"}
                )
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
            counts = period_coverage(
                snapshots, any_book, native_periods(self.config.league_enum)
            )
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
                "quarter_lines": (
                    per_period if native_labels == ["Q1", "Q2", "Q3", "Q4"] else {}
                ),
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
        if league in {League.NBA, League.NCAA_M, League.NCAA_W}:
            games = espn.discover_games(
                self.config.date_yyyymmdd, self.fetch, league
            )
        elif league is League.CBA:
            games = isportsapi.discover_games(
                self.config.date_yyyymmdd, self.fetch, league
            )
        elif league is League.EUROLEAGUE:
            games = euroleague.discover_games(
                _iso(self.config.date_yyyymmdd), self.fetch
            )
        else:
            raise ValueError(f"unsupported league for discovery: {league.value}")
        for game in games:
            self.store.upsert_game(game)
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
                pred
                for pred in self.store.get_predictions(game.game_id)
                if pred.pre_tip
            ]
            league = self.config.league_enum
            if league in {League.NBA, League.NCAA_M, League.NCAA_W}:
                try:
                    candidates.append(
                        espn.fetch_bpi(
                            game.provider_ids.get("espn", game.game_id),
                            self.fetch,
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
                    candidates.extend(
                        _hardwood_candidates(game, artifact, self.fetch, self.gaps)
                    )
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

    def _acquire_book_lines(self, games: List[Game]) -> Dict[str, List[BookSnapshot]]:
        raw_snapshots: List[BookSnapshot] = []
        for provider in _book_providers(self.config.league):
            module = _book_module(provider)
            if module is None or not module.is_enabled():
                continue
            try:
                options = (
                    {"tz_name": self.config.timezone}
                    if provider == "theoddsapi"
                    else {}
                )
                raw_snapshots.extend(
                    module.fetch_odds(
                        self.config.league.upper(),
                        self.config.date_yyyymmdd,
                        self.fetch_factory,
                        **options,
                    )
                )
                self.quota[provider] = (
                    module.last_quota() if hasattr(module, "last_quota") else {}
                )
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
            snapshot = replace(
                raw, game_id=game.game_id, provider_event_id=provider_event_id
            )
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
            chosen = newest_snapshot(
                select_total_snapshots(snapshots, self.config.book, period)
            )
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
                        reason=(
                            CoverageState.MISSING
                            if snapshots
                            else CoverageState.UNSUPPORTED
                        ),
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
        snapshot, price = max(
            candidates, key=lambda item: parse_utc(item[0].observed_utc)
        )
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
            if (
                not snapshot.pre_match
                or parse_utc(snapshot.observed_utc)
                >= parse_utc(game.scheduled_utc)
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

    def _fetch_finals(self) -> List[Game]:
        league = self.config.league_enum
        if league in {League.NBA, League.NCAA_M, League.NCAA_W}:
            return espn.fetch_recent_results(
                self.config.date_yyyymmdd, self.fetch, league
            )
        if league is League.CBA:
            return isportsapi.fetch_recent_results(
                self.config.date_yyyymmdd, self.fetch, league
            )
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
                    logger.warning(
                        "euroleague result failed for %s: %s", game.game_id, exc
                    )
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
                    "LIVE_VERIFIED"
                    if isinstance(fetch_target, HttpClient)
                    else "FIXTURE_VERIFIED"
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
            return predixsport.smoke(
                self.config.league.lower(), self.fetch_factory
            )
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
                options = (
                    {"tz_name": self.config.timezone}
                    if provider == "theoddsapi"
                    else {}
                )
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

    def _report(
        self, mode: str, fixtures_found: int, extra: Dict[str, Any]
    ) -> Dict[str, Any]:
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
        report_dict["persistence"] = {
            "kind": store_kind,
            "write_enabled": bool(getattr(self.store, "persist", False)),
            "unattended_state_configured": store_kind == "duckdb",
            "system_of_record": (
                "duckdb_csv" if store_kind == "duckdb" else "ephemeral"
            ),
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
    if snapshot.home_name and canonical_team_key(
        snapshot.home_name
    ) != canonical_team_key(game.home_name):
        return "home-team metadata contradicts the mapped game"
    if snapshot.away_name and canonical_team_key(
        snapshot.away_name
    ) != canonical_team_key(game.away_name):
        return "away-team metadata contradicts the mapped game"
    if snapshot.scheduled_utc:
        try:
            if not times_within(snapshot.scheduled_utc, game.scheduled_utc):
                return "start-time metadata contradicts the mapped game"
        except (TypeError, ValueError):
            return "invalid start-time metadata"
    return ""


def _resolve_book_game(
    snapshot: BookSnapshot, games: Sequence[Game]
) -> tuple[Optional[Game], str]:
    """Resolve a snapshot to one game without trusting an ID coincidence."""

    raw_id = str(snapshot.provider_event_id or snapshot.game_id)
    known = [
        game
        for game in games
        if str(game.provider_ids.get(snapshot.provider) or "") == raw_id
    ]
    if snapshot.provider_event_id is None:
        known.extend(
            game for game in games if game.game_id == raw_id and game not in known
        )
    if len(known) > 1:
        return None, "provider/canonical ID maps to multiple games"
    if known:
        mismatch = _book_identity_mismatch(snapshot, known[0])
        return (None, mismatch) if mismatch else (known[0], "")

    if not (
        snapshot.home_name and snapshot.away_name and snapshot.scheduled_utc
    ):
        return None, "unmapped provider event lacks home/away/start identity metadata"
    candidates = []
    for game in games:
        if _book_identity_mismatch(snapshot, game):
            continue
        if canonical_team_key(snapshot.home_name) != canonical_team_key(
            game.home_name
        ):
            continue
        if canonical_team_key(snapshot.away_name) != canonical_team_key(
            game.away_name
        ):
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
        "therundown": therundown,
        "sportsgameodds": sportsgameodds,
        "theoddsapi": theoddsapi,
        "isportsapi": isportsapi,
    }[provider]


def _hardwood_candidates(
    game: Game, artifact, fetch, gaps: List[CoverageGap]
) -> List[Prediction]:
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
    return (
        f"{date_yyyymmdd[:4]}-{date_yyyymmdd[4:6]}-{date_yyyymmdd[6:8]}"
    )


def _count_by(values: Sequence[str]) -> Dict[str, int]:
    output: Dict[str, int] = {}
    for value in values:
        output[value] = output.get(value, 0) + 1
    return output
