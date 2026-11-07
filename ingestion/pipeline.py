"""Pipeline orchestration: smoke, pregame, freeze, results, coverage.

Every command is dry-run first. Persistent writes require an explicit
``allow_writes=True`` after the CLI's environment guard. DuckDB is the
canonical store; Google Sheets worksheets are optional legacy exports only.

When a legacy worksheet export is explicitly requested, the old safety rules
still apply: never clear a clean tab, never write outside owned base columns,
and keep a single writer for ``UpcomingClean``/``ResultsClean`` A..R.
"""

from __future__ import annotations

import csv
import logging
from dataclasses import dataclass, field, replace
from typing import Any, Callable, Dict, List, Optional, Sequence

from .books import (
    BookConfig,
    ExistingCellValue,
    MergeDecision,
    merge_book_cell,
    quarter_coverage,
)
from .derive import bundle_rows
from .http import HttpClient
from .identity import canonical_team_key, now_utc_iso, parse_utc, times_within
from .models import (
    BookSnapshot,
    CoverageGap,
    CoverageReport,
    CoverageState,
    Game,
    GameStatus,
    League,
    Prediction,
    ProviderHealth,
)
from .predictions import (
    DEFAULT_PROVIDER_ORDER,
    SelectionResult,
    freeze_game_predictions,
    odds_for_prediction,
    resolve_frozen_bundle,
    select_prediction,
)
from .providers import espn, euroleague, hardwood
from .sheets_contract import (
    RESULTS_TAB,
    UPCOMING_TAB,
    COL,
    SheetContract,
    detect_contract,
    render_result_row,
    render_upcoming_row,
)
from .sheets_store import open_data_store
from .sheets_writer import (
    GspreadWorksheetAdapter,
    UpdatePlan,
    apply_plan,
    plan_update,
)

logger = logging.getLogger(__name__)

DEFAULT_TZ = {"NBA": "America/New_York", "EUROLEAGUE": "Europe/Paris"}

BOOK_PROVIDERS = ("therundown", "sportsgameodds", "theoddsapi")


def _noop_sleep(seconds: float) -> None:
    return None


@dataclass
class PipelineConfig:
    league: str
    date_yyyymmdd: str  # slate date in the league's local calendar
    tz_name: Optional[str] = None  # default per league
    sheet_id: Optional[str] = None  # required only for real sheet access
    book: BookConfig = field(default_factory=lambda: BookConfig(provider="none"))
    freshness_hours: float = 12.0
    h2h_depth: int = 10
    recent_depth: int = 15
    quarter_csv: Optional[str] = None  # manual quarter-line fallback path
    provider_order: Sequence[str] = DEFAULT_PROVIDER_ORDER
    # Deterministic observation clock for committed offline fixtures only.
    fixture_now_utc: Optional[str] = None

    @property
    def league_enum(self) -> League:
        return League[self.league.upper()]

    @property
    def timezone(self) -> str:
        return self.tz_name or DEFAULT_TZ[self.league.upper()]


class Pipeline:
    """One configured acquisition run against one store (+ optional sheet)."""

    def __init__(
        self,
        config: PipelineConfig,
        store,
        fetch: Optional[Callable[[str], Any]] = None,
        fetch_factory: Optional[Callable[[str], Callable[[str], Any]]] = None,
        allow_writes: bool = False,
        upcoming_ws: Optional[Any] = None,
        results_ws: Optional[Any] = None,
        sleep: Callable[[float], None] = _noop_sleep,
    ) -> None:
        self.config = config
        self.store = store
        self.fetch = fetch or HttpClient().get_json
        self.fetch_factory = fetch_factory or _default_fetch_factory
        self.allow_writes = allow_writes
        self.upcoming_ws = upcoming_ws
        self.results_ws = results_ws
        self.sleep = sleep
        self.gaps: List[CoverageGap] = []
        self.quota: Dict[str, Any] = {}

    # -- commands ------------------------------------------------------------

    def smoke(self) -> Dict[str, Any]:
        """Keyless live smoke: report ACTUAL coverage, never assumed."""
        league = self.config.league.upper()
        if league == "NBA":
            report = espn.smoke(self.config.date_yyyymmdd, self.fetch)
            hw = self._hardwood_smoke()
            report["hardwood"] = hw
            report["bookmakers"] = self._bookmakers_smoke()
            return report
        if league == "EUROLEAGUE":
            report = euroleague.smoke(self.fetch)
            report["predixsport"] = self._predixsport_smoke()
            report["bookmakers"] = self._bookmakers_smoke()
            return report
        raise ValueError(f"unsupported league for smoke: {self.config.league}")

    def pregame(self) -> Dict[str, Any]:
        """Acquire fixtures + predictions + book lines; render UpcomingClean."""
        games = self._discover_games()
        selections = self._acquire_predictions(games)
        book_by_game = self._acquire_book_lines(games)
        existing_lookup = _existing_lookup(self, games)

        rows = []
        coverage_rows = []
        for game in games:
            selection = selections[game.game_id]
            pred = selection.selected
            snapshots = book_by_game.get(game.game_id, [])
            odds = self._odds_for(snapshots, pred)
            if odds is None:
                odds = self._stored_odds_for(game, pred)
            quarter_lines, fg_line, merge_gaps = self._merge_book_cells(
                game, snapshots, existing_lookup=existing_lookup
            )
            if odds is None:
                # No fresh moneyline: preserve the existing Odds value (book
                # snapshot provenance or the sheet's manual price) instead of
                # blanking it.
                existing_odds = existing_lookup(game).get("Odds")
                if existing_odds is not None and str(existing_odds.raw or "").strip() != "":
                    try:
                        odds = float(str(existing_odds.raw).strip())
                    except ValueError:
                        odds = None
            self.gaps.extend(merge_gaps)
            rows.append(render_upcoming_row(
                game, pred, odds, quarter_lines, fg_line, self.config.timezone
            ))
            coverage_rows.append({
                "game_id": game.game_id,
                "matchup": f"{game.away_name} @ {game.home_name}",
                "tip_utc": game.scheduled_utc,
                "prediction_provider": selection.provider,
                "prediction_reason": selection.reason,
                "prob_pct": pred.display()["Prob %"] if pred else None,
                "avg": pred.display()["Avg"] if pred else None,
                "odds": odds,
                "full_game_line": fg_line,
                "quarters": {q: quarter_lines.get(q) for q in ("Q1", "Q2", "Q3", "Q4")},
            })

        plan = None
        if self.upcoming_ws is not None:
            plan = self._plan_for(UPCOMING_TAB, self.upcoming_ws, rows, "upcoming")
            if self.allow_writes:
                apply_plan(plan, self.upcoming_ws, allow_writes=True, sleep=self.sleep)

        report = self._report(
            mode="write" if self.allow_writes else "dry_run",
            fixtures_found=len(games),
            plan=plan,
            extra={
                "rows": coverage_rows,
                "predictions_selected": _count_by(
                    [s.provider for s in selections.values() if s.provider]
                ),
            },
        )
        return report

    def freeze(self) -> Dict[str, Any]:
        """Freeze latest valid pre-tip predictions and book snapshots."""
        tip_games = self._games_with_tip_passed()
        frozen: Dict[str, Dict[str, str]] = {}
        books_frozen = 0
        for game in tip_games:
            frozen[game.game_id] = freeze_game_predictions(self.store, game, game.scheduled_utc)
            books_frozen += self._freeze_books(game)
        report = self._report(
            mode="write" if self.allow_writes else "dry_run",
            fixtures_found=len(tip_games),
            plan=None,
            extra={"frozen_predictions": frozen, "book_snapshots_frozen": books_frozen},
        )
        return report

    def results(self) -> Dict[str, Any]:
        """Acquire finals; render ResultsClean from FROZEN pre-tip values."""
        finals = self._fetch_finals()

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
            statuses = freeze_game_predictions(self.store, game, game.scheduled_utc)
            self._freeze_books(game)
            pred, was_frozen = resolve_frozen_bundle(
                self.store, game.game_id, tuple(self.config.provider_order)
            )
            if was_frozen:
                used_frozen += 1
            if pred is None:
                self.gaps.append(CoverageGap(
                    game_id=game.game_id, required_field="pregame_prediction",
                    reason=CoverageState.MISSING,
                    detail="missing_pregame_prediction",
                ))
            snapshots = self.store.get_book_snapshots(game.game_id)
            odds = self._frozen_odds(snapshots, pred, game)
            rows.append(render_result_row(game, pred, odds, self.config.timezone))

        plan = None
        if self.results_ws is not None:
            plan = self._plan_for(RESULTS_TAB, self.results_ws, rows, "results")
            if self.allow_writes:
                apply_plan(plan, self.results_ws, allow_writes=True, sleep=self.sleep)

        report = self._report(
            mode="write" if self.allow_writes else "dry_run",
            fixtures_found=len(finals),
            plan=plan,
            extra={
                "finals_written": len(rows),
                "frozen_predictions_used": used_frozen,
                "quarantined": quarantined,
            },
        )
        return report

    def coverage(self) -> Dict[str, Any]:
        """Coverage summary from the persistent store (no network)."""
        from .books import BookConfig, select_total_snapshots

        games = self.store.list_games(league=self.config.league_enum)
        any_book = BookConfig(provider=None)
        with_preds = 0
        with_fg_line = 0
        per_quarter: Dict[str, int] = {q: 0 for q in ("Q1", "Q2", "Q3", "Q4")}
        finals = 0
        for g in games:
            if self.store.get_predictions(g.game_id):
                with_preds += 1
            snaps = self.store.get_book_snapshots(g.game_id)
            if select_total_snapshots(snaps, any_book):
                with_fg_line += 1
            qc = quarter_coverage(snaps, any_book)
            for q in per_quarter:
                per_quarter[q] += qc.get(q, 0)
            if g.status is GameStatus.FINAL and g.actual is not None:
                finals += 1
        report = self._report(
            mode="store", fixtures_found=len(games), plan=None,
            extra={
                "games_in_store": len(games),
                "games_with_prediction": with_preds,
                "games_with_full_game_line": with_fg_line,
                "quarter_lines": per_quarter,
                "finals_reconciled": finals,
                "gaps": [g.__dict__ for g in self.gaps],
            },
        )
        return report

    # -- acquisition helpers ---------------------------------------------------

    def _discover_games(self) -> List[Game]:
        if self.config.league.upper() == "NBA":
            games = espn.discover_games(self.config.date_yyyymmdd, self.fetch)
        else:
            games = euroleague.discover_games(_iso(self.config.date_yyyymmdd), self.fetch)
        for g in games:
            self.store.upsert_game(g)
        return games

    def _acquire_predictions(self, games: List[Game]) -> Dict[str, SelectionResult]:
        selections: Dict[str, SelectionResult] = {}
        now = self.config.fixture_now_utc or now_utc_iso()
        if self.config.league.upper() == "NBA":
            try:
                artifact = hardwood.load_forecasts(self.fetch)
            except Exception as e:
                artifact = None
                logger.warning("hardwood artifact unavailable: %s", e)
                self.gaps.append(CoverageGap(
                    game_id=None, required_field="hardwood_artifact",
                    reason=CoverageState.PROVIDER_ERROR, provider="hardwood",
                    detail=str(e)[:200],
                ))
        else:
            artifact = None

        for game in games:
            # A transient provider failure must not blank a still-fresh,
            # previously stored pre-tip forecast. New observations compete
            # with these candidates and win only under the selection policy.
            candidates: List[Prediction] = [
                p for p in self.store.get_predictions(game.game_id) if p.pre_tip
            ]
            if self.config.league.upper() == "NBA":
                try:
                    candidates.append(espn.fetch_bpi(
                        game.provider_ids.get("espn", game.game_id),
                        self.fetch,
                        game,
                        now_utc=self.config.fixture_now_utc,
                    ))
                except Exception as e:
                    self.gaps.append(CoverageGap(
                        game_id=game.game_id, required_field="espn_bpi",
                        reason=CoverageState.PROVIDER_ERROR, provider="espn",
                        detail=str(e)[:200],
                    ))
                if artifact is not None:
                    candidates.extend(_hardwood_candidates(game, artifact, self.fetch, self.gaps))
            else:
                # EuroLeague official facts only; predictions come from
                # Predixsport behind its feature flag (Milestone 4).
                candidates.extend(_predixsport_candidates(game, self.fetch_factory))

            selection = select_prediction(
                candidates, game, now,
                freshness_hours=self.config.freshness_hours,
                provider_order=tuple(self.config.provider_order),
            )
            # Stamp the pre-tip flag at storage time: an observation first
            # seen after tip is never a pre-tip prediction and must not
            # displace the stored pre-tip bundle (freezing depends on it).
            tip = parse_utc(game.scheduled_utc)
            for pred in [selection.selected, *selection.shadows]:
                if pred is not None:
                    pred.pre_tip = parse_utc(pred.observed_utc) <= tip
            for shadow in selection.shadows:
                try:
                    self.store.upsert_prediction(shadow, frozen=False)
                except Exception as e:
                    logger.warning("shadow store failed: %s", e)
            if selection.selected is not None:
                try:
                    self.store.upsert_prediction(selection.selected, frozen=False)
                except Exception as e:
                    logger.warning("prediction store failed: %s", e)
            selections[game.game_id] = selection
            self.gaps.extend(selection.gaps)
        return selections

    def _acquire_book_lines(self, games: List[Game]) -> Dict[str, List[BookSnapshot]]:
        raw_snapshots: List[BookSnapshot] = []
        for provider in BOOK_PROVIDERS:
            mod = _book_module(provider)
            if mod is None or not mod.is_enabled():
                continue
            try:
                options = (
                    {"tz_name": self.config.timezone}
                    if provider == "theoddsapi" else {}
                )
                raw_snapshots.extend(mod.fetch_odds(
                    self.config.league.upper(), self.config.date_yyyymmdd,
                    self.fetch_factory, **options,
                ))
                self.quota[provider] = (
                    mod.last_quota() if hasattr(mod, "last_quota") else {}
                )
            except Exception as e:
                logger.warning("%s fetch failed: %s", provider, e)
                self.gaps.append(CoverageGap(
                    game_id=None, required_field="book_lines",
                    reason=CoverageState.PROVIDER_ERROR, provider=provider,
                    detail=str(e)[:200],
                ))
        # Manual CSV fallback for quarter lines (never anonymous data).
        if self.config.quarter_csv:
            raw_snapshots.extend(_read_quarter_csv(self.config.quarter_csv))

        # Resolve every provider event before storage. A provider ID already
        # in the game map is authoritative only when any supplied team/time
        # metadata agrees; otherwise require exact canonical teams plus a
        # 30-minute start-time match. Never overwrite one provider's snapshots
        # with another's merely because their IDs differ.
        mapped: Dict[str, List[BookSnapshot]] = {}
        for raw in raw_snapshots:
            game, reason = _resolve_book_game(raw, games)
            if game is None:
                self.gaps.append(CoverageGap(
                    game_id=str(raw.game_id), required_field="book_identity",
                    reason=CoverageState.MISSING, provider=raw.provider,
                    detail=reason,
                ))
                continue
            provider_event_id = str(raw.provider_event_id or raw.game_id)
            if raw.provider != "manual_csv" and not game.provider_ids.get(raw.provider):
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
            except Exception as e:
                logger.warning("book store failed: %s", e)
        return mapped

    def _odds_for(
        self, snapshots: List[BookSnapshot], pred: Optional[Prediction]
    ) -> Optional[float]:
        from .books import select_moneyline_snapshots

        candidates = select_moneyline_snapshots(snapshots, self.config.book)
        if not candidates:
            return None
        newest = max(candidates, key=lambda s: s.observed_utc)
        return odds_for_prediction(newest, pred)

    def _stored_odds_for(
        self, game: Game, pred: Optional[Prediction]
    ) -> Optional[float]:
        """Fresh stored pre-tip price for the predicted side, never its opposite."""
        if pred is None:
            return None
        from datetime import timedelta

        now = parse_utc(now_utc_iso())
        tip = parse_utc(game.scheduled_utc)
        cutoff = min(now, tip)
        usable = []
        for snapshot in self.store.get_book_snapshots(game.game_id):
            try:
                observed = parse_utc(snapshot.observed_utc)
            except (TypeError, ValueError):
                continue
            if observed > cutoff or now - observed > timedelta(hours=24):
                continue
            usable.append(snapshot)
        return self._odds_for(usable, pred)

    def _merge_book_cells(
        self, game: Game, snapshots: List[BookSnapshot], existing_lookup
    ):
        from .models import Market, Period

        quarter_lines: Dict[str, Optional[float]] = {}
        gaps: List[CoverageGap] = []
        existing = existing_lookup(game)
        for q, period in (("Q1", Period.Q1), ("Q2", Period.Q2),
                          ("Q3", Period.Q3), ("Q4", Period.Q4)):
            decision = merge_book_cell(
                period,
                snapshots,
                existing.get(q),
                self.config.book,
                now_utc=now_utc_iso(),
                game_id=game.game_id,
            )
            quarter_lines[q] = decision.value
            if decision.gap:
                gaps.append(decision.gap)
        fg = merge_book_cell(
            Period.FULL_GAME, snapshots, existing.get("FT Score"),
            self.config.book, now_utc=now_utc_iso(), game_id=game.game_id,
        )
        if fg.gap:
            gaps.append(fg.gap)
        return quarter_lines, fg.value, gaps

    def _frozen_odds(
        self, snapshots: List[BookSnapshot], pred: Optional[Prediction], game: Game
    ) -> Optional[float]:
        """Pre-tip moneyline for the predicted side (frozen when available)."""
        from .books import select_moneyline_snapshots

        candidates = select_moneyline_snapshots(snapshots, self.config.book)
        tip = parse_utc(game.scheduled_utc)
        usable = [s for s in candidates if parse_utc(s.observed_utc) <= tip]
        if not usable:
            return None
        newest = max(usable, key=lambda s: s.observed_utc)
        return odds_for_prediction(newest, pred)

    def _freeze_books(self, game: Game) -> int:
        from .store import FrozenWriteError

        frozen = 0
        by_key: Dict[tuple, BookSnapshot] = {}
        for s in self.store.get_book_snapshots(game.game_id):
            if not s.pre_match or parse_utc(s.observed_utc) > parse_utc(game.scheduled_utc):
                continue
            key = (s.provider, s.book, s.market.value, s.period.value)
            prev = by_key.get(key)
            if prev is None or s.observed_utc > prev.observed_utc:
                by_key[key] = s
        for key, snap in sorted(by_key.items(), key=lambda kv: str(kv[0])):
            if self.store.is_book_snapshot_frozen(game.game_id, snap):
                continue
            try:
                self.store.upsert_book_snapshot(snap, frozen=True)
                frozen += 1
            except FrozenWriteError:
                pass
        return frozen

    def _games_with_tip_passed(self) -> List[Game]:
        now = now_utc_iso()
        return [
            g for g in self.store.list_games(league=self.config.league_enum)
            if g.scheduled_utc <= now
        ]

    def _fetch_finals(self) -> List[Game]:
        """Official finals for the slate (seam for offline tests)."""
        if self.config.league.upper() == "NBA":
            return espn.fetch_recent_results(self.config.date_yyyymmdd, self.fetch)
        return self._euroleague_finals(self._discover_games())

    def _euroleague_finals(self, games: List[Game]) -> List[Game]:
        out = []
        for g in games:
            if g.status is GameStatus.FINAL:
                try:
                    out.append(euroleague.fetch_game_result(g, self.fetch))
                except Exception as e:
                    logger.warning("euroleague result failed for %s: %s", g.game_id, e)
                    self.gaps.append(CoverageGap(
                        game_id=g.game_id, required_field="actual_result",
                        reason=CoverageState.PROVIDER_ERROR, provider="euroleague",
                        detail=str(e)[:200],
                    ))
        return out

    def _plan_for(self, tab: str, worksheet, rows: List[Dict[int, Any]], kind: str) -> UpdatePlan:
        contract = detect_contract(
            worksheet.get_all_values()[0], kind
        )
        plan = plan_update(tab, worksheet, contract, rows, kind)
        self._last_plan = plan  # object form for programmatic apply/tests
        return plan

    # -- smoke helpers -----------------------------------------------------

    def _hardwood_smoke(self) -> Dict[str, Any]:
        try:
            artifact = hardwood.load_forecasts(self.fetch)
            state = hardwood.freshness_state(artifact)
            fetch_target = getattr(self.fetch, "__self__", self.fetch)
            return {
                "verification": (
                    "LIVE_VERIFIED" if isinstance(fetch_target, HttpClient)
                    else "FIXTURE_VERIFIED"
                ),
                "generated_at": artifact.generated_at,
                "ingestion_status": artifact.ingestion_status,
                "freshness_state": state.value,
                "games_in_artifact": len(artifact.games_by_id),
            }
        except Exception as e:
            return {"verification": "PROVIDER_ERROR", "detail": str(e)[:200]}

    def _predixsport_smoke(self) -> Dict[str, Any]:
        from .providers import predixsport

        if not predixsport.is_enabled():
            return {"verification": "UNAUTHORIZED", "enabled": False,
                    "note": "PREDIXSPORT_API_KEY missing; provider disabled"}
        try:
            return predixsport.smoke(self.config.league.lower(), self.fetch_factory)
        except Exception as e:
            return {"verification": "PROVIDER_ERROR", "detail": str(e)[:200]}

    def _bookmakers_smoke(self) -> Dict[str, Any]:
        out = {}
        for provider in BOOK_PROVIDERS:
            mod = _book_module(provider)
            if mod is None:
                continue
            if not mod.is_enabled():
                out[provider] = {"enabled": False, "verification": "UNAUTHORIZED",
                                 "note": "API key missing; disabled (keyless pipeline unaffected)"}
                continue
            try:
                options = (
                    {"tz_name": self.config.timezone}
                    if provider == "theoddsapi" else {}
                )
                out[provider] = mod.smoke(
                    self.config.league.upper(), self.fetch_factory, **options
                )
            except Exception as e:
                out[provider] = {"enabled": True, "verification": "PROVIDER_ERROR",
                                 "detail": str(e)[:200]}
        return out

    # -- report ------------------------------------------------------------

    def _report(self, mode: str, fixtures_found: int, plan: Optional[UpdatePlan],
                extra: Dict[str, Any]) -> Dict[str, Any]:
        report = CoverageReport(
            league=self.config.league.upper(),
            generated_utc=now_utc_iso(),
            mode=mode,
            fixtures_found=fixtures_found,
            finals_written=extra.get("finals_written", 0),
            frozen_predictions_used=extra.get("frozen_predictions_used", 0),
            predictions_selected=extra.get("predictions_selected", {}),
            gaps=[g.__dict__.copy() for g in self.gaps],
            quota=self.quota,
        )
        report_dict = report.to_dict()
        store_kind = {
            "DuckDBStore": "duckdb",
            "SheetsDataStore": "legacy_google_sheets",
            "JsonFileStore": "json_file_local",
            "InMemoryStore": "in_memory_ephemeral",
        }.get(type(self.store).__name__, type(self.store).__name__)
        report_dict["persistence"] = {
            "kind": store_kind,
            "write_enabled": bool(getattr(self.store, "persist", False)),
            "unattended_state_configured": store_kind == "duckdb",
            "system_of_record": "duckdb_csv" if store_kind == "duckdb" else "legacy_or_development",
        }
        if plan is not None:
            report_dict["sheet_plan"] = plan.to_dict()
        report_dict.update({k: v for k, v in extra.items() if k not in report_dict})
        return report_dict


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------


def _book_identity_mismatch(snapshot: BookSnapshot, game: Game) -> str:
    """Return a contradiction, or ``''`` when supplied metadata agrees."""
    if snapshot.home_name and canonical_team_key(snapshot.home_name) != canonical_team_key(game.home_name):
        return "home-team metadata contradicts the mapped game"
    if snapshot.away_name and canonical_team_key(snapshot.away_name) != canonical_team_key(game.away_name):
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
    """Resolve one raw book snapshot to a unique canonical game.

    Known provider IDs and explicitly canonical manual/legacy IDs are
    accepted, subject to any supplied identity metadata. Otherwise all three
    of home, away and aware start time are required for a unique join.
    """
    raw_id = str(snapshot.provider_event_id or snapshot.game_id)
    known = [
        g for g in games
        if str(g.provider_ids.get(snapshot.provider) or "") == raw_id
    ]
    canonical_input = snapshot.provider == "manual_csv" or snapshot.provider_event_id is None
    if canonical_input:
        known.extend(g for g in games if g.game_id == raw_id and g not in known)
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
        # _book_identity_mismatch treats absent fields as neutral; this join
        # requires them all and also requires both names, not just the time.
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
    """Auth-aware fetch factory used in production (no secrets logged)."""
    import os

    if provider == "predixsport":
        key = os.environ.get("PREDIXSPORT_API_KEY", "")
        return HttpClient(default_headers={"Authorization": f"Bearer {key}"}).get_json if key else HttpClient().get_json
    if provider == "therundown":
        from .providers import therundown

        # v2: header auth only — the key must never appear in a URL.
        return (
            HttpClient(default_headers=therundown.auth_headers()).get_json
            if therundown.is_enabled()
            else HttpClient().get_json
        )
    if provider == "sportsgameodds":
        key = os.environ.get("SPORTSGAMEODDS_API_KEY", "")
        return HttpClient(default_headers={"x-api-key": key}).get_json if key else HttpClient().get_json
    if provider == "theoddsapi":
        return HttpClient().get_json  # apiKey goes in the URL by provider design
    return HttpClient().get_json


def _book_module(provider: str):
    from .providers import sportsgameodds, theoddsapi, therundown

    return {"therundown": therundown, "sportsgameodds": sportsgameodds,
            "theoddsapi": theoddsapi}[provider]


def _hardwood_candidates(
    game: Game, artifact, fetch, gaps: List[CoverageGap]
) -> List[Prediction]:
    out: List[Prediction] = []
    local_gaps: List[CoverageGap] = []
    pred = hardwood.fetch_prediction(game, fetch, gaps=local_gaps)
    if pred is not None:
        out.append(pred)
    gaps.extend(local_gaps)
    return out


def _predixsport_candidates(game: Game, fetch_factory) -> List[Prediction]:
    from .providers import predixsport

    if not predixsport.is_enabled():
        return []
    try:
        preds = predixsport.fetch_upcoming("nba" if game.league is League.NBA else "euroleague", 7, fetch_factory)
    except Exception as e:
        logger.warning("predixsport upcoming failed: %s", e)
        return []
    for p in preds:
        if p.game_id == game.provider_ids.get("predixsport") or p.game_id == game.game_id:
            return [p]
    return []


def _existing_from_store(store, game: Game):
    """Lookup of existing cell values with provenance, for merge policy."""
    def lookup(game_: Game) -> Dict[str, Optional[ExistingCellValue]]:
        out: Dict[str, Optional[ExistingCellValue]] = {}
        snaps = store.get_book_snapshots(game_.game_id)
        # Moneyline fallback is resolved against the current prediction in
        # _stored_odds_for. Never approximate it with whichever side happens
        # to be present first in a stored snapshot.
        for period in ("Q1", "Q2", "Q3", "Q4", "FT Score"):
            from .models import Market, Period

            wanted_period = {"Q1": Period.Q1, "Q2": Period.Q2, "Q3": Period.Q3,
                             "Q4": Period.Q4, "FT Score": Period.FULL_GAME}[period]
            with_line = [s for s in snaps if s.market is Market.TOTAL
                         and s.period is wanted_period and s.line is not None]
            if with_line:
                newest = max(with_line, key=lambda s: s.observed_utc)
                out[period] = ExistingCellValue(
                    raw=str(newest.line), source=newest.provider,
                    observed_utc=newest.observed_utc, is_manual=False,
                )
            else:
                out[period] = None
        return out
    return lookup


def _existing_lookup(pipeline: "Pipeline", games: List[Game]):
    """Sheet-aware existing-value lookup for the merge policy.

    Provenance order per field: a stored book snapshot (source + timestamp
    known) first; otherwise the CURRENT SHEET cell, treated as a manual
    value (is_manual=True) so the merge policy preserves it — a missing
    refresh never erases a line somebody entered by hand (section 9).
    """
    from .identity import canonical_team_key
    from .sheets_contract import COL, render_local
    from .sheets_writer import _parse_ambiguous

    rows_by_teams: Dict[tuple, List[List[Any]]] = {}
    if pipeline.upcoming_ws is not None:
        values = pipeline.upcoming_ws.get_all_values()
        for row in values[1:]:
            home = canonical_team_key(str(row[COL["Home"] - 1] if len(row) > COL["Home"] - 1 else "" or ""))
            away = canonical_team_key(str(row[COL["Away"] - 1] if len(row) > COL["Away"] - 1 else "" or ""))
            rows_by_teams.setdefault((home, away), []).append(row)

    store_lookup = _existing_from_store(pipeline.store, None)

    def lookup(game: Game) -> Dict[str, Optional[ExistingCellValue]]:
        merged = store_lookup(game)
        if pipeline.upcoming_ws is None:
            return merged
        local_date, _time = render_local(game.scheduled_utc, pipeline.config.timezone)
        candidates = rows_by_teams.get(
            (canonical_team_key(game.home_name), canonical_team_key(game.away_name)), []
        )
        matched = None
        for row in candidates:
            variants = _parse_ambiguous(row[COL["Date"] - 1] if len(row) > COL["Date"] - 1 else "")
            if local_date.isoformat() in variants:
                if matched is not None:
                    return merged  # ambiguous sheet: fall back to store provenance
                matched = row
        if matched is None:
            return merged
        for period in ("Odds", "Q1", "Q2", "Q3", "Q4", "FT Score"):
            if merged.get(period) is not None:
                continue  # store provenance wins
            raw = matched[COL[period] - 1] if len(matched) > COL[period] - 1 else None
            if str(raw or "").strip() != "":
                merged[period] = ExistingCellValue(
                    raw=str(raw), source=None, observed_utc=None, is_manual=True,
                )
        return merged

    return lookup


def _read_quarter_csv(path: str) -> List[BookSnapshot]:
    """Manual quarter-line CSV fallback.

    Required columns: game_id, book, period (Q1..Q4), line, observed_utc.
    Optional: provider (defaults to ``manual_csv``).  Without identity,
    book, period, line and observed timestamp the row is rejected — it must
    never become anonymous stale data (section 9).
    """
    snapshots: List[BookSnapshot] = []
    from .models import Market, Period

    with open(path, newline="") as f:
        for row in csv.DictReader(f):
            try:
                period = Period(str(row["period"]).strip().upper())
                if period not in (Period.Q1, Period.Q2, Period.Q3, Period.Q4):
                    raise ValueError(f"period {row['period']} not a quarter")
                snapshots.append(BookSnapshot(
                    game_id=str(row["game_id"]).strip(),
                    provider=str(row.get("provider") or "manual_csv"),
                    book=str(row["book"]).strip(),
                    market=Market.TOTAL,
                    period=period,
                    observed_utc=str(row["observed_utc"]).strip(),
                    line=float(str(row["line"]).strip()),
                    pre_match=True,
                    main_line=True,
                ))
            except (KeyError, ValueError) as e:
                logger.warning("quarter CSV row rejected: %s (%s)", row, e)
    return snapshots


def _iso(date_yyyymmdd: str) -> str:
    return f"{date_yyyymmdd[:4]}-{date_yyyymmdd[4:6]}-{date_yyyymmdd[6:8]}"


def _count_by(values: Sequence[str]) -> Dict[str, int]:
    out: Dict[str, int] = {}
    for v in values:
        out[v] = out.get(v, 0) + 1
    return out


def open_persistent_store(sheet_id: str, write: bool = False):
    """Open persistent state read-only by default (gspread).

    Only ``write=True`` uses write credentials, creates missing tables and
    persists mutations. Dry runs never create or rewrite store worksheets.
    """
    from auth.google_auth import get_client, get_write_client

    client, err = get_write_client() if write else get_client()
    if err:
        raise RuntimeError(f"Google auth unavailable: {err}")
    ss = client.open_by_key(sheet_id)
    return open_data_store(ss, create=write, persist=write)


def open_satellite_worksheets(sheet_id: str, write: bool):
    """Open UpcomingClean/ResultsClean adapters (auth reused from repo)."""
    from auth.google_auth import get_client, get_write_client

    client, err = (get_write_client() if write else get_client())
    if err:
        raise RuntimeError(f"Google auth unavailable: {err}")
    ss = client.open_by_key(sheet_id)
    upcoming = next((ws for ws in ss.worksheets() if ws.title == UPCOMING_TAB), None)
    results = next((ws for ws in ss.worksheets() if ws.title == RESULTS_TAB), None)
    return (
        GspreadWorksheetAdapter(upcoming) if upcoming else None,
        GspreadWorksheetAdapter(results) if results else None,
    )
