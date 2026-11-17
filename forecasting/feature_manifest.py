"""Auditable source/feature use and point-in-time coverage manifest.

Coverage is descriptive. Candidate evidence is deliberately kept separate from
active model inputs, and raw evidence inventory is not itself a claim of
point-in-time admissibility or predictive value.
"""

from __future__ import annotations

from collections import Counter
from datetime import timedelta
from typing import Any, Dict, Iterable, Mapping

from ingestion.identity import parse_utc
from ingestion.models import Game, League

SCHEMA_VERSION = "magolide.feature_use_manifest.v1"

_ESPN_LEAGUES = {League.NBA, League.WNBA, League.NCAA_M, League.NCAA_W}
_FEATURE_CATALOG = {
    "historical_team_scores": {
        "role": "active_model_input",
        "source": "reconciled prior game results",
        "guard": "scheduled tip and result-availability delay; unreconciled finals excluded",
    },
    "provider_prediction": {
        "role": "active_model_input_when_available",
        "source": "league/provider prediction records",
        "guard": (
            "latest valid pre-tip record per provider observed by report cutoff; "
            "separate forecast/calibration sources; no provider suppressed or averaged"
        ),
    },
    "book_market_price": {
        "role": "active_model_input_when_available",
        "source": "bookmaker snapshots",
        "guard": "pre-match snapshot observed by cutoff and before scheduled tip",
    },
    "historical_player_minutes_and_boxscore": {
        "role": "shadow_candidate_only",
        "source": "ESPN game summaries",
        "guard": "prior reconciled final, summary observed by cutoff, configured result-availability delay",
    },
    "recent_form_and_ats_context": {
        "role": "shadow_candidate_only",
        "source": "timestamped ESPN summary last-five and against-the-spread context",
        "guard": "summary observed by cutoff and before tip; last-five event dates strictly before target tip",
    },
    "current_injury_status": {
        "role": "shadow_candidate_only",
        "source": "timestamped ESPN summary or NBA league injury endpoint",
        "guard": "snapshot observed by cutoff and before tip; absent entry is unknown",
    },
    "espn_bpi_auxiliary_metrics": {
        "role": "retained_diagnostic_candidate_only",
        "source": "finite numeric non-core metrics in ESPN BPI predictor payload",
        "guard": "side identity and core prediction validated; never silently promoted",
    },
    "standings_recent_form_and_provider_detail": {
        "role": "evidence_inventory_only",
        "source": "league/provider evidence resources",
        "guard": "availability alone does not establish a pre-tip snapshot or value",
    },
    "timestamped_book_movement": {
        "role": "offline_descriptive_audit_only",
        "source": "immutable canonical slate ledgers",
        "guard": "same market/book/period; pre-match timestamps only; not CLV or a model input",
    },
}


def _prediction_rows(store, game: Game, cutoff: str) -> list[Any]:
    rows = []
    for prediction in store.get_predictions(game.game_id):
        if not prediction.pre_tip:
            continue
        try:
            if parse_utc(prediction.observed_utc) > parse_utc(cutoff):
                continue
            if parse_utc(prediction.observed_utc) >= parse_utc(game.scheduled_utc):
                continue
        except (TypeError, ValueError):
            continue
        rows.append(prediction)
    return rows


def _book_rows(store, game: Game, cutoff: str) -> list[Any]:
    rows = []
    for snapshot in store.get_book_snapshots(game.game_id):
        if not snapshot.pre_match:
            continue
        try:
            observed = parse_utc(snapshot.observed_utc)
            if observed > parse_utc(cutoff) or observed >= parse_utc(game.scheduled_utc):
                continue
        except (TypeError, ValueError):
            continue
        rows.append(snapshot)
    return rows


def _observed_field_presence(summary: Mapping[str, Any], field: str) -> bool | None:
    presence = summary.get("field_presence")
    if isinstance(presence, Mapping) and isinstance(presence.get(field), bool):
        return presence[field]
    if field not in summary:
        return False
    value = summary.get(field)
    if value not in (None, [], {}):
        return True
    # Older canonical summaries materialized empty values for absent fields.
    return None


def _resource_inventory(games: Iterable[Game]) -> Dict[str, Any]:
    providers: Dict[str, Counter] = {}
    summary_families: Counter = Counter()
    summary_counts: Counter = Counter()
    game_count = 0
    for game in games:
        game_count += 1
        evidence = game.evidence if isinstance(game.evidence, dict) else {}
        espn = evidence.get("espn") if isinstance(evidence.get("espn"), dict) else {}
        summary = espn.get("summary") if isinstance(espn.get("summary"), dict) else {}
        if summary:
            summary_fields = (
                "injuries", "leaders", "last_five_games", "against_the_spread",
                "game_info", "predictor", "team_boxscore", "player_boxscore",
            )
            for field in summary_fields:
                presence = _observed_field_presence(summary, field)
                state = "present" if presence is True else "absent" if presence is False else "unknown"
                summary_families[f"{field}_{state}"] += 1
                if presence is not True:
                    continue
                value = summary.get(field)
                if field != "injuries":
                    if isinstance(value, list):
                        summary_counts[f"{field}_entries"] += len(value)
                    elif isinstance(value, dict):
                        summary_counts[f"{field}_entries"] += len(value)
            injuries = summary.get("injuries")
            if _observed_field_presence(summary, "injuries") is True and isinstance(injuries, list):
                summary_counts["injury_teams"] += len(injuries)
                summary_counts["injury_entries"] += sum(
                    len(team.get("injuries"))
                    for team in injuries
                    if isinstance(team, dict) and isinstance(team.get("injuries"), list)
                )
            coverage = summary.get("context_coverage")
            if isinstance(coverage, dict):
                for field in ("player_boxscore_teams", "player_boxscore_athletes"):
                    value = coverage.get(field)
                    if isinstance(value, (int, float)) and value >= 0:
                        summary_counts[field] += int(value)
                athletes = coverage.get("player_boxscore_athletes")
                if isinstance(athletes, (int, float)) and athletes > 0:
                    summary_families["player_boxscore_nonempty_games"] += 1
        for provider, payload in evidence.items():
            if not isinstance(payload, dict):
                continue
            counts = providers.setdefault(str(provider), Counter())
            for resource, value in payload.items():
                if value is not None and value != {} and value != []:
                    counts[str(resource)] += 1
    return {
        "games_scanned": game_count,
        "provider_resources": {
            provider: {
                resource: {"games_with_resource": count, "coverage_fraction": round(count / game_count, 4) if game_count else None}
                for resource, count in sorted(counts.items())
            }
            for provider, counts in sorted(providers.items())
        },
        "espn_summary_context_families": {
            "games_with_field": {
                field.removesuffix("_present"): count
                for field, count in sorted(summary_families.items())
                if field.endswith("_present") and field != "player_boxscore_nonempty_games"
            },
            "field_presence_counts": {
                field: {
                    state: summary_families[f"{field}_{state}"]
                    for state in ("present", "absent", "unknown")
                }
                for field in (
                    "injuries", "leaders", "last_five_games", "against_the_spread",
                    "game_info", "predictor", "team_boxscore", "player_boxscore",
                )
            },
            "games_with_nonempty_player_boxscore": summary_families[
                "player_boxscore_nonempty_games"
            ],
            "entry_counts": dict(sorted(summary_counts.items())),
        },
        "interpretation": "inventory only; no point-in-time or model-use inference",
    }


def build_feature_use_manifest(
    *,
    league: League,
    games: Iterable[Game],
    upcoming: Iterable[Game],
    store,
    as_of_utc: str,
    candidate_contexts: Mapping[str, Mapping[str, Any]],
    result_delay_hours: float = 4.0,
) -> Dict[str, Any]:
    """Build per-report coverage counts from only visible upcoming records."""
    game_rows = list(games)
    upcoming_rows = list(upcoming)
    prediction_counts: Counter = Counter()
    book_counts: Counter = Counter()
    provider_prediction_games: Counter = Counter()
    provider_book_games: Counter = Counter()
    auxiliary_metric_counts: Counter = Counter()
    for game in upcoming_rows:
        predictions = _prediction_rows(store, game, as_of_utc)
        if predictions:
            prediction_counts["games_with_prediction"] += 1
        for prediction in predictions:
            provider_prediction_games[prediction.provider] += 1
            for field in (
                "prob_home", "prob_away", "expected_home_score",
                "expected_away_score", "expected_total", "expected_margin",
            ):
                if getattr(prediction, field) is not None:
                    prediction_counts[f"populated_{field}"] += 1
            if prediction.auxiliary_metrics:
                prediction_counts["predictions_with_auxiliary_metrics"] += 1
                for value in prediction.auxiliary_metrics.values():
                    if isinstance(value, Mapping):
                        auxiliary_metric_counts.update(
                            str(name) for name in value if value.get(name) is not None
                        )
        books = _book_rows(store, game, as_of_utc)
        if books:
            book_counts["games_with_pre_tip_book_snapshot"] += 1
        for snapshot in books:
            book_counts[f"market:{snapshot.market.value}"] += 1
            book_counts[f"period:{snapshot.period.value}"] += 1
            provider_book_games[snapshot.provider] += 1

    context_counts: Counter = Counter()
    for game in upcoming_rows:
        context = candidate_contexts.get(game.game_id)
        if not isinstance(context, Mapping):
            continue
        context_counts["games_with_candidate_record"] += 1
        if context.get("injury_snapshot_status") in {"snapshot_present", "no_entries_reported"}:
            context_counts["games_with_point_in_time_injury_observation"] += 1
        if context.get("reported_injury_entries"):
            context_counts["games_with_reported_injuries"] += 1
        teams = context.get("teams") if isinstance(context.get("teams"), dict) else {}
        team_values = [team for team in teams.values() if isinstance(team, dict)]
        if any(team.get("historical_boxscore_games", 0) for team in team_values):
            context_counts["games_with_historical_player_boxscore"] += 1
        if any(
            (team.get("recent_form_candidate") or {}).get("recent_results")
            for team in team_values
        ):
            context_counts["games_with_recent_form_candidate"] += 1
        if any(
            (team.get("recent_form_candidate") or {}).get("against_the_spread")
            for team in team_values
        ):
            context_counts["games_with_ats_candidate"] += 1

    cutoff = parse_utc(as_of_utc)
    result_boundary = cutoff - timedelta(hours=result_delay_hours)
    reconciled_prior_finals = [
        game
        for game in game_rows
        if game.status.value == "FINAL"
        and game.actual is not None
        and game.actual.reconciles()[0]
        and parse_utc(game.scheduled_utc) < cutoff
    ]
    available_prior_finals = [
        game for game in reconciled_prior_finals
        if parse_utc(game.scheduled_utc) <= result_boundary
    ]

    is_espn_league = league in _ESPN_LEAGUES
    catalog = {key: dict(value) for key, value in _FEATURE_CATALOG.items()}
    for key in (
        "historical_player_minutes_and_boxscore",
        "recent_form_and_ats_context",
        "current_injury_status",
        "espn_bpi_auxiliary_metrics",
    ):
        catalog[key]["league_support"] = (
            "provider_path_available" if is_espn_league else "not_supported_by_current_ESPN_path"
        )
    catalog["standings_recent_form_and_provider_detail"]["standings_support"] = (
        "smoke_coverage_only" if league in {League.NBA, League.EUROLEAGUE} else "not_supported_by_current_adapter"
    )
    catalog["current_injury_status"]["league_support"] = {
        League.NBA: "NBA_league_endpoint_and_summary; raw summary field presence measured",
        League.WNBA: "summary_path_configured_but_injury_field_coverage_unverified",
        League.NCAA_M: "summary_path_configured; injury field presence measured per payload",
        League.NCAA_W: "summary_path_configured; injury field presence measured per payload",
    }.get(league, "not_supported_by_current_ESPN_path")
    catalog["current_injury_status"]["league_endpoint_support"] = (
        "NBA_only" if league is League.NBA else "not_supported"
    )

    return {
        "schema_version": SCHEMA_VERSION,
        "league": league.value,
        "as_of_utc": as_of_utc,
        "catalog": catalog,
        "coverage_scope": {
            "upcoming_games": len(upcoming_rows),
            "candidate_records": len(candidate_contexts),
            "game_records_in_store": len(game_rows),
        },
        "historical_team_result_coverage": {
            "reconciled_prior_finals": len(reconciled_prior_finals),
            "available_before_report_cutoff_after_delay": len(available_prior_finals),
            "result_delay_hours": result_delay_hours,
            "scope_note": "report-cutoff count only; the model applies its history gate separately at each target tip",
        },
        "visible_upcoming_coverage": {
            "predictions": {
                "games_with_any_pre_tip_prediction": prediction_counts["games_with_prediction"],
                "pre_tip_prediction_rows_by_provider": dict(sorted(provider_prediction_games.items())),
                "populated_fields_across_visible_prediction_rows": {
                    key.removeprefix("populated_"): value
                    for key, value in sorted(prediction_counts.items())
                    if key.startswith("populated_")
                },
                "rows_with_auxiliary_metrics": prediction_counts["predictions_with_auxiliary_metrics"],
                "auxiliary_metric_rows_by_name": dict(sorted(auxiliary_metric_counts.items())),
            },
            "bookmakers": {
                "games_with_pre_tip_snapshot": book_counts["games_with_pre_tip_book_snapshot"],
                "snapshot_rows_by_provider": dict(sorted(provider_book_games.items())),
                "snapshot_rows_by_market": {
                    key.removeprefix("market:"): value
                    for key, value in sorted(book_counts.items()) if key.startswith("market:")
                },
                "snapshot_rows_by_period": {
                    key.removeprefix("period:"): value
                    for key, value in sorted(book_counts.items()) if key.startswith("period:")
                },
            },
            "espn_player_and_injury_candidates": {
                "games_with_candidate_record": context_counts["games_with_candidate_record"],
                "games_with_point_in_time_injury_observation": context_counts["games_with_point_in_time_injury_observation"],
                "games_with_reported_injuries": context_counts["games_with_reported_injuries"],
                "games_with_historical_player_boxscore": context_counts["games_with_historical_player_boxscore"],
                "games_with_recent_form_candidate": context_counts["games_with_recent_form_candidate"],
                "games_with_ats_candidate": context_counts["games_with_ats_candidate"],
            },
        },
        "evidence_resource_inventory": _resource_inventory(game_rows),
        "safety": {
            "manifest_changes_predictions": False,
            "manifest_changes_eligibility_or_stakes": False,
            "candidate_features_used_by_probability_model": False,
            "coverage_counts_are_not_predictive_validation": True,
        },
    }
