"""Official FIBA EuroLeague Women acquisition adapter.

EuroLeague Women is administered and published by FIBA, not by Euroleague
Basketball.  This module deliberately has its own FIBA game/team namespaces;
none of its IDs, schedules, or results share the men's ``euroleague`` adapter.

The official GDAP schedule is the canonical fixture/result index.  It includes
``gameDateTimeUTC`` (the only accepted canonical tip time), official team IDs,
status, final score and result/statistic availability flags.  A completed game
is enriched from FIBA's official game page payload, whose play-by-play contains
cumulative scores through every regulation and overtime period.  The latter is
converted to individual period scores only after the final and team orientation
are verified.

FIBA ships a public APIM subscription key in its web application.  The live
client reads the page-shipped key at runtime and retries a 401 once with a fresh
key; tests instead inject deterministic JSON fixtures through the provider's
``fetch`` callables.  The documented public fallback is used only when the
page cannot be read and is never logged as a credential.
"""

from __future__ import annotations

from datetime import datetime
import json
import re
from typing import Any, Callable, Dict, List, Optional
from urllib.parse import parse_qs, urlencode, urlparse
from zoneinfo import ZoneInfo

from ..http import HttpClient, HttpError
from ..identity import parse_utc, to_utc_iso
from ..models import (
    ActualResult,
    CoverageGap,
    CoverageState,
    Game,
    GameStatus,
    League,
    PeriodKind,
)
from . import FetchFn, IdentityError
from .fiba_euroleague_w_contract import (
    COMPETITION_CODE,
    COMPETITION_ID,
    GDAP_BASE_URL,
    KEY_HEADER,
    KEY_PAGE_URL,
    PUBLIC_FALLBACK_SUBSCRIPTION_KEY,
    SCHEDULE_ENDPOINT,
    game_datetime_utc,
    has_complete_team_identity,
    has_unassigned_team_slot,
    page_subscription_key,
    schedule_status,
)

# The observed competition payload uses the edition's ending year (2027), not
# the display season string previously frozen in this adapter.
COMPETITION_SEASON = "2027"
LEAGUE_TIMEZONE = "Europe/Paris"
GAME_PAGE_TEMPLATE = (
    "https://www.fiba.basketball/en/events/{event_slug}/games/"
    "{game_id}-{team_a_code}-{team_b_code}"
)

VerificationError = ValueError
PROVIDER = "fiba"

# The official competition schedule includes future bracket slots before teams
# are drawn. They are valid schedule rows but cannot become canonical games.
# Preserve every skip as telemetry instead of allowing one placeholder to abort
# the entire current/future slate.
_LAST_GAPS: List[CoverageGap] = []


def last_gaps() -> List[CoverageGap]:
    """Return non-fatal schedule rows skipped during the most recent discovery."""

    return list(_LAST_GAPS)


def _reset_gaps() -> None:
    _LAST_GAPS.clear()


def _unassigned_slots_gap(
    total: int, date_relevant: int, date_unknown: int, date_iso: Optional[str]
) -> None:
    """Expose benign, team-less future bracket slots as bounded telemetry."""

    if not total:
        return
    relevance = (
        f"; {date_relevant} fall on requested Europe/Paris slate {date_iso}"
        if date_iso is not None
        else ""
    )
    if date_unknown:
        relevance += f"; {date_unknown} had no parseable timestamp for date placement"
    _LAST_GAPS.append(
        CoverageGap(
            game_id=None,
            required_field="fiba_elw_unassigned_schedule_slots",
            reason=CoverageState.NOT_APPLICABLE,
            provider=PROVIDER,
            detail=(
                f"skipped {total} GDAP future bracket slot(s) without assigned teamA/teamB"
                f"{relevance}; they cannot become canonical fixtures"
            ),
        )
    )


def _unparseable_timestamp_gap(game_ids: List[str], date_iso: Optional[str]) -> None:
    """Record bounded partial-coverage telemetry without wedging a whole slate."""

    if not game_ids:
        return
    examples = ", ".join(game_ids[:5])
    suffix = "; target relevance could not be determined" if date_iso is not None else ""
    _LAST_GAPS.append(
        CoverageGap(
            game_id=None,
            required_field="fiba_elw_schedule_timestamp",
            reason=CoverageState.PROVIDER_ERROR,
            provider=PROVIDER,
            detail=(
                f"skipped {len(game_ids)} non-placeholder GDAP row(s) with invalid gameDateTimeUTC"
                f" (examples: {examples}){suffix}; no canonical game was emitted"
            ),
        )
    )


def schedule_url(competition_id: int = COMPETITION_ID) -> str:
    """Return the official GDAP route without ever placing APIM auth in a URL."""

    return f"{GDAP_BASE_URL}/{SCHEDULE_ENDPOINT}?{urlencode({'gdapCompetitionId': int(competition_id)})}"


def _event_slug(season: Any) -> str:
    """Derive FIBA's observed event slug from the GDAP edition-ending year."""

    text = str(season or "").strip()
    match = re.fullmatch(r"(?:20)?(\d{2})", text)
    if not match:
        # The live GDAP value is an integer ending year (2027). Accept a
        # display ``2026-27`` fallback only when it contains the same end year.
        match = re.fullmatch(r"20\d{2}-(\d{2})", text)
    if not match:
        return ""
    end_year = int(match.group(1))
    return f"euroleague-women-{(end_year - 1) % 100:02}-{end_year:02}"


def game_page_url(
    game_id: str, team_a_code: str, team_b_code: str, *, event_slug: str = ""
) -> str:
    """Build a season-derived official FIBA game page URL from GDAP fields."""

    if not str(game_id).isdigit():
        raise VerificationError("FIBA game id is invalid")
    codes = (str(team_a_code or "").upper(), str(team_b_code or "").upper())
    if not all(re.fullmatch(r"[A-Z0-9_-]+", code) for code in codes):
        raise VerificationError("FIBA game page requires both official team codes")
    if not re.fullmatch(r"euroleague-women-\d{2}-\d{2}", str(event_slug)):
        raise VerificationError(
            "FIBA event slug unavailable from GDAP season; check the game-page event route"
        )
    return GAME_PAGE_TEMPLATE.format(
        event_slug=event_slug, game_id=game_id, team_a_code=codes[0], team_b_code=codes[1]
    )


def _clean_name(team: Any, *, side: str, game_id: str) -> tuple[str, str, str]:
    if not isinstance(team, dict):
        raise VerificationError(f"FIBA game {game_id} has no {side} team object")
    team_id = team.get("teamId")
    name = team.get("officialName") or team.get("shortName")
    code = team.get("code")
    if team_id is None or not str(name or "").strip() or not str(code or "").strip():
        raise VerificationError(f"FIBA game {game_id} has incomplete {side} team identity")
    # Keep FIBA's official display spelling.  Shared name aliases would be an
    # unsafe cross-competition identity bridge; IDs below are namespaced.
    return str(team_id), re.sub(r"\s+", " ", str(name).strip()), str(code).upper()


def _required_competition(row: Dict[str, Any], competition_id: int) -> Dict[str, Any]:
    competition = row.get("competition")
    if not isinstance(competition, dict) or competition.get("competitionId") != competition_id:
        raise VerificationError("FIBA schedule row is outside the requested EuroLeague Women competition")
    if competition.get("competitionCode") not in {None, COMPETITION_CODE}:
        raise VerificationError("FIBA competition code is not EuroLeague Women")
    return competition


def _scheduled_utc(row: Dict[str, Any], game_id: str) -> str:
    try:
        return to_utc_iso(game_datetime_utc(row.get("gameDateTimeUTC")))
    except ValueError as exc:
        raise VerificationError(f"FIBA game {game_id} has invalid gameDateTimeUTC") from exc


def _game_status(row: Dict[str, Any]) -> GameStatus:
    """Map GDAP status through the contract shared with the adaptive gate."""

    return {
        "scheduled": GameStatus.SCHEDULED,
        "live": GameStatus.LIVE,
        "final": GameStatus.FINAL,
        "postponed": GameStatus.POSTPONED,
        "cancelled": GameStatus.CANCELLED,
        "abandoned": GameStatus.ABANDONED,
        "unknown": GameStatus.UNKNOWN,
    }[schedule_status(row)]


def game_from_schedule_row(
    row: Dict[str, Any], competition_id: int = COMPETITION_ID
) -> Game:
    """Map one strictly-scoped GDAP schedule row to an isolated canonical game."""

    if not isinstance(row, dict):
        raise VerificationError("FIBA schedule contains a non-object row")
    _required_competition(row, competition_id)
    game_id = row.get("gameId")
    if game_id is None or not str(game_id).isdigit():
        raise VerificationError("FIBA schedule row lacks an integer gameId")
    game_id = str(game_id)
    home_id, home_name, home_code = _clean_name(row.get("teamA"), side="teamA", game_id=game_id)
    away_id, away_name, away_code = _clean_name(row.get("teamB"), side="teamB", game_id=game_id)
    if home_id == away_id or home_code == away_code:
        raise IdentityError(f"FIBA game {game_id} has indistinguishable teamA/teamB identity")
    competition = _required_competition(row, competition_id)
    season = str(competition.get("season") or COMPETITION_SEASON)
    phase = str((row.get("round") or {}).get("roundName") or "EuroLeague Women")
    provider_ids = {
        "fiba_elw_game": game_id,
        "fiba_elw_competition": str(competition_id),
        "fiba_elw_team_a": home_id,
        "fiba_elw_team_b": away_id,
        "fiba_elw_team_a_code": home_code,
        "fiba_elw_team_b_code": away_code,
        "fiba_elw_event_slug": _event_slug(competition.get("season")),
        "fiba_elw_timezone": str(row.get("ianaTimeZone") or ""),
    }
    if _game_status(row) is GameStatus.FINAL:
        # Preserve GDAP's official schedule-index result as a compact
        # provenance check against the separately fetched game-page detail.
        provider_ids["fiba_elw_schedule_team_a_score"] = str(row.get("teamAScore"))
        provider_ids["fiba_elw_schedule_team_b_score"] = str(row.get("teamBScore"))
    return Game(
        game_id=f"fiba-elw-{competition_id}-{game_id}",
        league=League.EUROLEAGUE_W,
        season=season,
        # The namespace is intentional: a similarly named men's club must
        # never share an external team identity with this competition.
        home_team_id=f"fiba-elw:{competition_id}:{home_id}",
        away_team_id=f"fiba-elw:{competition_id}:{away_id}",
        home_name=home_name,
        away_name=away_name,
        scheduled_utc=_scheduled_utc(row, game_id),
        status=_game_status(row),
        phase=phase,
        provider_ids=provider_ids,
    )


def _schedule_rows(
    fetch: FetchFn, competition_id: int = COMPETITION_ID
) -> List[Dict[str, Any]]:
    payload = fetch(schedule_url(competition_id))
    if not isinstance(payload, list) or not payload:
        raise VerificationError("FIBA GDAP schedule has no non-empty game list")
    rows: List[Dict[str, Any]] = []
    for row in payload:
        if not isinstance(row, dict):
            raise VerificationError("FIBA GDAP schedule contains an invalid row")
        _required_competition(row, competition_id)
        rows.append(row)
    return rows


def discover_games(
    date_iso: Optional[str], fetch: FetchFn, competition_id: int = COMPETITION_ID
) -> List[Game]:
    """Read GDAP, date-filter before mapping, and skip unassigned bracket slots.

    FIBA publishes future playoff rows with ``teamA``/``teamB`` set to null.
    They are a truthful schedule fact but have no canonical fixture identity;
    mapping them must not prevent a valid earlier Paris slate from being read.
    In contrast, a date-relevant row with malformed *assigned* teams remains a
    hard contract failure rather than becoming a guessed fixture.
    """

    _reset_gaps()
    target = None
    if date_iso is not None:
        try:
            target = datetime.strptime(date_iso, "%Y-%m-%d").date()
        except ValueError as exc:
            raise ValueError("FIBA EuroLeague Women date must be YYYY-MM-DD") from exc
    zone = ZoneInfo(LEAGUE_TIMEZONE)
    games: List[Game] = []
    unassigned_total = 0
    unassigned_relevant = 0
    unassigned_date_unknown = 0
    unparseable_team_rows: List[str] = []
    for row in _schedule_rows(fetch, competition_id):
        game_id = str(row.get("gameId") or "<unknown>")
        if has_unassigned_team_slot(row):
            # Future bracket placeholders cannot become canonical fixtures.
            # Their timestamp may be absent while FIBA is still constructing a
            # round; skip/report them before time parsing so one placeholder
            # cannot revive the whole-schedule outage.
            unassigned_total += 1
            try:
                row_utc = _scheduled_utc(row, game_id)
            except VerificationError:
                unassigned_date_unknown += 1
                continue
            unassigned_relevant += int(
                target is None or parse_utc(row_utc).astimezone(zone).date() == target
            )
            continue
        try:
            row_utc = _scheduled_utc(row, game_id)
        except VerificationError:
            # The row cannot be safely placed on a requested league day. Keep
            # it out of canonical state but expose explicit provider-error
            # coverage instead of allowing one malformed season-index row to
            # block every date and the history cursor.
            unparseable_team_rows.append(game_id)
            continue
        is_relevant = target is None or parse_utc(row_utc).astimezone(zone).date() == target
        if not is_relevant:
            # Do not turn a malformed team identity on a different league day
            # into an outage for the requested slate.
            continue
        if not has_complete_team_identity(row):
            raise VerificationError(
                f"FIBA game {game_id} has incomplete assigned team identity"
            )
        games.append(game_from_schedule_row(row, competition_id))
    _unassigned_slots_gap(
        unassigned_total, unassigned_relevant, unassigned_date_unknown, date_iso
    )
    _unparseable_timestamp_gap(unparseable_team_rows, date_iso)
    return games


def _number(value: Any, *, field: str) -> int:
    if isinstance(value, bool) or not isinstance(value, (int, float)) or int(value) != value:
        raise VerificationError(f"FIBA game detail has invalid {field}")
    return int(value)


def _period_order(label: str) -> tuple[int, int]:
    upper = str(label).upper()
    match = re.fullmatch(r"Q([1-4])", upper)
    if match:
        return (0, int(match.group(1)))
    match = re.fullmatch(r"OT(\d+)", upper)
    if match:
        return (1, int(match.group(1)))
    raise VerificationError(f"FIBA detail has unsupported period {label!r}")


def _play_by_play_game_id(play_by_play: Any) -> str:
    """Return the explicit game binding carried by FIBA's PBP object."""

    if not isinstance(play_by_play, dict):
        return ""
    # Live RSC currently uses gameTechnicalName; the alternate spelling is
    # retained only as an observed-compatible fallback, never inferred from
    # scores, team names, or occurrence position.
    for field in ("gameTechnicalName", "gameTechnicalNameOrGameId"):
        value = play_by_play.get(field)
        if value is not None and str(value).strip():
            return str(value).strip()
    return ""


def _require_play_by_play_binding(play_by_play: Any, expected_game_id: Any) -> None:
    actual_game_id = _play_by_play_game_id(play_by_play)
    if not actual_game_id:
        raise VerificationError("FIBA play-by-play lacks an explicit game binding")
    if actual_game_id != str(expected_game_id):
        raise VerificationError("FIBA play-by-play belongs to another game")


def _period_scores(detail: Dict[str, Any]) -> tuple[List[int], List[int], int, int, int]:
    """Return individual regulation/OT scores from official cumulative PBP."""

    pbp = detail.get("playByPlay")
    items = pbp.get("items") if isinstance(pbp, dict) else None
    if not isinstance(items, dict):
        raise VerificationError("FIBA game detail has no play-by-play period scores")
    cumulative: List[tuple[tuple[int, int], str, int, int]] = []
    for name, period in items.items():
        if not isinstance(period, dict):
            continue
        label = str(period.get("name") or name).upper()
        try:
            order = _period_order(label)
        except VerificationError:
            continue
        if "scoreA" not in period or "scoreB" not in period:
            continue
        cumulative.append(
            (order, label, _number(period["scoreA"], field=f"{label}.scoreA"), _number(period["scoreB"], field=f"{label}.scoreB"))
        )
    cumulative.sort(key=lambda value: value[0])
    labels = [value[1] for value in cumulative]
    if labels[:4] != ["Q1", "Q2", "Q3", "Q4"]:
        raise VerificationError("FIBA game detail lacks complete Q1-Q4 play-by-play scores")
    expected = ["Q1", "Q2", "Q3", "Q4"] + [f"OT{index}" for index in range(1, len(labels) - 3)]
    if labels != expected:
        raise VerificationError("FIBA game detail has non-contiguous overtime periods")
    previous_a = previous_b = 0
    individual: List[tuple[str, int, int]] = []
    for _order, label, total_a, total_b in cumulative:
        if total_a < previous_a or total_b < previous_b:
            raise VerificationError("FIBA game detail cumulative score moves backwards")
        individual.append((label, total_a - previous_a, total_b - previous_b))
        previous_a, previous_b = total_a, total_b
    regulation = individual[:4]
    overtime = individual[4:]
    return (
        [score_a for _label, score_a, _score_b in regulation],
        [score_b for _label, _score_a, score_b in regulation],
        sum(score_a for _label, score_a, _score_b in overtime),
        sum(score_b for _label, _score_a, score_b in overtime),
        len(overtime),
    )


def result_from_game_detail(game: Game, detail: Dict[str, Any]) -> Game:
    """Attach a reconciled official final from a FIBA game-page detail object."""

    if game.league is not League.EUROLEAGUE_W:
        raise VerificationError("FIBA EuroLeague Women detail was given a game from another league")
    if not isinstance(detail, dict) or not isinstance(detail.get("game"), dict):
        raise VerificationError("FIBA game page has no game object")
    header = detail["game"]
    expected_id = game.provider_ids.get("fiba_elw_game")
    expected_competition = game.provider_ids.get("fiba_elw_competition")
    if str(header.get("gameId")) != str(expected_id):
        raise VerificationError("FIBA game detail belongs to another game")
    _require_play_by_play_binding(detail.get("playByPlay"), expected_id)
    competition = header.get("competition") or {}
    if str(competition.get("competitionId")) != str(expected_competition):
        raise VerificationError("FIBA game detail belongs to another competition")
    header_home = header.get("teamA")
    header_away = header.get("teamB")
    if not isinstance(header_home, dict) or not isinstance(header_away, dict):
        raise VerificationError("FIBA game detail has incomplete team identity")
    if (
        str(header_home.get("teamId")) != game.provider_ids.get("fiba_elw_team_a")
        or str(header_away.get("teamId")) != game.provider_ids.get("fiba_elw_team_b")
    ):
        raise IdentityError("FIBA game detail reverses or contradicts the GDAP team identity")
    if _game_status(header) is not GameStatus.FINAL:
        raise VerificationError("FIBA game page is not an official final")
    required = ("gameStatisticStatusCode", "gamePlayByPlayStatusCode", "gameResultStatusCode")
    if any(str(header.get(field) or "").upper() != "VALID" for field in required):
        raise VerificationError("FIBA final lacks validated result/statistics/play-by-play status")
    home_quarters, away_quarters, home_ot, away_ot, ot_periods = _period_scores(detail)
    home_final = _number(header.get("teamAScore"), field="teamAScore")
    away_final = _number(header.get("teamBScore"), field="teamBScore")
    indexed_scores = (
        game.provider_ids.get("fiba_elw_schedule_team_a_score"),
        game.provider_ids.get("fiba_elw_schedule_team_b_score"),
    )
    if indexed_scores != (None, None) and indexed_scores != (str(home_final), str(away_final)):
        raise VerificationError("FIBA game-page final contradicts the official GDAP schedule result")
    actual = ActualResult(
        period_kind=PeriodKind.QUARTERS,
        home_quarters=home_quarters,
        away_quarters=away_quarters,
        home_overtime=home_ot or None,
        away_overtime=away_ot or None,
        home_final=home_final,
        away_final=away_final,
        overtime_periods=ot_periods,
        source="fiba-game-page",
    )
    reconciled, reason = actual.reconciles()
    if not reconciled:
        raise VerificationError(f"FIBA game detail does not reconcile: {reason}")
    # Preserve an auditable, compact statement that usable official detail was
    # available without copying raw player PII into the canonical game record.
    # Period scores themselves are retained in ``actual`` for grading.
    fiba_evidence = dict(game.evidence.get("fiba") or {})
    fiba_evidence["game_detail"] = {
        "source": "official_fiba_game_page_rsc",
        "game_id": str(header.get("gameId")),
        "competition_id": str(competition.get("competitionId")),
        "result_status": str(header.get("gameResultStatusCode")),
        "statistics_status": str(header.get("gameStatisticStatusCode")),
        "play_by_play_status": str(header.get("gamePlayByPlayStatusCode")),
        "periods": ["Q1", "Q2", "Q3", "Q4"] + [f"OT{index}" for index in range(1, ot_periods + 1)],
        "player_boxscore_team_a_rows": len(detail.get("playersTeamA") or []),
        "player_boxscore_team_b_rows": len(detail.get("playersTeamB") or []),
        "team_detail_available": isinstance(detail.get("gameDetails"), dict),
    }
    # Copy rather than mutate the store's existing object in place.
    return Game(
        **{
            **game.__dict__,
            "status": GameStatus.FINAL,
            "actual": actual,
            "evidence": {**game.evidence, "fiba": fiba_evidence},
        }
    )


def fetch_game_result(game: Game, fetch_detail: Callable[[str], Dict[str, Any]]) -> Game:
    """Fetch and validate one official completed game's usable detail payload."""

    detail = fetch_detail(
        game_page_url(
            game.provider_ids.get("fiba_elw_game", ""),
            game.provider_ids.get("fiba_elw_team_a_code", ""),
            game.provider_ids.get("fiba_elw_team_b_code", ""),
            event_slug=game.provider_ids.get("fiba_elw_event_slug", ""),
        )
    )
    return result_from_game_detail(game, detail)


def _balanced_json(text: str, start: int) -> Any:
    """Decode one JSON object/array beginning at ``start`` in a Flight payload."""

    decoder = json.JSONDecoder()
    value, _end = decoder.raw_decode(text[start:])
    return value


def extract_game_page(text: str, game_id: str) -> Dict[str, Any]:
    """Extract the FIBA RSC objects required for official result parsing.

    FIBA's Next.js Flight stream includes JSON fragments after these anchors.
    Unknown/missing fragments fail closed; the schedule remains authoritative
    for fixtures, while results wait for usable official period detail.
    """

    if not isinstance(text, str) or not text:
        raise VerificationError("empty FIBA game-page response")
    extracted: Dict[str, Any] = {}
    # Flight can carry the same fragment once directly and once inside an
    # escaped ``self.__next_f.push`` string.  Inspect both forms; JSON decoding
    # below still rejects incomplete/invalid fragments.
    candidates = [text]
    normalized = text
    for _ in range(2):
        decoded = normalized.replace(r'\"', '"')
        if decoded == normalized:
            break
        candidates.append(decoded)
        normalized = decoded
    for key in ("game", "playByPlay", "gameDetails", "playersTeamA", "playersTeamB"):
        # Flight includes UI configuration objects such as
        # ``playByPlay: {timeout: ...}`` before the actual game payload. Walk
        # every occurrence and accept only the expected shape *and binding*.
        for candidate in candidates:
            for match in re.finditer(rf'"{re.escape(key)}"\s*:\s*([{{[])', candidate):
                try:
                    value = _balanced_json(candidate, match.start(1))
                except (ValueError, json.JSONDecodeError):
                    continue
                valid = (
                    isinstance(value, dict) and str(value.get("gameId")) == str(game_id)
                    if key == "game"
                    else (
                        isinstance(value, dict)
                        and isinstance(value.get("items"), dict)
                        and _play_by_play_game_id(value) == str(game_id)
                    )
                    if key == "playByPlay"
                    else isinstance(value, dict)
                    if key == "gameDetails"
                    else isinstance(value, list)
                )
                if valid:
                    extracted[key] = value
                    break
            if key in extracted:
                break
    if not isinstance(extracted.get("game"), dict) or str(extracted["game"].get("gameId")) != str(game_id):
        raise VerificationError("FIBA game page did not expose the requested game header")
    if not isinstance(extracted.get("playByPlay"), dict):
        raise VerificationError("FIBA game page did not expose requested-game play-by-play")
    _require_play_by_play_binding(extracted["playByPlay"], game_id)
    return extracted


class FibaClient:
    """Live transport that keeps FIBA's public subscription key out of URLs/logs."""

    def __init__(self, http: Optional[HttpClient] = None) -> None:
        # FIBA's RSC game pages can refuse a generic Python agent; use the
        # browser-shaped public-page agent FIBA serves for this read-only path.
        self.http = http or HttpClient(
            user_agent=(
                "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
                "(KHTML, like Gecko) Chrome/129.0 Safari/537.36"
            )
        )
        self._key: Optional[str] = None

    def _subscription_key(self, *, refresh: bool = False) -> str:
        if self._key and not refresh:
            return self._key
        try:
            page = self.http.get_text(KEY_PAGE_URL)
            self._key = page_subscription_key(page) or PUBLIC_FALLBACK_SUBSCRIPTION_KEY
        except Exception:
            self._key = PUBLIC_FALLBACK_SUBSCRIPTION_KEY
        return self._key

    def get_json(self, url: str) -> Any:
        """Read only this adapter's fixed GDAP competition through header auth."""

        parsed = urlparse(url)
        expected_path = f"/{SCHEDULE_ENDPOINT}"
        competition_ids = parse_qs(parsed.query).get("gdapCompetitionId", [])
        if (
            parsed.scheme != "https"
            or parsed.netloc != urlparse(GDAP_BASE_URL).netloc
            or parsed.path != f"/hapi{expected_path}"
            or competition_ids != [str(COMPETITION_ID)]
        ):
            raise VerificationError("FIBA client refused a non-ELW GDAP schedule URL")
        safe_url = (
            f"{GDAP_BASE_URL}/{SCHEDULE_ENDPOINT}?"
            + urlencode({"gdapCompetitionId": COMPETITION_ID})
        )
        # Do not retain query-key authentication in request URLs/logs. A 401
        # or other transient failure gets one retry after a public-page key
        # refresh; errors still propagate as explicit provider failures.
        for attempt in range(2):
            try:
                return self.http.get_json(
                    safe_url,
                    headers={
                        KEY_HEADER: self._subscription_key(refresh=bool(attempt)),
                        "Accept": "application/json",
                    },
                )
            except Exception as exc:
                if attempt:
                    raise exc
        raise AssertionError("unreachable")  # pragma: no cover

    def last_quota(self) -> Dict[str, Any]:
        """Return any quota/rate-limit headers captured by the shared HTTP client."""

        return dict(getattr(self.http, "last_quota", {}) or {})

    def game_detail(self, url: str) -> Dict[str, Any]:
        game_id = re.search(r"/games/(\d+)-", url)
        if not game_id:
            raise VerificationError("FIBA game-page URL has no game id")
        try:
            text = self.http.get_text(url, headers={"RSC": "1", "Accept": "text/x-component"})
        except HttpError as exc:
            if exc.status == 404:
                raise VerificationError(
                    "FIBA event game page returned 404; event slug not found — check the event route"
                ) from exc
            raise
        return extract_game_page(text, game_id.group(1))


def smoke(fetch: FetchFn, *, verification: str = "FIXTURE_VERIFIED") -> Dict[str, Any]:
    """Read official FIBA schedule coverage without persistence or false success."""

    games = discover_games(None, fetch)
    finals = sum(game.status is GameStatus.FINAL for game in games)
    return {
        "verification": verification,
        "provider": "fiba_gdap",
        "competition_id": COMPETITION_ID,
        "schedule_route": SCHEDULE_ENDPOINT,
        "fixtures_found": len(games),
        "official_finals_indexed": finals,
        "unassigned_schedule_slot_gaps": sum(
            gap.required_field == "fiba_elw_unassigned_schedule_slots" for gap in _LAST_GAPS
        ),
        "schedule_timestamp_gaps": sum(
            gap.required_field == "fiba_elw_schedule_timestamp" for gap in _LAST_GAPS
        ),
        "gaps": [
            {"required_field": gap.required_field, "reason": gap.reason.value, "detail": gap.detail}
            for gap in _LAST_GAPS
        ],
        "game_detail": "official_fiba_game_page_play_by_play_period_scores",
        "canonical_tip_time": "gameDateTimeUTC",
    }


__all__ = [
    "COMPETITION_ID",
    "FibaClient",
    "GAME_PAGE_TEMPLATE",
    "GDAP_BASE_URL",
    "LEAGUE_TIMEZONE",
    "PUBLIC_FALLBACK_SUBSCRIPTION_KEY",
    "VerificationError",
    "discover_games",
    "extract_game_page",
    "fetch_game_result",
    "game_from_schedule_row",
    "game_page_url",
    "last_gaps",
    "result_from_game_detail",
    "schedule_url",
    "smoke",
]
