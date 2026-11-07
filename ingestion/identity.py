"""Team/game/time identity: aliases, game keys, provider event map, UTC.

Rules (see docs/CONTRACT.md):
- Prefer provider event IDs; maintain a game map linking provider ids.
- Cross-provider joins require compatible canonical teams AND scheduled time
  within a strict tolerance; never team-only matching across dates.
- Store timezone-aware ISO 8601 internally and normalize to UTC for temporal
  comparisons. Slate/report time zones are applied only after UTC parsing.
- Never infer dates from locale-dependent display strings.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import datetime, timedelta, timezone
from typing import Dict, Iterable, List, Optional, Tuple

from .models import Game, League

UTC = timezone.utc

# Strict cross-provider join tolerance.
TIME_TOLERANCE = timedelta(minutes=30)

# Default alias table. Extendable via TeamAliases store table.
# Keys are canonical team names. Matching is case-insensitive via
# normalize_team_name, so provider casing (for example "REAL MADRID") still
# resolves.
DEFAULT_TEAM_ALIASES: Dict[str, List[str]] = {
    # Display name -> known provider variants (matching is case-insensitive).
    "Los Angeles Lakers": ["LA Lakers", "Lakers"],
    "Los Angeles Clippers": ["LA Clippers", "Clippers"],
    "Golden State Warriors": ["GS Warriors", "Warriors"],
    "New Orleans Pelicans": ["NO Pelicans", "Pelicans"],
    "Philadelphia 76ers": ["76ers", "Sixers", "Philadelphia Sixers"],
    "Portland Trail Blazers": ["Trail Blazers", "Blazers"],
    "San Antonio Spurs": ["Spurs"],
    "Oklahoma City Thunder": ["OKC Thunder", "Thunder"],
    "Real Madrid": ["RMB"],
    "Panathinaikos AKTOR Athens": ["Panathinaikos", "PAO", "Panathinaikos AKTOR"],
    "Besiktas Istanbul": ["Besiktas", "BJK"],
    "Zalgiris Kaunas": ["Zalgiris", "ZAL"],
    "FC Barcelona": ["Barca", "BAR"],
}


def normalize_team_name(name: str) -> str:
    """Lowercase, trim, and collapse whitespace."""
    return re.sub(r"\s+", " ", str(name or "").strip()).lower()


def canonical_team_name(name: str, aliases: Optional[Dict[str, List[str]]] = None) -> str:
    """Map a provider name to the canonical display name via the alias table."""
    table = aliases if aliases is not None else DEFAULT_TEAM_ALIASES
    norm = normalize_team_name(name)
    for canonical, variants in table.items():
        if norm == normalize_team_name(canonical) or norm in {normalize_team_name(v) for v in variants}:
            return canonical
    # No alias hit: return title-cased collapsed original.
    return re.sub(r"\s+", " ", str(name or "").strip())


def canonical_team_key(name: str) -> str:
    """Case-insensitive canonical key using the built-in alias registry."""
    return normalize_team_name(canonical_team_name(name))


# ---------------------------------------------------------------------------
# Time helpers
# ---------------------------------------------------------------------------


def parse_utc(value: str) -> datetime:
    """Parse an ISO 8601 timestamp into an aware UTC datetime."""
    if isinstance(value, datetime):
        dt = value
    else:
        text = str(value).strip()
        if text.endswith("Z"):
            text = text[:-1] + "+00:00"
        dt = datetime.fromisoformat(text)
    if dt.tzinfo is None:
        raise ValueError(f"naive datetime not allowed: {value!r}")
    return dt.astimezone(UTC)


def to_utc_iso(dt: datetime) -> str:
    return dt.astimezone(UTC).strftime("%Y-%m-%dT%H:%M:%SZ")


def now_utc_iso() -> str:
    return to_utc_iso(datetime.now(UTC))


def times_within(a: str, b: str, tolerance: timedelta = TIME_TOLERANCE) -> bool:
    return abs(parse_utc(a) - parse_utc(b)) <= tolerance


# ---------------------------------------------------------------------------
# Game map / provider event ids
# ---------------------------------------------------------------------------


@dataclass
class GameMapEntry:
    """Links one canonical game to all known provider event IDs."""

    game_id: str
    league: League
    scheduled_utc: str
    home_name: str
    away_name: str
    provider_ids: Dict[str, str] = field(default_factory=dict)

    def matches_fixture(self, home: str, away: str, scheduled_utc: str) -> bool:
        """Strict join: canonical teams (order-sensitive) + time tolerance."""
        return (
            normalize_team_name(self.home_name) == normalize_team_name(home)
            and normalize_team_name(self.away_name) == normalize_team_name(away)
            and times_within(self.scheduled_utc, scheduled_utc)
        )


class GameMap:
    """In-memory game map; persisted via the store's ``Games`` table."""

    def __init__(self, entries: Optional[Iterable[GameMapEntry]] = None) -> None:
        self._by_game_id: Dict[str, GameMapEntry] = {}
        self._by_provider: Dict[Tuple[str, str], GameMapEntry] = {}
        for e in entries or []:
            self.add(e)

    def add(self, entry: GameMapEntry) -> None:
        self._by_game_id[entry.game_id] = entry
        for provider, pid in entry.provider_ids.items():
            key = (provider, str(pid))
            existing = self._by_provider.get(key)
            if existing is not None and existing.game_id != entry.game_id:
                raise ValueError(
                    f"provider id conflict: {provider}:{pid} maps to "
                    f"{existing.game_id} and {entry.game_id}"
                )
            self._by_provider[key] = entry

    def by_game_id(self, game_id: str) -> Optional[GameMapEntry]:
        return self._by_game_id.get(game_id)

    def by_provider_id(self, provider: str, provider_id: str) -> Optional[GameMapEntry]:
        return self._by_provider.get((provider, str(provider_id)))

    def find_join(
        self, provider: str, provider_id: str, home: str, away: str, scheduled_utc: str
    ) -> Optional[GameMapEntry]:
        """Resolve a provider event into the map, or None.

        Raises ``ValueError`` when the provider id exists but the fixture
        identity contradicts the stored entry (reschedule handling must be
        explicit, never accidental).
        """
        entry = self.by_provider_id(provider, provider_id)
        if entry is None:
            return None
        if not entry.matches_fixture(home, away, scheduled_utc):
            raise ValueError(
                f"identity conflict for {provider}:{provider_id}: map has "
                f"{entry.home_name} vs {entry.away_name} at {entry.scheduled_utc}; "
                f"got {home} vs {away} at {scheduled_utc}. Explicit reschedule "
                "reconciliation is required."
            )
        return entry

    def all(self) -> List[GameMapEntry]:
        return list(self._by_game_id.values())


def game_from_map(entry: GameMapEntry) -> Game:
    return Game(
        game_id=entry.game_id,
        league=entry.league,
        season=entry.league.value,
        home_team_id=entry.provider_ids.get("espn", entry.game_id),
        away_team_id=entry.provider_ids.get("espn", entry.game_id),
        home_name=entry.home_name,
        away_name=entry.away_name,
        scheduled_utc=entry.scheduled_utc,
        provider_ids=dict(entry.provider_ids),
    )
