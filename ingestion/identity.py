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
import unicodedata
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
    "FC Barcelona": ["Barca", "BAR", "Barcelona"],
    # EuroLeague — BetExplorer display names (live-verified 2026-11-16
    # against the official api-live schedule for the same games).
    "Dubai Basketball": ["Dubai"],
    "Anadolu Efes Istanbul": ["Anadolu Efes"],
    "Maccabi Rapyd Tel Aviv": ["Maccabi Tel Aviv"],
    "Olympiacos Piraeus": ["Olympiacos"],
    "Fenerbahce Tarfin Istanbul": ["Fenerbahce"],
    "LDLC ASVEL Villeurbanne": ["ASVEL", "Lyon-Villeurbanne"],
    "FC Bayern Munich": ["Bayern", "Bayern Munich"],
    "Armani Olimpia Milan": ["Olimpia Milano", "Milan"],
    "Crvena Zvezda Meridianbet Belgrade": [
        "Crvena Zvezda",
        "Crvena Zvezda Meridianbet",
        "Red Star Belgrade",
    ],
    "Kosner Baskonia Vitoria-Gasteiz": ["Baskonia", "Saski Baskonia"],
    "Hapoel IBI Tel Aviv": ["Hapoel Tel Aviv"],
    "Partizan Mozzart Bet Belgrade": ["Partizan", "Partizan Mozzart Bet"],
    "Valencia Basket": ["Valencia"],
    "Paris Basketball": ["Paris"],
    # CBA — canonical names are the BetExplorer display spellings; the
    # Chinese variants are the official 2026-27 sponsor names observed live
    # on the league data API (every team seen 2026-10-24..2026-11-03). The
    # mapping is season-scoped evidence: sponsor names change, and unknown
    # Chinese names stay unmapped (they quarantine, never guess).
    "Beijing": ["北京首钢"],
    "Beijing Royal Fighters": ["北京控股"],
    "Fujian": ["福建晋江文旅"],
    "Guangdong": ["广东东阳光"],
    "Guangzhou": ["广州智都集团"],
    "Jiangsu Dragons": ["江苏肯帝亚"],
    "Jilin": ["长白山恩都里", "吉林东北虎"],
    "Liaoning": ["辽宁本钢"],
    "Nanjing Tongxi": ["南京同曦宙光"],
    "Ningbo Rockets": ["宁波町渥"],
    "Qingdao": ["青岛崂山啤酒"],
    "Shandong": ["山东高速"],
    "Shanghai": ["上海久事"],
    "Shanxi Zhongyu": ["山西汾酒"],
    "Shenzhen": ["深圳马可波罗"],
    "Sichuan": ["四川锦城"],
    "Tianjin": ["天津先行者"],
    "Xinjiang": ["新疆广汇能源"],
    "Zhejiang Chouzhou": ["浙江稠州金租"],
    "Zhejiang Guangsha": ["浙江浙商证券"],
}


def fold_diacritics(text: str) -> str:
    """Strip accents so ``Força`` == ``Forca``.

    Not cosmetic: BetExplorer publishes every club name with its diacritics
    already removed (``Forca Lleida``, ``Breogan``, ``Wurzburg``, verified across
    five competition pages on 2026-11-23) while league feeds keep them. Without
    folding, those rows never match a canonical game and the price silently
    disappears into "no matched odds".
    """

    decomposed = unicodedata.normalize("NFKD", str(text or ""))
    return "".join(ch for ch in decomposed if not unicodedata.combining(ch))


def normalize_team_name(name: str) -> str:
    """Lowercase, fold diacritics, trim, and collapse whitespace."""
    return re.sub(r"\s+", " ", fold_diacritics(name).strip()).lower()


def canonical_team_name(name: str, aliases: Optional[Dict[str, List[str]]] = None) -> str:
    """Map a provider name to the canonical display name via the alias table."""
    table = aliases if aliases is not None else DEFAULT_TEAM_ALIASES
    norm = normalize_team_name(name)
    for canonical, variants in table.items():
        if norm == normalize_team_name(canonical) or norm in {
            normalize_team_name(v) for v in variants
        }:
            return canonical
    # No alias hit: return title-cased collapsed original.
    return re.sub(r"\s+", " ", str(name or "").strip())


# Tokens that never distinguish two basketball teams.  Conference/school
# qualifiers are *leading* tokens ("LSU New Orleans Privateers" vs "New Orleans
# Privateers"), which is why team_names_agree requires a suffix relationship
# instead of a bag-of-words comparison.
_TEAM_NAME_FILLER = frozenset(
    {"university", "univ", "college", "of", "the", "at", "team", "men", "mens", "women", "womens"}
)


def _team_name_tokens(name: str) -> Tuple[str, ...]:
    words = re.split(r"[^0-9a-z]+", normalize_team_name(name))
    return tuple(word for word in words if word and word not in _TEAM_NAME_FILLER)


def team_names_agree(left: str, right: str) -> bool:
    """True when two team labels denote the same team.

    A provider may prepend a school qualifier to the same basketball team
    ("New Orleans Privateers" vs "LSU New Orleans Privateers") and may drop the
    mascot ("Troy" vs "Troy Trojans").  Only those two shapes are treated as
    agreement: equal alias-resolved names, or one token sequence being a suffix
    of the other.  A shared mascot with a *differing* qualifier is deliberately
    not agreement, because "Alabama Hornets" and "Alabama State Hornets" name
    different teams.
    """

    if not str(left or "").strip() or not str(right or "").strip():
        return False
    if normalize_team_name(canonical_team_name(left)) == normalize_team_name(
        canonical_team_name(right)
    ):
        return True
    first, second = _team_name_tokens(left), _team_name_tokens(right)
    if not first or not second:
        return False
    if first == second:
        return True
    shorter, longer = (first, second) if len(first) <= len(second) else (second, first)
    if len(shorter) >= 2:
        return longer[-len(shorter) :] == shorter
    return shorter[0] in longer


BOOK_CANONICALS: Dict[str, str] = {
    "draftkings": "DraftKings",
    "fanduel": "FanDuel",
    "betmgm": "BetMGM",
    "caesarssportsbook": "Caesars",
    "caesars": "Caesars",
    "pointsbet": "PointsBet",
    "betrivers": "BetRivers",
    "wynnbet": "WynnBET",
    "unibet": "Unibet",
    "barstool": "Barstool",
    "espnbet": "ESPN BET",
    "bet365": "bet365",
    "williamhill": "William Hill",
    "superbook": "SuperBook",
}


def canonical_book_name(name: str | None) -> str:
    """Normalize bookmaker display names to a canonical title.

    ESPN historically emitted both ``Draft Kings`` and ``DraftKings`` for the
    same market; both must map to one logical pick. Normalization is
    case/space/punctuation insensitive and falls back to the trimmed original
    when no known alias matches.
    """
    if not name:
        return ""
    raw = str(name).strip()
    if not raw:
        return ""
    key = re.sub(r"[^a-z0-9]", "", raw.lower())
    if key in BOOK_CANONICALS:
        return BOOK_CANONICALS[key]
    # Fallback: collapse whitespace but preserve original casing for unknown books.
    return re.sub(r"\s+", " ", raw)


def canonical_team_key(name: str) -> str:
    """Build a punctuation- and diacritic-insensitive team identity key.

    Display names remain untouched; this key only normalizes spelling-level
    Unicode and punctuation differences. It does not remove mascots, abbreviate
    names, reverse home/away, or perform fuzzy matching.
    """
    canonical = canonical_team_name(name)
    decomposed = unicodedata.normalize("NFKD", normalize_team_name(canonical)).casefold()
    return "".join(character for character in decomposed if character.isalnum())


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
