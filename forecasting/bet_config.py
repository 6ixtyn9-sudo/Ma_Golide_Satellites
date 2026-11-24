"""Bet configuration as data, ported from the retired system's config layer.

Why this exists
---------------
The retired system did not hard-code its gates. ``Config_Tier1`` was a sheet of
key/value rows read through a ``KEY_MAP`` that *typed and range-checked every key*
(``intOrDefault(val, def, min, max)``), falling back per key to ``DEFAULTS`` and
recording where each value came from -- ``_raw`` and ``_source: 'defaults'``.
``Config_Tier2`` turned ``even_target`` / ``medium_target`` / ``strong_target`` into
confidence tier cuts clamped and forced monotonic (``medium >= weak + 1``,
``strong >= medium + 1``). ``Config_Ledger`` stamped every slip row with
``"CFG_" + djb2(canonical config JSON)`` so a bet could always be traced back to the
exact configuration that produced it.

This module keeps that shape and moves it from a spreadsheet to a JSON file:

* :data:`KEYS` is the machine-readable replacement for ``KEY_MAP`` -- every gate
  carries its default, type, bounds and section.
* :class:`BetConfig` loads JSON overrides, validates each key against its bounds,
  falls back to the default with a warning (never silently), and enforces the tier
  ladder's invariants.
* :attr:`BetConfig.stamp_id` is the same ``CFG_`` stamp the ledger used: a djb2
  hash over the canonical key-sorted JSON, so two identical configurations stamp
  identically and any change stamps differently.

Nothing here decides a bet. It only decides what the numbers are, and remembers
where they came from.
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, Mapping, Optional, Tuple

#: Bumped whenever the set of keys or their meaning changes. Part of the stamp, so
#: a semantic change can never collide with an old configuration.
CONFIG_VERSION = "magolide-bets-v1"

#: Where the shipped baseline lives. Defaults in code and this file's values are
#: kept identical by a test, so the file is documentation a reader can diff.
SHIPPED_PATH = Path(__file__).resolve().parent.parent / "config" / "magolide_bets.json"


@dataclass(frozen=True)
class ConfigKey:
    """One configurable number, with the bounds the sheet loader used to enforce."""

    name: str
    default: Any
    kind: str  # "float" | "int" | "bool"
    minimum: float
    maximum: float
    section: str
    doc: str = ""


def _as_float(raw: Any) -> Optional[float]:
    if isinstance(raw, bool):
        return None
    if isinstance(raw, (int, float)):
        return float(raw)
    if isinstance(raw, str):
        try:
            return float(raw.strip())
        except ValueError:
            return None
    return None


def _as_bool(raw: Any) -> Optional[bool]:
    """The sheet loader's ``parseBoolean``: TRUE/YES/1/ON/ENABLED."""

    if isinstance(raw, bool):
        return raw
    if isinstance(raw, (int, float)):
        return raw != 0
    if raw is None:
        return None
    text = str(raw).strip().upper()
    if not text:
        return None
    return text in {"TRUE", "YES", "1", "ON", "ENABLED"}


def coerce_value(key: ConfigKey, raw: Any) -> Tuple[Any, Optional[str]]:
    """``(value, warning)`` -- the loader's one-key contract.

    A value that cannot be parsed, or that falls outside the key's bounds, becomes
    the default *and says so*. The retired loader behaved the same way; the only
    change is that the complaint is returned instead of logged.
    """

    if raw is None or (isinstance(raw, str) and not raw.strip()):
        # An empty cell is "unset", not a value: the sheet's boolOrDefault and
        # toNumber both returned the default silently in that case.
        return key.default, None

    if key.kind == "bool":
        parsed = _as_bool(raw)
        if parsed is None:
            return key.default, f"{key.name}: {raw!r} is not a boolean, using {key.default!r}"
        return parsed, None

    number = _as_float(raw)
    if number is None:
        return key.default, f"{key.name}: {raw!r} is not a number, using {key.default!r}"
    if key.kind == "int":
        number = float(int(number))
    if number < key.minimum or number > key.maximum:
        return (
            key.default,
            f"{key.name}: {number} outside [{key.minimum}, {key.maximum}], "
            f"using {key.default!r}",
        )
    return (int(number) if key.kind == "int" else number), None


def _key(name: str, default: Any, kind: str, minimum: float, maximum: float, section: str,
         doc: str = "") -> ConfigKey:
    return ConfigKey(name, default, kind, minimum, maximum, section, doc)


#: ``KEY_MAP``'s replacement. Defaults are the values the retired system shipped:
#: ``Contract_Enforcement.canonicalizeConfig`` for the bet gates, ``Data_Parser``'s
#: default config for the tier ladder and first-half/FT O/U knobs, and
#: ``H1_CONFIG`` for the combiner's own clamps.
KEYS: Dict[str, ConfigKey] = {
    k.name: k
    for k in (
        # -- banker (moneyline, the side the model banks) ----------------------
        _key("bankerThreshold", 60.0, "float", 0.0, 100.0, "banker",
             "Confidence a named pick must clear to be bucketed a banker."),
        _key("minBankerOdds", 1.01, "float", 1.0, 10.0, "banker"),
        _key("maxBankerOdds", 3.00, "float", 1.0, 50.0, "banker"),
        # -- robber (the dog) --------------------------------------------------
        _key("enableRobbers", True, "bool", 0, 1, "robber"),
        _key("robberMinConf", 54.0, "float", 0.0, 100.0, "robber"),
        _key("robberMaxConf", 45.0, "float", 0.0, 100.0, "robber"),
        _key("minOdds", 1.80, "float", 1.0, 100.0, "robber"),
        _key("maxOdds", 12.00, "float", 1.0, 1000.0, "robber"),
        _key("robberMinScore", 35.0, "float", 0.0, 100.0, "robber"),
        _key("robberMinEdge", 0.08, "float", 0.0, 1.0, "robber"),
        # -- first half 1x2 ----------------------------------------------------
        _key("enableFirstHalf", True, "bool", 0, 1, "first_half"),
        _key("firstHalfMinConf", 58.0, "float", 0.0, 100.0, "first_half"),
        _key("firstHalfTwoWay", True, "bool", 0, 1, "first_half"),
        _key("firstHalfDrawThreshold", 3.0, "float", 0.0, 30.0, "first_half",
             "Only consulted when firstHalfTwoWay is off."),
        _key("firstHalfHomeAdvantage", 1.5, "float", -20.0, 20.0, "first_half"),
        _key("firstHalfMinEdge", 0.03, "float", 0.0, 1.0, "first_half"),
        _key("firstHalfSignalDivisor", 100.0, "float", 1.0, 1000.0, "first_half",
             "Score units per unit of the combiner's [-1, 1] signal scale."),
        _key("firstHalfSignalScale", 2.0, "float", 0.1, 10.0, "first_half",
             "H1_CONFIG applies absSignal * 2 before the logistic."),
        _key("firstHalfMarginPerSignal", 8.0, "float", 0.0, 50.0, "first_half",
             "The script's own fallback: margin = combinedSignal * 8."),
        _key("firstHalfMinConfidence", 51.0, "float", 0.0, 100.0, "first_half"),
        _key("firstHalfMaxConfidence", 85.0, "float", 0.0, 100.0, "first_half"),
        # -- full time over/under ---------------------------------------------
        _key("enableFTOU", True, "bool", 0, 1, "ft_ou"),
        _key("ftOuMinConf", 55.0, "float", 0.0, 100.0, "ft_ou"),
        _key("ftOuMinEv", 0.005, "float", 0.0, 10.0, "ft_ou"),
        _key("ftouEdgeThreshold", 0.02, "float", 0.0, 1.0, "ft_ou"),
        _key("ftouSigmaFloor", 15.0, "float", 1.0, 100.0, "ft_ou"),
        _key("ftouSigmaScale", 1.0, "float", 0.1, 5.0, "ft_ou"),
        _key("ftouShrinkK", 8.0, "float", 0.0, 100.0, "ft_ou",
             "Shrinkage towards the observed total line: (n*mu + k*line) / (n+k)."),
        _key("ftouMinSamples", 3, "int", 1, 100, "ft_ou"),
        # -- tier ladder and confidence ---------------------------------------
        # -- ticket (one number, cross-cutting) --------------------------------
        #: A paper ticket must be able to pay for its own risk. Below this price a
        #: leg cannot: on 2026-11-24 the NCAA_M lane built five "3-leg accas" from
        #: ESPN's lopsided-game moneyline sentinel, the best of them priced at
        #: @1.003 -- one full unit risked to return 0.3%. 376 of that lane's 421
        #: eligible priced picks sat at 1.20 or better, so the floor costs nothing
        #: but the absurd tickets.
        _key("minTicketLegOdds", 1.20, "float", 1.0, 5.0, "ticket"),
        #: The most expensive leg a paper ticket may carry. A ceiling, not a
        #: target: it keeps the ranking inside prices the model can speak about.
        #: Above it the model's claimed edge explodes -- on 2026-11-24 it rated
        #: Rice p=72% against a 3.50 quote (+152% claimed), and a blend that was
        #: allowed to reach that leg assembled a 10.4 acca claiming +194% EV.
        _key("maxTicketLegOdds", 3.00, "float", 1.0, 20.0, "ticket"),
        #: How much of a ticket's ranking is price rather than probability.
        #: 0.0 reproduces "maximise joint probability", which walks to the
        #: shortest prices on the board; 1.0 chases the longest. 0.15 is measured
        #: on the 2026-11-24 NCAA_M slate: it keeps the top ticket at @2.09 with
        #: a 48.9% joint probability (EV +2%), where 0.35 produced @10.41 at a
        #: claimed +194% -- the model's overclaim, not an opportunity.
        _key("ticketRankingPriceWeight", 0.15, "float", 0.0, 1.0, "ticket"),
        _key("threshold", 25.0, "float", 0.0, 100.0, "tiers"),
        _key("tier_weak_min_score", 50.0, "float", 0.0, 200.0, "tiers"),
        _key("tier_medium_min_score", 60.0, "float", 0.0, 200.0, "tiers"),
        _key("tier_strong_min_score", 75.0, "float", 0.0, 200.0, "tiers"),
        _key("confMin", 50.0, "float", 0.0, 100.0, "tiers"),
        _key("confMax", 95.0, "float", 0.0, 100.0, "tiers"),
        _key("confidence_scale", 30.0, "float", 1.0, 200.0, "tiers"),
    )
}

DEFAULTS: Dict[str, Any] = {name: key.default for name, key in KEYS.items()}


def djb2(text: str) -> str:
    """``Config_Ledger._djb2Hash``: djb2 over UTF-16 code units, 8 uppercase hex."""

    h = 5381
    for ch in text:
        h = ((h << 5) + h) ^ ord(ch)
        h &= 0xFFFFFFFF
    return f"{h:08x}".upper()


@dataclass(frozen=True)
class BetConfig:
    """A validated configuration, its provenance, and its stamp."""

    version: str = CONFIG_VERSION
    values: Mapping[str, Any] = field(default_factory=lambda: dict(DEFAULTS))
    sources: Mapping[str, str] = field(default_factory=lambda: {k: "default" for k in DEFAULTS})
    league_weights: Mapping[str, float] = field(default_factory=dict)
    warnings: Tuple[str, ...] = ()

    # -- reading ------------------------------------------------------------
    def get(self, name: str) -> Any:
        return self.values[name]

    def as_dict(self) -> Dict[str, Any]:
        return {name: self.values[name] for name in sorted(self.values)}

    def canonical_json(self) -> str:
        """Key-sorted, separators-tight JSON -- what the stamp hashes.

        League weights are part of it: a run that discounts one league is a
        different configuration and must not share a stamp with one that does not.
        """

        return json.dumps(
            {
                "version": self.version,
                "keys": self.as_dict(),
                "leagues": {name: float(weight) for name, weight in sorted(self.league_weights.items())},
            },
            sort_keys=True,
            separators=(",", ":"),
        )

    def league_weight(self, league: str) -> float:
        """The discount for a league, 1.0 unless the tuner applied one."""

        return float(self.league_weights.get(league, 1.0))

    @property
    def stamp_id(self) -> str:
        """``CFG_`` + djb2 of the canonical JSON, exactly as the ledger stamped rows."""

        return "CFG_" + djb2(self.canonical_json())

    def provenance(self, name: str) -> str:
        return self.sources.get(name, "default")

    # -- deriving -----------------------------------------------------------
    def with_overrides(
        self, changes: Mapping[str, Any], *, source: str = "override"
    ) -> "BetConfig":
        """A new config with ``changes`` applied, validated and re-invariant-ed."""

        values = dict(self.values)
        sources = dict(self.sources)
        warnings = list(self.warnings)
        for name, raw in changes.items():
            key = KEYS.get(name)
            if key is None:
                warnings.append(f"{name}: unknown key ignored")
                continue
            value, warning = coerce_value(key, raw)
            if warning:
                warnings.append(warning)
            values[name] = value
            sources[name] = source
        values, invariant_warnings = enforce_invariants(values)
        warnings.extend(invariant_warnings)
        return BetConfig(self.version, values, sources, self.league_weights, tuple(warnings))

    def with_leagues(
        self, weights: Mapping[str, Any], *, source: str = "override"
    ) -> "BetConfig":
        """A new config with per-league discounts applied, clamped to [0.1, 1.0].

        The tuner only ever discounts (the sheet's advice was ``0.7``); a weight
        above 1.0 would let a noisy league raise confidence, which is the failure
        mode the discount exists to prevent.
        """

        merged = dict(self.league_weights)
        warnings = list(self.warnings)
        for name, raw in weights.items():
            number = _as_float(raw)
            if number is None:
                warnings.append(f"league weight {name}: {raw!r} is not a number, ignored")
                continue
            if number < 0.1 or number > 1.0:
                clamped = min(1.0, max(0.1, number))
                warnings.append(f"league weight {name}: {number} clamped to {clamped}")
                number = clamped
            merged[str(name)] = number
        return BetConfig(self.version, self.values, self.sources, merged, tuple(warnings))


def enforce_invariants(values: Mapping[str, Any]) -> Tuple[Dict[str, Any], list]:
    """Clamp the relationships, not just the keys.

    The retired loader forced ``weak < medium < strong`` and ``confMin < confMax``
    on every load. A config that could not produce a tier was not allowed to exist,
    so no downstream code ever had to defend against one.
    """

    out = dict(values)
    warnings: list = []

    weak = float(out["tier_weak_min_score"])
    medium = float(out["tier_medium_min_score"])
    strong = float(out["tier_strong_min_score"])
    if medium <= weak:
        medium = weak + 1.0
        warnings.append(f"tier_medium_min_score raised to {medium} to clear the weak floor")
    if strong <= medium:
        strong = medium + 1.0
        warnings.append(f"tier_strong_min_score raised to {strong} to clear the medium floor")
    out["tier_weak_min_score"], out["tier_medium_min_score"], out["tier_strong_min_score"] = (
        weak, medium, strong,
    )

    if float(out["threshold"]) > weak:
        out["threshold"] = weak
        warnings.append(f"threshold lowered to {weak}: a RISKY score cannot be tiered")

    if float(out["confMin"]) >= float(out["confMax"]):
        out["confMin"], out["confMax"] = 50.0, 95.0
        warnings.append("confMin/confMax were inverted or equal; reset to 50/95")

    return out, warnings


def _read_json(path: Path) -> Tuple[Dict[str, Any], Optional[str]]:
    try:
        payload = json.loads(path.read_text())
    except FileNotFoundError:
        return {}, None
    except (OSError, ValueError) as exc:
        return {}, f"{path}: unreadable ({exc})"
    keys = payload.get("keys", payload)
    if not isinstance(keys, Mapping):
        return {}, f"{path}: no key map found"
    cleaned = {name: value for name, value in keys.items() if not str(name).startswith("_")}
    cleaned["__league_weights__"] = payload.get("league_weights", {})
    return cleaned, None


def load_config(*paths: Path | str) -> BetConfig:
    """Load the shipped baseline plus any overrides, later files winning.

    Missing files are not an error -- a lane with no tuned config runs on the
    shipped defaults and stamps accordingly. Unreadable or malformed files are
    reported in ``warnings`` and contribute nothing.
    """

    values = dict(DEFAULTS)
    sources = {name: "default" for name in DEFAULTS}
    league_weights: Dict[str, float] = {}
    warnings: list = []

    for path in paths:
        resolved = Path(path)
        raw, problem = _read_json(resolved)
        if problem:
            warnings.append(problem)
        if not raw:
            continue
        for name, weight in (raw.pop("__league_weights__", {}) or {}).items():
            number = _as_float(weight)
            if number is None or number < 0.1 or number > 1.0:
                warnings.append(f"{resolved.name}: league weight {name} out of bounds, ignored")
                continue
            league_weights[str(name)] = number
        for name, raw_value in raw.items():
            key = KEYS.get(name)
            if key is None:
                warnings.append(f"{resolved.name}: unknown key {name!r} ignored")
                continue
            value, warning = coerce_value(key, raw_value)
            if warning:
                warnings.append(warning)
            values[name] = value
            sources[name] = f"file:{resolved.name}"

    values, invariant_warnings = enforce_invariants(values)
    warnings.extend(invariant_warnings)
    version = CONFIG_VERSION
    return BetConfig(version, values, sources, league_weights, tuple(warnings))


def default_config() -> BetConfig:
    return load_config(SHIPPED_PATH)


__all__ = [
    "CONFIG_VERSION",
    "DEFAULTS",
    "KEYS",
    "SHIPPED_PATH",
    "BetConfig",
    "ConfigKey",
    "coerce_value",
    "djb2",
    "default_config",
    "enforce_invariants",
    "load_config",
]
