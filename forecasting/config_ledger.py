"""The config ledger: one row per configuration stamp, exactly like the sheet's.

Why this exists
---------------
``Config_Ledger_Satellite.gs`` stamped every row of ``Bet_Slips`` with
``"CFG_" + djb2(canonical config JSON)`` and kept a ledger sheet with one row per
stamp: ``stamp_id``, ``version``, ``built_at``, the active leagues, and the key
numbers that defined that era of the system. Two runs on the same config shared a
stamp; a single changed key produced a new one. That is what made "why did the
system bet differently last Tuesday?" an answerable question.

This module keeps the same contract over a JSONL file:

* :class:`ConfigStamp` is the ledger row.
* :func:`record` appends a stamp the first time it is seen and is a no-op
  afterwards -- the "ensure one row per stamp" behaviour of ``_ensureLedgerRow``.
* :func:`stamps` reads the ledger back, newest last, for auditing.

A stamp is derived from :class:`forecasting.bet_config.BetConfig`, so the ledger can
never disagree with the configuration a card was rendered under: both compute the
same hash over the same key-sorted JSON.
"""

from __future__ import annotations

import json
from dataclasses import asdict, dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Mapping, Optional

from forecasting.bet_config import BetConfig

#: Ledger rows land here unless a caller names another path. Under ``localdata``
#: with the rest of the run state: a ledger is an audit trail, not source code.
DEFAULT_LEDGER_PATH = Path("localdata/config_ledger.jsonl")

#: The subset of keys worth showing at a glance in a ledger row, mirroring the
#: sheet's grouped settings blob.
HIGHLIGHT_KEYS = (
    "bankerThreshold",
    "robberMinConf",
    "firstHalfMinConf",
    "ftOuMinConf",
    "threshold",
    "tier_weak_min_score",
    "tier_medium_min_score",
    "tier_strong_min_score",
    "confMin",
    "confMax",
)


@dataclass(frozen=True)
class ConfigStamp:
    """One era of the configuration."""

    stamp_id: str
    version: str
    built_at: str
    highlights: Mapping[str, Any] = field(default_factory=dict)
    keys: Mapping[str, Any] = field(default_factory=dict)
    note: str = ""

    def as_dict(self) -> Dict[str, Any]:
        return asdict(self)


def stamp_for(config: BetConfig, *, note: str = "", built_at: Optional[str] = None) -> ConfigStamp:
    """Build the ledger row a card should carry alongside its rows."""

    return ConfigStamp(
        stamp_id=config.stamp_id,
        version=config.version,
        built_at=built_at or datetime.now(timezone.utc).isoformat(timespec="seconds"),
        highlights={name: config.get(name) for name in HIGHLIGHT_KEYS},
        keys=config.as_dict(),
        note=note,
    )


def _existing_ids(path: Path) -> set:
    if not path.exists():
        return set()
    ids = set()
    try:
        for line in path.read_text().splitlines():
            line = line.strip()
            if not line:
                continue
            try:
                ids.add(json.loads(line)["stamp_id"])
            except (ValueError, KeyError):
                continue
    except OSError:
        return set()
    return ids


def record(
    config: BetConfig,
    path: Path | str = DEFAULT_LEDGER_PATH,
    *,
    note: str = "",
) -> ConfigStamp:
    """Append this configuration's stamp if the ledger has not seen it.

    Returns the stamp either way, so a caller can put ``stamp.stamp_id`` on the rows
    it is about to write without caring whether this is the first time.
    """

    target = Path(path)
    stamp = stamp_for(config, note=note)
    if stamp.stamp_id in _existing_ids(target):
        return stamp
    target.parent.mkdir(parents=True, exist_ok=True)
    with target.open("a", encoding="utf-8") as handle:
        handle.write(json.dumps(stamp.as_dict(), sort_keys=True) + "\n")
    return stamp


def stamps(path: Path | str = DEFAULT_LEDGER_PATH) -> List[ConfigStamp]:
    """Every stamp in the ledger, oldest first. Unreadable lines are skipped."""

    target = Path(path)
    if not target.exists():
        return []
    out: List[ConfigStamp] = []
    for line in target.read_text().splitlines():
        line = line.strip()
        if not line:
            continue
        try:
            payload = json.loads(line)
            out.append(
                ConfigStamp(
                    stamp_id=payload["stamp_id"],
                    version=payload.get("version", ""),
                    built_at=payload.get("built_at", ""),
                    highlights=payload.get("highlights", {}),
                    keys=payload.get("keys", {}),
                    note=payload.get("note", ""),
                )
            )
        except (ValueError, KeyError):
            continue
    return out


__all__ = [
    "DEFAULT_LEDGER_PATH",
    "HIGHLIGHT_KEYS",
    "ConfigStamp",
    "record",
    "stamp_for",
    "stamps",
]
