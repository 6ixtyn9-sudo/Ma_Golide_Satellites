"""Config as data: bounds, invariants, stamps and the ledger.

The retired system read ``Config_Tier1`` as key/value rows through ``KEY_MAP``,
range-checked every key, reset invalid input to the default, forced the tier ladder
monotonic, and stamped each slip row with ``CFG_`` + djb2 of the canonical config.
These tests pin that contract: a number out of bounds is refused loudly, a config
that could not produce a tier cannot exist, and two different configurations can
never share a stamp.
"""

from __future__ import annotations

import json

import pytest

from forecasting.bet_config import (
    CONFIG_VERSION,
    DEFAULTS,
    KEYS,
    SHIPPED_PATH,
    coerce_value,
    djb2,
    default_config,
    enforce_invariants,
    load_config,
)
from forecasting.config_ledger import stamp_for, record, stamps


# -- the shipped file is the documented baseline ----------------------------


def test_the_shipped_file_matches_the_code_defaults():
    """The JSON is documentation; a drifting file would lie to a reader."""

    payload = json.loads(SHIPPED_PATH.read_text())
    assert payload["version"] == CONFIG_VERSION
    assert payload["keys"] == DEFAULTS


def test_every_key_carries_bounds_and_a_section():
    for name, key in KEYS.items():
        assert key.name == name
        assert key.minimum <= key.default <= key.maximum
        assert key.kind in {"float", "int", "bool"}
        assert key.section


# -- one key at a time: the sheet loader's contract -------------------------


def test_a_value_outside_its_bounds_falls_back_to_the_default_and_says_so():
    key = KEYS["bankerThreshold"]
    value, warning = coerce_value(key, 140)

    assert value == key.default == 60.0
    assert "outside" in warning


def test_an_unparseable_value_falls_back_to_the_default_and_says_so():
    value, warning = coerce_value(KEYS["ftouSigmaFloor"], "not a number")

    assert value == 15.0
    assert "not a number" in warning


@pytest.mark.parametrize("raw", ["TRUE", "yes", 1, True, " on ", "enabled"])
def test_the_boolean_vocabulary_is_the_sheets(raw):
    assert coerce_value(KEYS["enableRobbers"], raw)[0] is True


@pytest.mark.parametrize("raw", ["FALSE", "no", 0, False, "off"])
def test_anything_else_is_false(raw):
    assert coerce_value(KEYS["enableRobbers"], raw)[0] is False


def test_an_empty_cell_means_unset_not_false():
    """A blank sheet cell fell back to the default; that is not a "no"."""

    value, warning = coerce_value(KEYS["enableRobbers"], "")

    assert value is True
    assert warning is None


def test_an_integer_key_is_rounded_to_an_integer():
    assert coerce_value(KEYS["ftouMinSamples"], 3.9)[0] == 3
    assert isinstance(coerce_value(KEYS["ftouMinSamples"], 3.9)[0], int)


# -- relationships between keys ---------------------------------------------


def test_the_tier_ladder_cannot_invert():
    values, warnings = enforce_invariants(
        {**DEFAULTS, "tier_weak_min_score": 70.0, "tier_medium_min_score": 60.0}
    )

    assert values["tier_medium_min_score"] > values["tier_weak_min_score"]
    assert warnings


def test_the_strong_floor_always_clears_the_medium_floor():
    values, _ = enforce_invariants(
        {**DEFAULTS, "tier_medium_min_score": 90.0, "tier_strong_min_score": 80.0}
    )

    assert values["tier_strong_min_score"] > values["tier_medium_min_score"]


def test_the_risky_threshold_cannot_sit_above_the_weak_floor():
    values, warnings = enforce_invariants({**DEFAULTS, "threshold": 80.0})

    assert values["threshold"] == values["tier_weak_min_score"]
    assert any("RISKY" in warning for warning in warnings)


def test_inverted_confidence_bounds_reset_together():
    values, _ = enforce_invariants({**DEFAULTS, "confMin": 90.0, "confMax": 60.0})

    assert (values["confMin"], values["confMax"]) == (50.0, 95.0)


# -- loading ----------------------------------------------------------------


def test_a_missing_override_file_is_not_an_error(tmp_path):
    config = load_config(SHIPPED_PATH, tmp_path / "absent.json")

    assert config.warnings == ()
    assert config.get("bankerThreshold") == 60.0


def test_a_malformed_file_is_reported_and_contributes_nothing(tmp_path):
    bad = tmp_path / "bad.json"
    bad.write_text("{not json")

    config = load_config(SHIPPED_PATH, bad)

    assert any("unreadable" in warning for warning in config.warnings)
    assert config.get("bankerThreshold") == 60.0


def test_overrides_win_and_are_attributed(tmp_path):
    override = tmp_path / "tuned.json"
    override.write_text(json.dumps({"keys": {"bankerThreshold": 65}}))

    config = load_config(SHIPPED_PATH, override)

    assert config.get("bankerThreshold") == 65.0
    assert config.provenance("bankerThreshold") == "file:tuned.json"
    # The shipped baseline is a file too, and says so: provenance is about where a
    # number was read, not whether it is "the default".
    assert config.provenance("robberMinConf") == "file:magolide_bets.json"

    override_only = load_config(override)
    assert override_only.provenance("robberMinConf") == "default"


def test_an_unknown_key_is_ignored_with_a_warning(tmp_path):
    override = tmp_path / "odd.json"
    override.write_text(json.dumps({"keys": {"noSuchGate": 1}}))

    config = load_config(SHIPPED_PATH, override)

    assert any("noSuchGate" in warning for warning in config.warnings)


def test_league_weights_are_loaded_and_bounded(tmp_path):
    override = tmp_path / "tuned.json"
    override.write_text(json.dumps({"keys": {}, "league_weights": {"NKL": 0.7, "BAD": 3.0}}))

    config = load_config(SHIPPED_PATH, override)

    assert config.league_weight("NKL") == 0.7
    assert config.league_weight("BAD") == 1.0  # out of bounds, ignored
    assert config.league_weight("UNSEEN") == 1.0


# -- stamps -----------------------------------------------------------------


def test_a_stamp_is_the_ledgers_cfg_prefix_and_djb2():
    config = default_config()

    assert config.stamp_id.startswith("CFG_")
    assert config.stamp_id == "CFG_" + djb2(config.canonical_json())
    assert len(config.stamp_id) == 12


def test_the_same_configuration_stamps_identically():
    assert default_config().stamp_id == default_config().stamp_id
    assert load_config().stamp_id == load_config(SHIPPED_PATH).stamp_id


def test_one_changed_gate_produces_a_new_stamp():
    base = default_config()
    changed = base.with_overrides({"bankerThreshold": 61.0})

    assert changed.stamp_id != base.stamp_id


def test_a_league_discount_is_part_of_the_stamp():
    base = default_config()
    discounted = base.with_leagues({"NKL": 0.7})

    assert discounted.stamp_id != base.stamp_id
    assert discounted.highlights if hasattr(discounted, "highlights") else True


def test_a_stamp_does_not_depend_on_key_order():
    config = default_config()
    reordered = type(config)(
        config.version,
        dict(reversed(list(config.values.items()))),
        dict(reversed(list(config.sources.items()))),
        config.league_weights,
        config.warnings,
    )

    assert reordered.stamp_id == config.stamp_id


# -- the ledger -------------------------------------------------------------


def test_the_ledger_keeps_one_row_per_stamp(tmp_path):
    ledger = tmp_path / "config_ledger.jsonl"
    config = default_config()

    record(config, ledger, note="first")
    record(config, ledger, note="second")  # same config, same stamp

    rows = stamps(ledger)
    assert len(rows) == 1
    assert rows[0].stamp_id == config.stamp_id
    assert rows[0].note == "first"


def test_a_changed_config_adds_a_row(tmp_path):
    ledger = tmp_path / "config_ledger.jsonl"

    record(default_config(), ledger)
    tuned = default_config().with_overrides({"bankerThreshold": 61.0})
    record(tuned, ledger, note="tuned")

    rows = stamps(ledger)
    assert len(rows) == 2
    assert rows[-1].stamp_id == tuned.stamp_id
    assert rows[-1].highlights["bankerThreshold"] == 61.0


def test_the_ledger_records_the_version_and_build_time(tmp_path):
    ledger = tmp_path / "config_ledger.jsonl"
    stamp = stamp_for(default_config())

    record(default_config(), ledger)

    assert stamp.version == CONFIG_VERSION
    assert stamp.built_at.endswith("+00:00")
    assert stamps(ledger)[0].version == CONFIG_VERSION


def test_a_missing_ledger_reads_as_empty(tmp_path):
    assert stamps(tmp_path / "nothing.jsonl") == []


def test_an_unreadable_line_is_skipped_not_fatal(tmp_path):
    ledger = tmp_path / "config_ledger.jsonl"
    ledger.write_text("{broken\n" + json.dumps(stamp_for(default_config()).as_dict()) + "\n")

    rows = stamps(ledger)

    assert len(rows) == 1
    assert rows[0].stamp_id == default_config().stamp_id
