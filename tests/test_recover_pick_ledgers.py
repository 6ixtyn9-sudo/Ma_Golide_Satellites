from __future__ import annotations

import json
from pathlib import Path
import zipfile

import pytest

from scripts.recover_pick_ledgers import (
    missing_historical_pick_ledgers,
    restore_pick_ledgers_from_zip,
)


def write_card(localdata: Path, day: str, text: str) -> None:
    (localdata / f"picks_{day}.txt").write_text(text)


def make_artifact(path: Path, members: dict[str, bytes]) -> Path:
    with zipfile.ZipFile(path, "w") as bundle:
        for name, content in members.items():
            bundle.writestr(name, content)
    return path


def test_missing_historical_pick_ledgers_only_requests_prior_pending_cards(tmp_path):
    localdata = tmp_path / "nba"
    localdata.mkdir()
    write_card(localdata, "2026-11-17", "grade=✅ WIN\n")
    write_card(localdata, "2026-11-18", "grade=⏳ PENDING\n")
    write_card(localdata, "2026-11-19", "grade=⏳ PENDING\n")

    assert missing_historical_pick_ledgers(localdata, target_date="2026-11-19") == (
        "2026-11-18",
    )

    (localdata / "picks_2026-11-18.json").write_text("[]\n")
    assert missing_historical_pick_ledgers(localdata, target_date="2026-11-19") == ()


def test_restore_pick_ledgers_restores_only_requested_valid_json_arrays(tmp_path):
    localdata = tmp_path / "ncaa_m"
    localdata.mkdir()
    archive = make_artifact(
        tmp_path / "artifact.zip",
        {
            "localdata/ncaa_m/picks_2026-11-15.json": b'[{"pick_id":"old"}]\n',
            "localdata/ncaa_m/picks_2026-11-18.json": b'{"not":"a list"}\n',
            "localdata/ncaa_m/picks_2026-11-19.json": b'[{"pick_id":"future"}]\n',
            "../../outside.json": b"[]\n",
        },
    )

    restored = restore_pick_ledgers_from_zip(
        archive,
        localdata=localdata,
        dates=("2026-11-15", "2026-11-18"),
    )

    assert restored == ("2026-11-15",)
    assert json.loads((localdata / "picks_2026-11-15.json").read_text()) == [
        {"pick_id": "old"}
    ]
    assert not (localdata / "picks_2026-11-18.json").exists()
    assert not (localdata / "picks_2026-11-19.json").exists()
    assert not (tmp_path / "outside.json").exists()


def test_restore_pick_ledgers_never_rewrites_an_existing_ledger(tmp_path):
    localdata = tmp_path / "nba"
    localdata.mkdir()
    existing = localdata / "picks_2026-11-18.json"
    existing.write_text('[{"pick_id":"existing"}]\n')
    archive = make_artifact(
        tmp_path / "artifact.zip",
        {"picks_2026-11-18.json": b'[{"pick_id":"artifact"}]\n'},
    )

    assert restore_pick_ledgers_from_zip(
        archive, localdata=localdata, dates=("2026-11-18",)
    ) == ()
    assert json.loads(existing.read_text()) == [{"pick_id": "existing"}]


def test_missing_historical_pick_ledgers_rejects_invalid_target_date(tmp_path):
    with pytest.raises(ValueError, match="target_date"):
        missing_historical_pick_ledgers(tmp_path, target_date="not-a-date")
