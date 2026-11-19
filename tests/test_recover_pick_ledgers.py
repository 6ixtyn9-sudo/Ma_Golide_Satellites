from __future__ import annotations

import json
from pathlib import Path
import zipfile

import pytest

import scripts.recover_pick_ledgers as recovery
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
    # Ticket cards can contain audit-only shadow grades that depend on the
    # same artifact-only dated pick ledger. Plain explanatory prose is not a
    # grade and must not provoke an unnecessary/unsafe artifact requirement.
    (localdata / "auto_tickets_2026-11-16.txt").write_text(
        "The latest pending pre-tip observation remains mutable.\n"
    )
    (localdata / "auto_tickets_2026-11-17.txt").write_text("⏳ PENDING shadow\n")

    assert missing_historical_pick_ledgers(localdata, target_date="2026-11-19") == (
        "2026-11-17",
        "2026-11-18",
    )

    (localdata / "picks_2026-11-17.json").write_text("[]\n")
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


def test_artifact_urls_pages_until_it_finds_the_requested_lane(monkeypatch):
    calls: list[str] = []

    def fake_api(url, _token):
        calls.append(url)
        if url.endswith("page=1"):
            return {
                "total_count": 200,
                "artifacts": [
                    {
                        "name": "MaGolide-euroleague-1-1",
                        "archive_download_url": "https://api.github.com/artifacts/euroleague",
                    }
                ],
            }
        return {
            "total_count": 200,
            "artifacts": [
                {
                    "name": "MaGolide-nba-2-1",
                    "archive_download_url": "https://api.github.com/artifacts/nba",
                }
            ],
        }

    monkeypatch.setattr(recovery, "_api_json", fake_api)

    assert tuple(
        recovery._artifact_urls(
            "owner/repo", "nba", "token", max_artifacts=1
        )
    ) == ("https://api.github.com/artifacts/nba",)
    assert len(calls) == 2
    assert "page=1" in calls[0]
    assert "page=2" in calls[1]


def test_recovery_tries_later_artifacts_after_a_transient_download_failure(tmp_path, monkeypatch):
    localdata = tmp_path / "nba"
    localdata.mkdir()
    write_card(localdata, "2026-11-18", "grade=⏳ PENDING\n")
    archive = make_artifact(
        tmp_path / "working.zip",
        {"picks_2026-11-18.json": b'[{"pick_id":"recovered"}]\n'},
    )

    monkeypatch.setattr(
        recovery,
        "_artifact_urls",
        lambda *_args, **_kwargs: iter(("first", "second")),
    )

    def download(url, _token, destination):
        if url == "first":
            raise RuntimeError("GitHub Actions artifact download failed: ConnectionError")
        destination.write_bytes(archive.read_bytes())

    monkeypatch.setattr(recovery, "_download_artifact", download)

    report = recovery.recover_from_actions(
        localdata=localdata,
        target_date="2026-11-19",
        state_key="nba",
        repository="owner/repo",
        token="token",
        max_artifacts=2,
    )

    assert report.recovered_dates == ("2026-11-18",)
    assert report.inspected_artifacts == 1
    assert report.download_failures == 1
    assert json.loads((localdata / "picks_2026-11-18.json").read_text()) == [
        {"pick_id": "recovered"}
    ]


def test_missing_historical_pick_ledgers_rejects_invalid_target_date(tmp_path):
    with pytest.raises(ValueError, match="target_date"):
        missing_historical_pick_ledgers(tmp_path, target_date="not-a-date")
