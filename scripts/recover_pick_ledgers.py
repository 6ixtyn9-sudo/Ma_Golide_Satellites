#!/usr/bin/env python3
"""Restore missing historical pick ledgers from retained Actions artifacts.

Dated structured pick ledgers deliberately remain cache/artifact-only: they are
not published to Git.  A cold or evicted cache must therefore recover any
historical ledger whose public text card still contains pending rows.  This
utility downloads only the owning lane's retained workflow artifacts, restores
only the required ``picks_YYYY-MM-DD.json`` files, and leaves all Git-durable
cards and decisions untouched.

It is designed for GitHub Actions.  The pure discovery/extraction helpers are
also deliberately testable with local zip files.
"""

from __future__ import annotations

import argparse
from dataclasses import dataclass
from datetime import date
import json
import os
from pathlib import Path, PurePosixPath
import re
import tempfile
from typing import Any, Iterable
import requests
import zipfile


DATE_RE = re.compile(r"^picks_(\d{4}-\d{2}-\d{2})\.txt$")
LEDGER_RE = re.compile(r"^picks_(\d{4}-\d{2}-\d{2})\.json$")
ARTIFACT_NAME_RE = re.compile(r"^MaGolide-(?:lifecycle-)?([a-z0-9_]+)-\d+-\d+$")
MAX_LEDGER_BYTES = 25 * 1024 * 1024


@dataclass(frozen=True)
class RecoveryReport:
    """One bounded recovery attempt, safe to print in workflow logs."""

    needed_dates: tuple[str, ...]
    recovered_dates: tuple[str, ...]
    inspected_artifacts: int

    @property
    def missing_dates(self) -> tuple[str, ...]:
        return tuple(date for date in self.needed_dates if date not in self.recovered_dates)

    def as_dict(self) -> dict[str, Any]:
        return {
            "needed_dates": list(self.needed_dates),
            "recovered_dates": list(self.recovered_dates),
            "missing_dates": list(self.missing_dates),
            "inspected_artifacts": self.inspected_artifacts,
        }


def missing_historical_pick_ledgers(localdata: Path, *, target_date: str) -> tuple[str, ...]:
    """Return earlier card dates still showing pending data without raw evidence."""

    try:
        target = date.fromisoformat(target_date)
    except ValueError as exc:
        raise ValueError(f"target_date must be ISO yyyy-mm-dd: {target_date!r}") from exc
    needed: list[str] = []
    for card in sorted(Path(localdata).glob("picks_????-??-??.txt")):
        match = DATE_RE.fullmatch(card.name)
        if match is None:
            continue
        card_date = match.group(1)
        if date.fromisoformat(card_date) >= target:
            continue
        try:
            pending = "PENDING" in card.read_text(encoding="utf-8")
        except OSError as exc:
            raise ValueError(f"cannot read dated pick card: {card}") from exc
        if pending and not (Path(localdata) / f"picks_{card_date}.json").is_file():
            needed.append(card_date)
    return tuple(needed)


def _safe_ledger_member(member_name: str, *, wanted: set[str]) -> tuple[str, PurePosixPath] | None:
    """Return one wanted ledger member, rejecting path traversal and junk."""

    path = PurePosixPath(member_name)
    if path.is_absolute() or ".." in path.parts or path.name != member_name.split("/")[-1]:
        return None
    match = LEDGER_RE.fullmatch(path.name)
    if match is None or match.group(1) not in wanted:
        return None
    return match.group(1), path


def _atomic_write(path: Path, raw: bytes) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(path.name + ".tmp")
    try:
        temporary.write_bytes(raw)
        os.replace(temporary, path)
    finally:
        if temporary.exists():
            temporary.unlink()


def restore_pick_ledgers_from_zip(
    archive: Path,
    *,
    localdata: Path,
    dates: Iterable[str],
) -> tuple[str, ...]:
    """Restore missing, valid requested ledgers from one trusted artifact zip."""

    wanted = set(dates)
    restored: list[str] = []
    if not wanted:
        return ()
    try:
        bundle = zipfile.ZipFile(archive)
    except (OSError, zipfile.BadZipFile) as exc:
        raise ValueError(f"cannot read workflow artifact archive {archive}: {exc}") from exc
    with bundle:
        for info in bundle.infolist():
            selected = _safe_ledger_member(info.filename, wanted=wanted)
            if selected is None:
                continue
            ledger_date, _path = selected
            destination = Path(localdata) / f"picks_{ledger_date}.json"
            if destination.exists():
                wanted.remove(ledger_date)
                continue
            if info.is_dir() or info.file_size > MAX_LEDGER_BYTES:
                continue
            raw = bundle.read(info)
            if len(raw) > MAX_LEDGER_BYTES:
                continue
            try:
                payload = json.loads(raw)
            except (TypeError, ValueError, json.JSONDecodeError):
                continue
            if not isinstance(payload, list):
                continue
            _atomic_write(destination, raw)
            restored.append(ledger_date)
            wanted.remove(ledger_date)
    return tuple(sorted(restored))


def _api_json(url: str, token: str) -> dict[str, Any]:
    headers = {
        "Accept": "application/vnd.github+json",
        "Authorization": f"Bearer {token}",
        "User-Agent": "ma-golide-pick-ledger-recovery",
    }
    try:
        response = requests.get(url, headers=headers, timeout=30)
        response.raise_for_status()
    except requests.RequestException as exc:
        raise RuntimeError(f"GitHub Actions artifact lookup failed: {type(exc).__name__}") from exc
    try:
        decoded = response.json()
    except ValueError as exc:
        raise RuntimeError("GitHub Actions artifact lookup returned invalid JSON") from exc
    if not isinstance(decoded, dict):
        raise RuntimeError("GitHub Actions artifact lookup returned an invalid payload")
    return decoded


def _download_artifact(url: str, token: str, destination: Path) -> None:
    headers = {
        "Accept": "application/vnd.github+json",
        "Authorization": f"Bearer {token}",
        "User-Agent": "ma-golide-pick-ledger-recovery",
    }
    try:
        # requests removes Authorization when a GitHub API redirect crosses to
        # the signed storage host, so the workflow token is never forwarded.
        with requests.get(url, headers=headers, stream=True, timeout=180) as response:
            response.raise_for_status()
            with destination.open("wb") as handle:
                for chunk in response.iter_content(chunk_size=1024 * 1024):
                    if chunk:
                        handle.write(chunk)
    except (requests.RequestException, OSError) as exc:
        raise RuntimeError(f"GitHub Actions artifact download failed: {type(exc).__name__}") from exc


def _artifact_urls(repository: str, state_key: str, token: str) -> list[str]:
    if not re.fullmatch(r"[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+", repository):
        raise ValueError("repository must be owner/name")
    payload = _api_json(
        f"https://api.github.com/repos/{repository}/actions/artifacts?per_page=100",
        token,
    )
    artifacts = payload.get("artifacts")
    if not isinstance(artifacts, list):
        raise RuntimeError("GitHub Actions artifact lookup omitted artifacts")
    urls: list[str] = []
    for artifact in artifacts:
        if not isinstance(artifact, dict) or artifact.get("expired") is True:
            continue
        name = str(artifact.get("name") or "")
        match = ARTIFACT_NAME_RE.fullmatch(name)
        if match is None or match.group(1) != state_key:
            continue
        url = artifact.get("archive_download_url")
        if isinstance(url, str) and url.startswith("https://api.github.com/"):
            urls.append(url)
    return urls


def recover_from_actions(
    *,
    localdata: Path,
    target_date: str,
    state_key: str,
    repository: str,
    token: str,
    max_artifacts: int,
) -> RecoveryReport:
    """Recover missing historical ledgers from newest retained lane artifacts."""

    if not re.fullmatch(r"[a-z0-9_]+", state_key):
        raise ValueError("state_key must be a registered lowercase state key")
    if max_artifacts < 1:
        raise ValueError("max_artifacts must be positive")
    needed = missing_historical_pick_ledgers(localdata, target_date=target_date)
    if not needed:
        return RecoveryReport(needed, (), 0)
    if not token:
        raise RuntimeError("GITHUB_TOKEN is required to recover artifact-only pick ledgers")

    recovered: set[str] = set()
    inspected = 0
    with tempfile.TemporaryDirectory(prefix="magolide-pick-ledgers-") as temp_dir:
        for index, url in enumerate(_artifact_urls(repository, state_key, token)):
            if index >= max_artifacts or set(needed).issubset(recovered):
                break
            archive = Path(temp_dir) / f"artifact-{index}.zip"
            _download_artifact(url, token, archive)
            inspected += 1
            recovered.update(
                restore_pick_ledgers_from_zip(
                    archive,
                    localdata=localdata,
                    dates=set(needed) - recovered,
                )
            )
    report = RecoveryReport(needed, tuple(sorted(recovered)), inspected)
    if report.missing_dates:
        missing = ", ".join(report.missing_dates)
        raise RuntimeError(
            "historical pending pick cards lack retained artifact-only ledgers: " + missing
        )
    return report


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--localdata", type=Path, required=True)
    parser.add_argument("--target-date", required=True)
    parser.add_argument("--state-key", required=True)
    parser.add_argument("--repository", required=True)
    parser.add_argument("--max-artifacts", type=int, default=8)
    args = parser.parse_args()
    report = recover_from_actions(
        localdata=args.localdata,
        target_date=args.target_date,
        state_key=args.state_key,
        repository=args.repository,
        token=os.environ.get("GITHUB_TOKEN", ""),
        max_artifacts=args.max_artifacts,
    )
    print(json.dumps(report.as_dict(), sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
