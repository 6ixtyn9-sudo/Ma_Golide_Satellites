#!/usr/bin/env python3
"""Build and restore compact, verified GitHub Release evidence checkpoints.

Git remains the home for source code, current compact operator state, and tiny
release manifests. Full canonical CSV.GZ recoveries plus verbose run evidence
are packed as one deterministic, checksummed archive per league and uploaded as
an asset on the ``magolide-state-current`` GitHub Release by the workflow.

The archive never contains a DuckDB database, raw provider pages, credentials,
or arbitrary workspace files. A cold runner restores *only* its canonical CSV
checkpoint before the normal cache and live pipeline run.
"""

from __future__ import annotations

import argparse
import csv
from dataclasses import dataclass
import gzip
import hashlib
import io
import json
import os
from pathlib import Path
import re
import sys
import tarfile
from typing import Any, Iterable

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from ingestion.league_registry import SPECS  # noqa: E402

ARCHIVE_SCHEMA = "magolide.release_archive"
CONTENT_MANIFEST_NAME = "archive_content_manifest.json"
DEFAULT_RELEASE_TAG = "magolide-state-current"
STATE_KEYS = frozenset(spec.state_key for spec in SPECS)
# Tables whose loss means the run stopped being able to see history it already
# had. ``pipeline_runs`` is excluded because retention prunes it by design.
REGRESSION_GUARDED_TABLES = ("games", "predictions", "book_snapshots", "team_aliases")
DATE = r"\d{4}-\d{2}-\d{2}"


@dataclass(frozen=True)
class ArchiveEntry:
    archive_path: str
    source_path: Path
    role: str

    def manifest_dict(self) -> dict[str, Any]:
        raw = self.source_path.read_bytes()
        return {
            "path": self.archive_path,
            "role": self.role,
            "bytes": len(raw),
            "sha256": hashlib.sha256(raw).hexdigest(),
        }


def _sha256_path(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _json_bytes(value: object) -> bytes:
    return (json.dumps(value, indent=2, sort_keys=True) + "\n").encode()


def _safe_relative(path: Path, root: Path) -> str:
    relative = path.relative_to(root)
    if not relative.parts or any(part in {"", ".", ".."} for part in relative.parts):
        raise ValueError(f"unsafe archive path: {path}")
    return relative.as_posix()


def _regular_files(root: Path) -> Iterable[Path]:
    if not root.is_dir():
        return ()
    return (path for path in sorted(root.rglob("*")) if path.is_file() and not path.is_symlink())


def _artifact_state_dirs(artifacts: Path) -> dict[str, Path]:
    """Find one successful artifact state directory per registered league."""

    selected: dict[str, Path] = {}
    for localdata in sorted(artifacts.rglob("localdata")):
        if not localdata.is_dir():
            continue
        for state_dir in sorted(localdata.iterdir()):
            if not state_dir.is_dir() or state_dir.name not in STATE_KEYS:
                continue
            # Only a successful full lane has the canonical CSV materialization.
            csv = state_dir / "csv"
            if not any(csv.glob("*.csv.gz")):
                continue
            existing = selected.get(state_dir.name)
            if existing is not None:
                raise ValueError(
                    "multiple successful archive sources for "
                    f"{state_dir.name}: {existing} and {state_dir}"
                )
            selected[state_dir.name] = state_dir
    return selected


def _archive_artifact_file(relative: str) -> bool:
    """Whitelist only derived/canonical evidence emitted by the workflow."""

    name = Path(relative).name
    if relative.startswith("csv/"):
        return name.endswith(".csv.gz")
    if relative.startswith("ledgers/"):
        return name.endswith(".json")
    if relative.startswith("reports/"):
        return name.endswith(".json") and not name.startswith("daily_")
    if name.startswith("assay_") and name.endswith(".json"):
        return True
    if re.fullmatch(rf"picks_{DATE}(?:_lifecycle)?\.json", name):
        return True
    if re.fullmatch(rf"picks_{DATE}\.txt", name):
        return True
    if re.fullmatch(rf"picks_audit_{DATE}\.md", name):
        return True
    if re.fullmatch(r"picks_lifecycle_archive_\d{4}-\d{2}\.json", name):
        return True
    if name in {
        "pick_lifecycle_state.json",
        "auto_tickets_state.json",
        "auto_tickets_performance.json",
        "auto_tickets_performance.txt",
        "history_receipts.json",
    }:
        return True
    if re.fullmatch(rf"coverage_gap_{DATE}\.json", name):
        return True
    if re.fullmatch(rf"auto_tickets_{DATE}(?:_(?:decision|withdrawn))?\.(?:json|txt)", name):
        return True
    # Explicitly do not archive unknown payloads if artifact upload patterns
    # are widened later. This prevents accidental raw-page/credential capture.
    return False


def _operator_state_file(relative: str) -> bool:
    """Store Git-visible operator evidence beside canonical checkpoint data."""

    return _archive_artifact_file(relative) and not relative.startswith(("csv/", "ledgers/", "reports/"))


def _entries_for_state(*, state: str, artifact_dir: Path, repo: Path) -> list[ArchiveEntry]:
    entries: list[ArchiveEntry] = []
    canonical = artifact_dir / "csv"
    for path in _regular_files(canonical):
        relative = _safe_relative(path, artifact_dir)
        if not _archive_artifact_file(relative):
            continue
        entries.append(
            ArchiveEntry(
                archive_path=f"canonical/{state}/{relative}", source_path=path, role="canonical_csv",
            )
        )
    if not entries:
        raise ValueError(f"archive input has no canonical CSV.GZ files: {artifact_dir}")

    for path in _regular_files(artifact_dir):
        relative = _safe_relative(path, artifact_dir)
        if relative.startswith("csv/") or not _archive_artifact_file(relative):
            continue
        entries.append(
            ArchiveEntry(
                archive_path=f"artifact/{state}/{relative}", source_path=path, role="run_evidence",
            )
        )

    operator_root = repo / "localdata" / state
    for path in _regular_files(operator_root):
        relative = _safe_relative(path, operator_root)
        if not _operator_state_file(relative):
            continue
        entries.append(
            ArchiveEntry(
                archive_path=f"operator/{state}/{relative}", source_path=path, role="operator_state",
            )
        )

    archive_paths = [entry.archive_path for entry in entries]
    if len(archive_paths) != len(set(archive_paths)):
        raise ValueError(f"archive input contains duplicate destination paths for {state}")
    return sorted(entries, key=lambda entry: entry.archive_path)


def _csv_row_count(path: Path) -> int:
    """Count data rows in a canonical CSV.GZ, ignoring the header."""

    previous = csv.field_size_limit()
    limit = sys.maxsize
    while True:
        try:
            csv.field_size_limit(limit)
            break
        except OverflowError:  # pragma: no cover - 32-bit platforms
            limit //= 10
    try:
        with gzip.open(path, "rt", newline="") as handle:
            reader = csv.reader(handle)
            next(reader, None)
            return sum(1 for _ in reader)
    finally:
        csv.field_size_limit(previous)


def _canonical_row_counts(canonical_entries: list[ArchiveEntry]) -> dict[str, int]:
    counts: dict[str, int] = {}
    for entry in canonical_entries:
        name = Path(entry.archive_path).name
        if name.endswith(".csv.gz"):
            counts[name[: -len(".csv.gz")]] = _csv_row_count(entry.source_path)
    return counts


def _checkpoint_regressions(
    previous: dict[str, Any], current: dict[str, int]
) -> dict[str, dict[str, int]]:
    """Tables where the new checkpoint holds fewer rows than the committed one."""

    previous_counts = previous.get("canonical_row_counts") or {}
    regressions: dict[str, dict[str, int]] = {}
    for table in REGRESSION_GUARDED_TABLES:
        before = previous_counts.get(table)
        after = current.get(table)
        if not isinstance(before, int) or after is None:
            continue
        if after < before:
            regressions[table] = {"committed": before, "new": after}
    return regressions


def _write_deterministic_archive(path: Path, *, content_manifest: dict[str, Any], entries: list[ArchiveEntry]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(path.name + ".tmp")
    manifest_bytes = _json_bytes(content_manifest)
    with temporary.open("wb") as raw_handle:
        with gzip.GzipFile(fileobj=raw_handle, mode="wb", compresslevel=1, mtime=0) as gzip_handle:
            with tarfile.open(fileobj=gzip_handle, mode="w") as archive:
                manifest_info = tarfile.TarInfo(CONTENT_MANIFEST_NAME)
                manifest_info.size = len(manifest_bytes)
                manifest_info.mode = 0o644
                manifest_info.mtime = 0
                archive.addfile(manifest_info, io.BytesIO(manifest_bytes))
                for entry in entries:
                    info = tarfile.TarInfo(entry.archive_path)
                    info.size = entry.source_path.stat().st_size
                    info.mode = 0o644
                    info.mtime = 0
                    with entry.source_path.open("rb") as source:
                        archive.addfile(info, source)
    os.replace(temporary, path)


def build_release_archives(
    *,
    artifacts: Path,
    repo: Path,
    output: Path,
    release_tag: str = DEFAULT_RELEASE_TAG,
    allow_checkpoint_regression: bool = False,
) -> dict[str, Any]:
    """Create a deterministic canonical checkpoint for each successful lane."""

    artifacts = artifacts.resolve()
    repo = repo.resolve()
    output = output.resolve()
    if not artifacts.is_dir():
        raise ValueError(f"artifact directory does not exist: {artifacts}")
    state_dirs = _artifact_state_dirs(artifacts)
    records: list[dict[str, Any]] = []
    for state, artifact_dir in sorted(state_dirs.items()):
        entries = _entries_for_state(state=state, artifact_dir=artifact_dir, repo=repo)
        canonical_entries = [entry for entry in entries if entry.role == "canonical_csv"]
        row_counts = _canonical_row_counts(canonical_entries)
        state_manifest = repo / "localdata" / state / "archive_manifest.json"
        # A checkpoint is built from whatever CSV export the run produced, and
        # the asset upload plus manifest write are unconditional. Without a
        # comparison, a thin or freshly rebuilt warehouse silently re-packs
        # itself as the new canonical state, which is how a 234-prediction
        # corpus was replaced by 40. Check before anything is written so a
        # refusal leaves the published release and manifest intact.
        if state_manifest.exists() and not allow_checkpoint_regression:
            try:
                previous_manifest = json.loads(state_manifest.read_text())
            except (OSError, json.JSONDecodeError) as exc:
                raise ValueError(
                    f"invalid committed checkpoint manifest: {state_manifest}"
                ) from exc
            regressions = _checkpoint_regressions(previous_manifest, row_counts)
            if regressions:
                raise ValueError(
                    f"checkpoint_regression: refusing to replace the committed {state} "
                    f"checkpoint with a thinner one: "
                    f"{json.dumps(regressions, sort_keys=True)}; pass "
                    "--allow-checkpoint-regression once the loss is understood"
                )
        content_manifest = {
            "schema_version": ARCHIVE_SCHEMA,
            "state_key": state,
            "release_tag": release_tag,
            "entries": [entry.manifest_dict() for entry in entries],
        }
        asset_name = f"magolide-state-{state}.tar.gz"
        archive_path = output / asset_name
        _write_deterministic_archive(archive_path, content_manifest=content_manifest, entries=entries)
        archive_sha256 = _sha256_path(archive_path)
        manifest = {
            "schema_version": ARCHIVE_SCHEMA,
            "state_key": state,
            "release_tag": release_tag,
            "asset_name": asset_name,
            "archive_bytes": archive_path.stat().st_size,
            "archive_sha256": archive_sha256,
            "canonical_csv_files": len(canonical_entries),
            "canonical_row_counts": dict(sorted(row_counts.items())),
            "content_manifest_sha256": hashlib.sha256(_json_bytes(content_manifest)).hexdigest(),
        }
        state_manifest.parent.mkdir(parents=True, exist_ok=True)
        state_manifest.write_bytes(_json_bytes(manifest))
        records.append(manifest)
    index = {
        "schema_version": ARCHIVE_SCHEMA,
        "release_tag": release_tag,
        "archives": records,
    }
    output.mkdir(parents=True, exist_ok=True)
    (output / "release-index.json").write_bytes(_json_bytes(index))
    return index


def _load_json(path: Path, *, description: str) -> dict[str, Any]:
    try:
        data = json.loads(path.read_text())
    except (OSError, json.JSONDecodeError) as exc:
        raise ValueError(f"invalid {description}: {path}") from exc
    if not isinstance(data, dict):
        raise ValueError(f"invalid {description}: {path}")
    return data


def _safe_tar_member(name: str) -> bool:
    path = Path(name)
    return bool(name) and not path.is_absolute() and all(part not in {"", ".", ".."} for part in path.parts)


def restore_canonical_checkpoint(
    *, archive: Path, state_key: str, destination: Path, expected_manifest: Path | None = None
) -> dict[str, Any]:
    """Verify and restore only a state archive's canonical CSV.GZ payload."""

    if state_key not in STATE_KEYS:
        raise ValueError(f"unknown registered state key: {state_key}")
    if expected_manifest is not None and expected_manifest.exists():
        expected = _load_json(expected_manifest, description="release archive manifest")
        if expected.get("schema_version") != ARCHIVE_SCHEMA or expected.get("state_key") != state_key:
            raise ValueError(f"archive manifest identity mismatch: {expected_manifest}")
        if expected.get("archive_sha256") != _sha256_path(archive):
            raise ValueError(f"archive checksum mismatch: {archive}")

    try:
        with tarfile.open(archive, mode="r:gz") as bundle:
            try:
                manifest_member = bundle.getmember(CONTENT_MANIFEST_NAME)
            except KeyError as exc:
                raise ValueError(f"archive has no content manifest: {archive}") from exc
            manifest_handle = bundle.extractfile(manifest_member)
            if manifest_handle is None:
                raise ValueError(f"archive has unreadable content manifest: {archive}")
            manifest = json.loads(manifest_handle.read())
            if (
                not isinstance(manifest, dict)
                or manifest.get("schema_version") != ARCHIVE_SCHEMA
                or manifest.get("state_key") != state_key
                or not isinstance(manifest.get("entries"), list)
            ):
                raise ValueError(f"archive content manifest identity mismatch: {archive}")
            expected_entries = {
                str(entry.get("path")): entry
                for entry in manifest["entries"]
                if isinstance(entry, dict) and isinstance(entry.get("path"), str)
            }
            if len(expected_entries) != len(manifest["entries"]):
                raise ValueError(f"archive content manifest contains invalid entries: {archive}")

            restored: list[str] = []
            restored_archive_paths: set[str] = set()
            expected_canonical_paths = {
                path
                for path, entry in expected_entries.items()
                if entry.get("role") == "canonical_csv"
            }
            prefix = f"canonical/{state_key}/"
            for member in bundle.getmembers():
                if member.name == CONTENT_MANIFEST_NAME:
                    continue
                if not _safe_tar_member(member.name) or not member.isfile():
                    raise ValueError(f"archive contains unsafe member: {member.name}")
                expected = expected_entries.get(member.name)
                if expected is None:
                    raise ValueError(f"archive member missing from manifest: {member.name}")
                if not member.name.startswith(prefix):
                    continue
                relative = Path(member.name.removeprefix(prefix))
                if not relative.parts or relative.parts[0] != "csv" or relative.suffix != ".gz":
                    raise ValueError(f"archive has invalid canonical member: {member.name}")
                handle = bundle.extractfile(member)
                if handle is None:
                    raise ValueError(f"archive member cannot be read: {member.name}")
                raw = handle.read()
                if len(raw) != expected.get("bytes") or hashlib.sha256(raw).hexdigest() != expected.get("sha256"):
                    raise ValueError(f"archive member checksum mismatch: {member.name}")
                destination_path = destination / relative
                destination_path.parent.mkdir(parents=True, exist_ok=True)
                temporary = destination_path.with_name(destination_path.name + ".restore-tmp")
                temporary.write_bytes(raw)
                os.replace(temporary, destination_path)
                restored.append(relative.as_posix())
                restored_archive_paths.add(member.name)
            if restored_archive_paths != expected_canonical_paths:
                raise ValueError(f"archive canonical payload is incomplete: {archive}")
    except (tarfile.TarError, OSError, json.JSONDecodeError) as exc:
        raise ValueError(f"invalid release archive: {archive}") from exc
    if not restored:
        raise ValueError(f"archive has no canonical CSV payload for {state_key}: {archive}")
    return {"state_key": state_key, "restored": sorted(restored), "archive": str(archive)}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    subcommands = parser.add_subparsers(dest="command", required=True)
    build = subcommands.add_parser("build", help="build compact release checkpoint assets")
    build.add_argument("--artifacts", type=Path, required=True)
    build.add_argument("--repo", type=Path, default=Path.cwd())
    build.add_argument("--output", type=Path, required=True)
    build.add_argument("--release-tag", default=DEFAULT_RELEASE_TAG)
    build.add_argument(
        "--allow-checkpoint-regression",
        action="store_true",
        help="replace a committed checkpoint with one holding fewer canonical rows",
    )
    restore = subcommands.add_parser("restore", help="restore canonical CSV from a verified asset")
    restore.add_argument("--archive", type=Path, required=True)
    restore.add_argument("--state-key", required=True)
    restore.add_argument("--destination", type=Path, required=True)
    restore.add_argument("--expected-manifest", type=Path)
    args = parser.parse_args()
    if args.command == "build":
        report = build_release_archives(
            artifacts=args.artifacts,
            repo=args.repo,
            output=args.output,
            release_tag=args.release_tag,
            allow_checkpoint_regression=args.allow_checkpoint_regression,
        )
    else:
        report = restore_canonical_checkpoint(
            archive=args.archive,
            state_key=args.state_key,
            destination=args.destination,
            expected_manifest=args.expected_manifest,
        )
    print(json.dumps(report, indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
