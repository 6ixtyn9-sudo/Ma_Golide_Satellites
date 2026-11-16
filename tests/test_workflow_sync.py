from __future__ import annotations

import re
from pathlib import Path

# F1/R5-2: Guard against source vs installed workflow drift.
# The installed workflow (.github/workflows/ingestion.yml) requires
# `workflows` permission to update, so source (workflows/ingestion.yml)
# is the PR-reviewable copy. The cache-recovery allowlist must stay
# byte-identical in both files, otherwise a withdrawn sidecar survives
# incorrectly (fail-closed) or is deleted prematurely.


def _extract_allowlist_blocks(text: str) -> list[str]:
    blocks = [m.group(0) for m in re.finditer(r"allowed = tuple\(.*?\)\s*\)", text, re.DOTALL)]
    assert blocks, "allowlist block not found in workflow"
    return blocks


def test_workflow_allowlist_sync():
    repo = Path(__file__).resolve().parents[1]
    source = repo / "workflows" / "ingestion.yml"
    installed = repo / ".github" / "workflows" / "ingestion.yml"
    assert source.exists(), f"missing {source}"
    assert installed.exists(), f"missing {installed}"
    src_blocks = _extract_allowlist_blocks(source.read_text())
    inst_blocks = _extract_allowlist_blocks(installed.read_text())
    assert src_blocks == inst_blocks, (
        "workflows/ingestion.yml and .github/workflows/ingestion.yml allowlist blocks diverged — "
        "sync the source to installed via manual workflows-permission step"
    )
    # Pin the withdrawn sidecar is present in every block
    assert all("withdrawn" in b for b in src_blocks)
    assert all("withdrawn" in b for b in inst_blocks)


def test_operator_state_restore_uses_latest_main_not_event_sha():
    repo = Path(__file__).resolve().parents[1]
    for path in (
        repo / "workflows" / "ingestion.yml",
        repo / ".github" / "workflows" / "ingestion.yml",
    ):
        text = path.read_text()
        assert (
            "git fetch --no-tags origin +refs/heads/main:refs/remotes/origin/main" in text
        )
        assert 'MAIN_REF=refs/remotes/origin/main' in text
        assert 'git cat-file -e "${MAIN_REF}:${STATE_DIR}"' in text
        assert 'git restore --source="$MAIN_REF" --worktree -- "$STATE_DIR"' in text
        assert 'git cat-file -e "HEAD:${STATE_DIR}"' not in text


def test_scheduled_write_default_and_ncaaw_secret_gate_are_correct():
    repo = Path(__file__).resolve().parents[1]
    source = repo / "workflows" / "ingestion.yml"
    installed = repo / ".github" / "workflows" / "ingestion.yml"
    for path in (source, installed):
        text = path.read_text()
        assert "ALLOW_MAGOLIDE_WRITES: ${{ vars.ALLOW_MAGOLIDE_WRITES || 'true' }}" in text
        assert (
            'if [ "$GITHUB_EVENT_NAME" = "schedule" ] && '
            '[ "$ALLOW_MAGOLIDE_WRITES" = "true" ]; then'
        ) in text
        assert re.search(
            r"(?m)^          - label: NCAA women\n"
            r"            league: NCAA_W\n"
            r"            state_key: ncaa_w\n"
            r"            required_secret: \"\"$",
            text,
        ), f"NCAA_W must not require an iSports credential in {path}"

    source_text = source.read_text()
    assert re.search(
        r"(?m)^          - label: Chinese CBA\n"
        r"            league: CBA\n"
        r"            state_key: cba\n"
        r"            required_secret: \"\"$",
        source_text,
    ), "the PR-reviewable workflow source must enable keyless CBA ingestion"
    installed_text = installed.read_text()
    assert re.search(
        r"(?m)^          - label: Chinese CBA\n"
        r"            league: CBA\n"
        r"            state_key: cba\n"
        r"            required_secret: ISPORTS_API_KEY$",
        installed_text,
    ), "the installed workflow is manually updated after workflows permission is available"
