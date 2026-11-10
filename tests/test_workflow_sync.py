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
