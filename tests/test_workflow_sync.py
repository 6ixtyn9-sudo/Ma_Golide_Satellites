from __future__ import annotations

import re
from pathlib import Path

# F1/R5-2: Guard against source vs installed workflow drift.
# The installed workflow (.github/workflows/ingestion.yml) requires
# `workflows` permission to update, so source (workflows/ingestion.yml)
# is the PR-reviewable copy. The cache-recovery allowlist must stay
# byte-identical in both files, otherwise a withdrawn sidecar survives
# incorrectly (fail-closed) or is deleted prematurely.

def _extract_allowlist_block(text: str) -> str:
    m = re.search(r"allowed = tuple\(.*?\)\s*\)", text, re.DOTALL)
    assert m, "allowlist block not found in workflow"
    return m.group(0)

def test_workflow_allowlist_sync():
    repo = Path(__file__).resolve().parents[1]
    source = repo / "workflows" / "ingestion.yml"
    installed = repo / ".github" / "workflows" / "ingestion.yml"
    assert source.exists(), f"missing {source}"
    assert installed.exists(), f"missing {installed}"
    src_block = _extract_allowlist_block(source.read_text())
    inst_block = _extract_allowlist_block(installed.read_text())
    assert src_block == inst_block, (
        "workflows/ingestion.yml and .github/workflows/ingestion.yml allowlist blocks diverged — "
        "sync the source to installed via manual workflows-permission step"
    )
    # Pin the withdrawn sidecar is present in both
    assert "withdrawn" in src_block and "withdrawn" in inst_block
