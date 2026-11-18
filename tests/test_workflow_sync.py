from __future__ import annotations

import re
from pathlib import Path

# F1/R5-2: Guard against source vs installed workflow drift.
# The installed workflow (.github/workflows/ingestion.yml) requires
# `workflows` permission to update, so source (workflows/ingestion.yml)
# is the PR-reviewable copy. The complete files must remain byte-identical:
# a partial manual edit can otherwise make preflight or runtime semantics diverge.


def _extract_allowlist_blocks(text: str) -> list[str]:
    blocks = [m.group(0) for m in re.finditer(r"allowed = tuple\(.*?\)\s*\)", text, re.DOTALL)]
    assert blocks, "allowlist block not found in workflow"
    return blocks


def test_workflow_mirror_sync():
    repo = Path(__file__).resolve().parents[1]
    source = repo / "workflows" / "ingestion.yml"
    installed = repo / ".github" / "workflows" / "ingestion.yml"
    assert source.exists(), f"missing {source}"
    assert installed.exists(), f"missing {installed}"
    source_text = source.read_text()
    installed_text = installed.read_text()
    assert source_text == installed_text, (
        "workflows/ingestion.yml and .github/workflows/ingestion.yml diverged — "
        "replace both files with the same complete workflow YAML"
    )
    blocks = _extract_allowlist_blocks(source_text)
    # Pin the withdrawn/no-bet sidecars in every cache-recovery allowlist.
    assert all("withdrawn" in block for block in blocks)
    assert all("no_bet" in block for block in blocks)


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


def test_workflow_keeps_keyless_fallback_and_wires_optional_provider_secrets():
    repo = Path(__file__).resolve().parents[1]
    source = repo / "workflows" / "ingestion.yml"
    installed = repo / ".github" / "workflows" / "ingestion.yml"
    for path in (source, installed):
        text = path.read_text()
        assert 'ALLOW_MAGOLIDE_WRITES: "true"' in text
        assert 'SPORTRADAR_ACCESS_LEVEL: "trial"' in text
        assert 'ALLOW_NOTIFICATION_WRITES: "false"' in text
        for secret in (
            "THE_ODDS_API_KEY",
            "THERUNDOWN_API_KEY",
            "SPORTSGAMEODDS_API_KEY",
            "ISPORTS_API_KEY",
            "SPORTRADAR_API_KEY",
            "PREDIXSPORT_API_KEY",
        ):
            assert f"{secret}: ${{{{ secrets.{secret} }}}}" in text
        assert "${{ vars." not in text
        assert "required_secret" not in text
        assert text.count("state_key:") == 6
        assert '- cron: "11 * * * *"' in text
        assert "11 */3 * * *" not in text
        assert "available=true" in text
        assert "repository_dispatch:" in text
        assert "types: [magolide-scheduled-tick]" in text
        assert 'case "$EVENT_NAME:$EVENT_ACTION" in' in text
        assert "schedule:*|repository_dispatch:magolide-scheduled-tick)" in text
        assert '[ "$AUTOMATED_TICK" = "true" ] && [ "$ALLOW_MAGOLIDE_WRITES" = "true" ]' in text
        cache = text.split("name: Persist league-isolated acquisition and canonical state", 1)[1].split(
            "uses: actions/cache/save@v4", 1
        )[0]
        assert "github.event_name == 'workflow_dispatch'" in cache
        assert "github.event_name == 'repository_dispatch'" in cache
        assert "github.event.action == 'magolide-scheduled-tick'" in cache


def test_only_named_automated_triggers_can_publish_a_durable_operator_decision():
    repo = Path(__file__).resolve().parents[1]
    for path in (
        repo / "workflows" / "ingestion.yml",
        repo / ".github" / "workflows" / "ingestion.yml",
    ):
        text = path.read_text()
        publish = text.split("publish-localdata:", 1)[1].split("runs-on:", 1)[0]
        assert "github.event_name == 'schedule'" in publish
        assert "github.event_name == 'repository_dispatch'" in publish
        assert "github.event.action == 'magolide-scheduled-tick'" in publish
        assert "workflow_dispatch" not in publish
        assert "Manual dispatch uploads artifacts but cannot publish." in publish


def test_renamed_cache_key_falls_back_to_the_pre_rename_snapshot():
    # Source workflow only: the installed copy is synced by a manual
    # workflows-permission step and must keep its key until then.
    repo = Path(__file__).resolve().parents[1]
    for path in (repo / "workflows" / "ingestion.yml",):
        text = path.read_text()
        restore = text.split("restore-keys: |", 1)[1].splitlines()[1:3]
        assert [line.strip() for line in restore] == [
            "magolide-localdata-${{ matrix.state_key }}-",
            "magolide-localdata-v" + "4-${{ matrix.state_key }}-",
        ]
        assert text.count("key: magolide-localdata-${{ matrix.state_key }}-${{ github.run_id }}") == 2
