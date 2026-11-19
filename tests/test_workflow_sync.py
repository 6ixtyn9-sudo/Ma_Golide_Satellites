from __future__ import annotations

import re
import subprocess
from pathlib import Path

from ingestion.league_registry import SPECS

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
    # Withdrawn sidecars remain recoverable; no-bet decisions are inline in
    # auto_tickets_state.json and must not grow a per-date recovery surface.
    assert all("withdrawn" in block for block in blocks)
    assert all("no_bet" not in block for block in blocks)
    assert all("euroleague_consensus_paper" in block for block in blocks)
    # Structured pick JSON is cache/artifact-only for every league, so
    # restoration must not delete it before replacing Git-durable files.
    assert all("picks_{date}.json" not in block for block in blocks)
    assert all("picks_{date}_lifecycle" not in block for block in blocks)
    assert all("pick_freeze" not in block for block in blocks)


def test_gitignore_allows_every_git_durable_operator_artifact():
    repo = Path(__file__).resolve().parents[1]
    durable_paths = (
        "localdata/nba/picks_2026-11-18.txt",
        "localdata/nba/pick_lifecycle_state.json",
        "localdata/nba/coverage_gap_2026-11-18.json",
        "localdata/nba/auto_tickets_2026-11-18_decision.json",
        "localdata/euroleague/euroleague_consensus_paper_2026-11-18_roster.json",
        "localdata/euroleague/euroleague_consensus_paper_2026-11-18.txt",
    )
    for relative_path in durable_paths:
        result = subprocess.run(
            ["git", "check-ignore", "-q", relative_path],
            cwd=repo,
            check=False,
        )
        assert result.returncode == 1, f"durable publisher artifact is ignored: {relative_path}"


def test_gitignore_rejects_retired_ticket_and_pick_freeze_sidecars():
    repo = Path(__file__).resolve().parents[1]
    for relative_path in (
        "localdata/nba/auto_tickets_2026-11-18_preview.txt",
        "localdata/nba/pick_freeze_2026-11-18_0123456789abcdef.json",
    ):
        result = subprocess.run(
            ["git", "check-ignore", "-q", relative_path],
            cwd=repo,
            check=False,
        )
        assert result.returncode == 0


def test_no_fixture_pick_freeze_sidecars_are_tracked_in_any_league():
    """Inline lifecycle state is the sole durable home for fixture snapshots."""

    repo = Path(__file__).resolve().parents[1]
    tracked = subprocess.run(
        ["git", "ls-files", "-z"],
        cwd=repo,
        check=True,
        capture_output=True,
    ).stdout.decode().split("\0")
    sidecars = [
        path
        for path in tracked
        if re.fullmatch(
            r"localdata/[^/]+/pick_freeze_\d{4}-\d{2}-\d{2}_[0-9a-f]{16}\.json",
            path,
        )
    ]
    assert not sidecars, f"retired fixture-freeze sidecars are tracked: {sidecars}"


def test_gitignore_keeps_structured_pick_json_artifact_only_for_every_league():
    repo = Path(__file__).resolve().parents[1]
    states = tuple(spec.state_key for spec in SPECS)
    assert states
    for state in states:
        for name in (
            "picks_2026-11-18.json",
            "picks_2026-11-18_lifecycle.json",
            "pick_freeze_2026-11-18_0123456789abcdef.json",
        ):
            result = subprocess.run(
                ["git", "check-ignore", "-q", f"localdata/{state}/{name}"],
                cwd=repo,
                check=False,
            )
            assert result.returncode == 0, f"structured pick JSON is globally visible: {state}/{name}"


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
        restore = 'git restore --source="$MAIN_REF" --worktree -- "$STATE_DIR"'
        assert restore in text
        recovery = "Recover retained artifact-only pick ledgers after a cold cache"
        assert recovery in text
        assert text.index(restore) < text.index(recovery)
        assert "python scripts/recover_pick_ledgers.py" in text
        assert "--state-key \"$STATE_KEY\"" in text
        assert "--repository \"$GITHUB_REPOSITORY\"" in text
        assert 'actions: read' in text
        assert 'git cat-file -e "HEAD:${STATE_DIR}"' not in text


def test_workflow_keeps_keyless_fallback_and_wires_optional_provider_secrets():
    repo = Path(__file__).resolve().parents[1]
    source = repo / "workflows" / "ingestion.yml"
    installed = repo / ".github" / "workflows" / "ingestion.yml"
    for path in (source, installed):
        text = path.read_text()
        assert 'ALLOW_MAGOLIDE_WRITES: "true"' in text
        assert 'ALLOW_NOTIFICATION_WRITES: "false"' in text
        optional_provider_secrets = {
            "THE_ODDS_API_KEY",
            "THERUNDOWN_API_KEY",
            "SPORTSGAMEODDS_API_KEY",
            "PREDIXSPORT_API_KEY",
        }
        wired_secrets = set(re.findall(r"secrets\.([A-Z0-9_]+)", text))
        assert wired_secrets == optional_provider_secrets
        for secret in optional_provider_secrets:
            assert f"{secret}: ${{{{ secrets.{secret} }}}}" in text
        assert "${{ vars." not in text
        assert "required_secret" not in text
        # The static seven-row matrix is emitted by the read-only adaptive gate,
        # so idle external ticks do not start a full league job.
        assert "matrix: ${{ fromJSON(needs.gate.outputs.matrix) }}" in text
        assert "python scripts/adaptive_gate.py" in text
        assert "TARGET_DATE: ${{ matrix.target_date }}" in text
        assert "steps.target.outputs.date" not in text
        assert "schedule:" not in text
        assert "cron:" not in text
        assert "available=true" in text
        assert "repository_dispatch:" in text
        assert "types: [magolide-scheduled-tick]" in text
        assert 'case "$EVENT_NAME:$EVENT_ACTION" in' in text
        assert "repository_dispatch:magolide-scheduled-tick)" in text
        assert "schedule:*" not in text
        assert '[ "$AUTOMATED_TICK" = "true" ] && [ "$ALLOW_MAGOLIDE_WRITES" = "true" ]' in text
        cache = text.split("name: Persist league-isolated evidence-verified cache epoch", 1)[1].split(
            "uses: actions/cache/save@v4", 1
        )[0]
        assert "github.event_name == 'workflow_dispatch'" not in cache
        assert "github.event_name == 'repository_dispatch'" in cache
        assert "github.event.action == 'magolide-scheduled-tick'" in cache


def test_adaptive_gate_skips_idle_external_ticks_but_keeps_manual_rehearsals_full():
    repo = Path(__file__).resolve().parents[1]
    for path in (
        repo / "workflows" / "ingestion.yml",
        repo / ".github" / "workflows" / "ingestion.yml",
    ):
        text = path.read_text()
        gate = text.split("  gate:\n", 1)[1].split("\n  preflight:", 1)[0]
        assert "Adaptive schedule and operator-state gate" in gate
        assert "git restore --source=refs/remotes/origin/main --worktree -- localdata" in gate
        assert "latest_main_operator_state_unavailable" in gate
        assert "manual_read_only_rehearsal" in gate
        assert "python scripts/adaptive_gate.py" in gate
        preflight = text.split("  preflight:\n", 1)[1].split("\n  league:", 1)[0]
        assert "needs: gate" not in preflight
        assert "if: needs.gate.outputs.run_any" not in preflight
        assert "gate itself cannot start" in preflight
        league = text.split("  league:\n", 1)[1].split("\n  publish-localdata:", 1)[0]
        assert "needs: [gate, preflight]" in league
        assert "needs.gate.outputs.run_any == 'true'" in league
        assert "matrix: ${{ fromJSON(needs.gate.outputs.matrix) }}" in league


def test_successful_lanes_publish_despite_a_failed_matrix_sibling():
    repo = Path(__file__).resolve().parents[1]
    for path in (
        repo / "workflows" / "ingestion.yml",
        repo / ".github" / "workflows" / "ingestion.yml",
    ):
        text = path.read_text()
        artifact = text.split("name: Upload successful league reports and recovery evidence", 1)[1].split(
            "uses: actions/upload-artifact@v4", 1
        )[0]
        assert "if: success()" in artifact
        recovery = text.split(
            "name: Preserve durable pick lifecycle evidence from an interrupted lane", 1
        )[1].split("\n\n  publish-localdata:", 1)[0]
        assert "failure()" in recovery
        assert "MaGolide-lifecycle-${{ matrix.state_key }}" in text
        assert "pick_freeze_*.json" not in recovery
        assert "pick_lifecycle_state.json" in recovery
        assert "auto_tickets_" not in recovery
        assert "euroleague_consensus_paper_*_roster.json" in recovery
        assert "euroleague_consensus_paper_*.txt" not in recovery
        publish = text.split("publish-localdata:", 1)[1].split("runs-on:", 1)[0]
        assert "needs: [gate, preflight, league]" in text
        assert "needs.preflight.result == 'success'" in publish
        assert "needs.league.result == 'success'" not in publish
        assert "always()" in publish


def test_only_named_automated_triggers_can_publish_a_durable_operator_decision():
    repo = Path(__file__).resolve().parents[1]
    for path in (
        repo / "workflows" / "ingestion.yml",
        repo / ".github" / "workflows" / "ingestion.yml",
    ):
        text = path.read_text()
        publish = text.split("publish-localdata:", 1)[1].split("runs-on:", 1)[0]
        assert "github.event_name == 'schedule'" not in publish
        assert "github.event_name == 'repository_dispatch'" in publish
        assert "github.event.action == 'magolide-scheduled-tick'" in publish
        assert "needs.gate.outputs.run_any == 'true'" in publish
        assert "workflow_dispatch" not in publish
        assert "Manual dispatch uploads artifacts but cannot publish." in publish


def test_evidence_verified_cache_epoch_applies_to_every_league():
    repo = Path(__file__).resolve().parents[1]
    for path in (
        repo / "workflows" / "ingestion.yml",
        repo / ".github" / "workflows" / "ingestion.yml",
    ):
        text = path.read_text()
        assert "Restore league-isolated evidence-verified cache epoch" in text
        assert "Persist league-isolated evidence-verified cache epoch" in text
        assert text.count("magolide-localdata-evidence-v1-") == 3
        # No league may fall back to a pre-policy cache namespace.
        assert "magolide-localdata-v4-" not in text
        assert "magolide-localdata-ncaa_w-identity-v1-" not in text
        assert "matrix.league != 'NCAA_W'" not in text
        assert "matrix.league == 'NCAA_W'" not in text
