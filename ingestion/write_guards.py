"""Shared path guards for deliberately persisted fixture data."""

from __future__ import annotations

import os
from pathlib import Path

from .warehouse import DEFAULT_WAREHOUSE

REPOSITORY_STATE = Path(__file__).resolve().parent.parent / "localdata"


def assert_fixture_write_isolation(
    warehouse: str | os.PathLike[str],
    *,
    localdata: str | os.PathLike[str] | None = None,
) -> None:
    """Refuse to mix synthetic fixtures with the configured canonical state.

    A separate environment opt-in is still required by each caller. This path
    guard protects both the configured warehouse and the repository's default
    runtime directory, including when environment defaults are implicit.
    """

    target = Path(warehouse).expanduser().resolve()
    canonical = Path(os.environ.get("MAGOLIDE_WAREHOUSE", DEFAULT_WAREHOUSE)).expanduser().resolve()
    repository_state = REPOSITORY_STATE.resolve()
    runtime = Path(localdata).expanduser().resolve() if localdata else None

    if (
        target == canonical
        or target.is_relative_to(repository_state)
        or runtime == repository_state
        or (runtime is not None and canonical.is_relative_to(runtime))
    ):
        raise RuntimeError(
            "fixture persistence requires warehouse/localdata paths isolated from canonical state"
        )
