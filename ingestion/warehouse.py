"""SQL analytics layer over Ma Golide's canonical DuckDB tables.

The acquisition store owns writes.  This module opens independent connections
for research and exposes temporary SQL views, so mining cannot accidentally
mutate frozen operational facts.  JSON remains the complete source payload;
views expose only observed values and never derive missing provider facts.
"""

from __future__ import annotations

from pathlib import Path

import duckdb

from .duckdb_store import DuckDBStore

DEFAULT_WAREHOUSE = "localdata/warehouse.duckdb"
DEFAULT_CSV_DIR = "localdata/csv"


def connect(path: str = DEFAULT_WAREHOUSE, *, read_only: bool = True):
    """Open DuckDB and create temporary research views.

    A missing warehouse is an operational error, not an empty successful slate.
    Temporary views work on read-only connections and disappear on close.
    """
    if path != ":memory:" and read_only and not Path(path).exists():
        raise FileNotFoundError(f"warehouse does not exist: {path}")
    con = duckdb.connect(path, read_only=read_only)
    con.execute("""
        CREATE TEMP VIEW prediction_facts AS
        SELECT
            game_id, provider, slot, captured_at, model_version, pre_tip, frozen,
            TRY_CAST(json_extract_string(payload_json, '$.prob_home')
                     AS DOUBLE) AS prob_home,
            TRY_CAST(json_extract_string(payload_json, '$.prob_away')
                     AS DOUBLE) AS prob_away,
            TRY_CAST(json_extract_string(payload_json, '$.expected_home_score')
                     AS DOUBLE) AS expected_home_score,
            TRY_CAST(json_extract_string(payload_json, '$.expected_away_score')
                     AS DOUBLE) AS expected_away_score,
            TRY_CAST(json_extract_string(payload_json, '$.expected_margin')
                     AS DOUBLE) AS expected_margin,
            TRY_CAST(json_extract_string(payload_json, '$.expected_total')
                     AS DOUBLE) AS expected_total,
            payload_json
        FROM predictions;

        CREATE TEMP VIEW book_facts AS
        SELECT
            game_id, provider, book, market, period, observed_utc,
            source_modified_utc, line, over_price_decimal, under_price_decimal,
            home_price_decimal, away_price_decimal, pre_match, main_line,
            payload_hash, frozen, payload_json
        FROM book_snapshots;

        CREATE TEMP VIEW latest_pregame_predictions AS
        SELECT * EXCLUDE (rn)
        FROM (
            SELECT p.*,
                   row_number() OVER (
                       PARTITION BY game_id
                       ORDER BY captured_at DESC, provider ASC
                   ) AS rn
            FROM prediction_facts p
            WHERE pre_tip
        )
        WHERE rn = 1;

        CREATE TEMP VIEW game_results AS
        SELECT
            game_id, league, season, home_team_id, away_team_id,
            home_name, away_name, scheduled_utc, status, phase,
            actual_json,
            json_extract_string(actual_json, '$.period_kind') AS period_kind,
            json_extract_string(actual_json, '$.home_final') AS home_final,
            json_extract_string(actual_json, '$.away_final') AS away_final,
            json_extract_string(actual_json, '$.source') AS result_source
        FROM games
        WHERE actual_json IS NOT NULL;
    """)
    return con


def export_csv(
    warehouse: str = DEFAULT_WAREHOUSE,
    directory: str = DEFAULT_CSV_DIR,
) -> dict[str, str]:
    """Materialize canonical tables to CSV.GZ without rebuilding mutable state."""
    if warehouse != ":memory:" and not Path(warehouse).exists():
        raise FileNotFoundError(f"warehouse does not exist: {warehouse}")
    store = DuckDBStore(warehouse, persist=False)
    try:
        return store.export_csv(directory)
    finally:
        store.close()
