"""
Pipeline Writer — central writer to Bet_Slips sheet.
All pipelines must route through here to prevent prediction leakage.
Tracks predictions_generated vs predictions_written for Pipeline Coverage %.
"""

import logging
import time
from typing import List, Tuple

from assayer.bet_record import BetRecord

logger = logging.getLogger(__name__)

BET_SLIPS_TAB = "Bet_Slips"
RATE_LIMIT_DELAY = 1.1


class PipelineWriter:
    def __init__(self, gspread_client, spreadsheet_key: str):
        self.client = gspread_client
        self.spreadsheet_key = spreadsheet_key
        self._ss = None
        self._ws = None

    def _open(self):
        if self._ss is None:
            self._ss = self.client.open_by_key(self.spreadsheet_key)
        if self._ws is None:
            try:
                self._ws = self._ss.worksheet(BET_SLIPS_TAB)
            except Exception:
                self._ws = self._ss.add_worksheet(
                    title=BET_SLIPS_TAB, rows=5000, cols=40
                )
                self._ws.append_row(BetRecord.bet_slip_headers(), value_input_option="RAW")
                logger.info(f"Created new {BET_SLIPS_TAB} tab with headers")

    def _ensure_headers(self):
        self._open()
        existing = self._ws.row_values(1)
        if not existing or existing[0] != "Bet_Record_ID":
            self._ws.insert_row(BetRecord.bet_slip_headers(), index=1, value_input_option="RAW")
            logger.info("Inserted Bet_Slips headers")

    def write_records(
        self,
        records: List[BetRecord],
        predictions_generated: int = 0,
    ) -> dict:
        """
        Write accepted BetRecords to Bet_Slips.
        Returns a coverage summary dict.
        """
        self._ensure_headers()

        written = 0
        failed = 0
        for rec in records:
            try:
                row = rec.to_bet_slip_row()
                self._ws.append_row(row, value_input_option="USER_ENTERED")
                written += 1
                time.sleep(RATE_LIMIT_DELAY)
            except Exception as e:
                logger.error(f"Failed to write record {rec.bet_record_id}: {e}")
                failed += 1

        total_gen = max(predictions_generated, written)
        coverage_pct = round((written / total_gen * 100), 2) if total_gen > 0 else 0.0

        summary = {
            "predictions_generated": total_gen,
            "predictions_written": written,
            "write_failures": failed,
            "pipeline_coverage_pct": coverage_pct,
        }
        logger.info(f"PipelineWriter: {written} written, {failed} failed, coverage={coverage_pct}%")
        return summary

    def count_existing(self) -> int:
        try:
            self._open()
            all_vals = self._ws.get_all_values()
            return max(0, len(all_vals) - 1)
        except Exception as e:
            logger.warning(f"Could not count Bet_Slips rows: {e}")
            return 0


def write_bet_records(
    client,
    spreadsheet_key: str,
    records: List[BetRecord],
    predictions_generated: int = 0,
) -> Tuple[dict, str]:
    """
    Convenience wrapper. Returns (summary_dict, error_string_or_None).
    """
    try:
        writer = PipelineWriter(client, spreadsheet_key)
        summary = writer.write_records(records, predictions_generated)
        return summary, None
    except Exception as e:
        logger.exception(f"PipelineWriter error: {e}")
        return {}, str(e)


def count_bet_slips(client, spreadsheet_key: str) -> int:
    try:
        writer = PipelineWriter(client, spreadsheet_key)
        return writer.count_existing()
    except Exception:
        return 0
