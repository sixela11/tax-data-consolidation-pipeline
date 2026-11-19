"""Stage 1: consolidate Yearly, Micro and Sub-50M datasets by TIN.

For every taxpayer row in the Yearly workbook, the YEARLY TOTAL becomes the
highest of:

* the Yearly value itself,
* IT GROSS SALES from the Micro-business returns (if the TIN matches),
* TOTAL GROSS SALES from the Sub-50M returns (if the TIN matches).

Tie rules: Yearly is kept if it is >= both alternatives; Sub-50M beats Micro on
a tie. Audit columns record the original value and the winning source, and the
winning source's data is appended (the full Micro row, or Sub-50M columns G:AX).
"""

from __future__ import annotations

import logging
import time
from collections import Counter
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from .common import (
    PipelineError,
    clean_header,
    find_header_row,
    header_map,
    is_blank,
    last_header_col,
    last_row_in_col,
    normalize_tin,
    prepare_output,
    require_column,
    save_atomically,
    set_range_format,
    to_number,
)
from .config import PipelineConfig

log = logging.getLogger(__name__)

MONEY = "#,##0.00"


@dataclass
class SourceIndex:
    """Rows of a secondary dataset keyed by normalised TIN."""

    name: str
    header_row: int
    headers: list[Any]  # full header row values
    rows: dict[str, tuple[float, list[Any]]]  # tin -> (value, row values)
    duplicates: int


@dataclass
class ConsolidationResult:
    district: str
    output_path: Path
    taxpayers: int
    yearly_wins: int
    micro_wins: int
    sub50m_wins: int
    replaced: int
    original_total: float
    updated_total: float
    seconds: float


def index_source(path: Path, value_header: str, label: str, tin_headers: tuple[str, ...]) -> SourceIndex:
    """Load a secondary dataset once and index it by TIN.

    The header row is found by searching the first 20 rows for a TIN cell.
    Duplicate TINs keep the row with the highest value (first one on a tie).
    """
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb.worksheets[0]
        header_row = None
        for name in tin_headers:
            header_row = find_header_row(ws, name)
            if header_row:
                break
        if header_row is None:
            raise PipelineError(f"{label} {path.name}: TIN header not found in the first 20 rows.")

        all_rows = list(ws.iter_rows(min_row=header_row, values_only=True))
        headers = list(all_rows[0])
        lookup = {clean_header(h): i for i, h in enumerate(headers, 1) if not is_blank(h)}
        tin_col = require_column(lookup, f"{label} {path.name}", *tin_headers)
        value_col = require_column(lookup, f"{label} {path.name}", value_header)

        rows: dict[str, tuple[float, list[Any]]] = {}
        duplicates = 0
        for values in all_rows[1:]:
            values = list(values)
            if len(values) < len(headers):
                values += [None] * (len(headers) - len(values))
            tin = normalize_tin(values[tin_col - 1])
            if not tin:
                continue
            value = to_number(values[value_col - 1])
            if tin in rows:
                duplicates += 1
                if value > rows[tin][0]:
                    rows[tin] = (value, values)
            else:
                rows[tin] = (value, values)
        return SourceIndex(label, header_row, headers, rows, duplicates)
    finally:
        wb.close()


def _trim_headers(headers: list[Any]) -> list[Any]:
    """Header values up to the last non-empty one (Excel's End(xlToLeft))."""
    last = 0
    for i, h in enumerate(headers, 1):
        if not is_blank(h):
            last = i
    return headers[:last]


def consolidate_district(
    district: str, yearly_file: Path, micro_file: Path, sub50m_file: Path, config: PipelineConfig
) -> ConsolidationResult:
    start = time.perf_counter()
    output_path = prepare_output(config.combined_dir / f"{district}_COMBINED_BY_TIN.xlsx",
                                 config.overwrite)

    micro = index_source(micro_file, config.micro_value_header, "Micro", config.tin_headers)
    sub50m = index_source(sub50m_file, config.sub50m_value_header, "Sub-50M", config.tin_headers)
    log.info("  Indexed Micro: %s TINs (%d duplicates) | Sub-50M: %s TINs (%d duplicates)",
             f"{len(micro.rows):,}", micro.duplicates, f"{len(sub50m.rows):,}", sub50m.duplicates)

    wb = load_workbook(yearly_file)
    ws = wb[config.yearly_sheet] if config.yearly_sheet in wb.sheetnames else wb.worksheets[0]

    headers = header_map(ws, 1)
    label = f"Yearly {yearly_file.name}"
    tin_col = require_column(headers, label, *config.tin_headers)
    yearly_col = require_column(headers, label, config.yearly_total_header)
    last_row = last_row_in_col(ws, tin_col)
    original_last_col = last_header_col(ws, 1)

    # ---- appended column layout ------------------------------------------
    original_col = original_last_col + 1
    source_col = original_last_col + 2
    source_value_col = original_last_col + 3
    micro_headers = _trim_headers(micro.headers)
    micro_start = original_last_col + 4
    sub_first, sub_last = config.sub50m_cols
    sub_start = micro_start + len(micro_headers)
    final_col = sub_start + (sub_last - sub_first)

    ws.cell(1, original_col, "ORIGINAL YEARLY TOTAL")
    ws.cell(1, source_col, "YEARLY TOTAL SOURCE")
    ws.cell(1, source_value_col, "YEARLY TOTAL SOURCE VALUE")
    for i, h in enumerate(micro_headers, 1):
        ws.cell(1, micro_start + i - 1, f"MICRO | {h if h is not None else f'COLUMN_{i}'}")
    for src in range(sub_first, sub_last + 1):
        h = sub50m.headers[src - 1] if src <= len(sub50m.headers) else None
        ws.cell(1, sub_start + src - sub_first, f"50M | {h if h is not None else f'COLUMN_{src}'}")

    # ---- match every Yearly taxpayer ---------------------------------------
    taxpayers = yearly_wins = micro_wins = sub_wins = replaced = 0
    original_total = updated_total = 0.0
    seen: Counter[str] = Counter()

    for row in range(2, last_row + 1):
        tin = normalize_tin(ws.cell(row, tin_col).value)
        if not tin:
            continue
        taxpayers += 1
        seen[tin] += 1

        original = to_number(ws.cell(row, yearly_col).value)
        original_total += original

        micro_hit = micro.rows.get(tin)
        sub_hit = sub50m.rows.get(tin)
        micro_value = micro_hit[0] if micro_hit else 0.0
        sub_value = sub_hit[0] if sub_hit else 0.0

        # Only sources that actually matched take part (see README: fixes a
        # negative Yearly value being replaced by 0 when nothing matched).
        candidates = [original]
        if micro_hit:
            candidates.append(micro_value)
        if sub_hit:
            candidates.append(sub_value)
        highest = max(candidates)

        if sub_hit and sub_value == highest and sub_value > original:
            winner, source_text = "SUB50M", "50M - TOTAL GROSS SALES"
        elif micro_hit and micro_value == highest and micro_value > original:
            winner, source_text = "MICRO", "MICRO - IT GROSS SALES"
        else:
            winner, source_text = "YEARLY", "YEARLY"

        ws.cell(row, original_col, original)
        ws.cell(row, source_col, source_text)
        ws.cell(row, source_value_col, highest)
        ws.cell(row, yearly_col, highest)
        updated_total += highest
        if highest > original:
            replaced += 1

        if winner == "YEARLY":
            yearly_wins += 1
        elif winner == "MICRO":
            micro_wins += 1
            for i, value in enumerate(micro_hit[1][: len(micro_headers)]):
                ws.cell(row, micro_start + i, value)
        else:
            sub_wins += 1
            values = sub_hit[1]
            for src in range(sub_first, sub_last + 1):
                ws.cell(row, sub_start + src - sub_first,
                        values[src - 1] if src <= len(values) else None)

    dupes = {t: n for t, n in seen.items() if n > 1}
    if dupes:
        log.warning("  %d TIN(s) appear more than once in the Yearly file (each row processed).",
                    len(dupes))

    # ---- formatting -------------------------------------------------------
    for col in (yearly_col, original_col, source_value_col):
        set_range_format(ws, col, 2, last_row, MONEY)
    for col, width in ((original_col, 22), (source_col, 30), (source_value_col, 24)):
        ws.column_dimensions[get_column_letter(col)].width = width
    for col in range(micro_start, final_col + 1):  # approximate AutoFit, capped at 30
        longest = max((len(str(ws.cell(r, col).value or "")) for r in range(1, last_row + 1)),
                      default=8)
        ws.column_dimensions[get_column_letter(col)].width = min(max(longest + 2, 8), 30)

    save_atomically(wb, output_path)
    wb.close()

    return ConsolidationResult(district, output_path, taxpayers, yearly_wins, micro_wins,
                               sub_wins, replaced, original_total, updated_total,
                               time.perf_counter() - start)
