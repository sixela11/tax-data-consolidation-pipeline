"""Stage 2: find the taxpayers who together make up 90% of the yearly total.

Taxpayers are sorted by YEARLY TOTAL (highest first) and selected until the
running total reaches the target share (default 90%) of the overall total.
The sort is stable, so ties keep their original row order.
"""

from __future__ import annotations

import logging
import time
from dataclasses import dataclass
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font

from .common import (
    copy_column_widths,
    copy_row,
    copy_style,
    find_column,
    header_map,
    is_blank,
    is_grand_total_row,
    last_header_col,
    last_row_in_col,
    prepare_output,
    require_column,
    save_atomically,
    set_range_format,
    to_number,
    write_summary_sheet,
)
from .config import PipelineConfig

log = logging.getLogger(__name__)


@dataclass
class TopContributorsResult:
    district: str
    output_path: Path
    total_taxpayers: int
    selected_taxpayers: int
    original_total: float
    target: float
    selected_total: float
    coverage: float
    seconds: float


def select_top_contributors(values: list[float], target_percent: float) -> tuple[list[int], float, float]:
    """Return (indexes of selected items, target amount, selected total).

    Pure function so the cut-off rule can be tested on its own.
    """
    total = sum(values)
    target = total * target_percent
    order = sorted(range(len(values)), key=lambda i: values[i], reverse=True)  # stable
    selected, cumulative = [], 0.0
    for i in order:
        selected.append(i)
        cumulative += values[i]
        if cumulative >= target:
            break
    return selected, target, cumulative


def top_contributors_district(district: str, combined_file: Path, config: PipelineConfig) -> TopContributorsResult:
    start = time.perf_counter()
    sheet_name = f"{config.percent_label} COLLECTION"
    output_path = prepare_output(
        config.top_dir / f"{district}_{config.percent_label.rstrip('%')}_PERCENT_COLLECTION.xlsx",
        config.overwrite,
    )

    src_wb = load_workbook(combined_file)
    src = src_wb[config.yearly_sheet] if config.yearly_sheet in src_wb.sheetnames else src_wb.worksheets[0]
    headers = header_map(src, 1)
    label = f"Combined {combined_file.name}"
    yearly_col = require_column(headers, label, config.yearly_total_header)
    tin_col = require_column(headers, label, *config.tin_headers)
    taxpayer_col = find_column(headers, *config.taxpayer_headers)
    last_row = last_row_in_col(src, tin_col)
    last_col = last_header_col(src, 1)

    rows, values = [], []
    for row in range(2, last_row + 1):
        record = [src.cell(row, c).value for c in range(1, last_col + 1)]
        if is_grand_total_row(record, taxpayer_col - 1 if taxpayer_col else None):
            continue
        if is_blank(record[tin_col - 1]):
            continue
        rows.append(row)
        values.append(to_number(record[yearly_col - 1]))

    selected, target, selected_total = select_top_contributors(values, config.target_percent)
    original_total = sum(values)
    coverage = selected_total / original_total if original_total else 0.0

    out_wb = Workbook()
    out = out_wb.active
    out.title = sheet_name
    copy_row(src, 1, out, 1, last_col)
    for out_row, i in enumerate(selected, 2):
        copy_row(src, rows[i], out, out_row, last_col)

    gt_row = len(selected) + 2
    out.cell(gt_row, taxpayer_col or 1, "GRAND TOTAL")
    out.cell(gt_row, yearly_col, selected_total)
    if gt_row > 2:
        for col in range(1, last_col + 1):
            copy_style(out.cell(gt_row - 1, col), out.cell(gt_row, col))
    for col in range(1, last_col + 1):
        out.cell(gt_row, col).font = Font(bold=True)
    set_range_format(out, yearly_col, 2, gt_row, "#,##0.00")
    if selected:
        out.auto_filter.ref = f"A1:{out.cell(gt_row - 1, last_col).coordinate}"
    copy_column_widths(src, out, last_col, cap=40)

    write_summary_sheet(out_wb, [
        ("DISTRICT", district),
        ("Total Taxpayers", len(values)),
        ("Selected Taxpayers", len(selected)),
        ("Original Total Collection", original_total),
        (f"{config.percent_label} Target", target),
        ("Selected Collection", selected_total),
        ("Actual Coverage", coverage),
    ], {4: "#,##0.00", 5: "#,##0.00", 6: "#,##0.00", 7: "0.0000%"})

    save_atomically(out_wb, output_path)
    src_wb.close()

    return TopContributorsResult(district, output_path, len(values), len(selected), original_total,
                                 target, selected_total, coverage, time.perf_counter() - start)
