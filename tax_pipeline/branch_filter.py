"""Stage 3: from the top-contributor list, keep taxpayers with 7 or more branches."""

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
class BranchFilterResult:
    district: str
    output_path: Path
    count: int
    yearly_total: float
    min_branches_found: float | None
    max_branches_found: float | None
    seconds: float


def branch_filter_district(district: str, top_file: Path, config: PipelineConfig) -> BranchFilterResult:
    start = time.perf_counter()
    output_path = prepare_output(
        config.branches_dir / f"{district}_{config.min_branches}_OR_MORE_BRANCHES.xlsx",
        config.overwrite,
    )

    src_wb = load_workbook(top_file)
    wanted = f"{config.percent_label} COLLECTION"
    # Fallback: the data sheet is the last one (SUMMARY is inserted first).
    src = src_wb[wanted] if wanted in src_wb.sheetnames else src_wb.worksheets[-1]

    headers = header_map(src, 1)
    label = f"Top contributors {top_file.name}"
    branch_col = require_column(headers, label, *config.branch_headers)
    yearly_col = require_column(headers, label, config.yearly_total_header)
    tin_col = require_column(headers, label, *config.tin_headers)
    taxpayer_col = find_column(headers, *config.taxpayer_headers)
    last_row = last_row_in_col(src, tin_col)
    last_col = last_header_col(src, 1)

    selected_rows, branch_values = [], []
    total = 0.0
    for row in range(2, last_row + 1):
        record = [src.cell(row, c).value for c in range(1, last_col + 1)]
        if is_grand_total_row(record, taxpayer_col - 1 if taxpayer_col else None):
            continue
        if is_blank(record[tin_col - 1]):
            continue
        branches = to_number(record[branch_col - 1])
        if branches >= config.min_branches:
            selected_rows.append(row)
            branch_values.append(branches)
            total += to_number(record[yearly_col - 1])

    out_wb = Workbook()
    out = out_wb.active
    out.title = f"{config.min_branches}+ BRANCHES"
    copy_row(src, 1, out, 1, last_col)
    for out_row, src_row in enumerate(selected_rows, 2):
        copy_row(src, src_row, out, out_row, last_col)

    last_data_row = 1 + len(selected_rows)
    gt_row = last_data_row + 1
    out.cell(gt_row, taxpayer_col or 1, "GRAND TOTAL")
    out.cell(gt_row, yearly_col, total)
    if selected_rows:
        for col in range(1, last_col + 1):
            copy_style(out.cell(last_data_row, col), out.cell(gt_row, col))
    for col in range(1, last_col + 1):
        out.cell(gt_row, col).font = Font(bold=True)
    set_range_format(out, yearly_col, 2, gt_row, "#,##0.00")
    if selected_rows:
        set_range_format(out, branch_col, 2, last_data_row, "0")
        out.auto_filter.ref = f"A1:{out.cell(last_data_row, last_col).coordinate}"
    copy_column_widths(src, out, last_col, cap=40)

    write_summary_sheet(out_wb, [
        ("DISTRICT", district),
        ("Minimum Branches", config.min_branches),
        ("Taxpayers Extracted", len(selected_rows)),
        ("Total YEARLY TOTAL", total),
    ], {4: "#,##0.00"})

    save_atomically(out_wb, output_path)
    src_wb.close()

    return BranchFilterResult(
        district, output_path, len(selected_rows), total,
        min(branch_values) if branch_values else None,
        max(branch_values) if branch_values else None,
        time.perf_counter() - start,
    )
