"""Shared helpers: header lookup, number/TIN parsing, safe saving, style copying."""

from __future__ import annotations

import os
import re
import tempfile
from copy import copy
from pathlib import Path
from typing import Any, Iterable

from openpyxl import Workbook


class PipelineError(Exception):
    """Input content prevents a stage from running."""


def clean_header(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).strip().upper())


def to_number(value: Any) -> float:
    """Float conversion: blanks and non-numeric text become 0; commas/currency signs ignored."""
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if not text:
        return 0.0
    text = re.sub(r"[,\s₱$€£]", "", text)
    try:
        return float(text)
    except ValueError:
        return 0.0


def normalize_tin(value: Any) -> str:
    """Normalise a TIN to its 9-digit base for matching.

    123-456-789-000 -> 123456789   (branch suffix dropped)
    12345678 (number, leading zero lost) -> 012345678
    """
    if value is None:
        return ""
    if isinstance(value, float):
        value = str(int(value)) if value.is_integer() else str(value)
    else:
        value = str(value).strip()
    digits = re.sub(r"\D", "", value)
    if not digits:
        return ""
    if len(digits) < 9:
        digits = digits.zfill(9)
    return digits[:9]


def is_blank(value: Any) -> bool:
    return value is None or str(value).strip() == ""


def header_map(ws, header_row: int) -> dict[str, int]:
    """{CLEAN_HEADER: 1-based column} for the header row (last duplicate wins)."""
    result = {}
    for col in range(1, last_header_col(ws, header_row) + 1):
        header = clean_header(ws.cell(header_row, col).value)
        if header:
            result[header] = col
    return result


def last_header_col(ws, header_row: int) -> int:
    """Right-most non-empty column of the header row (Excel's End(xlToLeft))."""
    last = 0
    for cell in next(ws.iter_rows(min_row=header_row, max_row=header_row)):
        if not is_blank(cell.value):
            last = cell.column
    return last


def last_row_in_col(ws, col: int) -> int:
    """Bottom-most non-empty row of a column (Excel's End(xlUp))."""
    for row in range(ws.max_row, 0, -1):
        if not is_blank(ws.cell(row, col).value):
            return row
    return 1


def find_column(headers: dict[str, int], *names: str) -> int | None:
    for name in names:
        col = headers.get(clean_header(name))
        if col:
            return col
    return None


def require_column(headers: dict[str, int], label: str, *names: str) -> int:
    col = find_column(headers, *names)
    if col is None:
        raise PipelineError(f"{label}: column {' / '.join(names)} not found.")
    return col


def find_header_row(ws, text: str, max_rows: int = 20, max_cols: int = 120) -> int | None:
    """First row (top-down) in which any cell equals ``text``."""
    target = clean_header(text)
    for row in ws.iter_rows(min_row=1, max_row=max_rows, max_col=max_cols):
        for cell in row:
            if clean_header(cell.value) == target:
                return cell.row
    return None


def is_grand_total_row(values: list[Any], taxpayer_index: int | None) -> bool:
    """True if column A or the taxpayer column holds a TOTAL / GRAND TOTAL label."""
    indexes = [0] + ([taxpayer_index] if taxpayer_index is not None else [])
    for index in indexes:
        if index >= len(values) or values[index] is None:
            continue
        text = str(values[index]).strip().upper()
        if text in ("GRAND TOTAL", "TOTAL") or "GRAND TOTAL" in text:
            return True
    return False


def copy_style(src, dst) -> None:
    if src.has_style:
        dst.font = copy(src.font)
        dst.fill = copy(src.fill)
        dst.border = copy(src.border)
        dst.alignment = copy(src.alignment)
        dst.number_format = src.number_format
        dst.protection = copy(src.protection)


def copy_row(src_ws, src_row: int, dst_ws, dst_row: int, last_col: int, styles: bool = True) -> None:
    for col in range(1, last_col + 1):
        src = src_ws.cell(src_row, col)
        dst = dst_ws.cell(dst_row, col, src.value)
        if styles:
            copy_style(src, dst)


def copy_column_widths(src_ws, dst_ws, last_col: int, cap: float = 40) -> None:
    from openpyxl.utils import get_column_letter

    for col in range(1, last_col + 1):
        letter = get_column_letter(col)
        width = src_ws.column_dimensions[letter].width
        if width:
            dst_ws.column_dimensions[letter].width = min(width, cap)


def set_range_format(ws, col: int, first_row: int, last_row: int, number_format: str) -> None:
    for row in range(first_row, last_row + 1):
        ws.cell(row, col).number_format = number_format


def write_summary_sheet(wb: Workbook, rows: Iterable[tuple[str, Any]], formats: dict[int, str]) -> None:
    from openpyxl.styles import Font

    ws = wb.create_sheet("SUMMARY", 0)
    for r, (label, value) in enumerate(rows, 1):
        ws.cell(r, 1, label).font = Font(bold=True)
        ws.cell(r, 2, value)
    for r, fmt in formats.items():
        ws.cell(r, 2).number_format = fmt
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 24


def prepare_output(path: Path, overwrite: bool) -> Path:
    if path.exists() and not overwrite:
        raise FileExistsError(f"Output already exists (use --overwrite): {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    return path


def save_atomically(wb: Workbook, target: Path) -> None:
    """Write to a temp file then rename, so failures never leave partial outputs."""
    fd, tmp = tempfile.mkstemp(suffix=".xlsx", dir=target.parent)
    os.close(fd)
    try:
        wb.save(tmp)
        os.replace(tmp, target)
    except BaseException:
        Path(tmp).unlink(missing_ok=True)
        raise
