"""Pipeline settings."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from openpyxl.utils import column_index_from_string


@dataclass(frozen=True)
class PipelineConfig:
    input_dir: Path
    output_dir: Path
    overwrite: bool = False

    # Stage 1 - consolidation
    yearly_sheet: str = "YEARLY TOTAL"
    tin_headers: tuple[str, ...] = ("TIN", "TAXPAYER_TIN")
    taxpayer_headers: tuple[str, ...] = ("TAXPAYER_NAME", "TAXPAYER")
    yearly_total_header: str = "YEARLY TOTAL"
    micro_value_header: str = "IT GROSS SALES"
    sub50m_value_header: str = "TOTAL GROSS SALES"
    sub50m_copy_range: str = "G:AX"  # columns copied when the sub-50M source wins

    # Stage 2 - top contributors
    target_percent: float = 0.90

    # Stage 3 - branch filter
    min_branches: int = 7
    branch_headers: tuple[str, ...] = (
        "BRANCHES", "NO. OF BRANCHES", "NUMBER OF BRANCHES", "NO. OF REGISTERED BRANCHES",
    )

    @property
    def yearly_dir(self) -> Path:
        return self.input_dir / "yearly"

    @property
    def micro_dir(self) -> Path:
        return self.input_dir / "micro"

    @property
    def sub50m_dir(self) -> Path:
        return self.input_dir / "sub50m"

    @property
    def combined_dir(self) -> Path:
        return self.output_dir / "01_combined"

    @property
    def top_dir(self) -> Path:
        return self.output_dir / "02_top_contributors"

    @property
    def branches_dir(self) -> Path:
        return self.output_dir / "03_branch_filter"

    @property
    def sub50m_cols(self) -> tuple[int, int]:
        start, end = self.sub50m_copy_range.split(":")
        return column_index_from_string(start), column_index_from_string(end)

    @property
    def percent_label(self) -> str:
        return f"{self.target_percent:.0%}"
