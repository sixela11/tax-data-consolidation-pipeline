"""Command-line entry point.

    python -m tax_pipeline --input-dir sample_data/input --output-dir output
"""

from __future__ import annotations

import argparse
import logging
import time
from pathlib import Path

from .branch_filter import branch_filter_district
from .common import PipelineError
from .config import PipelineConfig
from .consolidate import consolidate_district
from .top_contributors import top_contributors_district

log = logging.getLogger("tax_pipeline")


def parse_args(argv: list[str] | None) -> argparse.Namespace:
    p = argparse.ArgumentParser(
        prog="python -m tax_pipeline",
        description="Consolidate yearly/micro/sub-50M data by TIN, find the top contributors "
                    "to 90%% of collections, and filter those with 7+ branches.",
    )
    p.add_argument("--input-dir", type=Path, required=True,
                   help="Folder containing yearly/, micro/ and sub50m/ subfolders.")
    p.add_argument("--output-dir", type=Path, required=True)
    p.add_argument("--districts", nargs="+",
                   help="District codes to process (default: every file in yearly/).")
    p.add_argument("--target-percent", type=float, default=0.90)
    p.add_argument("--min-branches", type=int, default=7)
    p.add_argument("--overwrite", action="store_true", help="Replace existing outputs.")
    p.add_argument("--log-file", type=Path)
    p.add_argument("-v", "--verbose", action="store_true")
    args = p.parse_args(argv)
    if not 0 < args.target_percent <= 1:
        p.error("--target-percent must be in (0, 1].")
    return args


def setup_logging(verbose: bool, log_file: Path | None) -> None:
    handlers: list[logging.Handler] = [logging.StreamHandler()]
    if log_file:
        log_file.parent.mkdir(parents=True, exist_ok=True)
        handlers.append(logging.FileHandler(log_file, encoding="utf-8"))
    logging.basicConfig(level=logging.DEBUG if verbose else logging.INFO,
                        format="%(asctime)s %(levelname)-7s %(message)s", datefmt="%H:%M:%S",
                        handlers=handlers, force=True)


def discover_districts(config: PipelineConfig) -> list[str]:
    return sorted(p.stem for p in config.yearly_dir.glob("*.xlsx") if not p.name.startswith("~$"))


def require_file(path: Path) -> Path:
    if not path.is_file():
        raise FileNotFoundError(f"Missing input file: {path}")
    return path


def run_district(district: str, config: PipelineConfig) -> dict:
    yearly = require_file(config.yearly_dir / f"{district}.xlsx")
    micro = require_file(config.micro_dir / f"{district}.xlsx")
    sub50m = require_file(config.sub50m_dir / f"{district}.xlsx")

    c = consolidate_district(district, yearly, micro, sub50m, config)
    log.info("  [1/3] Consolidated %s taxpayers: Yearly kept %s | Micro %s | Sub-50M %s | "
             "total %s -> %s (%.2fs)", f"{c.taxpayers:,}", f"{c.yearly_wins:,}",
             f"{c.micro_wins:,}", f"{c.sub50m_wins:,}", f"{c.original_total:,.2f}",
             f"{c.updated_total:,.2f}", c.seconds)

    t = top_contributors_district(district, c.output_path, config)
    log.info("  [2/3] %s of %s taxpayers make up %s of %s (target %s) (%.2fs)",
             f"{t.selected_taxpayers:,}", f"{t.total_taxpayers:,}", f"{t.coverage:.2%}",
             f"{t.original_total:,.2f}", config.percent_label, t.seconds)

    b = branch_filter_district(district, t.output_path, config)
    log.info("  [3/3] %s of them have %d+ branches, total %s (%.2fs)",
             f"{b.count:,}", config.min_branches, f"{b.yearly_total:,.2f}", b.seconds)
    log.info("  Outputs: %s | %s | %s", c.output_path, t.output_path, b.output_path)
    return {"district": district, "consolidation": c, "top": t, "branches": b}


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    setup_logging(args.verbose, args.log_file)
    config = PipelineConfig(input_dir=args.input_dir, output_dir=args.output_dir,
                            overwrite=args.overwrite, target_percent=args.target_percent,
                            min_branches=args.min_branches)

    for folder in (config.yearly_dir, config.micro_dir, config.sub50m_dir):
        if not folder.is_dir():
            log.error("Input folder not found: %s", folder)
            return 1

    districts = args.districts or discover_districts(config)
    if not districts:
        log.error("No yearly workbooks found in %s", config.yearly_dir)
        return 1

    log.info("Districts: %s", ", ".join(districts))
    start = time.perf_counter()
    ok, failed = [], []
    for district in districts:
        log.info("District %s", district)
        try:
            ok.append(run_district(district, config))
        except (PipelineError, FileNotFoundError, FileExistsError, KeyError) as exc:
            failed.append(district)
            log.error("  FAILED: %s", exc)
        except PermissionError as exc:
            failed.append(district)
            log.error("  FAILED: cannot access file (is it open in Excel?): %s", exc)
        except Exception:
            failed.append(district)
            log.exception("  FAILED with an unexpected error")

    log.info("Finished: %d succeeded, %d failed in %.2fs.", len(ok), len(failed),
             time.perf_counter() - start)
    if failed:
        log.info("Failed districts: %s", ", ".join(failed))
    return 0 if not failed else 2
