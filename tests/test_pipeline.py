"""Tests for the three-stage pipeline (all data is synthetic)."""

from __future__ import annotations

import sys
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "scripts"))

from tax_pipeline.branch_filter import branch_filter_district  # noqa: E402
from tax_pipeline.cli import main  # noqa: E402
from tax_pipeline.common import PipelineError, normalize_tin, to_number  # noqa: E402
from tax_pipeline.config import PipelineConfig  # noqa: E402
from tax_pipeline.consolidate import consolidate_district  # noqa: E402
from tax_pipeline.top_contributors import (  # noqa: E402
    select_top_contributors,
    top_contributors_district,
)

YEARLY_HEADERS = ["DISTRICT", "TIN", "TAXPAYER_NAME", "BRANCHES", "YEARLY TOTAL"]


def save(path: Path, rows: list[list], title: str | None = None) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    if title:
        ws.title = title
    for row in rows:
        ws.append(row)
    wb.save(path)
    return path


def sub50m_row(tin, gross, marker) -> list:
    # A:F metadata, G = TOTAL GROSS SALES, then 43 more fields up to AX
    return [1, tin, "NAME", "ADDR", "2025", "FORM"] + [gross, marker] + [None] * 42


SUB50M_HEADERS = ["SEQ", "TIN", "TAXPAYER_NAME", "ADDRESS", "PERIOD", "FORM",
                  "TOTAL GROSS SALES", "MARKER"] + [f"F{i}" for i in range(42)] + ["NOT COPIED"]


@pytest.fixture
def inputs(tmp_path: Path) -> PipelineConfig:
    cfg = PipelineConfig(input_dir=tmp_path / "in", output_dir=tmp_path / "out")
    save(cfg.yearly_dir / "T.xlsx", [
        YEARLY_HEADERS,
        ["D", "111-111-111-000", "Micro Wins", 7, 100],
        ["D", "222-222-222-000", "Sub50M Wins Tie", 9, 100],
        ["D", "333-333-333-000", "Yearly Kept On Tie", 2, 300],
        ["D", "444-444-444-000", "Negative No Match", 12, -50],
        ["D", None, "No TIN", 8, 999],
        ["D", "055-555-555-000", "Duplicate Micro", 7, 10],
        ["D", "666-666-666", "Text Number", "8", "1,000.50"],
        [None, None, "GRAND TOTAL", None, 1460.5],
    ], title="YEARLY TOTAL")
    save(cfg.micro_dir / "T.xlsx", [
        ["MICRO RETURNS"], [],
        ["TIN", "TAXPAYER_NAME", "IT GROSS SALES"],
        ["111111111", "MICRO WINS", 150],
        [222222222, "SUB50M WINS TIE", 200],
        ["333-333-333-00000", "YEARLY KEPT", 300],
        [55555555, "DUP LOW", 20],          # lost leading zero
        ["055555555", "DUP HIGH", 500],
        ["055-555-555-001", "DUP LOWER", 400],
    ])
    save(cfg.sub50m_dir / "T.xlsx", [
        ["SUB-50M"],
        SUB50M_HEADERS,
        sub50m_row("111-111-111-000", 120, "A"),
        sub50m_row("222222222", 200, "B"),
    ])
    return cfg


def run_stage1(cfg: PipelineConfig):
    return consolidate_district("T", cfg.yearly_dir / "T.xlsx", cfg.micro_dir / "T.xlsx",
                                cfg.sub50m_dir / "T.xlsx", cfg)


# ---------------------------------------------------------------------------
@pytest.mark.parametrize("raw, expected", [
    ("007-265-475", "007265475"),
    (7265475.0, "007265475"),
    (7265475, "007265475"),
    ("123-456-789-00000", "123456789"),
    ("", ""),
    (None, ""),
    ("n/a", ""),
])
def test_normalize_tin(raw, expected):
    assert normalize_tin(raw) == expected


def test_to_number():
    assert to_number("1,000.50") == 1000.5
    assert to_number("$ 25") == 25
    assert to_number("abc") == 0
    assert to_number(None) == 0


def test_stage1_highest_value_rules_and_audit_columns(inputs):
    r = run_stage1(inputs)

    assert (r.taxpayers, r.yearly_wins, r.micro_wins, r.sub50m_wins, r.replaced) == (6, 3, 2, 1, 3)
    assert r.original_total == pytest.approx(100 + 100 + 300 - 50 + 10 + 1000.5)
    assert r.updated_total == pytest.approx(150 + 200 + 300 - 50 + 500 + 1000.5)

    ws = load_workbook(r.output_path)["YEARLY TOTAL"]
    rows = list(ws.iter_rows(values_only=True))
    header = rows[0]
    assert header[5:8] == ("ORIGINAL YEARLY TOTAL", "YEARLY TOTAL SOURCE", "YEARLY TOTAL SOURCE VALUE")
    assert header[8:11] == ("MICRO | TIN", "MICRO | TAXPAYER_NAME", "MICRO | IT GROSS SALES")
    assert header[11] == "50M | TOTAL GROSS SALES" and header[12] == "50M | MARKER"
    assert len(header) == 11 + 44  # G:AX = 44 columns, 'NOT COPIED' (AY) excluded

    by_name = {row[2]: row for row in rows[1:]}
    micro = by_name["Micro Wins"]
    assert micro[4] == 150 and micro[5] == 100 and micro[6] == "MICRO - IT GROSS SALES"
    assert micro[8:11] == ("111111111", "MICRO WINS", 150) and micro[11] is None

    tie = by_name["Sub50M Wins Tie"]
    assert tie[4] == 200 and tie[6] == "50M - TOTAL GROSS SALES"
    assert tie[8] is None and tie[11:13] == (200, "B")

    assert by_name["Yearly Kept On Tie"][4:7] == (300, 300, "YEARLY")
    # Fixed edge case: no match must not turn a negative yearly value into 0.
    assert by_name["Negative No Match"][4:8] == (-50, -50, "YEARLY", -50)
    assert by_name["No TIN"][4] == 999 and by_name["No TIN"][5] is None
    dup = by_name["Duplicate Micro"]
    assert dup[4] == 500 and dup[9] == "DUP HIGH"
    assert by_name["Text Number"][4] == 1000.5
    assert by_name["GRAND TOTAL"][5] is None  # rows without TIN are left alone


def test_stage1_failure_writes_no_output(inputs):
    save(inputs.micro_dir / "T.xlsx", [["TIN", "WRONG COLUMN"], ["111111111", 1]])
    with pytest.raises(PipelineError, match="IT GROSS SALES"):
        run_stage1(inputs)
    assert not (inputs.combined_dir / "T_COMBINED_BY_TIN.xlsx").exists()


@pytest.mark.parametrize("values, expected_idx", [
    ([50, 30, 10, 5, 5], [0, 1, 2]),          # reaches exactly 90
    ([10, 30, 50, 5, 5], [2, 1, 0]),          # order of input does not matter
    ([40, 40, 10, 10], [0, 1, 2]),            # 80 < 90 -> needs one of the 10s (first one, stable)
    ([100], [0]),
    ([0, 0], [0]),                             # zero total: first record satisfies 0 >= 0
    ([], []),
])
def test_select_top_contributors(values, expected_idx):
    idx, target, total = select_top_contributors(values, 0.9)
    assert idx == expected_idx
    assert target == pytest.approx(sum(values) * 0.9)
    assert total == pytest.approx(sum(values[i] for i in idx))


def test_stage2_and_stage3_outputs(inputs):
    combined = run_stage1(inputs).output_path
    t = top_contributors_district("T", combined, inputs)
    # updated values: 150,200,300,-50,500,1000.5 -> total 2100.5, target 1890.45
    # sorted: 1000.5, 500, 300, 200 -> 2000.5 >= target after 4
    assert (t.total_taxpayers, t.selected_taxpayers) == (6, 4)
    assert t.selected_total == pytest.approx(2000.5)
    assert t.coverage == pytest.approx(2000.5 / 2100.5)

    wb = load_workbook(t.output_path)
    assert wb.sheetnames == ["SUMMARY", "90% COLLECTION"]
    data = list(wb["90% COLLECTION"].iter_rows(values_only=True))
    assert [r[2] for r in data[1:5]] == ["Text Number", "Duplicate Micro", "Yearly Kept On Tie",
                                         "Sub50M Wins Tie"]
    assert data[5][2] == "GRAND TOTAL" and data[5][4] == pytest.approx(2000.5)
    assert wb["90% COLLECTION"]["C6"].font.bold
    summary = {r[0]: r[1] for r in wb["SUMMARY"].iter_rows(values_only=True)}
    assert summary["Selected Taxpayers"] == 4 and summary["Actual Coverage"] == pytest.approx(t.coverage)

    b = branch_filter_district("T", t.output_path, inputs)
    # branches among selected: Text Number "8", Duplicate Micro 7, Yearly Kept 2, Sub50M 9
    assert b.count == 3
    assert b.yearly_total == pytest.approx(1000.5 + 500 + 200)
    assert (b.min_branches_found, b.max_branches_found) == (7, 9)
    wb = load_workbook(b.output_path)
    assert wb.sheetnames == ["SUMMARY", "7+ BRANCHES"]
    rows = list(wb["7+ BRANCHES"].iter_rows(values_only=True))
    assert len(rows) == 5 and rows[-1][2] == "GRAND TOTAL"


def test_cli_end_to_end_matches_independent_reference(tmp_path):
    import generate_sample_data as gen

    gen.generate(tmp_path / "in", taxpayers=150, districts=["X1"], seed=11)
    assert main(["--input-dir", str(tmp_path / "in"), "--output-dir", str(tmp_path / "out")]) == 0

    # Independent re-implementation of the business rules.
    def index(path, header_row, value_col_name):
        rows = list(load_workbook(path).active.iter_rows(min_row=header_row, values_only=True))
        h = list(rows[0])
        tin_i, val_i = h.index("TIN"), h.index(value_col_name)
        out = {}
        for r in rows[1:]:
            tin = normalize_tin(r[tin_i])
            if tin and (tin not in out or r[val_i] > out[tin]):
                out[tin] = r[val_i]
        return out

    micro = index(tmp_path / "in/micro/X1.xlsx", 3, "IT GROSS SALES")
    sub = index(tmp_path / "in/sub50m/X1.xlsx", 2, "TOTAL GROSS SALES")
    people = []
    for r in list(load_workbook(tmp_path / "in/yearly/X1.xlsx").active.iter_rows(values_only=True))[1:]:
        if not r[1]:
            continue
        tin = normalize_tin(r[1])
        value = max([r[6]] + [d[tin] for d in (micro, sub) if tin in d])
        people.append((value, r[4]))

    total = sum(v for v, _ in people)
    ranked = sorted(people, key=lambda p: p[0], reverse=True)
    selected, running = [], 0.0
    for p in ranked:
        selected.append(p)
        running += p[0]
        if running >= total * 0.9:
            break
    with_branches = [p for p in selected if p[1] >= 7]

    s1 = {r[0]: r[1] for r in load_workbook(tmp_path / "out/02_top_contributors/X1_90_PERCENT_COLLECTION.xlsx")["SUMMARY"].iter_rows(values_only=True)}
    assert s1["Original Total Collection"] == pytest.approx(total)
    assert s1["Selected Taxpayers"] == len(selected)
    assert s1["Selected Collection"] == pytest.approx(running)
    s2 = {r[0]: r[1] for r in load_workbook(tmp_path / "out/03_branch_filter/X1_7_OR_MORE_BRANCHES.xlsx")["SUMMARY"].iter_rows(values_only=True)}
    assert s2["Taxpayers Extracted"] == len(with_branches)
    assert s2["Total YEARLY TOTAL"] == pytest.approx(sum(v for v, _ in with_branches))


def test_cli_overwrite_protection_and_exit_codes(tmp_path, inputs):
    args = ["--input-dir", str(inputs.input_dir), "--output-dir", str(inputs.output_dir)]
    assert main(args) == 0
    assert main(args) == 2                      # outputs exist, not overwritten
    assert main(args + ["--overwrite"]) == 0
    assert main(["--input-dir", str(tmp_path / "nope"), "--output-dir", str(tmp_path)]) == 1
    assert main(args + ["--districts", "MISSING"]) == 2
