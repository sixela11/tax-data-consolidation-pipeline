"""Generate fictional input data for the pipeline.

Creates, per district:
    <out>/yearly/<D>.xlsx   sheet "YEARLY TOTAL", header in row 1, a GRAND TOTAL row
    <out>/micro/<D>.xlsx    title rows above the header, IT GROSS SALES
    <out>/sub50m/<D>.xlsx   title row, metadata in A:F, 44 return fields in G:AX

TINs are written in different formats across files (dashed with branch suffix,
plain digits, numbers that lost their leading zero) so matching has to normalise.
Some TINs appear twice in the secondary files. Every name, address and amount is random.

Usage:
    python scripts/generate_sample_data.py --output-dir sample_data/input
"""

from __future__ import annotations

import argparse
import random
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

DISTRICTS = ["NORTH", "SOUTH"]
WORDS_A = ["Bluefin", "Golden", "Silver", "Evergreen", "Summit", "Harbor", "Crescent", "Maple",
           "Coral", "Granite", "Sunrise", "Pioneer", "Riverside", "Orchid", "Atlas", "Cobalt"]
WORDS_B = ["Trading", "Foods", "Logistics", "Hardware", "Textiles", "Pharmacy", "Builders",
           "Motors", "Bakery", "Electronics", "Farms", "Printing", "Supplies", "Ventures"]
HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(color="FFFFFF", bold=True)
SUB50M_FIELDS = ["TOTAL GROSS SALES", "COST OF SALES", "GROSS INCOME", "OTHER INCOME",
                 "TOTAL DEDUCTIONS", "TAXABLE INCOME", "TAX DUE", "TAX CREDITS", "TAX PAYABLE"]


def style_header(ws, row: int) -> None:
    for cell in ws[row]:
        if cell.value is not None:
            cell.fill, cell.font = HEADER_FILL, HEADER_FONT


def tin_variant(rng: random.Random, digits: str):
    roll = rng.random()
    if roll < 0.4:
        return f"{digits[:3]}-{digits[3:6]}-{digits[6:]}-000"
    if roll < 0.7:
        return digits
    if roll < 0.85:
        return int(digits)  # leading zero lost when stored as a number
    return f"{digits}00000"  # 14-digit form with branch code


def generate(output_dir: Path, taxpayers: int, districts: list[str], seed: int) -> None:
    rng = random.Random(seed)
    for sub in ("yearly", "micro", "sub50m"):
        (output_dir / sub).mkdir(parents=True, exist_ok=True)

    for district in districts:
        people = []
        used: set[str] = set()
        while len(people) < taxpayers:
            digits = f"{rng.randint(0, 999_999_999):09d}"
            if digits in used:
                continue
            used.add(digits)
            # Heavy-tailed amounts, so a minority of taxpayers makes up most of the total.
            yearly = round(rng.paretovariate(1.2) * 50_000, 2)
            branches = rng.choices([0, 1, 2, 3, 5, 7, 9, 12, 20],
                                   weights=[25, 30, 15, 10, 8, 5, 3, 2, 2])[0]
            people.append({
                "digits": digits,
                "name": f"{rng.choice(WORDS_A)} {rng.choice(WORDS_B)} {len(people) + 1:04d}",
                "address": f"{rng.randint(1, 999)} Example St., Sample City",
                "branches": branches,
                "pos": branches + rng.randint(0, 3) if branches else rng.randint(0, 2),
                "yearly": yearly,
            })

        # ---- yearly ---------------------------------------------------
        wb = Workbook()
        ws = wb.active
        ws.title = "YEARLY TOTAL"
        ws.append(["DISTRICT", "TIN", "TAXPAYER_NAME", "ADDRESS", "BRANCHES", "POS_UNITS",
                   "YEARLY TOTAL"])
        style_header(ws, 1)
        for p in people:
            ws.append([f"DISTRICT {district}", f"{p['digits'][:3]}-{p['digits'][3:6]}-{p['digits'][6:]}-000",
                       p["name"], p["address"], p["branches"], p["pos"], p["yearly"]])
        ws.append([None, None, "GRAND TOTAL", None, None, None,
                   round(sum(p["yearly"] for p in people), 2)])
        for col, width in zip("ABCDEFG", (16, 20, 30, 34, 11, 11, 18)):
            ws.column_dimensions[col].width = width
        wb.save(output_dir / "yearly" / f"{district}.xlsx")

        # ---- micro ----------------------------------------------------
        wb = Workbook()
        ws = wb.active
        ws.append(["MICRO-BUSINESS RETURNS (SYNTHETIC DATA)"])
        ws.append([])
        ws.append(["TIN", "TAXPAYER_NAME", "RETURN_PERIOD", "IT GROSS SALES", "IT TAX DUE"])
        style_header(ws, 3)
        for p in rng.sample(people, k=int(len(people) * 0.4)):
            for _ in range(2 if rng.random() < 0.1 else 1):  # some duplicates
                sales = round(p["yearly"] * rng.uniform(0.3, 1.8), 2)
                ws.append([tin_variant(rng, p["digits"]), p["name"].upper(), "2025",
                           sales, round(sales * 0.03, 2)])
        wb.save(output_dir / "micro" / f"{district}.xlsx")

        # ---- sub-50M --------------------------------------------------
        wb = Workbook()
        ws = wb.active
        ws.append(["SUB-50M RETURNS (SYNTHETIC DATA)"])
        fields = SUB50M_FIELDS + [f"LINE_{i:02d}" for i in range(len(SUB50M_FIELDS) + 1, 45)]
        ws.append(["SEQ", "TIN", "TAXPAYER_NAME", "ADDRESS", "RETURN_PERIOD", "FORM"] + fields)
        style_header(ws, 2)
        for seq, p in enumerate(rng.sample(people, k=int(len(people) * 0.3)), 1):
            for _ in range(2 if rng.random() < 0.1 else 1):
                gross = round(p["yearly"] * rng.uniform(0.3, 2.0), 2)
                cost = round(gross * rng.uniform(0.4, 0.8), 2)
                taxable = round((gross - cost) * 0.7, 2)
                tax_due = round(taxable * 0.25, 2)
                line_items = [gross, cost, round(gross - cost, 2), 0, round((gross - cost) * 0.3, 2),
                              taxable, tax_due, 0, tax_due]
                line_items += [round(rng.uniform(0, 1000), 2) for _ in range(44 - len(line_items))]
                ws.append([seq, tin_variant(rng, p["digits"]), p["name"], p["address"], "2025",
                           "FORM-X"] + line_items)
        wb.save(output_dir / "sub50m" / f"{district}.xlsx")
        print(f"Created district {district}: {taxpayers} taxpayers")


def main() -> None:
    p = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    p.add_argument("--output-dir", type=Path, default=Path("sample_data/input"))
    p.add_argument("--taxpayers", type=int, default=400)
    p.add_argument("--districts", nargs="+", default=DISTRICTS)
    p.add_argument("--seed", type=int, default=2025)
    args = p.parse_args()
    generate(args.output_dir, args.taxpayers, args.districts, args.seed)


if __name__ == "__main__":
    main()
