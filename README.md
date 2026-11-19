# Tax Data Consolidation & Analysis Pipeline

A three-stage Python pipeline that:
1. combines three taxpayer datasets by TIN, keeping the highest reported figure for each taxpayer;
2. finds the taxpayers who together make up 90% of total collections;
3. from those, picks out the taxpayers with 7 or more branches.

> **About this repository.** This is a portfolio rewrite of three scripts I wrote and used
> together as one workflow. The originals automated Excel through Windows COM. This
> version uses `openpyxl`, so it runs on any OS without Excel installed. All data is
> synthetic, and no organisation-specific paths, codes or records are included.

## Demo video

▶ **[Watch the demo (about 2 minutes, Google Drive)](https://drive.google.com/file/d/1LJjWh5YUM_TpuZs6vwcTZS_Xl_fzVg1L/view?usp=sharing)**

The video follows one fictional taxpayer through all three stages: TIN consolidation, the
90% cumulative collection cut-off, and the 7-or-more-branches filter. It ends with the
generated reports and the test results. Every dataset shown is fictional, and the
automation is started by hand with a single command; it does not run on a schedule.
Spreadsheet and terminal views are labelled reconstructions that use values from the
actual files and captured output. The narration is AI-generated.

## Business problem

Collections for each district come from three separate datasets:

| Dataset | Contents | Value used |
|---|---|---|
| `yearly/` | Main yearly list, one row per taxpayer, with branch and POS counts | `YEARLY TOTAL` |
| `micro/` | Micro-business returns. Title rows sit above the header. | `IT GROSS SALES` |
| `sub50m/` | Returns from businesses under 50M. Metadata is in A:F and the return fields in G:AX. | `TOTAL GROSS SALES` |

The same taxpayer can appear in all three files with its TIN written differently
(`123-456-789-000`, `123456789`, `12345678` after Excel drops the leading zero, or a
14-digit form with a branch code). Matching these by hand with lookups, picking the
largest figure, ranking taxpayers by contribution and filtering by branch count takes
many steps and is easy to get wrong. The pipeline does all of it in one command.

## Pipeline

**Stage 1: Consolidate** (`tax_pipeline/consolidate.py`) → `output/01_combined/<D>_COMBINED_BY_TIN.xlsx`
- **TIN normalisation:** keep digits only, left-pad to 9 digits, and use the first 9 (the head-office TIN).
- **Duplicate TINs** in a secondary file: the row with the highest value is kept.
- **Value per taxpayer:** the new `YEARLY TOTAL` is the highest of Yearly, Micro and Sub-50M. Only sources whose TIN matched are compared.
- **Tie-breaking:**
  - Yearly is kept if it is greater than or equal to both other values.
  - If Micro and Sub-50M tie, Sub-50M wins.
- **Audit columns added:** `ORIGINAL YEARLY TOTAL`, `YEARLY TOTAL SOURCE`, `YEARLY TOTAL SOURCE VALUE`.
- **Source data copied:** the winning row's data is appended — the full Micro row, or the Sub-50M columns G:AX.

**Stage 2: Top contributors** (`top_contributors.py`) → `output/02_top_contributors/<D>_90_PERCENT_COLLECTION.xlsx`
- Skips `TOTAL` / `GRAND TOTAL` rows and rows without a TIN.
- Sorts by `YEARLY TOTAL`, highest first. The sort is stable, so ties keep their original row order.
- Selects taxpayers until the running total is at least 90% of the overall total.
- Output has a `SUMMARY` sheet (total, target, selected total, actual coverage) and a data sheet with formatting copied from the source and a GRAND TOTAL row.

**Stage 3: Branch filter** (`branch_filter.py`) → `output/03_branch_filter/<D>_7_OR_MORE_BRANCHES.xlsx`
- Keeps the stage 2 taxpayers whose `BRANCHES` value is 7 or more.
- Also accepts the headers `NO. OF BRANCHES`, `NUMBER OF BRANCHES` and `NO. OF REGISTERED BRANCHES`.
- Adds a GRAND TOTAL row and a `SUMMARY` sheet.

**Across all stages:**
- Required columns are checked; a clear error names any that are missing.
- Each district is processed independently. If one fails, the error is logged and the rest still run.
- Files are written to a temporary file first, so a failure never leaves a partial output.
- Existing outputs are kept unless you pass `--overwrite`.
- Every stage logs its counts, totals and timing.
- Exit codes: `0` success, `1` setup error, `2` at least one district failed.

## Quick start

```bash
python -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -r requirements.txt

python scripts/generate_sample_data.py --output-dir sample_data/input
python -m tax_pipeline --input-dir sample_data/input --output-dir output
```

Options: `--districts NORTH SOUTH`, `--target-percent 0.9`, `--min-branches 7`,
`--overwrite`, `--log-file PATH`, `-v`.

Expected input layout: `<input>/yearly/<D>.xlsx`, `<input>/micro/<D>.xlsx` and
`<input>/sub50m/<D>.xlsx`, using the same district code `<D>` for all three.

## Example result (included sample data)

From [`examples/output/run.log`](examples/output/run.log), two districts with 400 synthetic taxpayers each:

| District | Values replaced (Micro / Sub-50M) | Selected for 90% | Coverage | 7+ branches |
|---|---:|---:|---:|---:|
| NORTH | 81 / 70 | 232 of 400 | 90.06% | 27 |
| SOUTH | 72 / 61 | 245 of 400 | 90.03% | 32 |

All three stages for both districts took 1.4 s in total. In a separate run, three districts of
5,000 synthetic taxpayers each took 25 s (Linux container, 4 vCPU). Most of that time goes on
copying cell formatting in stage 2. These are processing times on fictional data, not measurements of
time saved in real work.

## Testing

```bash
pip install -r requirements-dev.txt
python -m pytest -v
```

19 tests cover:
- TIN normalisation edge cases;
- every stage 1 rule: Micro win, Sub-50M tie priority, Yearly kept on a tie, negative value with no match, rows without a TIN, duplicate TINs, numbers stored as text, the G:AX copy range, and audit columns;
- no output written when a stage fails;
- the 90% cut-off, including exact hits, ties, a zero total and empty input;
- stage 2 and stage 3 outputs;
- an end-to-end run compared with an independent re-implementation of the rules;
- overwrite protection and exit codes.

During development, I compared the rewrite's outputs with those of my original scripts,
run on the same synthetic data, and the cell values in all three outputs matched exactly.
That was a one-off comparison made earlier. The original scripts are not in this
repository, so it cannot be repeated from here. The automated tests above check the
rewrite against an independent re-implementation of the rules instead.

## Changes from the original workflow

- **Fixed:** a taxpayer with a negative Yearly value and no Micro/Sub-50M match was set to `0`
  while still being labelled `YEARLY`. Now only sources that actually matched are compared.
- **Fixed:** a failed run could save a half-updated workbook. Outputs are now written atomically.
- Each secondary file is read once instead of twice.
- Duplicate TINs in the Yearly file are reported in the log.
- Paths, district codes, thresholds and column names are configurable instead of hard-coded.

## Known limitations

- Truncating to 9 digits deliberately merges branch TINs into the head-office TIN. Branch-level
  matching would need a different key.
- Taxpayers that exist only in Micro or Sub-50M are not added. The Yearly file defines who is included.
- The 7+ branch filter runs only on the 90% subset, not on every taxpayer.
- Branch counts stored as text such as `"7 branches"` are read as 0.
- Headers in the Yearly file must be in row 1. Secondary files must have a TIN header within the first 20 rows.
- Formatting is copied cell by cell, so very large files are slower than a values-only export would be.
- Column widths in stage 1 are estimated from text length, not measured the way Excel's AutoFit does.
