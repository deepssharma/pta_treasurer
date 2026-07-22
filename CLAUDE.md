# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this repo does

Generates a PTA's monthly treasurer report (Excel, 6 tabs) and a whole-fiscal-year
Debits & Credits ledger from three input files per month:
- QuickBooks Transaction Detail export (CSV)
- Chase bank statement (PDF)
- Givebacks payout export(s) (CSV)

The pipeline lives in `PTA_Treasurer_Report_v4.ipynb`, driven either interactively
(Jupyter) or headlessly via `run_all_months.sh` / `jupyter nbconvert`.

## Commands

Run the test suite:
```
pytest
```
Run a single test file or test:
```
pytest tests/test_builders.py
pytest tests/test_builders.py::test_name -q
```

Run one month headlessly (no Jupyter UI):
```
BATCH_MODE=1 INPUT_MONTH=March FISCAL_YEAR=2026 \
  jupyter nbconvert --to notebook --execute \
    --ExecutePreprocessor.timeout=300 \
    --output output/executed_March_2026.ipynb \
    PTA_Treasurer_Report_v4.ipynb
```
`BATCH_MODE=1` suppresses the interactive Givebacks-URL prompt (Cell 2) and the
GitHub-push cell (Cell 11) — both only make sense in an interactive run.

Run every configured month in sequence (edit the `MONTHS` array in the script to
change which months are processed):
```
./run_all_months.sh
```
Each month writes `output/Treasurer_Report_{Month}_{Year}.xlsx` and a log to
`logs/{Month}_{Year}.log`; months whose `input/{Month}_{Year}/` folder is missing
or has no QuickBooks CSV are skipped with a warning, not a failure.

Try the pipeline safely against the synthetic dataset (`sample_data/July_1999/` —
fake org, fake year, no real data) instead of writing test prompts from scratch —
see `.claude/skills/generate-demo-report/SKILL.md` for the exact steps.

## Architecture

**Pipeline shape:** `parsers.py` (read raw exports → structured dicts) →
notebook Cell 5 (merge current month into `data/history/{Month}_{Year}.json`,
then rebuild fiscal-year-to-date actuals from all history files) → `builders.py`
(structured data → styled `openpyxl` worksheets) → notebook Cells 9/12 (assemble
worksheets into the two output workbooks).

**`parsers.py`** — one function per input file type, each returns plain
dicts/lists with no Excel or org-specific knowledge:
- `parse_quickbooks_detail(folder, month, year)` → transactions, each tagged
  `is_income` (used downstream to split into credits/debits)
- `parse_givebacks_files(file_info_list)` — merges multiple payout CSVs by item
- `parse_chase_pdf(bank_file)` — bank statement for reconciliation

**`builders.py`** — one `build_*` function per worksheet tab, each takes an
`openpyxl` worksheet plus already-parsed data and org name, and does no parsing.
Two independent helper families back these, because the sheets have two
different shapes:
- `_sec_hdr`/`_col_hdrs`/`_data_row`/`_total_row` — fixed 4-column sheets
  (Treasurer Report, Budget vs Actuals, Giveback Reconciliation, YTD Summary)
- `_wide_hdr_row`/`_wide_col_hdrs`/`_month_band`/`_wide_data_row`/`_wide_total_row`
  — wide (5–9 column), month-sectioned ledger sheets (Credits, Debits,
  MemberHub_Summary in the Debits & Credits workbook), sectioned with a gold
  month band and a running-total column

**Two output workbooks per run:**
1. `output/Treasurer_Report_{Month}_{Year}.xlsx` (Cell 9) — one month's report:
   Treasurer Report (with bank reconciliation), Income/Expense Budget vs
   Actuals, Giveback Reconciliation, File Manifest, YTD Summary (accumulates
   across all months processed so far via `data/history/`).
2. `output/Debits_and_Credits_{fy_start}_to_{fy_end}.xlsx` (Cell 12) — rebuilt
   from scratch every run from *every* `input/{Month}_{Year}/` folder that has
   a QuickBooks file (independent of the `INPUT_MONTH`/`FISCAL_YEAR` config);
   a running check-register-style ledger, sectioned by month. `NOTES` is
   intentionally left blank — there's no source field for hand-written
   reconciliation notes.

**Category mapping:** `QB_TO_BUDGET_MAP` (notebook Cell 6) maps QuickBooks
category strings to budget line names; this dict plus `INCOME_BUDGET`/
`EXPENSE_BUDGET` hold this PTA's real budget figures and are edited once per
fiscal year, in the notebook itself, not in `builders.py`/`parsers.py`.

**Config (notebook Cell 1):** `ORG_NAME`, `INPUT_MONTH`, `FISCAL_YEAR`, and
Givebacks credentials come from environment variables / a local `.env` file,
defaulting to this PTA's own values — this is what makes the notebook
org-specific even though `parsers.py`/`builders.py` themselves take `org_name`
as a parameter and have no hardcoded org strings.

## Data safety — read before touching input/output/data/files

This is a **public** GitHub repo. `input/`, `output/`, `data/`, `files/`,
`logs/` all hold real financial data (names, bank account numbers, dollar
amounts) and are gitignored, including as symlinks (`data` and `files` are
local symlinks to storage outside the repo — matched by bare `data`/`files`
patterns in `.gitignore`, not just `data/`/`files/`, since that gap is exactly
how they got committed once before). Never commit, push, or paste the
contents of these paths into a chat, PR, or issue.

Only `sample_data/July_1999/` (synthetic, fake org/year) and
`tests/fixtures/*` (synthetic, anonymized) are real data-shaped files that are
meant to be tracked in git.

Before committing `PTA_Treasurer_Report_v4.ipynb` after running it against
real data, clear cell outputs first (`Edit → Clear Outputs of All Cells` in
Jupyter) — saved outputs are part of the `.ipynb` file and are not covered by
`.gitignore`.
