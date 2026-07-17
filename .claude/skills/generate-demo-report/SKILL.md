---
name: generate-demo-report
description: Generate a sample PTA Treasurer monthly report + YTD summary using the synthetic demo dataset (sample_data/July_1999), to verify the report pipeline works or show someone the workflow, without touching real financial data and without pushing to GitHub. Use when asked to demo, test, or verify the report generator, or run the pipeline safely.
user-invocable: true
---

# Generate a demo report (synthetic data only)

This repo's real monthly inputs (`input/`), generated reports (`output/`), run logs
(`logs/`), and history (`data/`) all contain real PTA financial data (names, bank
account numbers, dollar amounts) and are gitignored on purpose — never commit or
push them, and never paste their contents into a chat, PR, or issue.

For demoing the tool or smoke-testing the pipeline after a code change, use the
synthetic dataset instead — it's fabricated (fake org "DEMO SCHOOL PTA", fake
names/addresses/account numbers) but shaped exactly like a real month, so it
exercises the same code paths safely.

## Synthetic dataset

`sample_data/July_1999/` — tracked in git:
- `quickbooks_july_1999.csv`
- `Chase_july_1999.pdf`
- `givebacks/givebacks_july_po_DEMO0000000000000000.csv`

The year 1999 is deliberately not a real fiscal year used by this PTA, so it can
never be confused with an actual report.

## Steps

1. Stage the synthetic month where the pipeline expects it — a local, gitignored
   copy under `input/`. Never edit files under `sample_data/` in place; treat
   that directory as the source of truth and copy from it.
   ```sh
   mkdir -p input/July_1999
   cp -r sample_data/July_1999/* input/July_1999/
   ```
2. Run the notebook headlessly against it, in batch mode (suppresses the
   interactive Givebacks-URL prompt and the GitHub-push cell):
   ```sh
   BATCH_MODE=1 INPUT_MONTH=July FISCAL_YEAR=1999 \
     jupyter nbconvert --to notebook --execute \
       --ExecutePreprocessor.timeout=300 \
       --output output/executed_demo_July_1999.ipynb \
       PTA_Treasurer_Report_v4.ipynb
   ```
3. Confirm `output/Treasurer_Report_July_1999.xlsx` was created with all 6 tabs
   (Treasurer Report, Income Budget vs Actuals, Expense Budget vs Actuals,
   Giveback Reconciliation, File Manifest, YTD Summary).
4. Optionally clean up afterwards — harmless to leave since it's gitignored:
   ```sh
   rm -rf input/July_1999 output/Treasurer_Report_July_1999.xlsx output/executed_demo_July_1999.nbconvert.ipynb data/history/July_1999.json
   ```

## Hard rules

- **Never** push, commit, or share anything produced by a run against a *real*
  month's `input/{Month}_{Year}/` folder — `output/`, `logs/`, and `data/`
  always stay local-only, demo run or not.
- **Never** invoke the notebook's `push_to_github()` git-push cell as part of
  this skill. `BATCH_MODE=1` already suppresses it automatically — don't
  override that.
- Code changes (to `parsers.py`, `builders.py`, the notebook, this skill,
  `README.md`, etc.) go through the normal git workflow — `git add` /
  `git commit` / `git push` — as a deliberate, separate step. Running a demo
  report is not a trigger to push code, and pushing code is not a trigger to
  run a demo report; keep the two independent.
