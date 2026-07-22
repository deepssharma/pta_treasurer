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

`sample_data/output/` — also tracked in git, the expected output of running the
pipeline on the synthetic month above:
- `Treasurer_Report_July_1999.xlsx`
- `Debits_and_Credits_1999_to_1999.xlsx`

These serve two purposes: a reference example of what correct output looks like
(for someone setting this up for the first time), and a regression baseline —
after a code change to `parsers.py`/`builders.py`/the notebook, regenerate the
demo and compare against these to catch unintended changes to the output shape.
If a change is intentional, regenerate these files the same way (isolated run,
see the warning below) and commit the update alongside the code change.

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
   **Warning:** the notebook's Debits & Credits cell (Cell 12) scans *every*
   `input/{Month}_{Year}/` folder it finds, not just the one you just staged —
   if real months are also sitting in `input/` at the same time (the normal
   case in this repo, since real monthly data lives there too), the resulting
   `Debits_and_Credits_*.xlsx` will mix demo and real data together. If you
   need a clean demo-only result (e.g. to regenerate the `sample_data/output/`
   reference files), run this in an isolated git worktree instead, where only
   tracked files exist and no real `input/` data is present:
   ```sh
   git worktree add --detach /tmp/pta-demo-worktree main
   cd /tmp/pta-demo-worktree
   mkdir -p input/July_1999 output data/history
   cp -r sample_data/July_1999/* input/July_1999/
   BATCH_MODE=1 INPUT_MONTH=July FISCAL_YEAR=1999 \
     jupyter nbconvert --to notebook --execute \
       --ExecutePreprocessor.timeout=300 \
       --output output/executed_demo_July_1999.ipynb \
       PTA_Treasurer_Report_v4.ipynb
   cd - && git worktree remove /tmp/pta-demo-worktree --force
   ```
3. Confirm both workbooks were created with the right tabs:
   - `output/Treasurer_Report_July_1999.xlsx` — 6 tabs (Treasurer Report,
     Income Budget vs Actuals, Expense Budget vs Actuals, Giveback
     Reconciliation, File Manifest, YTD Summary)
   - `output/Debits_and_Credits_1999_to_1999.xlsx` — 3 tabs (Credits, Debits,
     MemberHub_Summary) — only if generated in isolation per the warning above;
     otherwise skip comparing this one
4. Optionally clean up afterwards — harmless to leave since it's gitignored:
   ```sh
   rm -rf input/July_1999 output/Treasurer_Report_July_1999.xlsx output/Debits_and_Credits_1999_to_1999.xlsx output/executed_demo_July_1999.nbconvert.ipynb data/history/July_1999.json
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
