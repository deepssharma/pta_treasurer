# 📊 Setauket School PTA — Treasurer Report Generator

This tool generates your monthly treasurer reports from three input files:
- **Givebacks export** (CSV)
- **QuickBooks P&L / Transaction Detail export** (CSV)
- **Chase bank statement** (PDF)

It produces a single Excel file with 6 tabs:
1. Treasurer Report (with bank reconciliation)
2. Income Budget vs Actuals
3. Expense Budget vs Actuals
4. Giveback Reconciliation
5. File Manifest
6. YTD Summary (accumulates automatically across months you've already run)

---

## 🖥️ One-Time Setup (do this once)

### Step 1 — Install Python

1. Go to **python.org/downloads**
2. Click the big yellow **"Download Python"** button
3. Run the installer
   - ✅ On the first screen, check **"Add Python to PATH"** before clicking Install
4. Click **Install Now** and wait for it to finish

To verify it worked, open Terminal (Mac) or Command Prompt (Windows) and type:
```
python --version
```
You should see something like `Python 3.12.0`

---

### Step 2 — Download this project

Download the `pta_treasurer` folder and save it somewhere easy to find, like your Desktop or Documents folder.

---

### Step 3 — Install the required libraries

Open Terminal (Mac) or Command Prompt (Windows), navigate to the project folder, then run:
```
pip install -r requirements.txt
playwright install chromium
```
Wait for it to finish. You only need to do this once.

---

### Step 4 — Configure your org (optional but recommended)

Create a file named `.env` in the project folder (no file extension) with:
```
ORG_NAME=Your PTA Name
GIVEBACKS_ORG_URL=https://yourschool.givebacks.com
GIVEBACKS_EMAIL=you@example.com
GIVEBACKS_PASSWORD=your-password
```
`GIVEBACKS_EMAIL`/`GIVEBACKS_PASSWORD` are only needed if you want the tool to
auto-download Givebacks payout CSVs for you (Step 2 of the notebook, via a
Playwright browser automation with a one-time login). If you'd rather export
Givebacks CSVs yourself, skip those two and just place the files manually (see
below).

---

## 📁 Every Month — Running the Report

### Step 1 — Add your input files

Create a folder for the month under `input/`, named `MonthName_Year` (e.g.
`input/March_2026/`), and place inside it:

| File | Where to get it | Where it goes |
|---|---|---|
| QuickBooks Transaction Detail export | QuickBooks → Reports → Transaction Detail by Account → Export to CSV | `input/March_2026/quickbooks_march_2026.csv` |
| Chase statement | Chase.com → Statements → Download PDF | `input/March_2026/Chase_march_2026.pdf` |
| Givebacks payout export(s) | Givebacks → Payouts → Export CSV (or let the notebook auto-download them) | `input/March_2026/givebacks/givebacks_march_po_<id>.csv` |

Filenames just need to contain the month name somewhere — exact casing doesn't matter.

### Step 2 — Run the notebook

Open `PTA_Treasurer_Report_v4.ipynb` in Jupyter, set `INPUT_MONTH` and
`FISCAL_YEAR` near the top (Cell 1), and run all cells top to bottom.

```
jupyter notebook PTA_Treasurer_Report_v4.ipynb
```

You'll see progress printed as it parses each file, saves month history to
`data/history/`, and builds the workbook.

### Step 3 — Open your report

Find `Treasurer_Report_{Month}_{Year}.xlsx` in the `output/` folder and open it
in Excel or Google Sheets.

---

## 🔁 Running an entire fiscal year at once

`./run_all_months.sh` drives the notebook headlessly (no Jupyter UI needed) for
a fixed list of months, skipping any month whose `input/` folder doesn't exist
yet or has no QuickBooks file:

```
./run_all_months.sh
```

Each month's `output/Treasurer_Report_{Month}_{Year}.xlsx` is generated in
turn, and console output is captured to `logs/{Month}_{Year}.log`. Batch runs
set `BATCH_MODE=1`, which skips the interactive Givebacks-URL prompt and the
GitHub-push step at the end of the notebook — those only make sense when
running interactively.

---

## 🧪 Trying it out safely / running the test suite

The repo ships a small **synthetic** demo dataset (`sample_data/July_1999/` —
fake org, fake year, no real names or account numbers) so you can see the
whole workflow run without touching real financial data. See
`.claude/skills/generate-demo-report/SKILL.md` for the exact steps, or just:

```
mkdir -p input/July_1999 && cp -r sample_data/July_1999/* input/July_1999/
BATCH_MODE=1 INPUT_MONTH=July FISCAL_YEAR=1999 \
  jupyter nbconvert --to notebook --execute --output output/executed_demo_July_1999.ipynb PTA_Treasurer_Report_v4.ipynb
```

To run the pytest suite (uses the same kind of synthetic fixtures under
`tests/fixtures/`, safe to run anywhere):
```
pytest
```

**Real input files, generated reports, run logs, and history (`input/`,
`output/`, `logs/`, `data/`) are gitignored on purpose and should never be
committed or pushed** — they contain real names, account numbers, and dollar
amounts.

⚠️ **Before committing `PTA_Treasurer_Report_v4.ipynb`**, clear its cell
outputs first if you've just run it against real data — saved outputs are
part of the `.ipynb` file itself and are **not** covered by `.gitignore`.
In Jupyter: *Edit → Clear Outputs of All Cells* (or `Kernel → Restart & Clear
Output`), then save, before `git add`/`git commit`.

---

## ❓ Troubleshooting

**"python is not recognized" or "command not found"**
→ Python wasn't added to PATH during install. Re-run the Python installer and check "Add Python to PATH"

**"No module named openpyxl" / "No module named playwright" / etc.**
→ Run `pip install -r requirements.txt` again from inside the project folder

**"MONTH MISMATCH" or `FileNotFoundError` when running a month**
→ One of the three input files (QuickBooks CSV, Chase PDF, Givebacks CSV)
doesn't match the `INPUT_MONTH`/`FISCAL_YEAR` you set, or is missing entirely.
Check `input/{Month}_{Year}/` has all three, each named for the right month.

**Numbers look wrong on the bank reconciliation**
→ Make sure the Chase PDF is the full statement (all pages), not a partial download

**The report ran but I can't find the output file**
→ Look in the `output/` folder inside the project directory.

---

## 📞 Help

If something isn't working, take a screenshot of the error message in the terminal and send it to whoever set this up for you.
