# 📊 Setauket School PTA — Treasurer Report Generator

This tool automatically generates your monthly treasurer reports from three input files:
- **Givebacks export** (CSV)
- **QuickBooks P&L export** (CSV)
- **Chase bank statement** (PDF)

It produces a single Excel file with 4 tabs:
1. Monthly Treasurer Report
2. Income Budget vs Actuals
3. Expense Budget vs Actuals
4. Giveback Reconciliation

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

Open Terminal (Mac) or Command Prompt (Windows).

**On Mac**, navigate to the project folder:
```
cd ~/Desktop/pta_treasurer
```

**On Windows**, navigate to the project folder:
```
cd C:\Users\YourName\Desktop\pta_treasurer
```

Then run:
```
pip install -r requirements.txt
```

Wait for it to finish. You only need to do this once.

---

## 📁 Every Month — Running the Report

### Step 1 — Add your input files

Place the following files in the `input/` folder inside the project:

| File | Where to get it | What to name it |
|---|---|---|
| Givebacks export | Givebacks → Reports → Export CSV | `givebacks_example.csv` |
| QuickBooks P&L | QuickBooks → Reports → Profit & Loss → Export to CSV | `PTA_feb_quickbooks.csv` |
| Chase statement | Chase.com → Statements → Download PDF | `chase_feb_statement.pdf` |

> 💡 **Tip:** You can rename the files to anything — just update the filenames at the top of `generate_report.py` (look for the section marked `# ── MAIN ──`)

---

### Step 2 — Run the script

Open Terminal (Mac) or Command Prompt (Windows) and navigate to the project folder:

**Mac:**
```
cd ~/Desktop/pta_treasurer
python generate_report.py
```

**Windows:**
```
cd C:\Users\YourName\Desktop\pta_treasurer
python generate_report.py
```

You'll see:
```
Parsing Givebacks...
Parsing QuickBooks...
Parsing Chase PDF...
Building Excel workbook...
Saved → output/PTA_Treasurer_Report_February_2026.xlsx
```

---

### Step 3 — Open your report

Find `PTA_Treasurer_Report_February_2026.xlsx` in the `output/` folder and open it in Excel or Google Sheets.

---

## 🔄 Updating for a New Month

At the top of `generate_report.py`, find this section:

```python
gb_path  = base / "givebacks_example.csv"
qb_path  = base / "PTA_feb_quickbooks.csv"
pdf_path = base / "chase_feb_statement.pdf"
```

Update the filenames to match your new files. For example, for March:

```python
gb_path  = base / "givebacks_march.csv"
qb_path  = base / "PTA_march_quickbooks.csv"
pdf_path = base / "chase_march_statement.pdf"
```

Also update the month label on this line:
```python
build_treasurer_report(ws1, qb, bank, "February 2026")
```
Change `"February 2026"` to `"March 2026"` (or whatever month you're running).

---

## ❓ Troubleshooting

**"python is not recognized" or "command not found"**
→ Python wasn't added to PATH during install. Re-run the Python installer and check "Add Python to PATH"

**"No module named openpyxl" or "No module named pdfplumber"**
→ Run `pip install -r requirements.txt` again from inside the project folder

**Numbers look wrong on the bank reconciliation**
→ Make sure the Chase PDF is the full statement (all pages), not a partial download

**The script ran but I can't find the output file**
→ Look in the `output/` folder inside the project directory. On Mac you can also search for `PTA_Treasurer_Report` in Finder.

---

## 📞 Help

If something isn't working, take a screenshot of the error message in the terminal and send it to whoever set this up for you.
