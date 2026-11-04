# Deployment and Setup Guide

This guide covers running the project locally, publishing it to GitHub, the automated checks, and how to refresh the dashboard.

## 1. Requirements

| Tool | Version | Notes |
|---|---|---|
| Python | 3.10 or newer | Tested on 3.11, 3.12 and 3.13 |
| Git | any recent | To publish to GitHub |
| Microsoft Excel | 2016 or newer / Microsoft 365 | To open the dashboard. LibreOffice Calc also works |

Python packages are listed in `requirements.txt`: pandas, NumPy, openpyxl, matplotlib and pytest.

## 2. Run it locally

### macOS / Linux

```bash
cd hospitality-procurement-dashboard
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python run_pipeline.py
```

### Windows (PowerShell)

```powershell
cd hospitality-procurement-dashboard
py -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python run_pipeline.py
```

If PowerShell blocks the activation script, run `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned` once, then activate again.

### What the pipeline produces

| Output | Location |
|---|---|
| Raw synthetic tables (9 CSVs) | `data/raw/` |
| KPI tables (10 CSVs) | `data/processed/` |
| Excel dashboard | `output/Procurement_Intelligence_Dashboard.xlsx` |
| README charts (4 PNGs) | `docs/images/` |

Options:

```bash
python run_pipeline.py --seed 7        # a different random dataset
python run_pipeline.py --skip-charts   # skip the PNG charts
```

Each step can also be run on its own:

```bash
python -m src.generate_data
python -m src.analysis
python -m src.build_dashboard
python -m src.charts
```

## 3. Run the tests

```bash
pytest
```

28 tests check referential integrity, date and quantity logic, contract pricing rules, stock reconciliation arithmetic, KPI reconciliation, scorecard bounds and the workbook structure. Expect about a minute, most of it building the workbook.

## 4. Open the dashboard

Open `output/Procurement_Intelligence_Dashboard.xlsx` in Excel. The workbook is set to calculate fully on open, so values appear after a second or two.

- Change the **Fiscal year** cell on the Dashboard (yellow, top left) to FY25, FY26 or All. Every tab follows it.
- Yellow cells on **Notes & Assumptions** are editable inputs: compliance target, scorecard weights and rating bands.
- Blue numbers are hardcoded inputs (hotel room counts, contract prices); black numbers are formulas.

If you see zeros or blanks, check that calculation is set to Automatic (Formulas → Calculation Options → Automatic) and press `Ctrl+Alt+F9`.

## 5. Publish to GitHub

### Option A: GitHub website (no command line)

1. Sign in at github.com and click **New repository**.
2. Name it `hospitality-procurement-dashboard`, set it to **Public**, and do not add a README (the project has one).
3. On the next page choose **uploading an existing file**, drag in the contents of the project folder (not the `.venv` folder), and commit.

### Option B: command line

```bash
cd hospitality-procurement-dashboard
git init
git add .
git commit -m "Hospitality procurement intelligence dashboard"
git branch -M main
git remote add origin https://github.com/<your-username>/hospitality-procurement-dashboard.git
git push -u origin main
```

With the GitHub CLI installed you can replace the last three lines with:

```bash
gh repo create hospitality-procurement-dashboard --public --source=. --push
```

### What gets committed

`.gitignore` excludes the raw CSVs (about 9 MB, regenerated in seconds), the virtual environment and Python caches. The processed KPI tables, the Excel dashboard and the chart images **are** committed, so visitors can see results without running anything.

### Polish the repository page

- **About** (gear icon, top right of the repo): add a one-line description and topics such as `procurement`, `data-analytics`, `python`, `pandas`, `excel`, `hospitality`, `supply-chain`.
- **Releases**: create release `v1.0` and attach `Procurement_Intelligence_Dashboard.xlsx` so recruiters can download it in one click.
- **Pin** the repository on your GitHub profile.

## 6. Continuous integration

`.github/workflows/ci.yml` runs on every push and pull request. It installs dependencies on Python 3.11 and 3.12, runs the full pipeline and the test suite, and uploads the freshly built dashboard as a workflow artifact. A green tick on the repository page shows the project builds from a clean checkout.

## 7. Using real data (optional)

The pipeline can analyse a real purchasing export instead of synthetic data. Replace `data/raw/*.csv` with files that follow the schema in `docs/DATA_DICTIONARY.md`, then run:

```bash
python -m src.analysis
python -m src.build_dashboard
```

Typical mappings from an e-procurement and finance export:

| Project column | Typical source field |
|---|---|
| `po_line_id`, `po_number` | Order number + line number |
| `order_date`, `expected_delivery_date` | Order date, requested delivery date |
| `qty_received`, `delivery_date` | Goods received / delivery note |
| `invoice_unit_price`, `qty_invoiced` | Invoice line from the finance package |
| `contract_unit_price` | Agreed price from the supplier catalogue or price file |

> **Never commit real company or supplier data to a public repository.** Keep real exports outside the project folder, or add their folder to `.gitignore`.

## 8. Troubleshooting

| Symptom | Fix |
|---|---|
| `ModuleNotFoundError: No module named 'src'` | Run commands from the project root folder, not from inside `src/` |
| `TypeError` mentioning `include_groups` | Upgrade pandas: `pip install -U "pandas>=2.2"` |
| Dashboard shows zeros | Set Excel calculation to Automatic and press `Ctrl+Alt+F9` |
| Workbook opens slowly | Normal on first open: about 1,200 formulas recalculate over 46,000 rows |
| `python` not found on Windows | Use `py` instead of `python`, or reinstall Python with "Add to PATH" ticked |
