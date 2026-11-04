# Hospitality Procurement Intelligence Dashboard

An end-to-end procurement analytics project for a multi-property hotel group: a Python pipeline that generates a realistic purchasing dataset, joins purchase orders, goods receipts and supplier invoices, calculates the KPIs a hotel procurement team actually tracks, and publishes them in a formula-driven Excel dashboard.

> **All data is synthetic.** The hotel group ("Harbourline"), its hotels and suppliers are fictional. The data model mirrors how a multi-property Irish hotel group buys food, beverage and operating supplies through an e-procurement system and finance package (PO → GRN → invoice), but no real company data is used.

![Dashboard preview](docs/images/dashboard_preview.png)

## The questions it answers

| Question | Where |
|---|---|
| How much are we spending, where, and is it seasonal? | Dashboard · Calc |
| What share of spend goes through agreed contracts, and which hotels buy off-contract? | Dashboard · Hotel Performance · Off-Contract |
| Are contracted suppliers charging what we agreed? | Price Variance |
| Which invoices fail three-way match (PO vs GRN vs invoice), and by how much? | Invoice Exceptions |
| Which suppliers deliver on time and in full, invoice accurately and hold price? | Supplier Scorecard |
| How much bar stock is unaccounted for between deliveries, till sales and stock counts? | Stock Variance |
| What is the total value leaking out, and from which source? | Dashboard |

## Headline findings (FY26, Oct 2025 – Sep 2026)

| KPI | FY26 | vs FY25 |
|---|---|---|
| Total spend | €5.97M | +€160k |
| On-contract spend | 88.7% | +1.2 pp |
| Price variance vs contract | €152k | +€9.8k |
| Three-way invoice match | 88.9% | +0.1 pp |
| Supplier OTIF | 87.6% | −0.3 pp |
| Bar stock shrinkage | 2.2% of units sold | flat |
| **Value leakage identified** | **€293k** | |

- **Compliance improved but is uneven.** The group moved from 87.4% to 88.7% on-contract, but Limerick sits at 70.7%, nearly 20 points below the 90% target. Off-contract buying cost about €111k above contract prices.
- **Contract prices drift.** Contracted suppliers' invoiced prices creep up through each contract year and reset at the October renewal (chart below). LinenCare Ireland and Atlantic Seafood Traders score lowest on the supplier scorecard.
- **Invoice control matters as much as sourcing.** 11% of invoice lines fail three-way match, and €107k was billed above what was received at agreed prices.
- **Bar shrinkage concentrates in spirits** at the sites with the weakest purchasing discipline.

<p>
<img src="docs/images/monthly_spend.png" width="49%">
<img src="docs/images/price_drift.png" width="49%">
</p>
<p>
<img src="docs/images/hotel_compliance.png" width="49%">
<img src="docs/images/value_leakage.png" width="49%">
</p>

## How it works

```
src/generate_data.py   ──►  data/raw/*.csv            hotels, suppliers, products, contract prices,
                                                      46k PO lines, deliveries, invoices, bar stock counts
src/analysis.py        ──►  data/processed/*.csv      fact table + 10 KPI tables
src/build_dashboard.py ──►  output/*.xlsx             formula-driven Excel dashboard (11 tabs, 6 charts)
src/charts.py          ──►  docs/images/*.png         README charts
run_pipeline.py             runs all four steps
tests/                      28 data-quality, logic and workbook tests (pytest)
```

**The dataset** covers 6 hotels (1,405 rooms), 18 suppliers (13 contracted, 5 spot/cash-and-carry), 57 products across 15 subcategories, and 24 months. It deliberately contains the problems a procurement team looks for: off-contract buying that varies by property, contracted suppliers drifting above agreed prices, late and short deliveries, invoice price and quantity errors, and bar stock shrinkage. All assumptions (supplier reliability, inflation by category, seasonality, compliance by site) are in [`src/config.py`](src/config.py).

**The workbook is live, not pasted.** The line-level extract sits on the `Data` and `Stock_Data` tabs, and every KPI, scorecard row and chart series is a `SUMIFS` / `COUNTIFS` / `AVERAGEIFS` formula over it. A fiscal-year selector on the Dashboard (FY25 / FY26 / All) drives every tab. Scorecard weights, rating bands and the compliance target are editable inputs on `Notes & Assumptions`. The Excel figures are tested against the Python results to the cent.

| Tab | What it shows |
|---|---|
| Dashboard | KPI cards with year-on-year change, value-leakage banner, six charts, key findings |
| Supplier Scorecard | Spend, on-time, in-full, OTIF, invoice accuracy, price gap, composite score and A/B/C rating |
| Hotel Performance | Spend per room, on-contract %, off-contract premium, OTIF, match rate, shrinkage |
| Price Variance | Invoiced vs contract price by product and contract year |
| Off-Contract | Hotel × subcategory heatmap of off-contract spend, plus top offenders |
| Invoice Exceptions | Every three-way-match failure, sorted by € discrepancy, filterable |
| Stock Variance | Bar shrinkage % by hotel and drinks category |
| Notes & Assumptions | Editable inputs, KPI definitions and method |

## Quick start

```bash
git clone https://github.com/<anuragratho>/hospitality-procurement-dashboard.git
cd hospitality-procurement-dashboard
python -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -r requirements.txt
python run_pipeline.py             # ~30 seconds
pytest                             # optional: run the tests
```

Open `output/Procurement_Intelligence_Dashboard.xlsx` in Excel. Formulas calculate on open. Full setup, GitHub publishing and troubleshooting notes are in [`docs/DEPLOYMENT.md`](docs/DEPLOYMENT.md); every table and column is described in [`docs/DATA_DICTIONARY.md`](docs/DATA_DICTIONARY.md).

## KPI definitions

| KPI | Definition |
|---|---|
| Spend | Invoiced value (qty invoiced × invoice unit price) |
| On-contract spend % | Share of spend with the contracted supplier at the contract price |
| Price variance vs contract | (invoice price − contract price) × qty invoiced |
| Three-way match | Invoice price = PO price (±€0.005) and invoiced qty = received qty |
| Invoice over-billing | Sum of positive (invoiced value − received qty × PO price) |
| OTIF | Delivered by the expected date **and** in full, per PO line |
| Bar stock shrinkage | (opening + received − sold per till) − counted closing stock |
| Supplier score | 40% OTIF + 30% invoice accuracy + 30% price discipline (weights editable) |
| Value leakage | Off-contract premium + contract creep + over-billing + shrinkage (indicative; components can overlap) |

## Tech

Python 3.10+ · pandas · NumPy · openpyxl · matplotlib · pytest · GitHub Actions

## Possible extensions

- Swap the generator for a real export (e.g. a Procure Wizard purchase history and invoice listing) by mapping its columns to the schema in the data dictionary.
- Load the processed tables into Power BI or a SQL database.
- Add demand forecasting for par-level ordering, using the seasonal pattern already in the data.

## Author

Built by Anurag: procurement and stores professional in Irish hospitality with an MSc in Data Analytics (National College of Ireland).

## Licence

MIT. See [LICENSE](LICENSE).
