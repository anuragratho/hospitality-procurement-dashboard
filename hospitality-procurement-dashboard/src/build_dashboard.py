"""
Excel dashboard builder.

Writes a formatted, formula-driven workbook. The line-level data extract sits on
the `Data` and `Stock_Data` sheets; every KPI, scorecard and chart series is an
Excel formula (SUMIFS / COUNTIFS / AVERAGEIFS) over that extract, so the workbook
recalculates when the fiscal-year selector on the Dashboard changes, or when the
data sheets are refreshed with new rows.

Run directly:  python -m src.build_dashboard
"""

from __future__ import annotations

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.drawing.text import CharacterProperties, ParagraphProperties
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, DataBarRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

from src import analysis
from src import config as cfg

# --------------------------------------------------------------------------- #
# Styling
# --------------------------------------------------------------------------- #
FONT = "Arial"
NAVY = "1F3864"
TEAL = "1A7F8E"
SLATE = "44546A"
LIGHT = "F2F5F9"
MID = "D9E1EC"
WHITE = "FFFFFF"
INPUT_FILL = "FFF2CC"     # pale yellow: cells the user may change
RED = "C0392B"
GREEN = "2E7D32"
# Chart series colours: validated categorical pair (colour-blind safe, ΔE ≥ 8)
SERIES_1 = "2A78D6"   # blue
SERIES_2 = "EB6834"   # orange

EUR = '€#,##0;(€#,##0);"-"'
EUR2 = '€#,##0.00;(€#,##0.00);"-"'
INT = '#,##0;(#,##0);"-"'
PCT = '0.0%;(0.0%);"-"'
PCT_DELTA = '+0.0%;-0.0%;0.0%'
EUR_DELTA = '+€#,##0;-€#,##0;€0'

thin = Side(style="thin", color=MID)
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)


def font(size=10, bold=False, color="000000", italic=False):
    return Font(name=FONT, size=size, bold=bold, color=color, italic=italic)


def fill(hex_color):
    return PatternFill("solid", start_color=hex_color, end_color=hex_color)


def style_header_row(ws, row, col_start, col_end, color=NAVY):
    for c in range(col_start, col_end + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = font(10, True, WHITE)
        cell.fill = fill(color)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER


def sheet_title(ws, title, subtitle=None, width_cols=12):
    ws.sheet_view.showGridLines = False
    ws["B2"] = title
    ws["B2"].font = font(16, True, NAVY)
    if subtitle:
        ws["B3"] = subtitle
        ws["B3"].font = font(9, False, SLATE, italic=True)
    ws.column_dimensions["A"].width = 2


def set_widths(ws, widths: dict[str, float]):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


def body(cell, fmt=None, bold=False, color="000000", align=None):
    cell.font = font(10, bold, color)
    cell.border = BORDER
    if fmt:
        cell.number_format = fmt
    if align:
        cell.alignment = Alignment(horizontal=align, vertical="center")


# --------------------------------------------------------------------------- #
# Data sheets
# --------------------------------------------------------------------------- #
DATA_COLS = [
    # (header, source column, number format)
    ("po_line_id", "po_line_id", None),
    ("order_date", "order_date", "yyyy-mm-dd"),
    ("month", "month", None),
    ("fiscal_year", "fy_num", "0"),
    ("contract_year", "contract_year", None),
    ("hotel", "hotel_name", None),
    ("supplier", "supplier_name", None),
    ("category", "category", None),
    ("subcategory", "subcategory", None),
    ("product", "product_name", None),
    ("on_contract", "on_contract_txt", None),
    ("qty_ordered", "qty_ordered", INT),
    ("qty_received", "qty_received", INT),
    ("qty_invoiced", "qty_invoiced", INT),
    ("po_unit_price", "po_unit_price", EUR2),
    ("contract_unit_price", "contract_unit_price", EUR2),
    ("invoice_unit_price", "invoice_unit_price", EUR2),
    ("spend", "spend", EUR2),
    ("benchmark_value", "benchmark_value", EUR2),
    ("ppv_value", "ppv_value", EUR2),
    ("contract_creep_value", "creep_value", EUR2),
    ("overcharge_value", "overcharge_value", EUR2),
    ("invoice_discrepancy", "invoice_discrepancy", EUR2),
    ("overbilled_value", "overbilled_value", EUR2),
    ("on_time", "on_time_n", "0"),
    ("in_full", "in_full_n", "0"),
    ("otif", "otif_n", "0"),
    ("three_way_match", "match_n", "0"),
]

STOCK_COLS = [
    ("hotel", "hotel_name", None),
    ("subcategory", "subcategory", None),
    ("product", "product_name", None),
    ("month", "month", None),
    ("fiscal_year", "fy_num", "0"),
    ("opening_qty", "opening_qty", INT),
    ("received_qty", "received_qty", INT),
    ("sold_qty_pos", "sold_qty_pos", INT),
    ("theoretical_closing_qty", "theoretical_closing_qty", INT),
    ("counted_closing_qty", "counted_closing_qty", INT),
    ("variance_qty", "variance_qty", '#,##0;(#,##0);"-"'),
    ("unit_cost", "unit_cost", EUR2),
    ("variance_value", "variance_value", EUR2),
]


def prepare_fact(f: pd.DataFrame) -> pd.DataFrame:
    f = f.copy()
    f["fy_num"] = 2000 + f["fiscal_year"].str[2:].astype(int)
    f["on_contract_txt"] = f["on_contract"].map({True: "Yes", False: "No"})
    f["creep_value"] = f["ppv_value"].where(f["on_contract"], 0).clip(lower=0).round(2)
    f["overcharge_value"] = ((f["invoice_unit_price"] - f["po_unit_price"]).clip(lower=0)
                             * f["qty_invoiced"]).round(2)
    f["overbilled_value"] = f["invoice_discrepancy"].clip(lower=0)
    f["on_time_n"] = f["on_time"].astype(int)
    f["in_full_n"] = f["in_full"].astype(int)
    f["otif_n"] = f["otif"].astype(int)
    f["match_n"] = f["three_way_match"].astype(int)
    f["order_date"] = f["order_date"].dt.date
    return f


def write_table_sheet(wb, name, df, cols, table_name, tab_color=SLATE):
    ws = wb.create_sheet(name)
    ws.sheet_properties.tabColor = tab_color
    for j, (header, _, _) in enumerate(cols, start=1):
        ws.cell(row=1, column=j, value=header)
    data = df[[src for _, src, _ in cols]].itertuples(index=False, name=None)
    for row in data:
        ws.append(list(row))
    n = len(df) + 1
    for j, (_, _, fmt) in enumerate(cols, start=1):
        letter = get_column_letter(j)
        ws.column_dimensions[letter].width = max(12, len(cols[j - 1][0]) + 3)
        if fmt:
            for r in range(2, n + 1):
                ws.cell(row=r, column=j).number_format = fmt
    style_header_row(ws, 1, 1, len(cols))
    ref = f"A1:{get_column_letter(len(cols))}{n}"
    t = Table(displayName=table_name, ref=ref)
    t.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True)
    ws.add_table(t)
    ws.freeze_panes = "A2"
    return ws, n


class Ref:
    """Absolute column references into a data sheet, e.g. R['spend'] -> Data!$R$2:$R$n"""

    def __init__(self, sheet: str, cols, last_row: int):
        self.map = {}
        for j, (header, _, _) in enumerate(cols, start=1):
            letter = get_column_letter(j)
            self.map[header] = f"{sheet}!${letter}$2:${letter}${last_row}"

    def __getitem__(self, key):
        return self.map[key]


# --------------------------------------------------------------------------- #
# Builder
# --------------------------------------------------------------------------- #
def build(results: dict[str, pd.DataFrame], tables: dict[str, pd.DataFrame], path=cfg.DASHBOARD_FILE):
    f = prepare_fact(results["fact_lines"])
    stock = (tables["stock_counts"]
             .merge(tables["products"][["product_id", "product_name", "subcategory"]], on="product_id")
             .merge(tables["hotels"][["hotel_id", "hotel_name"]], on="hotel_id"))
    stock["fy_num"] = 2000 + analysis.fiscal_year(pd.to_datetime(stock["month"])).str[2:].astype(int)

    wb = Workbook()
    dash = wb.active
    dash.title = "Dashboard"
    calc = wb.create_sheet("Calc")
    sc_ws = wb.create_sheet("Supplier Scorecard")
    hotel_ws = wb.create_sheet("Hotel Performance")
    pv_ws = wb.create_sheet("Price Variance")
    oc_ws = wb.create_sheet("Off-Contract")
    ex_ws = wb.create_sheet("Invoice Exceptions")
    st_ws = wb.create_sheet("Stock Variance")
    notes = wb.create_sheet("Notes & Assumptions")

    _, n_data = write_table_sheet(wb, "Data", f, DATA_COLS, "POLines")
    _, n_stock = write_table_sheet(wb, "Stock_Data", stock, STOCK_COLS, "BarStock")
    D = Ref("Data", DATA_COLS, n_data)
    S = Ref("Stock_Data", STOCK_COLS, n_stock)

    # Fiscal-year criteria used by every formula: fiscal_year between lo and hi
    FYLO, FYHI, MONTHS_SEL = "Calc!$C$4", "Calc!$C$5", "Calc!$C$6"

    def fy(rng):        # fiscal-year criteria pair for a data range
        return f'{rng},">="&{FYLO},{rng},"<="&{FYHI}'

    def fyd(extra=""):  # fiscal-year criteria on the Data sheet
        return fy(D["fiscal_year"]) + (f",{extra}" if extra else "")

    def fys(extra=""):
        return fy(S["fiscal_year"]) + (f",{extra}" if extra else "")

    build_notes(notes)
    build_calc(calc, D, S, results, fyd, fys, FYLO, FYHI)
    build_scorecard(sc_ws, D, results, fyd)
    build_hotels(hotel_ws, D, S, results, fyd, fys, MONTHS_SEL)
    build_price_variance(pv_ws, D, results)
    build_off_contract(oc_ws, D, results, fyd)
    build_exceptions(ex_ws, results)
    build_stock(st_ws, S, results, fys)
    build_dashboard(dash, results)

    # Tab colours
    dash.sheet_properties.tabColor = NAVY
    for ws in (sc_ws, hotel_ws, pv_ws, oc_ws, ex_ws, st_ws):
        ws.sheet_properties.tabColor = TEAL
    calc.sheet_properties.tabColor = "A6A6A6"
    notes.sheet_properties.tabColor = "A6A6A6"

    # Print setup: landscape, one page wide
    for ws in (dash, sc_ws, hotel_ws, pv_ws, oc_ws, st_ws, notes):
        ws.page_setup.orientation = "landscape"
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
    dash.print_area = "A1:S75"
    dash.page_setup.fitToHeight = 1

    wb.calculation.fullCalcOnLoad = True
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


# --------------------------------------------------------------------------- #
# Calc sheet: KPI block and chart series
# --------------------------------------------------------------------------- #
CALC_KPI_ROW = 10        # first KPI row on Calc
KPI_ROWS = {}            # kpi key -> Calc row (filled by build_calc)


def build_calc(ws, D, S, results, fyd, fys, FYLO, FYHI):
    sheet_title(ws, "Calculation sheet",
                "Formula engine behind the Dashboard. Every number is a SUMIFS / COUNTIFS over "
                "the Data and Stock_Data sheets. Do not type over formulas.")
    set_widths(ws, {"B": 38, "C": 16, "D": 16, "E": 16, "F": 16, "G": 16, "H": 16, "I": 16, "J": 16})

    # Selector translation
    ws["B4"], ws["C4"] = "Fiscal year from (numeric)", '=IF(Dashboard!$C$4="All",2000,2000+VALUE(RIGHT(Dashboard!$C$4,2)))'
    ws["B5"], ws["C5"] = "Fiscal year to (numeric)", '=IF(Dashboard!$C$4="All",2100,2000+VALUE(RIGHT(Dashboard!$C$4,2)))'
    ws["B6"], ws["C6"] = "Months in selection", None  # filled once the monthly table exists
    ws["B7"], ws["C7"] = "Prior fiscal year (for deltas)", '=IF(Dashboard!$C$4="All","",C4-1)'
    for r in range(4, 8):
        ws.cell(row=r, column=2).font = font(10)
        body(ws.cell(row=r, column=3), "0", align="center")

    # ---- KPI block: Selected | FY25 | FY26 | Total ---------------------------
    r0 = CALC_KPI_ROW - 1
    for j, h in enumerate(["KPI", "Selected", "FY25", "FY26", "All years"], start=2):
        ws.cell(row=r0, column=j, value=h)
    style_header_row(ws, r0, 2, 6)

    def crit_sel(rng):
        return f'{rng},">="&$C$4,{rng},"<="&$C$5'

    def crit_fixed(rng, year):
        return f"{rng},{year}" if year else None

    sp, bm, ppv = D["spend"], D["benchmark_value"], D["ppv_value"]
    oc, fyD, fyS = D["on_contract"], D["fiscal_year"], S["fiscal_year"]

    def sumifs(rng, crit_parts):
        parts = [p for p in crit_parts if p]
        return f"SUMIFS({rng},{','.join(parts)})" if parts else f"SUM({rng})"

    def countifs(crit_parts, fallback_rng):
        parts = [p for p in crit_parts if p]
        return f"COUNTIFS({','.join(parts)})" if parts else f"COUNTA({fallback_rng})"

    def averageifs(rng, crit_parts):
        parts = [p for p in crit_parts if p]
        return f"AVERAGEIFS({rng},{','.join(parts)})" if parts else f"AVERAGE({rng})"

    # (key, label, builder(criteria for Data fy, criteria for Stock fy), format)
    kpis = [
        ("spend", "Total spend (EUR)", lambda c, s: sumifs(sp, [c]), EUR),
        ("lines", "PO lines", lambda c, s: countifs([c], D["po_line_id"]), INT),
        ("offspend", "Off-contract spend (EUR)", lambda c, s: sumifs(sp, [c, f'{oc},"No"']), EUR),
        ("oncpct", "On-contract spend %", None, PCT),
        ("ppv", "Price variance vs contract (EUR)", lambda c, s: sumifs(ppv, [c]), EUR),
        ("ppvpct", "Price variance vs contract %", lambda c, s: f"{sumifs(ppv, [c])}/{sumifs(bm, [c])}", PCT),
        ("match", "Three-way match rate %", lambda c, s: averageifs(D["three_way_match"], [c]), PCT),
        ("overbill", "Invoice over-billing (EUR)", lambda c, s: sumifs(D["overbilled_value"], [c]), EUR),
        ("otif", "Supplier OTIF %", lambda c, s: averageifs(D["otif"], [c]), PCT),
        ("shrink", "Bar stock shrinkage (EUR)", lambda c, s: "-" + sumifs(S["variance_value"], [s]), EUR),
        ("shrinkpct", "Bar stock shrinkage %",
         lambda c, s: f"-{sumifs(S['variance_qty'], [s])}/{sumifs(S['sold_qty_pos'], [s])}", PCT),
        ("offprem", "Off-contract premium paid (EUR)", lambda c, s: sumifs(ppv, [c, f'{oc},"No"']), EUR),
        ("creep", "Contract price creep (EUR)", lambda c, s: sumifs(D["contract_creep_value"], [c]), EUR),
        ("leak", "Value leakage identified (EUR)", None, EUR),
    ]
    variants = [
        (3, crit_sel(fyD), crit_sel(fyS)),
        (4, crit_fixed(fyD, 2025), crit_fixed(fyS, 2025)),
        (5, crit_fixed(fyD, 2026), crit_fixed(fyS, 2026)),
        (6, None, None),
    ]
    for i, (key, label, builder, fmt) in enumerate(kpis):
        r = CALC_KPI_ROW + i
        KPI_ROWS[key] = r
        ws.cell(row=r, column=2, value=label).font = font(10)
        ws.cell(row=r, column=2).border = BORDER
        for col, cD, cS in variants:
            L = get_column_letter(col)
            if key == "oncpct":
                formula = f"=IFERROR(1-{L}{KPI_ROWS['offspend']}/{L}{KPI_ROWS['spend']},0)"
            elif key == "leak":
                formula = (f"={L}{KPI_ROWS['offprem']}+{L}{KPI_ROWS['creep']}"
                           f"+{L}{KPI_ROWS['overbill']}+{L}{KPI_ROWS['shrink']}")
            else:
                formula = f"=IFERROR({builder(cD, cS)},0)"
            ws.cell(row=r, column=col, value=formula)
            body(ws.cell(row=r, column=col), fmt)
    ws.cell(row=KPI_ROWS["leak"], column=2).comment = Comment(
        "Value leakage = off-contract premium over contract price + contracted suppliers "
        "charging above the agreed price + net over-billing versus goods received + bar stock "
        "shrinkage at cost. Components can overlap where a creep line is also over-billed; "
        "treat it as an indicative opportunity size, not booked savings.", "Model")

    # ---- Monthly series (full 24 months, not filtered) -----------------------
    mt = results["monthly_trend"]
    mrow = CALC_KPI_ROW + len(kpis) + 3
    ws.cell(row=mrow - 1, column=2, value="Monthly series (all months)").font = font(11, True, NAVY)
    heads = ["Month", "Fiscal year", "On-contract spend", "Off-contract spend", "Total spend",
             "Contract price drift %", "OTIF %", "3-way match %"]
    for j, h in enumerate(heads, start=2):
        ws.cell(row=mrow, column=j, value=h)
    style_header_row(ws, mrow, 2, 1 + len(heads))
    months = list(mt["month"])
    for i, m in enumerate(months):
        r = mrow + 1 + i
        ws.cell(row=r, column=2, value=m)
        yr, mo = int(m[:4]), int(m[5:])
        ws.cell(row=r, column=3, value=yr + (1 if mo >= 10 else 0))
        crit = f'{D["month"]},$B{r}'
        ws.cell(row=r, column=4, value=f'=SUMIFS({sp},{crit},{oc},"Yes")')
        ws.cell(row=r, column=5, value=f'=SUMIFS({sp},{crit},{oc},"No")')
        ws.cell(row=r, column=6, value=f"=D{r}+E{r}")
        ws.cell(row=r, column=7, value=f'=IFERROR(SUMIFS({sp},{crit},{oc},"Yes")/SUMIFS({bm},{crit},{oc},"Yes")-1,0)')
        ws.cell(row=r, column=8, value=f'=AVERAGEIFS({D["otif"]},{crit})')
        ws.cell(row=r, column=9, value=f'=AVERAGEIFS({D["three_way_match"]},{crit})')
        for c, fmt in zip(range(2, 10), [None, "0", EUR, EUR, EUR, PCT, PCT, PCT]):
            body(ws.cell(row=r, column=c), fmt, align="center" if c <= 3 else None)
    m_first, m_last = mrow + 1, mrow + len(months)
    ws["C6"] = f'=COUNTIFS($C${m_first}:$C${m_last},">="&$C$4,$C${m_first}:$C${m_last},"<="&$C$5)'
    CALC_RANGES["monthly"] = (mrow, m_first, m_last)

    # ---- Spend by subcategory (selected FY) ----------------------------------
    cat = results["spend_by_category"].sort_values("spend", ascending=True)
    srow = m_last + 3
    ws.cell(row=srow - 1, column=2, value="Spend by subcategory (selected year)").font = font(11, True, NAVY)
    for j, h in enumerate(["Subcategory", "Category", "Spend"], start=2):
        ws.cell(row=srow, column=j, value=h)
    style_header_row(ws, srow, 2, 4)
    for i, rr in enumerate(cat.itertuples()):
        r = srow + 1 + i
        ws.cell(row=r, column=2, value=rr.subcategory)
        ws.cell(row=r, column=3, value=rr.category)
        ws.cell(row=r, column=4, value=f"=SUMIFS({sp},{fyd()},{D['subcategory']},$B{r})")
        for c, fmt in zip(range(2, 5), [None, None, EUR]):
            body(ws.cell(row=r, column=c), fmt)
    CALC_RANGES["subcat"] = (srow, srow + 1, srow + len(cat))

    # ---- Hotel compliance (selected FY) --------------------------------------
    hs = results["hotel_summary"].sort_values("on_contract_pct")
    hrow = srow + len(cat) + 3
    ws.cell(row=hrow - 1, column=2, value="On-contract % by hotel (selected year)").font = font(11, True, NAVY)
    for j, h in enumerate(["Hotel", "On-contract %", "Target"], start=2):
        ws.cell(row=hrow, column=j, value=h)
    style_header_row(ws, hrow, 2, 4)
    for i, rr in enumerate(hs.itertuples()):
        r = hrow + 1 + i
        ws.cell(row=r, column=2, value=rr.hotel_name.replace("Harbourline ", ""))
        hotel_full = rr.hotel_name
        ws.cell(row=r, column=3, value=(f'=IFERROR(1-SUMIFS({sp},{fyd()},{D["hotel"]},"{hotel_full}",{oc},"No")'
                                        f'/SUMIFS({sp},{fyd()},{D["hotel"]},"{hotel_full}"),0)'))
        ws.cell(row=r, column=4, value="='Notes & Assumptions'!$C$8")
        for c, fmt in zip(range(2, 5), [None, PCT, PCT]):
            body(ws.cell(row=r, column=c), fmt)
    CALC_RANGES["hotel"] = (hrow, hrow + 1, hrow + len(hs))

    # ---- Value leakage breakdown (selected FY) -------------------------------
    lrow = hrow + len(hs) + 3
    ws.cell(row=lrow - 1, column=2, value="Value leakage breakdown (selected year)").font = font(11, True, NAVY)
    for j, h in enumerate(["Component", "EUR"], start=2):
        ws.cell(row=lrow, column=j, value=h)
    style_header_row(ws, lrow, 2, 3)
    comps = [("Off-contract premium", "offprem"), ("Contract price creep", "creep"),
             ("Invoice over-billing", "overbill"), ("Bar stock shrinkage", "shrink")]
    for i, (lab, key) in enumerate(comps):
        r = lrow + 1 + i
        ws.cell(row=r, column=2, value=lab)
        ws.cell(row=r, column=3, value=f"=C{KPI_ROWS[key]}")
        body(ws.cell(row=r, column=2))
        body(ws.cell(row=r, column=3), EUR)
    CALC_RANGES["leak"] = (lrow, lrow + 1, lrow + len(comps))


CALC_RANGES: dict[str, tuple[int, int, int]] = {}


# --------------------------------------------------------------------------- #
# Dashboard
# --------------------------------------------------------------------------- #
def build_dashboard(ws, results):
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 90
    widths = {"A": 2}
    for c in range(2, 20):
        widths[get_column_letter(c)] = 2.5 if (c - 2) % 3 == 2 else 13
    set_widths(ws, widths)

    # Banner
    for c in range(2, 20):
        ws.cell(row=2, column=c).fill = fill(NAVY)
        ws.cell(row=3, column=c).fill = fill(NAVY)
    ws["B2"] = "Hospitality Procurement Intelligence Dashboard"
    ws["B2"].font = font(18, True, WHITE)
    ws["B3"] = (f"{cfg.GROUP_NAME}  ·  6 hotels  ·  Oct 2024 – Sep 2026  ·  "
                "Synthetic data for portfolio demonstration")
    ws["B3"].font = font(9, False, "D9E1EC", italic=True)
    ws.row_dimensions[2].height = 30
    ws.row_dimensions[3].height = 18

    # Selector
    ws["B4"] = "Fiscal year ▸"
    ws["B4"].font = font(10, True, SLATE)
    ws["B4"].alignment = Alignment(horizontal="right", vertical="center")
    ws["C4"] = "FY26"
    ws["C4"].fill = fill(INPUT_FILL)
    ws["C4"].font = font(11, True, "0000FF")
    ws["C4"].alignment = Alignment(horizontal="center", vertical="center")
    ws["C4"].border = BORDER
    dv = DataValidation(type="list", formula1='"FY25,FY26,All"', allow_blank=False)
    dv.prompt, dv.promptTitle = "Choose FY25, FY26 or All", "Fiscal year"
    ws.add_data_validation(dv)
    dv.add("C4")
    ws["E4"] = "◂ change to FY25, FY26 or All. Fiscal year runs Oct–Sep (FY26 = Oct 2025 – Sep 2026)."
    ws["E4"].font = font(9, False, SLATE, italic=True)
    ws.row_dimensions[4].height = 22

    # KPI cards: (label, kpi key, format, delta format, higher_is_better)
    cards = [
        ("Total spend", "spend", EUR, EUR_DELTA, None),
        ("On-contract spend", "oncpct", PCT, PCT_DELTA, True),
        ("Price variance vs contract", "ppv", EUR, EUR_DELTA, False),
        ("3-way invoice match", "match", PCT, PCT_DELTA, True),
        ("Supplier OTIF", "otif", PCT, PCT_DELTA, True),
        ("Bar stock shrinkage", "shrinkpct", PCT, PCT_DELTA, False),
    ]
    card_w, gap, start_col, top = 1, 1, 2, 6   # card spans 2 columns, then a narrow gap column
    # Wider band for the leakage card below
    for i, (label, key, fmt, dfmt, better_up) in enumerate(cards):
        c0 = start_col + i * (card_w + 1 + gap)
        c1 = c0 + card_w
        L0, L1 = get_column_letter(c0), get_column_letter(c1)
        kr = KPI_ROWS[key]
        for r in range(top, top + 4):
            for c in range(c0, c1 + 1):
                ws.cell(row=r, column=c).fill = fill(LIGHT)
        for c in range(c0, c1 + 1):
            ws.cell(row=top, column=c).border = Border(top=Side(style="thick", color=TEAL))
        ws.merge_cells(f"{L0}{top}:{L1}{top}")
        ws.cell(row=top, column=c0, value=label).font = font(9, True, SLATE)
        ws.cell(row=top, column=c0).alignment = Alignment(horizontal="left", vertical="bottom", indent=1)
        ws.merge_cells(f"{L0}{top+1}:{L1}{top+2}")
        v = ws.cell(row=top + 1, column=c0, value=f"=Calc!$C${kr}")
        v.font = font(18, True, NAVY)
        v.number_format = fmt
        v.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.merge_cells(f"{L0}{top+3}:{L1}{top+3}")
        # delta vs prior year only meaningful when FY26 is selected
        scale = "*100" if fmt == PCT else ""
        d = ws.cell(row=top + 3, column=c0,
                    value=f'=IF(Dashboard!$C$4="FY26",(Calc!$E${kr}-Calc!$D${kr}){scale},"")')
        if fmt == PCT:
            d.number_format = '"vs FY25 "+0.0" pp";"vs FY25 "-0.0" pp";"vs FY25 no change"'
        else:
            d.number_format = '"vs FY25 "+€#,##0;"vs FY25 "-€#,##0;"vs FY25 no change"'
        d.font = font(9, False, SLATE)
        d.alignment = Alignment(horizontal="left", vertical="top", indent=1)
        if better_up is not None:
            good, bad = ("greaterThan", "lessThan") if better_up else ("lessThan", "greaterThan")
            rng = f"{L0}{top+3}"
            ws.conditional_formatting.add(rng, CellIsRule(operator=good, formula=["0"], font=Font(name=FONT, size=9, color=GREEN, bold=True)))
            ws.conditional_formatting.add(rng, CellIsRule(operator=bad, formula=["0"], font=Font(name=FONT, size=9, color=RED, bold=True)))
    ws.row_dimensions[top].height = 20
    ws.row_dimensions[top + 1].height = 18
    ws.row_dimensions[top + 2].height = 18
    ws.row_dimensions[top + 3].height = 16

    # Leakage banner
    lr = top + 5
    for c in range(2, 20):
        ws.cell(row=lr, column=c).fill = fill("FBEAEA")
    ws.merge_cells(f"B{lr}:F{lr}")
    ws.cell(row=lr, column=2, value="Value leakage identified (selected year)").font = font(10, True, RED)
    ws.cell(row=lr, column=2).alignment = Alignment(vertical="center", indent=1)
    ws.merge_cells(f"G{lr}:H{lr}")
    lv = ws.cell(row=lr, column=7, value=f"=Calc!$C${KPI_ROWS['leak']}")
    lv.font = font(14, True, RED)
    lv.number_format = EUR
    lv.alignment = Alignment(vertical="center")
    ws.merge_cells(f"J{lr}:S{lr}")
    ws.cell(row=lr, column=10,
            value=(f'="Off-contract premium "&TEXT(Calc!$C${KPI_ROWS["offprem"]},"€#,##0")'
                   f'&"  ·  Contract creep "&TEXT(Calc!$C${KPI_ROWS["creep"]},"€#,##0")'
                   f'&"  ·  Over-billing "&TEXT(Calc!$C${KPI_ROWS["overbill"]},"€#,##0")'
                   f'&"  ·  Shrinkage "&TEXT(Calc!$C${KPI_ROWS["shrink"]},"€#,##0")')
            ).font = font(9, False, SLATE)
    ws.cell(row=lr, column=10).alignment = Alignment(vertical="center")
    ws.row_dimensions[lr].height = 26

    # ---- Charts ---------------------------------------------------------------
    calc = ws.parent["Calc"]
    mh, m1, m2 = CALC_RANGES["monthly"]

    def style_chart(ch, title, w=16.2, h=7.4):
        ch.title = title
        para = ch.title.tx.rich.p[0]
        props = CharacterProperties(sz=1100, b=True, solidFill=NAVY, latin=None)
        para.pPr = ParagraphProperties(defRPr=props)
        for run in para.r or []:
            run.rPr = CharacterProperties(sz=1100, b=True, solidFill=NAVY)
        ch.width, ch.height = w, h
        if ch.legend is not None:
            ch.legend.position = "b"
        return ch

    # 1 Monthly spend stacked
    c1 = BarChart()
    c1.type, c1.grouping, c1.overlap = "col", "stacked", 100
    c1.add_data(Reference(calc, min_col=4, max_col=5, min_row=mh, max_row=m2), titles_from_data=True)
    c1.set_categories(Reference(calc, min_col=2, min_row=m1, max_row=m2))
    c1.y_axis.numFmt = '€#,##0,"k"'
    c1.y_axis.majorGridlines = None
    c1.gapWidth = 40
    c1.series[0].graphicalProperties.solidFill = SERIES_1
    c1.series[1].graphicalProperties.solidFill = SERIES_2
    for s in c1.series:
        s.graphicalProperties.line.noFill = True
    style_chart(c1, "Monthly spend: on- vs off-contract (all months)")
    ws.add_chart(c1, f"B{lr+2}")

    # 2 Contract price drift line
    c2 = LineChart()
    c2.add_data(Reference(calc, min_col=7, min_row=mh, max_row=m2), titles_from_data=True)
    c2.set_categories(Reference(calc, min_col=2, min_row=m1, max_row=m2))
    c2.y_axis.numFmt = "0.0%"
    c2.series[0].graphicalProperties.line.solidFill = SERIES_2
    c2.series[0].graphicalProperties.line.width = 28000
    c2.series[0].smooth = False
    c2.legend = None
    style_chart(c2, "Contracted suppliers: invoiced price above contract (resets at Oct renewal)")
    ws.add_chart(c2, f"K{lr+2}")

    # 3 Spend by subcategory
    sh, s1, s2 = CALC_RANGES["subcat"]
    c3 = BarChart()
    c3.type = "bar"
    c3.add_data(Reference(calc, min_col=4, min_row=sh, max_row=s2), titles_from_data=True)
    c3.set_categories(Reference(calc, min_col=2, min_row=s1, max_row=s2))
    c3.x_axis.numFmt = "General"
    c3.y_axis.numFmt = '€#,##0,"k"'
    c3.y_axis.majorGridlines = None
    c3.series[0].graphicalProperties.solidFill = SERIES_1
    c3.gapWidth = 35
    c3.legend = None
    style_chart(c3, "Spend by subcategory (selected year)", h=9.0)
    ws.add_chart(c3, f"B{lr+18}")

    # 4 Hotel compliance with target
    hh, h1, h2 = CALC_RANGES["hotel"]
    c4 = BarChart()
    c4.type = "bar"
    c4.add_data(Reference(calc, min_col=3, min_row=hh, max_row=h2), titles_from_data=True)
    c4.set_categories(Reference(calc, min_col=2, min_row=h1, max_row=h2))
    c4.y_axis.numFmt = "0%"
    c4.y_axis.scaling.min = 0.5
    c4.y_axis.scaling.max = 1.0
    c4.y_axis.majorGridlines = None
    c4.series[0].graphicalProperties.solidFill = SERIES_1
    c4.series[0].dLbls = DataLabelList()
    c4.series[0].dLbls.showVal = True
    c4.series[0].dLbls.showSerName = False
    c4.series[0].dLbls.showCatName = False
    c4.series[0].dLbls.showLegendKey = False
    c4.series[0].dLbls.numFmt = "0.0%"
    c4.gapWidth = 45
    c4.legend = None
    style_chart(c4, "On-contract spend % by hotel (selected year; target 90%)", h=9.0)
    ws.add_chart(c4, f"K{lr+18}")

    # 5 Leakage breakdown
    lh, l1, l2 = CALC_RANGES["leak"]
    c5 = BarChart()
    c5.type = "col"
    c5.add_data(Reference(calc, min_col=3, min_row=lh, max_row=l2), titles_from_data=True)
    c5.set_categories(Reference(calc, min_col=2, min_row=l1, max_row=l2))
    c5.y_axis.numFmt = '€#,##0,"k"'
    c5.y_axis.majorGridlines = None
    c5.series[0].graphicalProperties.solidFill = SERIES_2
    c5.series[0].dLbls = DataLabelList()
    c5.series[0].dLbls.showVal = True
    c5.series[0].dLbls.showSerName = False
    c5.series[0].dLbls.showCatName = False
    c5.series[0].dLbls.showLegendKey = False
    c5.series[0].dLbls.numFmt = '€#,##0,"k"'
    c5.gapWidth = 60
    c5.legend = None
    style_chart(c5, "Where value leaks (selected year)")
    ws.add_chart(c5, f"B{lr+37}")

    # 6 OTIF & match trend
    c6 = LineChart()
    c6.add_data(Reference(calc, min_col=8, max_col=9, min_row=mh, max_row=m2), titles_from_data=True)
    c6.set_categories(Reference(calc, min_col=2, min_row=m1, max_row=m2))
    c6.y_axis.numFmt = "0%"
    c6.y_axis.scaling.min = 0.80
    c6.y_axis.scaling.max = 0.95
    c6.series[0].graphicalProperties.line.solidFill = SERIES_1
    c6.series[1].graphicalProperties.line.solidFill = SERIES_2
    for s in c6.series:
        s.graphicalProperties.line.width = 22000
        s.smooth = False
    style_chart(c6, "Supplier OTIF and 3-way invoice match rate by month")
    ws.add_chart(c6, f"K{lr+37}")

    # Key findings (static narrative computed at build time)
    kr = lr + 53
    ws.cell(row=kr, column=2, value="Key findings (FY26, generated when the workbook was built)").font = font(11, True, NAVY)
    for i, line in enumerate(key_findings(results)):
        cell = ws.cell(row=kr + 1 + i, column=2, value=f"•  {line}")
        cell.font = font(10)
        ws.merge_cells(start_row=kr + 1 + i, start_column=2, end_row=kr + 1 + i, end_column=19)
    ws.cell(row=kr + 8, column=2,
            value="Detail: Supplier Scorecard · Hotel Performance · Price Variance · Off-Contract · "
                  "Invoice Exceptions · Stock Variance tabs. Method and definitions: Notes & Assumptions."
            ).font = font(9, False, SLATE, italic=True)
    ws.freeze_panes = "A5"


def key_findings(results) -> list[str]:
    f = results["fact_lines"]
    cur = f[f["fiscal_year"] == "FY26"]
    prev = f[f["fiscal_year"] == "FY25"]
    onc = lambda d: 1 - d.loc[~d["on_contract"], "spend"].sum() / d["spend"].sum()
    hotels = (cur.groupby("hotel_name")
                 .apply(lambda d: onc(d), include_groups=False).sort_values())
    worst_h, worst_v = hotels.index[0], hotels.iloc[0]
    off_prem = cur.loc[~cur["on_contract"], "ppv_value"].sum()
    sc = analysis.supplier_scorecard(cur)
    review = sc[sc["rating"].str.startswith("C") & sc["is_contracted"]]
    pv = analysis.price_variance(cur).head(1).iloc[0]
    exc = cur[~cur["three_way_match"]]
    over = exc.loc[exc["invoice_discrepancy"] > 0, "invoice_discrepancy"].sum()
    stock = results["stock_variance"]
    worst_stock = stock.sort_values("shrink_pct", ascending=False).iloc[0]
    return [
        f"On-contract spend rose from {onc(prev):.1%} in FY25 to {onc(cur):.1%} in FY26, but "
        f"{worst_h.replace('Harbourline ', '')} lags at {worst_v:.1%}, well below the 90% target.",
        f"Off-contract buying cost an estimated €{off_prem:,.0f} in FY26 above contract prices.",
        f"Contracted suppliers to review: {', '.join(review['supplier_name'])} "
        f"(lowest composite scores on OTIF, invoice accuracy and price discipline).",
        f"Largest price drift: {pv['product_name']} from {pv['supplier_name']}, "
        f"€{pv['ppv_value']:,.0f} above contract ({pv['variance_pct']:.1%} average) in FY26.",
        f"{len(exc):,} invoice lines ({len(exc)/len(cur):.1%}) failed three-way match; "
        f"€{over:,.0f} was billed above goods received at agreed prices.",
        f"Highest bar shrinkage: {worst_stock['subcategory']} at "
        f"{worst_stock['hotel_name'].replace('Harbourline ', '')} ({worst_stock['shrink_pct']:.1%} of units sold, both years).",
    ]


# --------------------------------------------------------------------------- #
# Detail sheets
# --------------------------------------------------------------------------- #
def selector_note(ws, row=4):
    ws.cell(row=row, column=2, value='="Showing: "&Dashboard!$C$4&"  (change the selector on the Dashboard tab)"')
    ws.cell(row=row, column=2).font = font(10, True, TEAL)


def build_scorecard(ws, D, results, fyd):
    sheet_title(ws, "Supplier Scorecard",
                "Composite score = weighted OTIF + invoice accuracy + price discipline. "
                "Weights and rating bands are editable on Notes & Assumptions.")
    selector_note(ws)
    heads = ["Supplier", "Contracted?", "Spend", "Share of spend", "PO lines", "On time %",
             "In full %", "OTIF %", "Invoice accuracy %", "Overcharge vs PO",
             "Price vs contract %", "Price gap %", "Score (0-100)", "Rating"]
    hr = 6
    for j, h in enumerate(heads, start=2):
        ws.cell(row=hr, column=j, value=h)
    style_header_row(ws, hr, 2, 1 + len(heads))
    ws.row_dimensions[hr].height = 32
    set_widths(ws, {"B": 28, "C": 11, "D": 13, "E": 10, "F": 9, "G": 9, "H": 9, "I": 9,
                    "J": 11, "K": 13, "L": 11, "M": 10, "N": 11, "O": 14})
    sc = results["supplier_scorecard"]
    n = len(sc)
    first, last = hr + 1, hr + n
    N = "'Notes & Assumptions'"
    for i, s in enumerate(sc.itertuples()):
        r = hr + 1 + i
        sup = f'{D["supplier"]},$B{r}'
        cr = f"{fyd()},{sup}"
        ws.cell(row=r, column=2, value=s.supplier_name)
        ws.cell(row=r, column=3, value="Yes" if s.is_contracted else "No")
        ws.cell(row=r, column=4, value=f"=SUMIFS({D['spend']},{cr})")
        ws.cell(row=r, column=5, value=f"=IFERROR(D{r}/SUM($D${first}:$D${last}),0)")
        ws.cell(row=r, column=6, value=f"=COUNTIFS({cr})")
        ws.cell(row=r, column=7, value=f"=IFERROR(AVERAGEIFS({D['on_time']},{cr}),0)")
        ws.cell(row=r, column=8, value=f"=IFERROR(AVERAGEIFS({D['in_full']},{cr}),0)")
        ws.cell(row=r, column=9, value=f"=IFERROR(AVERAGEIFS({D['otif']},{cr}),0)")
        ws.cell(row=r, column=10, value=f"=IFERROR(AVERAGEIFS({D['three_way_match']},{cr}),0)")
        ws.cell(row=r, column=11, value=f"=SUMIFS({D['overcharge_value']},{cr})")
        ws.cell(row=r, column=12, value=f"=IFERROR(SUMIFS({D['ppv_value']},{cr})/SUMIFS({D['benchmark_value']},{cr}),0)")
        ws.cell(row=r, column=13, value=f"=IFERROR(MAX(K{r}/D{r},MAX(0,L{r})),0)")
        ws.cell(row=r, column=14, value=(f"=ROUND(100*({N}!$C$12*I{r}+{N}!$C$13*J{r}"
                                         f"+{N}!$C$14*MAX(0,MIN(1,1-M{r}/{N}!$C$15))),1)"))
        ws.cell(row=r, column=15, value=(f'=IF(N{r}>={N}!$C$16,"A - Preferred",'
                                         f'IF(N{r}>={N}!$C$17,"B - Monitor","C - Review"))'))
        fmts = [None, None, EUR, PCT, INT, PCT, PCT, PCT, PCT, EUR, PCT, PCT, "0.0", None]
        for c, fmt in zip(range(2, 16), fmts):
            body(ws.cell(row=r, column=c), fmt, align="center" if c in (3, 15) else None)
    # Totals row
    tr = last + 1
    ws.cell(row=tr, column=2, value="All suppliers")
    ws.cell(row=tr, column=4, value=f"=SUM(D{first}:D{last})")
    ws.cell(row=tr, column=5, value=f"=SUM(E{first}:E{last})")
    ws.cell(row=tr, column=6, value=f"=SUM(F{first}:F{last})")
    ws.cell(row=tr, column=9, value=f"=SUMPRODUCT(I{first}:I{last},F{first}:F{last})/F{tr}")
    ws.cell(row=tr, column=10, value=f"=SUMPRODUCT(J{first}:J{last},F{first}:F{last})/F{tr}")
    ws.cell(row=tr, column=11, value=f"=SUM(K{first}:K{last})")
    for c, fmt in zip(range(2, 16), [None, None, EUR, PCT, INT, None, None, PCT, PCT, EUR, None, None, None, None]):
        body(ws.cell(row=tr, column=c), fmt, bold=True)
        ws.cell(row=tr, column=c).fill = fill(MID)

    ws.conditional_formatting.add(f"N{first}:N{last}", ColorScaleRule(
        start_type="num", start_value=55, start_color="F8696B",
        mid_type="num", mid_value=80, mid_color="FFEB84",
        end_type="num", end_value=95, end_color="63BE7B"))
    ws.conditional_formatting.add(f"D{first}:D{last}", DataBarRule(
        start_type="min", end_type="max", color="5B9BD5", showValue=True))
    ws.conditional_formatting.add(f"O{first}:O{last}", FormulaRule(
        formula=[f'LEFT(O{first},1)="C"'], fill=fill("F8D7DA"), font=Font(name=FONT, color=RED, bold=True)))
    ws.conditional_formatting.add(f"O{first}:O{last}", FormulaRule(
        formula=[f'LEFT(O{first},1)="A"'], fill=fill("D4EDDA"), font=Font(name=FONT, color=GREEN, bold=True)))
    ws.conditional_formatting.add(f"C{first}:C{last}", CellIsRule(
        operator="equal", formula=['"No"'], font=Font(name=FONT, color="E07B39", bold=True)))
    ws.cell(row=tr + 2, column=2, value=(
        "Price gap = larger of overcharge vs PO price and premium over the group contract price. "
        "Off-contract suppliers score low on price because their quotes sit above contract.")).font = font(9, False, SLATE, True)
    ws.freeze_panes = ws.cell(row=hr + 1, column=3)


def build_hotels(ws, D, S, results, fyd, fys, MONTHS_SEL):
    sheet_title(ws, "Hotel Performance",
                "Spend, compliance and control indicators by property.")
    selector_note(ws)
    heads = ["Hotel", "City", "Rooms", "Spend", "Annualised spend per room", "Off-contract spend",
             "On-contract %", "Off-contract premium", "OTIF %", "3-way match %",
             "Bar shrinkage (EUR)", "Bar shrinkage %"]
    hr = 6
    for j, h in enumerate(heads, start=2):
        ws.cell(row=hr, column=j, value=h)
    style_header_row(ws, hr, 2, 1 + len(heads))
    ws.row_dimensions[hr].height = 32
    set_widths(ws, {"B": 28, "C": 11, "D": 8, "E": 14, "F": 14, "G": 14, "H": 11, "I": 13,
                    "J": 9, "K": 11, "L": 13, "M": 11})
    hs = results["hotel_summary"]
    first, last = hr + 1, hr + len(hs)
    for i, h in enumerate(hs.itertuples()):
        r = hr + 1 + i
        crit = f"{fyd()},{D['hotel']},$B{r}"
        scrit = f"{fys()},{S['hotel']},$B{r}"
        ws.cell(row=r, column=2, value=h.hotel_name)
        ws.cell(row=r, column=3, value=h.city)
        c = ws.cell(row=r, column=4, value=int(h.rooms))
        ws.cell(row=r, column=5, value=f"=SUMIFS({D['spend']},{crit})")
        ws.cell(row=r, column=6, value=f"=IFERROR(E{r}/D{r}/({MONTHS_SEL}/12),0)")
        ws.cell(row=r, column=7, value=f'=SUMIFS({D["spend"]},{crit},{D["on_contract"]},"No")')
        ws.cell(row=r, column=8, value=f"=IFERROR(1-G{r}/E{r},0)")
        ws.cell(row=r, column=9, value=f'=SUMIFS({D["ppv_value"]},{crit},{D["on_contract"]},"No")')
        ws.cell(row=r, column=10, value=f"=IFERROR(AVERAGEIFS({D['otif']},{crit}),0)")
        ws.cell(row=r, column=11, value=f"=IFERROR(AVERAGEIFS({D['three_way_match']},{crit}),0)")
        ws.cell(row=r, column=12, value=f"=-SUMIFS({S['variance_value']},{scrit})")
        ws.cell(row=r, column=13, value=f"=IFERROR(-SUMIFS({S['variance_qty']},{scrit})/SUMIFS({S['sold_qty_pos']},{scrit}),0)")
        fmts = [None, None, INT, EUR, EUR, EUR, PCT, EUR, PCT, PCT, EUR, PCT]
        for col, fmt in zip(range(2, 14), fmts):
            body(ws.cell(row=r, column=col), fmt)
        c.font = font(10, False, "0000FF")
    tr = last + 1
    ws.cell(row=tr, column=2, value="Group")
    ws.cell(row=tr, column=4, value=f"=SUM(D{first}:D{last})")
    ws.cell(row=tr, column=5, value=f"=SUM(E{first}:E{last})")
    ws.cell(row=tr, column=6, value=f"=IFERROR(E{tr}/D{tr}/({MONTHS_SEL}/12),0)")
    ws.cell(row=tr, column=7, value=f"=SUM(G{first}:G{last})")
    ws.cell(row=tr, column=8, value=f"=IFERROR(1-G{tr}/E{tr},0)")
    ws.cell(row=tr, column=9, value=f"=SUM(I{first}:I{last})")
    ws.cell(row=tr, column=12, value=f"=SUM(L{first}:L{last})")
    for col, fmt in zip(range(2, 14), [None, None, INT, EUR, EUR, EUR, PCT, EUR, None, None, EUR, None]):
        body(ws.cell(row=tr, column=col), fmt, bold=True)
        ws.cell(row=tr, column=col).fill = fill(MID)
    ws.conditional_formatting.add(f"H{first}:H{last}", ColorScaleRule(
        start_type="min", start_color="F8696B", mid_type="percentile", mid_value=50,
        mid_color="FFEB84", end_type="max", end_color="63BE7B"))
    ws.conditional_formatting.add(f"M{first}:M{last}", ColorScaleRule(
        start_type="min", start_color="63BE7B", mid_type="percentile", mid_value=50,
        mid_color="FFEB84", end_type="max", end_color="F8696B"))
    ws.cell(row=tr + 2, column=2, value="Rooms (blue) are inputs. Spend per room is annualised for the months in the selected year.").font = font(9, False, SLATE, True)
    ws.freeze_panes = ws.cell(row=hr + 1, column=3)


def build_price_variance(ws, D, results):
    sheet_title(ws, "Price Variance — contracted lines",
                "Invoiced price versus agreed contract price, by product and contract year "
                "(CY1 = Oct 2024 – Sep 2025, CY2 = Oct 2025 – Sep 2026). Sorted by € variance.")
    heads = ["Product", "Subcategory", "Supplier", "Contract year", "Contract price",
             "Qty invoiced", "Paid value", "Avg paid price", "Variance %", "Variance (EUR)"]
    hr = 5
    for j, h in enumerate(heads, start=2):
        ws.cell(row=hr, column=j, value=h)
    style_header_row(ws, hr, 2, 1 + len(heads))
    set_widths(ws, {"B": 34, "C": 20, "D": 26, "E": 10, "F": 12, "G": 11, "H": 14, "I": 12,
                    "J": 10, "K": 13})
    pv = results["price_variance"]
    first, last = hr + 1, hr + len(pv)
    for i, p in enumerate(pv.itertuples()):
        r = hr + 1 + i
        crit = (f'{D["product"]},$B{r},{D["contract_year"]},$E{r},{D["on_contract"]},"Yes"')
        ws.cell(row=r, column=2, value=p.product_name)
        ws.cell(row=r, column=3, value=p.subcategory)
        ws.cell(row=r, column=4, value=p.supplier_name)
        ws.cell(row=r, column=5, value=p.contract_year)
        ws.cell(row=r, column=6, value=float(p.contract_price))
        ws.cell(row=r, column=7, value=f"=SUMIFS({D['qty_invoiced']},{crit})")
        ws.cell(row=r, column=8, value=f"=SUMIFS({D['spend']},{crit})")
        ws.cell(row=r, column=9, value=f"=IFERROR(H{r}/G{r},0)")
        ws.cell(row=r, column=10, value=f"=IFERROR(I{r}/F{r}-1,0)")
        ws.cell(row=r, column=11, value=f"=SUMIFS({D['ppv_value']},{crit})")
        fmts = [None, None, None, None, EUR2, INT, EUR, EUR2, PCT, EUR]
        for c, fmt in zip(range(2, 12), fmts):
            body(ws.cell(row=r, column=c), fmt, align="center" if c == 5 else None)
        ws.cell(row=r, column=6).font = font(10, False, "0000FF")
    ws.conditional_formatting.add(f"J{first}:J{last}", ColorScaleRule(
        start_type="num", start_value=0, start_color="FFFFFF", end_type="max", end_color="F8696B"))
    ws.conditional_formatting.add(f"K{first}:K{last}", DataBarRule(
        start_type="num", start_value=0, end_type="max", color="F4B183", showValue=True))
    ws.auto_filter.ref = f"B{hr}:K{last}"
    ws.cell(row=last + 2, column=2, value="Contract prices (blue) come from the contract price list (data/raw/contract_prices.csv).").font = font(9, False, SLATE, True)
    ws.freeze_panes = ws.cell(row=hr + 1, column=3)


def build_off_contract(ws, D, results, fyd):
    sheet_title(ws, "Off-Contract Spend — hotel × subcategory",
                "Spend placed with non-contracted suppliers. Darker = more leakage.")
    selector_note(ws)
    hotels = list(results["hotel_summary"].sort_values("on_contract_pct")["hotel_name"])
    subs = list(results["spend_by_category"].query("off_contract_spend > 0")
                .sort_values("off_contract_spend", ascending=False)["subcategory"])
    hr = 6
    ws.cell(row=hr, column=2, value="Hotel")
    for j, s in enumerate(subs, start=3):
        ws.cell(row=hr, column=j, value=s)
    tot_col = 3 + len(subs)
    ws.cell(row=hr, column=tot_col, value="Total off-contract")
    ws.cell(row=hr, column=tot_col + 1, value="Premium paid vs contract")
    style_header_row(ws, hr, 2, tot_col + 1)
    ws.row_dimensions[hr].height = 32
    ws.column_dimensions["B"].width = 28
    for j in range(3, tot_col + 2):
        ws.column_dimensions[get_column_letter(j)].width = 14
    for i, h in enumerate(hotels):
        r = hr + 1 + i
        ws.cell(row=r, column=2, value=h)
        body(ws.cell(row=r, column=2))
        for j in range(3, tot_col):
            L = get_column_letter(j)
            ws.cell(row=r, column=j, value=(f'=SUMIFS({D["spend"]},{fyd()},{D["hotel"]},$B{r},'
                                            f'{D["subcategory"]},{L}${hr},{D["on_contract"]},"No")'))
            body(ws.cell(row=r, column=j), EUR)
        ws.cell(row=r, column=tot_col, value=f"=SUM(C{r}:{get_column_letter(tot_col-1)}{r})")
        body(ws.cell(row=r, column=tot_col), EUR, bold=True)
        ws.cell(row=r, column=tot_col + 1, value=(f'=SUMIFS({D["ppv_value"]},{fyd()},{D["hotel"]},$B{r},'
                                                  f'{D["on_contract"]},"No")'))
        body(ws.cell(row=r, column=tot_col + 1), EUR, bold=True, color=RED)
    last = hr + len(hotels)
    tr = last + 1
    ws.cell(row=tr, column=2, value="Group")
    for j in range(3, tot_col + 2):
        L = get_column_letter(j)
        ws.cell(row=tr, column=j, value=f"=SUM({L}{hr+1}:{L}{last})")
    for j in range(2, tot_col + 2):
        body(ws.cell(row=tr, column=j), EUR, bold=True)
        ws.cell(row=tr, column=j).fill = fill(MID)
    ws.conditional_formatting.add(
        f"C{hr+1}:{get_column_letter(tot_col-1)}{last}",
        ColorScaleRule(start_type="min", start_color="FFFFFF", end_type="max", end_color="E07B39"))

    # Top supplier/hotel combinations (static extract)
    oc = results["off_contract"].head(15)
    r0 = tr + 3
    ws.cell(row=r0 - 1, column=2, value="Top 15 off-contract combinations by premium paid (all years)").font = font(11, True, NAVY)
    heads = ["Hotel", "Subcategory", "Supplier", "Off-contract spend", "Premium paid", "Premium %"]
    for j, h in enumerate(heads, start=2):
        ws.cell(row=r0, column=j, value=h)
    style_header_row(ws, r0, 2, 7)
    for i, o in enumerate(oc.itertuples()):
        r = r0 + 1 + i
        vals = [o.hotel_name, o.subcategory, o.supplier_name, o.off_contract_spend, o.premium_paid, o.premium_pct]
        for j, (v, fmt) in enumerate(zip(vals, [None, None, None, EUR, EUR, PCT]), start=2):
            ws.cell(row=r, column=j, value=v)
            body(ws.cell(row=r, column=j), fmt)
    ws.freeze_panes = ws.cell(row=hr + 1, column=3)


def build_exceptions(ws, results):
    sheet_title(ws, "Invoice Exceptions — three-way match failures",
                "Lines where the invoiced price differs from the PO price, or the invoiced "
                "quantity differs from goods received. Sorted by € discrepancy (over-billing first).")
    ex = results["invoice_exceptions"].copy()
    ex["invoice_date"] = pd.to_datetime(ex["invoice_date"]).dt.date
    hr = 7
    ws["B4"], ws["C4"] = "Exception lines", f"=COUNTA(B{hr+1}:B{hr+len(ex)})"
    ws["B5"], ws["C5"] = "Net discrepancy (EUR)", f"=SUM(P{hr+1}:P{hr+len(ex)})"
    ws["E4"], ws["F4"] = "Over-billed (EUR)", f'=SUMIF(P{hr+1}:P{hr+len(ex)},">0")'
    ws["E5"], ws["F5"] = "Under-billed (EUR)", f'=SUMIF(P{hr+1}:P{hr+len(ex)},"<0")'
    for c in ("B4", "B5", "E4", "E5"):
        ws[c].font = font(10, True, SLATE)
    for c, fmt in (("C4", INT), ("C5", EUR), ("F4", EUR), ("F5", EUR)):
        body(ws[c], fmt, bold=True)
    heads = ["Invoice", "PO line", "Invoice date", "Hotel", "Supplier", "Product", "Exception",
             "Qty ordered", "Qty received", "Qty invoiced", "PO price", "Invoice price",
             "Expected value", "Invoiced value", "Discrepancy"]
    for j, h in enumerate(heads, start=2):
        ws.cell(row=hr, column=j, value=h)
    style_header_row(ws, hr, 2, 1 + len(heads))
    ws.row_dimensions[hr].height = 30
    set_widths(ws, {"B": 20, "C": 13, "D": 11, "E": 25, "F": 25, "G": 30, "H": 16, "I": 9,
                    "J": 9, "K": 9, "L": 10, "M": 10, "N": 12, "O": 12, "P": 12})
    fmts = [None, None, "yyyy-mm-dd", None, None, None, None, INT, INT, INT, EUR2, EUR2, EUR2, EUR2, EUR2]
    for i, row in enumerate(ex.itertuples(index=False)):
        r = hr + 1 + i
        for j, v in enumerate(row, start=2):
            ws.cell(row=r, column=j, value=v)
    for j, fmt in enumerate(fmts, start=2):
        for r in range(hr + 1, hr + 1 + len(ex)):
            cell = ws.cell(row=r, column=j)
            cell.font = font(9)
            if fmt:
                cell.number_format = fmt
    last = hr + len(ex)
    ws.conditional_formatting.add(f"P{hr+1}:P{last}", CellIsRule(operator="greaterThan", formula=["0"], font=Font(name=FONT, size=9, color=RED)))
    ws.conditional_formatting.add(f"P{hr+1}:P{last}", CellIsRule(operator="lessThan", formula=["0"], font=Font(name=FONT, size=9, color=GREEN)))
    ws.auto_filter.ref = f"B{hr}:P{last}"
    ws.freeze_panes = ws.cell(row=hr + 1, column=3)


def build_stock(ws, S, results, fys):
    sheet_title(ws, "Bar Stock Variance — beer, wine and spirits",
                "Monthly stock count reconciliation: counted closing stock vs theoretical "
                "(opening + received − sold through the till). Negative = unexplained loss.")
    selector_note(ws)
    hotels = list(results["hotel_summary"]["hotel_name"])
    subs = cfg.STOCK_COUNTED_SUBCATEGORIES
    hr = 6
    heads = ["Hotel"] + [f"{s} shrink %" for s in subs] + ["Total shrink %", "Variance (EUR)"]
    for j, h in enumerate(heads, start=2):
        ws.cell(row=hr, column=j, value=h)
    style_header_row(ws, hr, 2, 1 + len(heads))
    ws.row_dimensions[hr].height = 32
    set_widths(ws, {"B": 28, "C": 14, "D": 14, "E": 14, "F": 14, "G": 15})
    for i, h in enumerate(hotels):
        r = hr + 1 + i
        ws.cell(row=r, column=2, value=h)
        body(ws.cell(row=r, column=2))
        for j, s in enumerate(subs, start=3):
            cr = f'{fys()},{S["hotel"]},$B{r},{S["subcategory"]},"{s}"'
            ws.cell(row=r, column=j, value=f"=IFERROR(-SUMIFS({S['variance_qty']},{cr})/SUMIFS({S['sold_qty_pos']},{cr}),0)")
            body(ws.cell(row=r, column=j), PCT)
        cr = f'{fys()},{S["hotel"]},$B{r}'
        ws.cell(row=r, column=6, value=f"=IFERROR(-SUMIFS({S['variance_qty']},{cr})/SUMIFS({S['sold_qty_pos']},{cr}),0)")
        body(ws.cell(row=r, column=6), PCT, bold=True)
        ws.cell(row=r, column=7, value=f"=SUMIFS({S['variance_value']},{cr})")
        body(ws.cell(row=r, column=7), EUR, bold=True, color=RED)
    last = hr + len(hotels)
    ws.conditional_formatting.add(f"C{hr+1}:F{last}", ColorScaleRule(
        start_type="num", start_value=0, start_color="FFFFFF", mid_type="num", mid_value=0.025,
        mid_color="FFEB84", end_type="num", end_value=0.06, end_color="F8696B"))
    ws.cell(row=last + 2, column=2, value=(
        "Industry rule of thumb: bar shrinkage above 2% of units sold warrants a spot-check "
        "of pouring practice, comps and wastage logging.")).font = font(9, False, SLATE, True)


def build_notes(ws):
    sheet_title(ws, "Notes & Assumptions")
    set_widths(ws, {"B": 40, "C": 14, "D": 90})
    ws["B5"] = "Editable assumptions (yellow cells feed the formulas)"
    ws["B5"].font = font(11, True, NAVY)
    rows = [
        (8, "On-contract spend target", 0.90, PCT, "Group policy target used on the Dashboard hotel chart."),
        (12, "Score weight: OTIF", 0.40, PCT, "Share of the supplier composite score from on-time-in-full delivery."),
        (13, "Score weight: invoice accuracy", 0.30, PCT, "Share from clean three-way match."),
        (14, "Score weight: price discipline", 0.30, PCT, "Share from price discipline."),
        (15, "Price gap for zero price points", 0.05, PCT, "A supplier with a price gap at or above this scores 0 on price."),
        (16, "Rating A threshold (score ≥)", 85, "0", "A - Preferred."),
        (17, "Rating B threshold (score ≥)", 75, "0", "B - Monitor. Below this: C - Review."),
    ]
    ws["B7"], ws["B11"] = "Targets", "Supplier scorecard"
    ws["B7"].font = ws["B11"].font = font(10, True, SLATE)
    for r, label, val, fmt, note in rows:
        ws.cell(row=r, column=2, value=label).font = font(10)
        c = ws.cell(row=r, column=3, value=val)
        c.number_format, c.fill, c.font, c.border = fmt, fill(INPUT_FILL), font(10, True, "0000FF"), BORDER
        c.alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=4, value=note).font = font(9, False, SLATE)
    ws["B18"] = "Weights total (should be 100%)"
    ws["C18"] = "=C12+C13+C14"
    ws["C18"].number_format = PCT
    ws["B18"].font = font(10)
    body(ws["C18"], PCT, bold=True)

    defs = [
        ("About the data", None),
        ("Synthetic dataset", "Every hotel, supplier and transaction is fictional, generated by src/generate_data.py "
         "(seed 42) to mirror how a multi-property Irish hotel group buys food, beverage and operating supplies. "
         "No real company data is used."),
        ("Period", "1 Oct 2024 – 30 Sep 2026. Fiscal and contract years run October to September."),
        ("Scenario built in", "A group compliance push cuts off-contract buying at most hotels in FY26, while one property slips; "
         "contracted suppliers drift above agreed prices as each contract year ages."),
        ("KPI definitions", None),
        ("Spend", "Invoiced value (qty invoiced × invoice unit price)."),
        ("On-contract spend %", "Share of spend placed with the contracted supplier at the contract price."),
        ("Price variance vs contract", "(Invoice unit price − contract unit price) × qty invoiced. Positive = paid more than contract."),
        ("Contract price creep", "Positive price variance on contracted lines only: the supplier charging above the agreed price."),
        ("Three-way match", "PO, goods received note and invoice agree: invoice price = PO price (±€0.005) and invoiced qty = received qty."),
        ("Invoice over-billing", "Sum of positive (invoiced value − received qty × PO price)."),
        ("OTIF", "Delivered on or before the expected date AND full quantity received, measured per PO line."),
        ("Bar stock shrinkage", "Theoretical closing stock (opening + received − sold per till) minus counted closing stock, at average cost."),
        ("Value leakage identified", "Off-contract premium + contract creep + over-billing + shrinkage. An indicative opportunity size; "
         "components can overlap, so it is not a savings forecast."),
        ("How the workbook works", None),
        ("Formulas", "Every KPI, table and chart series is a live SUMIFS/COUNTIFS/AVERAGEIFS over the Data and Stock_Data sheets. "
         "Change the fiscal-year selector on the Dashboard, or paste refreshed rows into the data sheets, and everything recalculates."),
        ("Refresh", "Run `python run_pipeline.py` to regenerate the data, rerun the analysis and rebuild this workbook."),
    ]
    r = 21
    for label, text in defs:
        if text is None:
            r += 1
            ws.cell(row=r, column=2, value=label).font = font(11, True, NAVY)
        else:
            ws.cell(row=r, column=2, value=label).font = font(10, True)
            ws.cell(row=r, column=2).alignment = Alignment(vertical="top")
            t = ws.cell(row=r, column=4, value=text)
            t.font = font(10)
            t.alignment = Alignment(wrap_text=True, vertical="top")
            ws.row_dimensions[r].height = 15 * max(1, len(text) // 95 + 1)
        r += 1


def main():
    tables = analysis.load_raw()
    results = analysis.run(tables)
    path = build(results, tables)
    print(f"Dashboard written to {path}")


if __name__ == "__main__":
    main()
