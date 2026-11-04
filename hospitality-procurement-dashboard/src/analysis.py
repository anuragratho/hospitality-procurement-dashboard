"""
Procurement analytics pipeline.

Joins purchase orders, goods receipts and invoices into one line-level fact table,
then derives the KPI tables that feed the Excel dashboard:

    kpi_summary          headline KPIs, FY26 vs FY25
    monthly_trend        spend, compliance, delivery and match performance by month
    spend_by_category    spend by category / subcategory
    hotel_summary        spend and controls by property
    supplier_scorecard   OTIF, invoice accuracy, price discipline, composite score
    supplier_pareto      cumulative spend concentration
    price_variance       paid price vs contract price by product
    off_contract         maverick spend and premium paid, by hotel and subcategory
    invoice_exceptions   three-way-match failures with over-billed amounts
    stock_variance       bar stock shrinkage by hotel and drinks category

Fiscal year runs October to September (FY25 = Oct 2024 - Sep 2025), matching the
annual supplier contract cycle.

Run directly:  python -m src.analysis
"""

from __future__ import annotations

import numpy as np
import pandas as pd

from src import config as cfg

PRICE_TOL = 0.005   # euro tolerance when comparing unit prices


# --------------------------------------------------------------------------- #
# Loading and modelling
# --------------------------------------------------------------------------- #
def load_raw() -> dict[str, pd.DataFrame]:
    date_cols = {
        "purchase_orders": ["order_date", "expected_delivery_date"],
        "deliveries": ["delivery_date"],
        "invoices": ["invoice_date"],
    }
    tables = {}
    for name in ["hotels", "suppliers", "products", "contract_prices", "purchase_orders",
                 "deliveries", "invoices", "stock_counts", "market_prices"]:
        tables[name] = pd.read_csv(cfg.RAW_DIR / f"{name}.csv",
                                   parse_dates=date_cols.get(name, False))
    return tables


def fiscal_year(dates: pd.Series) -> pd.Series:
    """Oct-Sep fiscal year label, e.g. 2025-03-01 -> 'FY25', 2025-10-01 -> 'FY26'."""
    fy = dates.dt.year + (dates.dt.month >= 10).astype(int)
    return "FY" + (fy % 100).astype(str).str.zfill(2)


def build_fact_lines(t: dict[str, pd.DataFrame]) -> pd.DataFrame:
    """One row per PO line with delivery, invoice and master-data attributes."""
    f = (
        t["purchase_orders"]
        .merge(t["deliveries"][["po_line_id", "grn_number", "delivery_date", "qty_received"]],
               on="po_line_id", how="left", validate="one_to_one")
        .merge(t["invoices"][["po_line_id", "invoice_number", "invoice_date", "qty_invoiced",
                              "invoice_unit_price", "invoice_line_value"]],
               on="po_line_id", how="left", validate="one_to_one")
        .merge(t["products"][["product_id", "product_name", "category", "subcategory", "unit"]],
               on="product_id", how="left", validate="many_to_one")
        .merge(t["hotels"][["hotel_id", "hotel_name", "city", "rooms"]],
               on="hotel_id", how="left", validate="many_to_one")
        .merge(t["suppliers"][["supplier_id", "supplier_name", "is_contracted"]],
               on="supplier_id", how="left", validate="many_to_one")
    )

    f["month"] = f["order_date"].dt.to_period("M").astype(str)
    f["fiscal_year"] = fiscal_year(f["order_date"])

    # Delivery performance
    f["on_time"] = f["delivery_date"] <= f["expected_delivery_date"]
    f["in_full"] = f["qty_received"] >= f["qty_ordered"]
    f["otif"] = f["on_time"] & f["in_full"]
    f["days_late"] = (f["delivery_date"] - f["expected_delivery_date"]).dt.days.clip(lower=0)

    # Three-way match: PO price vs invoice price, received qty vs invoiced qty
    f["price_match"] = (f["invoice_unit_price"] - f["po_unit_price"]).abs() <= PRICE_TOL
    f["qty_match"] = f["qty_invoiced"] == f["qty_received"]
    f["three_way_match"] = f["price_match"] & f["qty_match"]
    f["expected_invoice_value"] = (f["qty_received"] * f["po_unit_price"]).round(2)
    f["invoice_discrepancy"] = (f["invoice_line_value"] - f["expected_invoice_value"]).round(2)

    # Purchase price variance vs the contract benchmark (positive = paid more)
    f["ppv_value"] = ((f["invoice_unit_price"] - f["contract_unit_price"]) * f["qty_invoiced"]).round(2)
    f["benchmark_value"] = (f["contract_unit_price"] * f["qty_invoiced"]).round(2)
    f["spend"] = f["invoice_line_value"]
    return f


# --------------------------------------------------------------------------- #
# KPI tables
# --------------------------------------------------------------------------- #
def _kpis_for(f: pd.DataFrame, stock: pd.DataFrame) -> dict[str, float]:
    spend = f["spend"].sum()
    off = f.loc[~f["on_contract"], "spend"].sum()
    overbilled = f.loc[f["invoice_discrepancy"] > 0, "invoice_discrepancy"].sum()
    off_premium = f.loc[~f["on_contract"], "ppv_value"].sum()
    contract_creep = f.loc[f["on_contract"], "ppv_value"].clip(lower=0).sum()
    shrink_value = -stock["variance_value"].sum()
    return {
        "Total spend (EUR)": spend,
        "Purchase orders": f["po_number"].nunique(),
        "PO lines": len(f),
        "Active suppliers": f["supplier_id"].nunique(),
        "On-contract spend %": 1 - off / spend,
        "Off-contract spend (EUR)": off,
        "Price variance vs contract (EUR)": f["ppv_value"].sum(),
        "Price variance vs contract %": f["ppv_value"].sum() / f["benchmark_value"].sum(),
        "Three-way match rate %": f["three_way_match"].mean(),
        "Invoice over-billing (EUR)": overbilled,
        "Supplier OTIF %": f["otif"].mean(),
        "Bar stock shrinkage (EUR)": shrink_value,
        "Bar stock shrinkage %": -stock["variance_qty"].sum() / stock["sold_qty_pos"].sum(),
        "Value leakage identified (EUR)": off_premium + contract_creep + overbilled + shrink_value,
    }


def kpi_summary(f: pd.DataFrame, stock: pd.DataFrame) -> pd.DataFrame:
    stock = stock.assign(fiscal_year=fiscal_year(pd.to_datetime(stock["month"])))
    rows = {}
    for fy in sorted(f["fiscal_year"].unique()):
        rows[fy] = _kpis_for(f[f["fiscal_year"] == fy], stock[stock["fiscal_year"] == fy])
    rows["Total"] = _kpis_for(f, stock)
    out = pd.DataFrame(rows)
    out.index.name = "kpi"
    fys = [c for c in out.columns if c.startswith("FY")]
    if len(fys) >= 2:
        prev, curr = fys[-2], fys[-1]
        out["change"] = out[curr] - out[prev]
    return out.reset_index()


def monthly_trend(f: pd.DataFrame) -> pd.DataFrame:
    g = f.groupby("month")
    out = pd.DataFrame({
        "spend": g["spend"].sum(),
        "on_contract_spend": f[f["on_contract"]].groupby("month")["spend"].sum(),
        "off_contract_spend": f[~f["on_contract"]].groupby("month")["spend"].sum(),
        "ppv_value": g["ppv_value"].sum(),
        "otif_pct": g["otif"].mean(),
        "three_way_match_pct": g["three_way_match"].mean(),
        "po_lines": g.size(),
    }).fillna(0)
    out["on_contract_pct"] = out["on_contract_spend"] / out["spend"]
    contracted = f[f["on_contract"]]
    out["contract_price_drift_pct"] = (
        contracted.groupby("month").apply(
            lambda x: (x["invoice_unit_price"] * x["qty_invoiced"]).sum() / x["benchmark_value"].sum() - 1,
            include_groups=False)
    )
    return out.reset_index()


def spend_by_category(f: pd.DataFrame) -> pd.DataFrame:
    out = (f.groupby(["category", "subcategory"])
             .agg(spend=("spend", "sum"), po_lines=("po_line_id", "count"),
                  off_contract_spend=("spend", lambda s: s[~f.loc[s.index, "on_contract"]].sum()),
                  ppv_value=("ppv_value", "sum"))
             .reset_index())
    total = out["spend"].sum()
    out["share_of_spend"] = out["spend"] / total
    out["on_contract_pct"] = 1 - out["off_contract_spend"] / out["spend"]
    return out.sort_values(["category", "spend"], ascending=[True, False]).reset_index(drop=True)


def hotel_summary(f: pd.DataFrame, stock: pd.DataFrame) -> pd.DataFrame:
    g = f.groupby(["hotel_id", "hotel_name", "city", "rooms"])
    out = g.agg(
        spend=("spend", "sum"),
        off_contract_spend=("spend", lambda s: s[~f.loc[s.index, "on_contract"]].sum()),
        ppv_value=("ppv_value", "sum"),
        otif_pct=("otif", "mean"),
        three_way_match_pct=("three_way_match", "mean"),
    ).reset_index()
    months = f["month"].nunique()
    out["annual_spend_per_room"] = out["spend"] / out["rooms"] / (months / 12)
    out["on_contract_pct"] = 1 - out["off_contract_spend"] / out["spend"]
    shrink = stock.groupby("hotel_id").agg(shrink_value=("variance_value", lambda s: -s.sum()),
                                           var_qty=("variance_qty", "sum"),
                                           sold=("sold_qty_pos", "sum"))
    shrink["shrink_pct"] = -shrink["var_qty"] / shrink["sold"]
    out = out.merge(shrink[["shrink_value", "shrink_pct"]], on="hotel_id", how="left")
    return out.sort_values("spend", ascending=False).reset_index(drop=True)


def supplier_scorecard(f: pd.DataFrame) -> pd.DataFrame:
    """Composite score (0-100): 40% OTIF, 30% invoice accuracy, 30% price discipline.

    Price discipline is scored on the price gap: the larger of (a) the share of
    invoice value charged above the agreed PO price and (b) the premium paid over
    the group contract benchmark. 0% gap -> 100 points, 5% or more -> 0 points.
    Off-contract suppliers therefore score low on price even when they invoice
    exactly what was quoted, because the quote itself is above contract.
    """
    f = f.assign(overcharge=(f["invoice_unit_price"] - f["po_unit_price"]).clip(lower=0) * f["qty_invoiced"])
    g = f.groupby(["supplier_id", "supplier_name", "is_contracted"])
    out = g.agg(
        spend=("spend", "sum"),
        po_lines=("po_line_id", "count"),
        hotels_served=("hotel_id", "nunique"),
        otif_pct=("otif", "mean"),
        on_time_pct=("on_time", "mean"),
        in_full_pct=("in_full", "mean"),
        avg_days_late=("days_late", lambda s: s[s > 0].mean() if (s > 0).any() else 0.0),
        invoice_accuracy_pct=("three_way_match", "mean"),
        overcharge=("overcharge", "sum"),
        ppv_value=("ppv_value", "sum"),
        benchmark_value=("benchmark_value", "sum"),
    ).reset_index()
    out["share_of_spend"] = out["spend"] / out["spend"].sum()
    out["overcharge_pct"] = out["overcharge"] / out["spend"]
    out["price_vs_contract_pct"] = out["ppv_value"] / out["benchmark_value"]
    out["price_gap_pct"] = np.maximum(out["overcharge_pct"], out["price_vs_contract_pct"].clip(lower=0))
    price_score = (1 - out["price_gap_pct"] / 0.05).clip(0, 1)
    out["score"] = (100 * (0.4 * out["otif_pct"] + 0.3 * out["invoice_accuracy_pct"]
                           + 0.3 * price_score)).round(1)
    out["rating"] = pd.cut(out["score"], [-1, 75, 85, 101], labels=["C - Review", "B - Monitor", "A - Preferred"]).astype(str)
    out = out.drop(columns=["benchmark_value"])
    return out.sort_values("spend", ascending=False).reset_index(drop=True)


def supplier_pareto(scorecard: pd.DataFrame) -> pd.DataFrame:
    p = scorecard[["supplier_name", "spend"]].sort_values("spend", ascending=False).reset_index(drop=True)
    p["cumulative_share"] = p["spend"].cumsum() / p["spend"].sum()
    p["rank"] = np.arange(1, len(p) + 1)
    return p


def price_variance(f: pd.DataFrame) -> pd.DataFrame:
    """Contracted lines only: how far invoiced prices drift above the agreed price."""
    c = f[f["on_contract"]]
    g = c.groupby(["product_id", "product_name", "subcategory", "supplier_name", "contract_year"])
    out = g.agg(
        contract_price=("contract_unit_price", "first"),
        qty=("qty_invoiced", "sum"),
        paid_value=("spend", "sum"),
        ppv_value=("ppv_value", "sum"),
        lines=("po_line_id", "count"),
        lines_above_contract=("invoice_unit_price",
                              lambda s: (s > c.loc[s.index, "contract_unit_price"] + PRICE_TOL).sum()),
    ).reset_index()
    out["avg_paid_price"] = out["paid_value"] / out["qty"]
    out["variance_pct"] = out["avg_paid_price"] / out["contract_price"] - 1
    out["pct_lines_above_contract"] = out["lines_above_contract"] / out["lines"]
    return out.sort_values("ppv_value", ascending=False).reset_index(drop=True)


def off_contract(f: pd.DataFrame) -> pd.DataFrame:
    o = f[~f["on_contract"]]
    out = (o.groupby(["hotel_name", "subcategory", "supplier_name"])
             .agg(off_contract_spend=("spend", "sum"), lines=("po_line_id", "count"),
                  premium_paid=("ppv_value", "sum"))
             .reset_index())
    out["premium_pct"] = out["premium_paid"] / (out["off_contract_spend"] - out["premium_paid"])
    return out.sort_values("premium_paid", ascending=False).reset_index(drop=True)


def invoice_exceptions(f: pd.DataFrame) -> pd.DataFrame:
    e = f[~f["three_way_match"]].copy()
    e["exception_type"] = np.select(
        [~e["price_match"] & ~e["qty_match"], ~e["price_match"], ~e["qty_match"]],
        ["Price and quantity", "Price", "Quantity"], default="")
    cols = ["invoice_number", "po_line_id", "invoice_date", "hotel_name", "supplier_name",
            "product_name", "exception_type", "qty_ordered", "qty_received", "qty_invoiced",
            "po_unit_price", "invoice_unit_price", "expected_invoice_value",
            "invoice_line_value", "invoice_discrepancy"]
    return e[cols].sort_values("invoice_discrepancy", ascending=False).reset_index(drop=True)


def stock_variance(stock: pd.DataFrame, t: dict[str, pd.DataFrame]) -> pd.DataFrame:
    s = (stock.merge(t["products"][["product_id", "product_name", "subcategory"]], on="product_id")
              .merge(t["hotels"][["hotel_id", "hotel_name"]], on="hotel_id"))
    out = (s.groupby(["hotel_name", "subcategory"])
             .agg(sold_qty=("sold_qty_pos", "sum"), variance_qty=("variance_qty", "sum"),
                  variance_value=("variance_value", "sum"))
             .reset_index())
    out["shrink_pct"] = -out["variance_qty"] / out["sold_qty"]
    return out.sort_values("variance_value").reset_index(drop=True)


# --------------------------------------------------------------------------- #
# Orchestration
# --------------------------------------------------------------------------- #
def run(tables: dict[str, pd.DataFrame] | None = None) -> dict[str, pd.DataFrame]:
    t = tables if tables is not None else load_raw()
    # Normalise dtypes when tables come straight from the generator
    for name, cols in {"purchase_orders": ["order_date", "expected_delivery_date"],
                       "deliveries": ["delivery_date"], "invoices": ["invoice_date"]}.items():
        for c in cols:
            t[name][c] = pd.to_datetime(t[name][c])

    f = build_fact_lines(t)
    stock = t["stock_counts"]
    sc = supplier_scorecard(f)
    results = {
        "fact_lines": f,
        "kpi_summary": kpi_summary(f, stock),
        "monthly_trend": monthly_trend(f),
        "spend_by_category": spend_by_category(f),
        "hotel_summary": hotel_summary(f, stock),
        "supplier_scorecard": sc,
        "supplier_pareto": supplier_pareto(sc),
        "price_variance": price_variance(f),
        "off_contract": off_contract(f),
        "invoice_exceptions": invoice_exceptions(f),
        "stock_variance": stock_variance(stock, t),
    }
    cfg.PROCESSED_DIR.mkdir(parents=True, exist_ok=True)
    for name, df in results.items():
        if name != "fact_lines":
            df.to_csv(cfg.PROCESSED_DIR / f"{name}.csv", index=False)
    return results


if __name__ == "__main__":
    print("Running procurement analysis ...")
    res = run()
    pd.set_option("display.width", 160)
    pd.set_option("display.float_format", lambda v: f"{v:,.3f}")
    print(res["kpi_summary"].to_string(index=False))
