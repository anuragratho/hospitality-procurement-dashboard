"""
Synthetic data generator.

Creates a realistic, relational procurement dataset for a fictional multi-property
hotel group: hotels, suppliers, products, contract price lists, purchase-order
lines, goods receipts (deliveries), supplier invoices and monthly bar stock counts.

The data deliberately contains the problems a procurement team looks for:
  * off-contract ("maverick") buying, worse at some properties than others
  * contracted suppliers drifting above agreed prices during the contract year
  * late and short deliveries
  * invoice price and quantity mismatches (three-way match failures)
  * bar stock shrinkage (counted stock below theoretical stock)

Run directly:  python -m src.generate_data
"""

from __future__ import annotations

import numpy as np
import pandas as pd

from src import config as cfg

FRESH_SUBCATS = {"Fresh Produce", "Meat & Poultry", "Seafood", "Dairy & Eggs", "Bakery"}
VOLATILE_SUBCATS = {"Seafood", "Fresh Produce"}


# --------------------------------------------------------------------------- #
# Master data
# --------------------------------------------------------------------------- #
def build_hotels() -> pd.DataFrame:
    return pd.DataFrame(
        cfg.HOTELS,
        columns=["hotel_id", "hotel_name", "city", "rooms", "compliance_discipline"],
    )


def build_suppliers() -> pd.DataFrame:
    rows = []
    for sid, name, subcats, contracted, lead, reliab, inv_acc, premium in cfg.SUPPLIERS:
        rows.append(
            {
                "supplier_id": sid,
                "supplier_name": name,
                "subcategories": "; ".join(subcats),
                "is_contracted": contracted,
                "lead_time_days": lead,
                "delivery_reliability": reliab,
                "invoice_accuracy": inv_acc,
                "price_creep_tendency": premium,
                "payment_terms_days": 30 if contracted else 0,
            }
        )
    return pd.DataFrame(rows)


def build_products() -> pd.DataFrame:
    sub_to_cat = {s: c for c, subs in cfg.CATEGORIES.items() for s in subs}
    # preferred (contracted) supplier for each subcategory
    preferred = {}
    for sid, _, subcats, contracted, *_ in cfg.SUPPLIERS:
        if contracted:
            for s in subcats:
                preferred.setdefault(s, sid)

    rows = []
    for i, (name, sub, unit, price, qty) in enumerate(cfg.PRODUCTS, start=1):
        rows.append(
            {
                "product_id": f"P{i:03d}",
                "product_name": name,
                "category": sub_to_cat[sub],
                "subcategory": sub,
                "unit": unit,
                "base_price": price,
                "base_weekly_qty_per_100_rooms": qty,
                "preferred_supplier_id": preferred[sub],
            }
        )
    return pd.DataFrame(rows)


# --------------------------------------------------------------------------- #
# Prices
# --------------------------------------------------------------------------- #
def build_market_prices(products: pd.DataFrame, rng: np.random.Generator) -> pd.DataFrame:
    """Monthly open-market price index per product (inflation + random walk)."""
    months = pd.period_range(cfg.START_DATE, cfg.END_DATE, freq="M")
    rows = []
    for p in products.itertuples():
        infl = cfg.CATEGORY_INFLATION[p.category]
        vol = 0.025 if p.subcategory in VOLATILE_SUBCATS else 0.008
        walk = 1.0
        for m_idx, month in enumerate(months):
            walk *= 1 + rng.normal(0, vol)
            trend = (1 + infl) ** (m_idx / 12)
            rows.append(
                {
                    "product_id": p.product_id,
                    "month": str(month),
                    "market_price": round(p.base_price * trend * walk, 4),
                }
            )
    return pd.DataFrame(rows)


def build_contract_prices(products: pd.DataFrame, market: pd.DataFrame,
                          rng: np.random.Generator) -> pd.DataFrame:
    """Two annual contract price lists (Oct-Sep). Year 2 is renegotiated off
    the market price at the end of year 1, with a small negotiated discount."""
    start = pd.Timestamp(cfg.START_DATE)
    cy1_from, cy1_to = start, start + pd.DateOffset(years=1) - pd.Timedelta(days=1)
    cy2_from, cy2_to = cy1_to + pd.Timedelta(days=1), pd.Timestamp(cfg.END_DATE)
    renegotiation_month = str((cy1_to).to_period("M"))

    mkt = market.set_index(["product_id", "month"])["market_price"]
    rows = []
    for p in products.itertuples():
        rows.append(
            {
                "product_id": p.product_id,
                "supplier_id": p.preferred_supplier_id,
                "contract_year": "CY1",
                "valid_from": cy1_from.date(),
                "valid_to": cy1_to.date(),
                "contract_price": round(p.base_price, 2),
            }
        )
        cy2_price = mkt[(p.product_id, renegotiation_month)] * rng.uniform(0.965, 0.995)
        rows.append(
            {
                "product_id": p.product_id,
                "supplier_id": p.preferred_supplier_id,
                "contract_year": "CY2",
                "valid_from": cy2_from.date(),
                "valid_to": cy2_to.date(),
                "contract_price": round(cy2_price, 2),
            }
        )
    return pd.DataFrame(rows)


# --------------------------------------------------------------------------- #
# Transactions
# --------------------------------------------------------------------------- #
def _order_dates() -> tuple[pd.DatetimeIndex, pd.DatetimeIndex]:
    days = pd.date_range(cfg.START_DATE, cfg.END_DATE, freq="D")
    mondays = days[days.dayofweek == 0]
    thursdays = days[days.dayofweek == 3]
    return mondays, thursdays


def build_purchase_orders(hotels, suppliers, products, contracts, market,
                          rng: np.random.Generator) -> pd.DataFrame:
    """One row per PO line. Lines are grouped into POs by hotel/supplier/date."""
    mondays, thursdays = _order_dates()
    sup = suppliers.set_index("supplier_id")
    spot_by_subcat: dict[str, list[str]] = {}
    for s in suppliers.itertuples():
        if not s.is_contracted:
            for sc in s.subcategories.split("; "):
                spot_by_subcat.setdefault(sc, []).append(s.supplier_id)

    contract_lookup = {
        (r.product_id, r.contract_year): r.contract_price for r in contracts.itertuples()
    }
    cy2_start = pd.Timestamp(contracts.loc[contracts.contract_year == "CY2", "valid_from"].iloc[0])
    mkt = market.set_index(["product_id", "month"])["market_price"].to_dict()

    lines = []
    for h in hotels.itertuples():
        off_contract_base = (1 - h.compliance_discipline) * 0.95
        for p in products.itertuples():
            fresh = p.subcategory in FRESH_SUBCATS
            dates = mondays.union(thursdays) if fresh else mondays
            per_order_share = 0.5 if fresh else 1.0
            for d in dates:
                if not fresh and rng.random() < 0.12:      # skipped week (stock on hand)
                    continue
                season = cfg.SEASONALITY[d.month - 1]
                expected = (p.base_weekly_qty_per_100_rooms * h.rooms / 100 * season
                            * per_order_share * cfg.VOLUME_SCALE)
                qty = max(1, int(round(expected * rng.gamma(9, 1 / 9))))

                contract_year = "CY2" if d >= cy2_start else "CY1"
                contract_price = contract_lookup[(p.product_id, contract_year)]
                month = str(d.to_period("M"))
                market_price = mkt[(p.product_id, month)]

                # Off-contract buying: more likely in peak months and at weaker sites
                peak_bump = 0.04 if season >= 1.15 else 0.0
                spot_options = spot_by_subcat.get(p.subcategory, [])
                year2_mult = cfg.YEAR2_OFF_CONTRACT_MULTIPLIER[h.hotel_id] if contract_year == "CY2" else 1.0
                off_prob = off_contract_base * year2_mult + peak_bump
                off_contract = bool(spot_options) and rng.random() < off_prob

                if off_contract:
                    supplier_id = rng.choice(spot_options)
                    prem = max(0.02, rng.normal(*cfg.OFF_CONTRACT_PREMIUM))
                    po_price = max(market_price, contract_price) * (1 + prem)
                else:
                    supplier_id = p.preferred_supplier_id
                    po_price = contract_price

                lines.append(
                    {
                        "hotel_id": h.hotel_id,
                        "supplier_id": supplier_id,
                        "product_id": p.product_id,
                        "order_date": d,
                        "contract_year": contract_year,
                        "on_contract": not off_contract,
                        "qty_ordered": qty,
                        "po_unit_price": round(po_price, 2),
                        "contract_unit_price": contract_price,
                        # Off-contract orders are placed by department heads, not the desk
                        "buyer": (rng.choice(cfg.BUYERS[1:]) if off_contract
                                  else (cfg.BUYERS[0] if rng.random() < 0.7
                                        else rng.choice(cfg.BUYERS[1:]))),
                    }
                )

    po = pd.DataFrame(lines)
    po = po.sort_values(["order_date", "hotel_id", "supplier_id", "product_id"]).reset_index(drop=True)

    # PO header ids
    keys = po[["hotel_id", "supplier_id", "order_date"]].drop_duplicates().reset_index(drop=True)
    keys["po_number"] = [f"PO{100000 + i}" for i in range(len(keys))]
    po = po.merge(keys, on=["hotel_id", "supplier_id", "order_date"], how="left")
    po["po_line_id"] = po["po_number"] + "-" + (po.groupby("po_number").cumcount() + 1).astype(str).str.zfill(2)

    lead = po["supplier_id"].map(sup["lead_time_days"]).clip(lower=1)
    po["expected_delivery_date"] = po["order_date"] + pd.to_timedelta(lead, unit="D")
    po["po_line_value"] = (po["qty_ordered"] * po["po_unit_price"]).round(2)

    cols = ["po_line_id", "po_number", "hotel_id", "supplier_id", "product_id", "order_date",
            "expected_delivery_date", "contract_year", "on_contract", "buyer", "qty_ordered",
            "po_unit_price", "contract_unit_price", "po_line_value"]
    return po[cols]


def build_deliveries(po: pd.DataFrame, suppliers: pd.DataFrame,
                     rng: np.random.Generator) -> pd.DataFrame:
    """Goods received notes. Delivery date is per PO; quantity can be short per line."""
    sup = suppliers.set_index("supplier_id")
    headers = po[["po_number", "supplier_id", "expected_delivery_date"]].drop_duplicates("po_number")
    reliab = headers["supplier_id"].map(sup["delivery_reliability"]).to_numpy()
    late = rng.random(len(headers)) > reliab
    delay = np.where(late, rng.integers(1, 5, len(headers)), 0)
    headers = headers.assign(
        delivery_date=headers["expected_delivery_date"] + pd.to_timedelta(delay, unit="D")
    )

    d = po[["po_line_id", "po_number", "supplier_id", "qty_ordered"]].merge(
        headers[["po_number", "delivery_date"]], on="po_number"
    )
    # Short deliveries are more likely with less reliable suppliers
    short_p = cfg.SHORT_DELIVERY_PROB * (1 + (1 - d["supplier_id"].map(sup["delivery_reliability"])) * 6)
    short = rng.random(len(d)) < short_p
    frac = rng.uniform(0.5, 0.9, len(d))
    received = np.where(short, np.floor(d["qty_ordered"] * frac), d["qty_ordered"]).astype(int)
    d["qty_received"] = np.maximum(received, 0)
    d["grn_number"] = "GRN" + d["po_number"].str[2:]
    return d[["grn_number", "po_line_id", "po_number", "delivery_date", "qty_received"]]


def build_invoices(po: pd.DataFrame, deliveries: pd.DataFrame, suppliers: pd.DataFrame,
                   rng: np.random.Generator) -> pd.DataFrame:
    """Supplier invoices. Errors: price creep above PO price on contract lines,
    random keying errors, and invoicing ordered rather than received quantity."""
    sup = suppliers.set_index("supplier_id")
    inv = po[["po_line_id", "po_number", "supplier_id", "order_date", "contract_year",
              "on_contract", "qty_ordered", "po_unit_price"]].merge(
        deliveries[["po_line_id", "delivery_date", "qty_received"]], on="po_line_id"
    )
    n = len(inv)

    # 1) Price creep: contracted suppliers drift above agreed price as the contract year ages
    creep_tend = inv["supplier_id"].map(sup["price_creep_tendency"]).to_numpy()
    cy1_start = pd.Timestamp(cfg.START_DATE)
    cy2_start = cy1_start + pd.DateOffset(years=1)
    year_start = np.where(inv["contract_year"] == "CY1", cy1_start, cy2_start).astype("datetime64[ns]")
    months_in = ((inv["order_date"].to_numpy().astype("datetime64[ns]") - year_start)
                 / np.timedelta64(1, "D") / 30.44)
    creep_prob = np.where(creep_tend > 0, cfg.PRICE_CREEP_PROB + creep_tend * 3, 0.0)
    creep_prob = np.where(inv["on_contract"], creep_prob * (0.4 + months_in / 10), 0.0)
    creep = rng.random(n) < creep_prob
    creep_pct = np.where(creep, creep_tend * 4.5 * (0.6 + months_in / 8) * rng.lognormal(0, 0.25, n), 0)

    # 2) Random keying errors (both directions), driven by supplier invoice accuracy
    acc = inv["supplier_id"].map(sup["invoice_accuracy"]).to_numpy()
    keying = rng.random(n) > acc
    key_pct = np.where(keying, rng.choice([-1, 1], n, p=[0.3, 0.7]) * rng.uniform(0.03, 0.12, n), 0)

    inv["invoice_unit_price"] = (inv["po_unit_price"] * (1 + creep_pct + key_pct)).round(2)

    # 3) Quantity: most suppliers bill what was delivered; some bill what was ordered
    bills_ordered = (inv["qty_received"] < inv["qty_ordered"]) & (rng.random(n) < 0.55)
    inv["qty_invoiced"] = np.where(bills_ordered, inv["qty_ordered"], inv["qty_received"])

    inv["invoice_line_value"] = (inv["qty_invoiced"] * inv["invoice_unit_price"]).round(2)
    inv["invoice_date"] = inv["delivery_date"] + pd.to_timedelta(rng.integers(0, 6, n), unit="D")
    inv["invoice_number"] = "INV-" + inv["supplier_id"] + "-" + inv["po_number"].str[2:]
    return inv[["invoice_number", "po_line_id", "po_number", "supplier_id", "invoice_date",
                "qty_invoiced", "invoice_unit_price", "invoice_line_value"]]


def build_stock_counts(hotels, products, po, deliveries, rng: np.random.Generator) -> pd.DataFrame:
    """Monthly bar-store stock reconciliation for beer, wine and spirits.

    theoretical closing = opening + received - sold (from the till)
    counted closing     = theoretical closing - unexplained loss
    """
    bar = products[products["subcategory"].isin(cfg.STOCK_COUNTED_SUBCATEGORIES)]
    rec = po[["po_line_id", "hotel_id", "product_id", "po_unit_price"]].merge(
        deliveries[["po_line_id", "delivery_date", "qty_received"]], on="po_line_id"
    )
    rec = rec[rec["product_id"].isin(bar["product_id"])]
    rec["month"] = rec["delivery_date"].dt.to_period("M").astype(str)
    months = [str(m) for m in pd.period_range(cfg.START_DATE, cfg.END_DATE, freq="M")]
    rec = rec[rec["month"].isin(months)]
    rec["value"] = rec["qty_received"] * rec["po_unit_price"]

    qty = rec.groupby(["hotel_id", "product_id", "month"])["qty_received"].sum()
    totals = rec.groupby(["hotel_id", "product_id"])[["value", "qty_received"]].sum()
    avg_cost_map = (totals["value"] / totals["qty_received"].clip(lower=1)).to_dict()

    # Shrinkage rate: weaker-discipline sites lose more; spirits lose most
    sub_factor = {"Beer & Cider": 0.8, "Wine": 1.0, "Spirits": 1.6}
    sub_of = bar.set_index("product_id")["subcategory"]
    disc = hotels.set_index("hotel_id")["compliance_discipline"]

    rows = []
    for hid, pid in totals.index:
        received_by_month = qty.loc[(hid, pid)].reindex(months, fill_value=0)
        base_shrink = 0.004 + (1 - disc[hid]) * 0.11
        shrink = base_shrink * sub_factor[sub_of[pid]]
        avg_cost = avg_cost_map[(hid, pid)]
        opening = int(round(received_by_month.iloc[:2].mean() * 0.5))
        for m in months:
            received = int(received_by_month[m])
            sold = int(round((opening + received) * rng.uniform(0.72, 0.82)))
            theoretical = opening + received - sold
            loss = rng.binomial(max(sold, 0), min(shrink * rng.lognormal(0, 0.35), 0.5))
            loss = min(loss, theoretical)
            counted = theoretical - loss
            rows.append(
                {
                    "hotel_id": hid,
                    "product_id": pid,
                    "month": m,
                    "opening_qty": opening,
                    "received_qty": received,
                    "sold_qty_pos": sold,
                    "theoretical_closing_qty": theoretical,
                    "counted_closing_qty": counted,
                    "variance_qty": counted - theoretical,
                    "unit_cost": round(avg_cost, 2),
                    "variance_value": round((counted - theoretical) * avg_cost, 2),
                }
            )
            opening = counted
    return pd.DataFrame(rows)


# --------------------------------------------------------------------------- #
# Orchestration
# --------------------------------------------------------------------------- #
def generate_all(seed: int = cfg.RANDOM_SEED) -> dict[str, pd.DataFrame]:
    rng = np.random.default_rng(seed)
    hotels = build_hotels()
    suppliers = build_suppliers()
    products = build_products()
    market = build_market_prices(products, rng)
    contracts = build_contract_prices(products, market, rng)
    po = build_purchase_orders(hotels, suppliers, products, contracts, market, rng)
    deliveries = build_deliveries(po, suppliers, rng)
    invoices = build_invoices(po, deliveries, suppliers, rng)
    stock = build_stock_counts(hotels, products, po, deliveries, rng)
    return {
        "hotels": hotels,
        "suppliers": suppliers,
        "products": products,
        "market_prices": market,
        "contract_prices": contracts,
        "purchase_orders": po,
        "deliveries": deliveries,
        "invoices": invoices,
        "stock_counts": stock,
    }


def save_all(tables: dict[str, pd.DataFrame]) -> None:
    cfg.RAW_DIR.mkdir(parents=True, exist_ok=True)
    for name, df in tables.items():
        df.to_csv(cfg.RAW_DIR / f"{name}.csv", index=False)


def main() -> dict[str, pd.DataFrame]:
    tables = generate_all()
    save_all(tables)
    for name, df in tables.items():
        print(f"  {name:<16} {len(df):>8,} rows")
    return tables


if __name__ == "__main__":
    print("Generating synthetic procurement data ...")
    main()
