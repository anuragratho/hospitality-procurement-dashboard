"""
Data-quality and logic tests.

    pytest -q

The generator is run once per test session (in memory, nothing written to disk
except the processed outputs the analysis always saves).
"""

from __future__ import annotations

import numpy as np
import pandas as pd
import pytest

from src import analysis, generate_data
from src import config as cfg


@pytest.fixture(scope="session")
def tables():
    return generate_data.generate_all(seed=cfg.RANDOM_SEED)


@pytest.fixture(scope="session")
def results(tables):
    return analysis.run({k: v.copy() for k, v in tables.items()})


# --------------------------------------------------------------------------- #
# Generator: structure and integrity
# --------------------------------------------------------------------------- #
def test_generation_is_reproducible(tables):
    again = generate_data.generate_all(seed=cfg.RANDOM_SEED)
    pd.testing.assert_frame_equal(tables["purchase_orders"], again["purchase_orders"])
    pd.testing.assert_frame_equal(tables["invoices"], again["invoices"])


def test_different_seed_gives_different_data(tables):
    other = generate_data.generate_all(seed=7)
    assert not tables["purchase_orders"]["po_line_value"].equals(other["purchase_orders"]["po_line_value"])


def test_primary_keys_unique(tables):
    assert tables["purchase_orders"]["po_line_id"].is_unique
    assert tables["hotels"]["hotel_id"].is_unique
    assert tables["suppliers"]["supplier_id"].is_unique
    assert tables["products"]["product_id"].is_unique


def test_referential_integrity(tables):
    po = tables["purchase_orders"]
    assert po["hotel_id"].isin(tables["hotels"]["hotel_id"]).all()
    assert po["supplier_id"].isin(tables["suppliers"]["supplier_id"]).all()
    assert po["product_id"].isin(tables["products"]["product_id"]).all()
    assert tables["deliveries"]["po_line_id"].isin(po["po_line_id"]).all()
    assert tables["invoices"]["po_line_id"].isin(po["po_line_id"]).all()


def test_one_delivery_and_one_invoice_per_line(tables):
    n = len(tables["purchase_orders"])
    assert len(tables["deliveries"]) == n and tables["deliveries"]["po_line_id"].is_unique
    assert len(tables["invoices"]) == n and tables["invoices"]["po_line_id"].is_unique


def test_dates_are_in_window_and_ordered(tables):
    po = tables["purchase_orders"]
    assert po["order_date"].min() >= pd.Timestamp(cfg.START_DATE)
    assert po["order_date"].max() <= pd.Timestamp(cfg.END_DATE)
    d = po.merge(tables["deliveries"], on="po_line_id")
    assert (d["delivery_date"] >= d["order_date"]).all()
    i = d.merge(tables["invoices"][["po_line_id", "invoice_date"]], on="po_line_id")
    assert (i["invoice_date"] >= i["delivery_date"]).all()


def test_quantities_valid(tables):
    d = tables["purchase_orders"].merge(tables["deliveries"], on="po_line_id")
    assert (d["qty_ordered"] > 0).all()
    assert (d["qty_received"] >= 0).all()
    assert (d["qty_received"] <= d["qty_ordered"]).all()


def test_on_contract_lines_use_preferred_supplier_at_contract_price(tables):
    po = tables["purchase_orders"].merge(tables["products"][["product_id", "preferred_supplier_id"]])
    on = po[po["on_contract"]]
    assert (on["supplier_id"] == on["preferred_supplier_id"]).all()
    assert np.allclose(on["po_unit_price"], on["contract_unit_price"])


def test_off_contract_lines_cost_more_than_contract(tables):
    po = tables["purchase_orders"]
    off = po[~po["on_contract"]]
    assert len(off) > 0
    assert (off["po_unit_price"] > off["contract_unit_price"]).all()
    off_sup = set(tables["suppliers"].query("not is_contracted")["supplier_id"])
    assert set(off["supplier_id"]) <= off_sup


def test_compliance_worse_at_low_discipline_sites(tables):
    po = tables["purchase_orders"]
    share = po.groupby("hotel_id").apply(
        lambda g: (g.loc[~g["on_contract"], "po_line_value"].sum() / g["po_line_value"].sum()),
        include_groups=False)
    assert share["H06"] > share["H01"]


def test_contract_prices_cover_both_years(tables):
    cp = tables["contract_prices"]
    assert set(cp["contract_year"]) == {"CY1", "CY2"}
    assert cp.groupby("product_id").size().eq(2).all()
    assert (cp["contract_price"] > 0).all()


def test_stock_reconciliation_arithmetic(tables):
    s = tables["stock_counts"]
    assert (s["theoretical_closing_qty"] == s["opening_qty"] + s["received_qty"] - s["sold_qty_pos"]).all()
    assert (s["variance_qty"] == s["counted_closing_qty"] - s["theoretical_closing_qty"]).all()
    assert (s["variance_qty"] <= 0).all()          # only losses are simulated
    assert (s["counted_closing_qty"] >= 0).all()


def test_spend_is_in_a_realistic_range(tables):
    total = tables["invoices"]["invoice_line_value"].sum()
    rooms = tables["hotels"]["rooms"].sum()
    per_room_year = total / rooms / 2
    assert 2_000 < per_room_year < 8_000


# --------------------------------------------------------------------------- #
# Analysis logic
# --------------------------------------------------------------------------- #
def test_fiscal_year_boundaries():
    s = pd.Series(pd.to_datetime(["2024-10-01", "2025-09-30", "2025-10-01", "2026-09-30"]))
    assert list(analysis.fiscal_year(s)) == ["FY25", "FY25", "FY26", "FY26"]


def test_fact_table_one_row_per_line(tables, results):
    assert len(results["fact_lines"]) == len(tables["purchase_orders"])


def test_total_spend_reconciles_to_invoices(tables, results):
    k = results["kpi_summary"].set_index("kpi")
    assert k.loc["Total spend (EUR)", "Total"] == pytest.approx(tables["invoices"]["invoice_line_value"].sum())
    assert k.loc["Total spend (EUR)", "FY25"] + k.loc["Total spend (EUR)", "FY26"] == pytest.approx(
        k.loc["Total spend (EUR)", "Total"])


def test_three_way_match_flags(results):
    f = results["fact_lines"]
    matched = f[f["three_way_match"]]
    assert (matched["qty_invoiced"] == matched["qty_received"]).all()
    assert ((matched["invoice_unit_price"] - matched["po_unit_price"]).abs() <= analysis.PRICE_TOL).all()
    assert 0.80 < f["three_way_match"].mean() < 0.97


def test_invoice_exceptions_are_exactly_the_unmatched_lines(results):
    f = results["fact_lines"]
    assert len(results["invoice_exceptions"]) == (~f["three_way_match"]).sum()


def test_supplier_scorecard_bounds_and_ratings(results):
    sc = results["supplier_scorecard"]
    assert sc["score"].between(0, 100).all()
    assert set(sc["rating"]) <= {"A - Preferred", "B - Monitor", "C - Review"}
    assert sc["share_of_spend"].sum() == pytest.approx(1.0)
    # Off-contract suppliers should never be rated Preferred: their prices sit above contract
    assert not (sc.loc[~sc["is_contracted"], "rating"] == "A - Preferred").any()


def test_pareto_is_monotonic(results):
    p = results["supplier_pareto"]
    assert p["cumulative_share"].is_monotonic_increasing
    assert p["cumulative_share"].iloc[-1] == pytest.approx(1.0)


def test_value_leakage_is_sum_of_components(results):
    f = results["fact_lines"]
    k = results["kpi_summary"].set_index("kpi")
    expected = (f.loc[~f["on_contract"], "ppv_value"].sum()
                + f.loc[f["on_contract"], "ppv_value"].clip(lower=0).sum()
                + f.loc[f["invoice_discrepancy"] > 0, "invoice_discrepancy"].sum()
                + k.loc["Bar stock shrinkage (EUR)", "Total"])
    assert k.loc["Value leakage identified (EUR)", "Total"] == pytest.approx(expected)


def test_price_creep_grows_through_contract_year(results):
    mt = results["monthly_trend"].set_index("month")["contract_price_drift_pct"]
    # Late in each contract year drift should exceed the first month of that year
    assert mt["2025-09"] > mt["2024-10"]
    assert mt["2026-09"] > mt["2025-10"]
