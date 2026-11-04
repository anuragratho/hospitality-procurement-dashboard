"""
Workbook structure tests: the dashboard builds, has the expected tabs, and its
KPIs are formulas rather than pasted numbers.
"""

from __future__ import annotations

import pytest
from openpyxl import load_workbook

from src import analysis, build_dashboard, generate_data
from src import config as cfg

EXPECTED_SHEETS = [
    "Dashboard", "Calc", "Supplier Scorecard", "Hotel Performance", "Price Variance",
    "Off-Contract", "Invoice Exceptions", "Stock Variance", "Notes & Assumptions",
    "Data", "Stock_Data",
]


@pytest.fixture(scope="module")
def workbook(tmp_path_factory):
    tables = generate_data.generate_all(seed=cfg.RANDOM_SEED)
    results = analysis.run({k: v.copy() for k, v in tables.items()})
    path = tmp_path_factory.mktemp("out") / "dashboard.xlsx"
    build_dashboard.build(results, tables, path=path)
    return load_workbook(path), results


def test_sheets_present(workbook):
    wb, _ = workbook
    assert wb.sheetnames == EXPECTED_SHEETS


def test_selector_has_validation(workbook):
    wb, _ = workbook
    ws = wb["Dashboard"]
    assert ws["C4"].value == "FY26"
    dvs = ws.data_validations.dataValidation
    assert any("C4" in str(dv.sqref) and "FY25" in dv.formula1 for dv in dvs)


def test_kpis_are_formulas(workbook):
    wb, _ = workbook
    calc = wb["Calc"]
    kpi_cells = [calc.cell(row=r, column=c).value for r in range(10, 24) for c in range(3, 7)]
    assert all(isinstance(v, str) and v.startswith("=") for v in kpi_cells)
    for addr in ("B7", "E7", "H7", "K7", "N7", "Q7"):
        assert str(wb["Dashboard"][addr].value).startswith("=Calc!")


def test_charts_present(workbook):
    wb, _ = workbook
    assert len(wb["Dashboard"]._charts) == 6


def test_data_sheet_row_count(workbook):
    wb, results = workbook
    assert wb["Data"].max_row == len(results["fact_lines"]) + 1


def test_scorecard_lists_every_supplier(workbook):
    wb, results = workbook
    ws = wb["Supplier Scorecard"]
    names = {ws.cell(row=r, column=2).value for r in range(7, 7 + len(results["supplier_scorecard"]))}
    assert names == set(results["supplier_scorecard"]["supplier_name"])
