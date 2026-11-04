"""
Static PNG charts for the README and portfolio write-ups.

Run directly:  python -m src.charts
"""

from __future__ import annotations

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import pandas as pd

from src import analysis
from src import config as cfg

BLUE = "#2a78d6"     # categorical slot 1 (validated pair, colour-blind safe)
ORANGE = "#eb6834"   # categorical slot 2
INK = "#1f2933"
MUTED = "#6b7785"
GRID = "#e6e9ee"

plt.rcParams.update({
    "font.family": "DejaVu Sans",
    "font.size": 10,
    "axes.edgecolor": GRID,
    "axes.labelcolor": MUTED,
    "axes.titlecolor": INK,
    "axes.titlesize": 12,
    "axes.titleweight": "bold",
    "axes.titlelocation": "left",
    "xtick.color": MUTED,
    "ytick.color": MUTED,
    "axes.spines.top": False,
    "axes.spines.right": False,
    "figure.dpi": 150,
})

EUR_K = mticker.FuncFormatter(lambda v, _: f"€{v/1000:,.0f}k")


def _finish(fig, name):
    cfg.IMAGES_DIR.mkdir(parents=True, exist_ok=True)
    fig.tight_layout()
    fig.savefig(cfg.IMAGES_DIR / name, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def monthly_spend(mt: pd.DataFrame):
    fig, ax = plt.subplots(figsize=(10, 4))
    x = range(len(mt))
    ax.bar(x, mt["on_contract_spend"], color=BLUE, width=0.72, label="On-contract",
           edgecolor="white", linewidth=1)
    ax.bar(x, mt["off_contract_spend"], bottom=mt["on_contract_spend"], color=ORANGE, width=0.72,
           label="Off-contract", edgecolor="white", linewidth=1)
    ax.set_xticks(list(x)[::2], [m for m in mt["month"]][::2], rotation=45, ha="right")
    ax.yaxis.set_major_formatter(EUR_K)
    ax.grid(axis="y", color=GRID, linewidth=0.8)
    ax.set_axisbelow(True)
    ax.set_title("Monthly spend, on- vs off-contract")
    ax.set_ylim(0, mt["spend"].max() * 1.15)
    ax.legend(frameon=False, loc="upper left", ncols=2)
    ax.axvline(11.5, color=MUTED, linewidth=1, linestyle=":")
    ax.text(11.6, ax.get_ylim()[1] * 0.97, "FY26 →", color=MUTED, fontsize=9, va="top")
    _finish(fig, "monthly_spend.png")


def price_drift(mt: pd.DataFrame):
    fig, ax = plt.subplots(figsize=(10, 3.6))
    ax.plot(mt["month"], mt["contract_price_drift_pct"], color=BLUE, linewidth=2, marker="o", markersize=4)
    ax.yaxis.set_major_formatter(mticker.PercentFormatter(1.0, decimals=1))
    ax.set_xticks(range(0, len(mt), 2), mt["month"][::2], rotation=45, ha="right")
    ax.grid(axis="y", color=GRID, linewidth=0.8)
    ax.set_title("Contracted suppliers: invoiced price above agreed contract price")
    renew = list(mt["month"]).index("2025-10")
    ax.annotate("Contract renewal\n(Oct 2025) resets prices", xy=(renew, mt["contract_price_drift_pct"].iloc[renew]),
                xytext=(renew + 1.5, mt["contract_price_drift_pct"].max() * 0.85),
                arrowprops=dict(arrowstyle="->", color=MUTED), color=INK, fontsize=9)
    _finish(fig, "price_drift.png")


def hotel_compliance(f: pd.DataFrame):
    cur = f[f["fiscal_year"] == "FY26"]
    onc = (cur.groupby("hotel_name")
              .apply(lambda d: 1 - d.loc[~d["on_contract"], "spend"].sum() / d["spend"].sum(), include_groups=False)
              .sort_values())
    labels = [h.replace("Harbourline ", "") for h in onc.index]
    fig, ax = plt.subplots(figsize=(8, 3.6))
    ax.barh(labels, onc.values, color=BLUE, height=0.6)
    for i, v in enumerate(onc.values):
        ax.text(v + 0.004, i, f"{v:.1%}", va="center", color=INK, fontsize=9, zorder=3,
                bbox=dict(facecolor="white", edgecolor="none", pad=1))
    ax.axvline(0.90, color=INK, linewidth=1, linestyle="--", zorder=2)
    ax.text(0.902, len(onc) - 0.45, "90% target", color=INK, fontsize=9)
    ax.set_xlim(0.5, 1.02)
    ax.xaxis.set_major_formatter(mticker.PercentFormatter(1.0, decimals=0))
    ax.grid(axis="x", color=GRID, linewidth=0.8)
    ax.set_axisbelow(True)
    ax.set_title("On-contract spend by hotel, FY26")
    _finish(fig, "hotel_compliance.png")


def leakage(kpi: pd.DataFrame, f: pd.DataFrame, stock_variance: pd.DataFrame):
    cur = f[f["fiscal_year"] == "FY26"]
    shrink = kpi.set_index("kpi").loc["Bar stock shrinkage (EUR)", "FY26"]
    parts = {
        "Off-contract premium": cur.loc[~cur["on_contract"], "ppv_value"].sum(),
        "Invoice over-billing": cur.loc[cur["invoice_discrepancy"] > 0, "invoice_discrepancy"].sum(),
        "Contract price creep": cur.loc[cur["on_contract"], "ppv_value"].clip(lower=0).sum(),
        "Bar stock shrinkage": shrink,
    }
    s = pd.Series(parts).sort_values()
    fig, ax = plt.subplots(figsize=(8, 3.2))
    ax.barh(s.index, s.values, color=ORANGE, height=0.6)
    for i, v in enumerate(s.values):
        ax.text(v + s.max() * 0.01, i, f"€{v:,.0f}", va="center", color=INK, fontsize=9)
    ax.xaxis.set_major_formatter(EUR_K)
    ax.set_xlim(0, s.max() * 1.18)
    ax.grid(axis="x", color=GRID, linewidth=0.8)
    ax.set_axisbelow(True)
    ax.set_title(f"Where value leaks, FY26 (total €{s.sum():,.0f})")
    _finish(fig, "value_leakage.png")


def main(results=None):
    res = results or analysis.run()
    monthly_spend(res["monthly_trend"])
    price_drift(res["monthly_trend"])
    hotel_compliance(res["fact_lines"])
    leakage(res["kpi_summary"], res["fact_lines"], res["stock_variance"])
    print(f"Charts written to {cfg.IMAGES_DIR}")


if __name__ == "__main__":
    main()
