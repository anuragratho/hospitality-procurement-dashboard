"""
One-command pipeline:  python run_pipeline.py

  1. generate the synthetic dataset           -> data/raw/*.csv
  2. run the procurement analysis             -> data/processed/*.csv
  3. build the Excel dashboard                -> output/Procurement_Intelligence_Dashboard.xlsx
  4. render README charts                     -> docs/images/*.png

Options:
  --seed N        change the random seed (default 42 in src/config.py)
  --skip-charts   skip the PNG charts (faster)
"""

from __future__ import annotations

import argparse
import time

from src import analysis, build_dashboard, charts, generate_data
from src import config as cfg


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--seed", type=int, default=cfg.RANDOM_SEED)
    parser.add_argument("--skip-charts", action="store_true")
    args = parser.parse_args()

    t0 = time.time()
    print(f"[1/4] Generating synthetic data (seed {args.seed}) ...")
    tables = generate_data.generate_all(seed=args.seed)
    generate_data.save_all(tables)
    for name, df in tables.items():
        print(f"      {name:<16} {len(df):>8,} rows")

    print("[2/4] Running analysis ...")
    tables = analysis.load_raw()
    results = analysis.run(tables)
    k = results["kpi_summary"].set_index("kpi")
    print(f"      Total spend         €{k.loc['Total spend (EUR)', 'Total']:,.0f}")
    print(f"      On-contract (FY26)  {k.loc['On-contract spend %', 'FY26']:.1%}")
    print(f"      Value leakage (FY26) €{k.loc['Value leakage identified (EUR)', 'FY26']:,.0f}")

    print("[3/4] Building Excel dashboard ...")
    path = build_dashboard.build(results, tables)
    print(f"      {path.relative_to(cfg.PROJECT_ROOT)}")

    if not args.skip_charts:
        print("[4/4] Rendering charts ...")
        charts.main(results)
    else:
        print("[4/4] Charts skipped")

    print(f"Done in {time.time() - t0:.0f}s")


if __name__ == "__main__":
    main()
