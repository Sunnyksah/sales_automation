import argparse
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from src.extractor import extract
from src.transformer import transform
from src.loader import load


def run(report_date: str = None):
    t0 = time.time()
    print("=" * 55)
    print("  SALES REPORT AUTOMATION PIPELINE")
    print("=" * 55)

    print("\n[1/3] Extracting data from Excel files...")
    products, regions, sales = extract()
    print(f"      products : {len(products)} rows")
    print(f"      regions  : {len(regions)} rows")
    print(f"      sales    : {len(sales)} rows")

    print("\n[2/3] Transforming — cleaning, merging, aggregating...")
    data = transform(products, regions, sales)
    kpis = data["kpis"]
    print(f"      Completed orders : {len(data['raw_completed'])}")
    print(f"      Months covered   : {len(data['monthly'])}")
    print(f"      Total revenue    : ${kpis['Total Revenue']:,.2f}")
    print(f"      Top region       : {kpis['Top Region']}")
    print(f"      Top category     : {kpis['Top Category']}")

    print("\n[3/3] Loading — writing styled Excel report...")
    out_path = load(data, report_date)

    elapsed = time.time() - t0
    print(f"\n{'=' * 55}")
    print(f"  Report saved to : {out_path}")
    print(f"  Completed in    : {elapsed:.1f}s")
    print(f"{'=' * 55}\n")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--month", default=None, help="e.g. 2024_12")
    args = parser.parse_args()
    run(args.month)