import pandas as pd
from pathlib import Path

DATA_DIR = Path(__file__).parent.parent / "data"


def extract() -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    products = pd.read_excel(DATA_DIR / "products.xlsx", dtype={"Product ID": str})

    regions = pd.read_excel(DATA_DIR / "regions.xlsx", dtype={"Region ID": str})

    sales = pd.read_excel(
        DATA_DIR / "sales_raw.xlsx",
        parse_dates=["Order Date"],
        dtype={"Product ID": str, "Region ID": str},
    )

    return products, regions, sales