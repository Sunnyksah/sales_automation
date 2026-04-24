import pandas as pd


def transform(
    products: pd.DataFrame,
    regions: pd.DataFrame,
    sales: pd.DataFrame,
) -> dict[str, pd.DataFrame]:

    # --- Clean: keep only Completed orders ---
    sales = sales[sales["Status"] == "Completed"].copy()
    sales["Month"] = sales["Order Date"].dt.to_period("M")
    sales["Month_Label"] = sales["Order Date"].dt.strftime("%b %Y")

    # --- Merge dimension tables ---
    df = sales.merge(
        products[["Product ID", "Cost Price", "Margin %"]],
        on="Product ID",
        how="left",
    ).merge(
        regions[["Region ID", "Manager", "Target Revenue"]],
        on="Region ID",
        how="left",
    )

    # ── Monthly Summary ──────────────────────────────────────────────────────
    monthly = (
        df.groupby(["Month", "Month_Label"], sort=True)
        .agg(
            Total_Revenue=("Net Revenue", "sum"),
            Total_Units=("Units", "sum"),
            Total_Orders=("Order ID", "nunique"),
            Avg_Order_Value=("Net Revenue", "mean"),
        )
        .reset_index()
        .sort_values("Month")
    )
    monthly["MoM_Growth"] = monthly["Total_Revenue"].pct_change()
    monthly = monthly.drop(columns=["Month"])

    # ── Region Performance ───────────────────────────────────────────────────
    region_perf = (
        df.groupby(["Region ID", "Region Name", "Manager", "Target Revenue"], sort=False)
        .agg(
            Actual_Revenue=("Net Revenue", "sum"),
            Total_Units=("Units", "sum"),
            Total_Orders=("Order ID", "nunique"),
        )
        .reset_index()
        .sort_values("Actual_Revenue", ascending=False)
        .reset_index(drop=True)
    )
    region_perf["Rank"] = region_perf["Actual_Revenue"].rank(ascending=False).astype(int)
    region_perf["Target_Achievement"] = (
        region_perf["Actual_Revenue"] / region_perf["Target Revenue"]
    )

    # ── Product Breakdown ────────────────────────────────────────────────────
    product_breakdown = (
        df.groupby(["Category", "Product ID", "Product Name", "Unit Price", "Margin %"])
        .agg(
            Units_Sold=("Units", "sum"),
            Revenue=("Net Revenue", "sum"),
            Orders=("Order ID", "nunique"),
        )
        .reset_index()
        .sort_values("Revenue", ascending=False)
    )
    product_breakdown["Revenue_Share"] = (
        product_breakdown["Revenue"] / product_breakdown["Revenue"].sum()
    )

    # ── Category Summary ─────────────────────────────────────────────────────
    category_summary = (
        df.groupby("Category")
        .agg(
            Revenue=("Net Revenue", "sum"),
            Units=("Units", "sum"),
            Orders=("Order ID", "nunique"),
        )
        .reset_index()
        .sort_values("Revenue", ascending=False)
    )
    category_summary["Revenue_Share"] = (
        category_summary["Revenue"] / category_summary["Revenue"].sum()
    )

    # ── KPI Cards ────────────────────────────────────────────────────────────
    kpis = {
        "Total Revenue": df["Net Revenue"].sum(),
        "Total Orders": df["Order ID"].nunique(),
        "Total Units": df["Units"].sum(),
        "Avg Order Value": df["Net Revenue"].mean(),
        "Top Region": region_perf.iloc[0]["Region Name"],
        "Top Category": category_summary.iloc[0]["Category"],
        "Best Month": monthly.loc[monthly["Total_Revenue"].idxmax(), "Month_Label"],
        "Best Month Revenue": monthly["Total_Revenue"].max(),
    }

    return {
        "monthly": monthly,
        "region_perf": region_perf,
        "product_breakdown": product_breakdown,
        "category_summary": category_summary,
        "kpis": kpis,
        "raw_completed": df,
    }