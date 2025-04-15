import pandas as pd

def calculate_kpis(df):
    df["Year"] = df["Order Date"].dt.year
    df["Month"] = df["Order Date"].dt.month

    total_sales = df["Sales"].sum()
    total_profit = df["Profit"].sum()
    total_orders = df["Order ID"].nunique()
    profit_margin = total_profit / total_sales if total_sales else 0

    top_cities = df.groupby("City")["Sales"].sum().sort_values(ascending=False).head(5)
    top_customers = df.groupby("Customer Name")["Sales"].sum().sort_values(ascending=False).head(5)
    last_customers = df.groupby("Customer Name")["Sales"].sum().sort_values().head(5)

    return {
        "total_sales": total_sales,
        "total_profit": total_profit,
        "total_orders": total_orders,
        "profit_margin": profit_margin,
        "top_cities": top_cities,
        "top_customers": top_customers,
        "last_customers": last_customers,
    }
