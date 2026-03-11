"""
Retail Performance & Inventory Analysis
Author: Your Name
Description:
This script performs data cleaning, transformation,
EDA and visualization on retail inventory datasets.
"""

import pandas as pd
import matplotlib.pyplot as plt

# -------------------------------
# 1. Load datasets
# -------------------------------

opening = pd.read_excel("Opening Stock.xlsx")
closing = pd.read_excel("Closing Stock.xlsx")
sales = pd.read_excel("Sales.xlsx")
shipments = pd.read_excel("Stock Transfer Data.xlsx")

print("Datasets loaded successfully")

# -------------------------------
# 2. Data Cleaning
# -------------------------------

# Standardize column names
opening.columns = opening.columns.str.strip()
closing.columns = closing.columns.str.strip()
sales.columns = sales.columns.str.strip()
shipments.columns = shipments.columns.str.strip()

# Handle missing values
opening.fillna(0, inplace=True)
closing.fillna(0, inplace=True)
sales.fillna(0, inplace=True)
shipments.fillna(0, inplace=True)

print("Missing values handled")

# -------------------------------
# 3. Data Transformation
# -------------------------------

# Aggregate by Category
category_open = opening.groupby("Category")["Total Stock On Hand"].sum()
category_close = closing.groupby("Category")["Total Stock On Hand"].sum()
category_sales = sales.groupby("Category")["Qty"].sum()
category_ship = shipments.groupby("Category")["Transaction Qty"].sum()

category_df = pd.concat(
    [category_open, category_ship, category_sales, category_close],
    axis=1
)

category_df.columns = [
    "Opening Stock",
    "Shipments",
    "Sales",
    "Closing Stock"
]

# Sell Through
category_df["Sell Through %"] = (
    category_df["Sales"] /
    (category_df["Opening Stock"] + category_df["Shipments"])
) * 100

# Average Stock
category_df["Average Stock"] = (
    category_df["Opening Stock"] +
    category_df["Closing Stock"]
) / 2

# Stock to Sales Ratio
category_df["Stock Sales Ratio"] = (
    category_df["Average Stock"] /
    category_df["Sales"]
)

print("Data transformation complete")

# -------------------------------
# 4. Exploratory Data Analysis
# -------------------------------

print("\nSummary Statistics\n")
print(category_df.describe())

# Region revenue
region_revenue = sales.groupby("Region")["Invoice Value"].sum()

print("\nRevenue by Region\n")
print(region_revenue)

# Store sell-through
store_sales = sales.groupby("Store")["Qty"].sum()
store_ship = shipments.groupby("Store")["Transaction Qty"].sum()
store_open = opening.groupby("Store")["Total Stock On Hand"].sum()

store_df = pd.concat(
    [store_sales, store_ship, store_open],
    axis=1
)

store_df.columns = ["Sales", "Shipments", "Opening"]

store_df["Sell Through"] = (
    store_df["Sales"] /
    (store_df["Opening"] + store_df["Shipments"])
) * 100

print("\nTop 10 Stores by Sell Through\n")
print(store_df.sort_values("Sell Through", ascending=False).head(10))

# -------------------------------
# 5. Dead Stock Detection
# -------------------------------

dead_stock = opening.merge(
    sales.groupby(["Store", "Product Code"])["Qty"]
    .sum()
    .reset_index(),
    on=["Store", "Product Code"],
    how="left"
)

dead_stock["Qty"] = dead_stock["Qty"].fillna(0)

dead_stock = dead_stock[
    (dead_stock["Total Stock On Hand"] > 0) &
    (dead_stock["Qty"] == 0)
]

print("\nDead Stock Items\n")
print(dead_stock.head())

# -------------------------------
# 6. Visualizations
# -------------------------------

# Sales by Category
plt.figure()
category_df["Sales"].plot(kind="bar")
plt.title("Sales by Category")
plt.xlabel("Category")
plt.ylabel("Units Sold")
plt.tight_layout()
plt.show()

# Revenue by Region
plt.figure()
region_revenue.plot(kind="bar")
plt.title("Revenue by Region")
plt.xlabel("Region")
plt.ylabel("Revenue")
plt.tight_layout()
plt.show()

# Sell Through by Store
plt.figure()
store_df["Sell Through"].sort_values(
    ascending=False
).head(10).plot(kind="bar")
plt.title("Top 10 Stores by Sell Through")
plt.xlabel("Store")
plt.ylabel("Sell Through %")
plt.tight_layout()
plt.show()

print("\nVisualizations generated successfully")