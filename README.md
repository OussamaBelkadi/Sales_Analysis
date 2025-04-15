# 📊 Sales Data Analysis Project

This project analyzes sales data from an Excel file (`sales_data.xlsx`) and generates insightful visualizations, performance indicators, and comparisons across time, products, cities, and customers. The output includes charts saved as images, useful for business reporting or strategic planning.

---

## 📁 Table of Contents

1. [Overview](#overview)
2. [Functional Objectives](#functional-objectives)
3. [Technical Implementation](#technical-implementation)
4. [Installation & Usage](#installation--usage)
5. [Generated Outputs](#generated-outputs)
6. [Dependencies](#dependencies)
7. [Project Structure](#project-structure)
8. [Author](#author)

---

## 🔎 Overview

The goal of this project is to analyze historical sales data and produce relevant metrics and visualizations, including:

- Sales performance by product and region.
- Profitability across sales channels.
- Customer segmentation based on purchasing volume.
- Key Performance Indicators (KPIs) like total sales, total profit, and profit margin.

---

## ✅ Functional Objectives

This project performs the following operations:

- ✅ Clean and preprocess Excel sales data.
- ✅ Generate calculated columns: Sales, Cost, Profit, and Profit Margin.
- ✅ Group and aggregate data to extract key trends.
- ✅ Create the following visualizations:

  | Visualisation                          | Description                                 |
  |----------------------------------------|---------------------------------------------|
  | Bar chart: Sales by product/year       | Compare sales across years for each product |
  | Pie chart: Top 5 product sales         | Share of total sales by top products        |
  | Line chart: Monthly sales trends       | Evolution of sales across months & years    |
  | Bar chart: Sales by region             | Top 5 cities with highest sales             |
  | Bar chart: Profit by channel/year      | Profitability by sales channel              |
  | Bar chart: Top 5 and bottom 5 clients  | Best and worst performing customers         |
  | KPI cards                              | Summary: total sales, profit, margin, qty   |

---

## ⚙️ Technical Implementation

### 💾 Data Preprocessing

- Load Excel file: `sales_data.xlsx`
- Clean column names and parse dates
- Create new columns:
  - `Sales = Order Quantity × Unit Selling Price`
  - `Cost = Order Quantity × Unit Cost`
  - `Profit = Sales - Cost`
  - `Profit Margin = Profit / Sales`

### 🧮 Aggregation

The data is grouped using `groupby()` to calculate:

- Total sales by product, city, month, year, customer
- Profits by channel and year
- Top & bottom 5 customers by sales

### 📈 Visualizations

Generated using:
- `matplotlib` for bar and line charts
- `seaborn` for color palettes
- `plt.savefig()` to export charts as `.png` files in `output/`

---

## 🖥️ Installation & Usage

```bash
# Clone the project
git clone https://github.com/your-username/sales-analysis-project.git
cd sales-analysis-project

# Create a virtual environment
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Place your sales_data.xlsx file in the root folder

# Run the analysis
python main.py
