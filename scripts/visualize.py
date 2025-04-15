import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt

# Existing function: plot_sales_by_month
def plot_sales_by_month(df):
    # Ensure 'UnitSellingPrice' is numeric, force errors to NaN and drop them
    df['UnitSellingPrice'] = pd.to_numeric(df['UnitSellingPrice'], errors='coerce')
    
    # Ensure 'OrderDate' is in datetime format, and extract the month
    df['OrderDate'] = pd.to_datetime(df['OrderDate'], errors='coerce')
    df['Month'] = df['OrderDate'].dt.month_name()
    
    # Calculate monthly sales (Total Sales = UnitSellingPrice * OrderQuantity)
    monthly_sales = df.groupby('Month').agg({'UnitSellingPrice': 'sum'}).reset_index()
    
    # Plot the data
    plt.figure(figsize=(10, 6))
    sns.lineplot(data=monthly_sales, x='Month', y='UnitSellingPrice')
    plt.title('Sales by Month')
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.show()

# New function: plot_sales_by_product
def plot_sales_by_product(df):
    # Calculate sales by product (Total Sales = UnitSellingPrice * OrderQuantity)
    df['TotalSales'] = df['UnitSellingPrice'] * df['OrderQuantity']
    product_sales = df.groupby('Product Description Index').agg({'TotalSales': 'sum'}).reset_index()
    
    # Sort the data by TotalSales in descending order and select top 10 products
    top_products = product_sales.sort_values(by='TotalSales', ascending=False).head(10)
    
    # Plot the data
    plt.figure(figsize=(12, 6))
    sns.barplot(data=top_products, x='Product Description Index', y='TotalSales')
    plt.title('Top 10 Products by Sales')
    plt.xticks(rotation=90)
    plt.tight_layout()
    plt.show()

# Function to plot the top 5 cities by total sales
def plot_top_5_cities(df):
    # Group the data by city and calculate total sales for each city
    city_sales = df.groupby('City')['UnitSellingPrice'].sum().reset_index()
    
    # Sort the cities by total sales in descending order and get the top 5
    top_5_cities = city_sales.sort_values('UnitSellingPrice', ascending=False).head(5)
    
    # Plot the top 5 cities
    plt.figure(figsize=(10, 6))
    sns.barplot(data=top_5_cities, x='City', y='UnitSellingPrice', palette='Blues_d')
    plt.title('Top 5 Cities by Total Sales')
    plt.xlabel('City')
    plt.ylabel('Total Sales')
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    plt.show()
