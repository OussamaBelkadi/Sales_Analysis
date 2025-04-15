import pandas as pd

# Existing function: load_and_clean_data
def load_and_clean_data(data_path):
    # Load the data from the provided Excel file
    df = pd.read_excel(data_path)
    
    # Clean the data:
    # Drop rows with missing critical columns
    required_columns = ['OrderDate', 'UnitSellingPrice', 'Profit', 'Channel', 'City', 'Customer Name']
    df = df.dropna(subset=required_columns)
    
    # Ensure that 'UnitSellingPrice' and 'Profit' are numeric, coerce errors to NaN
    df['UnitSellingPrice'] = pd.to_numeric(df['UnitSellingPrice'], errors='coerce')
    df['Profit'] = pd.to_numeric(df['Profit'], errors='coerce')
    
    # Drop any rows with NaN values after coercion
    df = df.dropna()
    
    # Optionally, reset index after cleaning
    df = df.reset_index(drop=True)
    
    return df

# New function: create_date_table
def create_date_table(start_date, end_date):
    """
    Create a date table with additional columns such as Year, Month, Weekday, etc.
    from a range of dates between start_date and end_date.

    Parameters:
    - start_date (str): The start date in 'YYYY-MM-DD' format.
    - end_date (str): The end date in 'YYYY-MM-DD' format.

    Returns:
    - pd.DataFrame: A DataFrame with columns 'Date', 'Year', 'Month', 'Day', 'Weekday', and 'MonthName'.
    """
    # Generate a range of dates
    date_range = pd.date_range(start=start_date, end=end_date, freq='D')
    
    # Create a DataFrame
    date_table = pd.DataFrame(date_range, columns=['Date'])
    
    # Add additional columns for year, month, day, etc.
    date_table['Year'] = date_table['Date'].dt.year
    date_table['Month'] = date_table['Date'].dt.month
    date_table['Day'] = date_table['Date'].dt.day
    date_table['Weekday'] = date_table['Date'].dt.weekday
    date_table['MonthName'] = date_table['Date'].dt.month_name()
    
    return date_table
