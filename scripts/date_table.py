import pandas as pd

def create_date_table(start_date, end_date):
    date_range = pd.date_range(start=start_date, end=end_date, freq="D")
    df = pd.DataFrame({"Date": date_range})
    df["Year"] = df["Date"].dt.year
    df["Month"] = df["Date"].dt.month
    df["Month Name"] = df["Date"].dt.strftime("%B")
    df["Quarter"] = df["Date"].dt.quarter
    df["Day of Week"] = df["Date"].dt.day_name()
    return df
