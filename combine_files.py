import calendar
from pathlib import Path

import pandas as pd
import plotly.express as px

BASE_DIR = Path(__file__).parent

# Get all excel file paths
files = list(BASE_DIR.rglob("*.xlsx*"))

# Create empty DataFrame
combined = pd.DataFrame()

for file in files:
    df = pd.read_excel(file)
    df["Date"] = df["Date"].dt.date
    df["Day"] = pd.DatetimeIndex(df["Date"]).day
    df["Month"] = pd.DatetimeIndex(df["Date"]).month
    df["Year"] = pd.DatetimeIndex(df["Date"]).year
    df.dropna(inplace=True)
    df["Month_Name"] = df["Month"].apply(lambda x: calendar.month_abbr[int(x)])
    combined = pd.concat([combined, df], ignore_index=True)

# Create Bar Chart
fig = px.bar(combined, x="Month_Name", y="Sales", title="Sales 1Q 2020")

# Save Bar Chart and Export to HTML
fig.write_html(str(BASE_DIR / "Sales_1Q_2020.html"))

# Send Report as attached mail
# Create more charts
# ...

# Export combined files to Excel
combined.to_excel(BASE_DIR / "Sales_1Q2020.xlsx", index=False, sheet_name="1Q 2020 Sales")
