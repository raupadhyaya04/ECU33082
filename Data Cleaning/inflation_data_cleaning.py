import pandas as pd

# Load raw CPI data
df = pd.read_csv("Collected Data/CPM01.20260304T160359.csv")

# Extract the year from the 'Month' column (format: "YYYY MonthName")
df["Year"] = df["Month"].str.split(" ").str[0].astype(int)

# Filter for years 2007–2023
df = df[(df["Year"] >= 2007) & (df["Year"] <= 2023)]

# Filter for the Housing commodity group only
df = df[df["Commodity Group"] == "Housing, water, electricity, gas and other fuels"]

# Keep only the two relevant statistic series:
# - Base Dec 2023=100 index level (for context/deflating)
# - YoY % change (the inflation rate to compare against income)
keep_stats = [
    "Consumer Price Index (Base Dec 2023=100)",
    "Percentage Change over 12 months for Consumer Price Index",
]
df = df[df["Statistic Label"].isin(keep_stats)]

# Drop rows with missing CPI values
df = df.dropna(subset=["VALUE"])
df = df[df["VALUE"].astype(str).str.strip() != ""]

# Convert VALUE to numeric
df["VALUE"] = pd.to_numeric(df["VALUE"])

# Pivot so each statistic is its own column
df = df.pivot_table(
    index=["Month", "Year", "Commodity Group"],
    columns="Statistic Label",
    values="VALUE",
).reset_index()

# Rename columns for clarity
df.columns.name = None
df = df.rename(columns={
    "Consumer Price Index (Base Dec 2023=100)": "CPI_Index",
    "Percentage Change over 12 months for Consumer Price Index": "YoY_Pct_Change",
})

# Sort chronologically (parse "YYYY MonthName" as a proper date)
df["Month_dt"] = pd.to_datetime(df["Month"], format="%Y %B")
df = df.sort_values("Month_dt").drop(columns="Month_dt").reset_index(drop=True)

# Save to Cleaned Data folder
df.to_csv("Cleaned Data/CPI_Housing_2007_2023.csv", index=False)

print(f"Rows saved: {len(df)}")
print(df.head())

