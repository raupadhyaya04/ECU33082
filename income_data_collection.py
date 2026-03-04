import requests
import pandas as pd
from itertools import product

# ---------------------------------------------------------------------------
# Helper: fetch a CSO PxStat table and return a tidy DataFrame
# ---------------------------------------------------------------------------

def fetch_cso_table(table_id: str) -> pd.DataFrame:
    url = (
        f"https://ws.cso.ie/public/api.restful/"
        f"PxStat.Data.Cube_API.ReadDataset/{table_id}/JSON-stat/2.0/en"
    )
    r = requests.get(url, timeout=30)
    r.raise_for_status()
    js = r.json()

    dims = js["dimension"]
    values = js["value"]

    # Build every combination of dimension categories in order
    dim_ids = list(dims.keys())
    cat_lists = [list(dims[d]["category"]["label"].items()) for d in dim_ids]

    rows = []
    for i, combo in enumerate(product(*cat_lists)):
        row = {dim_ids[j]: combo[j][1] for j in range(len(dim_ids))}
        row["VALUE"] = values[i]
        rows.append(row)

    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# EIHC01 — Estimated Inflation by Equivalised Gross Household Income Deciles
# Monthly, 2016-Dec onwards. Broken down by income decile.
# This is the core table linking inflation rates directly to income level.
# ---------------------------------------------------------------------------

print("Fetching EIHC01 — Inflation by Income Decile...")
eihc01 = fetch_cso_table("EIHC01")

# Rename dimension columns to readable names
eihc01 = eihc01.rename(columns={
    "STATISTIC": "Statistic",
    "TLIST(M1)": "Month",
    "C02305V02776": "Income_Decile",
})

# Keep only the YoY % change statistic (most useful for comparison)
eihc01 = eihc01[
    eihc01["Statistic"] == "Percentage change over 12 months for Consumer Price Index"
].copy()

# Parse Month to datetime and filter 2007–2023
eihc01["Month_dt"] = pd.to_datetime(eihc01["Month"], format="%Y %B")
eihc01 = eihc01[
    (eihc01["Month_dt"].dt.year >= 2017) &   # EIHC01 starts Dec 2016
    (eihc01["Month_dt"].dt.year <= 2023)
]

eihc01["VALUE"] = pd.to_numeric(eihc01["VALUE"], errors="coerce")
eihc01 = eihc01.dropna(subset=["VALUE"])
eihc01 = eihc01.sort_values("Month_dt").reset_index(drop=True)

eihc01 = eihc01[["Month", "Month_dt", "Income_Decile", "VALUE"]].rename(
    columns={"VALUE": "YoY_Inflation_Pct"}
)

eihc01.to_csv("Cleaned Data/Inflation_by_Income_Decile_2017_2023.csv", index=False)
print(f"  Saved: {len(eihc01)} rows")
print(eihc01.head())


# ---------------------------------------------------------------------------
# SIA01 — Persons at Risk of Poverty and Consistent Poverty Rates
# Annual, 2005 onwards. Breakdown by sex.
# Used as a proxy for energy poverty transitions over time.
# ---------------------------------------------------------------------------

print("\nFetching SIA01 — Poverty Rates...")
sia01 = fetch_cso_table("SIA01")

sia01 = sia01.rename(columns={
    "C02199V02655": "Sex",
    "TLIST(A1)": "Year",
    "STATISTIC": "Statistic",
})

# Keep only both sexes combined and the consistent poverty measure
sia01 = sia01[sia01["Sex"] == "Both sexes"].copy()
sia01["Year"] = pd.to_numeric(sia01["Year"], errors="coerce")
sia01 = sia01[(sia01["Year"] >= 2007) & (sia01["Year"] <= 2023)]
sia01["VALUE"] = pd.to_numeric(sia01["VALUE"], errors="coerce")
sia01 = sia01.dropna(subset=["VALUE"])
sia01 = sia01.sort_values(["Year", "Statistic"]).reset_index(drop=True)

sia01 = sia01[["Year", "Statistic", "VALUE"]]

sia01.to_csv("Cleaned Data/Poverty_Rates_2007_2023.csv", index=False)
print(f"  Saved: {len(sia01)} rows")
print(sia01.head())


# ---------------------------------------------------------------------------
# SIA52 — Household Disposable Income by Tenure Type
# Annual, 2004 onwards. Owner-occupied vs renter split.
# Useful for splitting energy burden by housing tenure.
# ---------------------------------------------------------------------------

print("\nFetching SIA52 — Disposable Income by Tenure...")
sia52 = fetch_cso_table("SIA52")

sia52 = sia52.rename(columns={
    "STATISTIC": "Statistic",
    "TLIST(A1)": "Year",
    "C01783V03119": "Tenure_Type",
})

# Keep Median Real Household Disposable Income only
sia52 = sia52[
    sia52["Statistic"] == "Median Real Household Disposable Income"
].copy()

sia52["Year"] = pd.to_numeric(sia52["Year"], errors="coerce")
sia52 = sia52[(sia52["Year"] >= 2007) & (sia52["Year"] <= 2023)]
sia52["VALUE"] = pd.to_numeric(sia52["VALUE"], errors="coerce")
sia52 = sia52.dropna(subset=["VALUE"])
sia52 = sia52.sort_values(["Year", "Tenure_Type"]).reset_index(drop=True)

sia52 = sia52[["Year", "Tenure_Type", "VALUE"]].rename(
    columns={"VALUE": "Median_Real_Disposable_Income_EUR"}
)

sia52.to_csv("Cleaned Data/Disposable_Income_by_Tenure_2007_2023.csv", index=False)
print(f"  Saved: {len(sia52)} rows")
print(sia52.head())

print("\nAll done. Files saved to 'Cleaned Data/'.")
