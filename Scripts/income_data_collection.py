import requests
import pandas as pd
from itertools import product
from pathlib import Path

ROOT     = Path(__file__).resolve().parent.parent
DATA_DIR = ROOT / "Cleaned Data"
DATA_DIR.mkdir(exist_ok=True)

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
# EIHC01: Estimated Inflation by Equivalised Gross Household Income Deciles
# Monthly, 2016-Dec onwards. Broken down by income decile.
# This is the core table linking inflation rates directly to income level.
# ---------------------------------------------------------------------------

print("Fetching EIHC01: Inflation by Income Decile...")
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

eihc01.to_csv(DATA_DIR / "Inflation_by_Income_Decile_2017_2023.csv", index=False)
print(f"  Saved: {len(eihc01)} rows")
print(eihc01.head())


# ---------------------------------------------------------------------------
# Poverty Rates 2004-2024  (national totals, Ireland)
# ---------------------------------------------------------------------------
# Sources:
#   SIA24  (CSO PxStat, 2004-2019): national totals table — provides the
#          real annually-varying "At Risk of Poverty Rate" (60% median
#          income threshold, after all social transfers) and "Consistent
#          Poverty Rate" as single State-level rows. This is the definitive
#          CSO API source for the pre-2020 period.
#
#   2020-2024: CSO SILC press-release headline figures (national totals).
#          The PxStat API for the 2020+ period (SIA70/SIA81) only stores
#          breakdowns by household type with no national total row; the
#          headline rates below are taken directly from the annual CSO SILC
#          publications and match Eurostat table ilc_li02 for Ireland.
#
#          Year | At Risk of Poverty Rate | Consistent Poverty Rate
#          2020 |         13.9            |          5.5
#          2021 |         11.6            |          4.3
#          2022 |         12.8            |          5.5
#          2023 |         12.1            |          4.2
#          2024 |         12.7            |          4.0
# ---------------------------------------------------------------------------

print("\nFetching poverty rates (SIA24 2004-2019 + CSO SILC 2020-2024)...")

# ── SIA24: 2004-2019 national totals ────────────────────────────────────────
ATRISK_STAT  = (
    "At Risk of Poverty Rate: Equivalised Total Disposable Income: "
    "Including all Social Transfers (60% Median Income Threshold)"
)
CONSIST_STAT = "Consistent Poverty Rate (60% Median Income Threshold)"

sia24_raw = fetch_cso_table("SIA24")
sia24_raw["Year"]  = pd.to_numeric(sia24_raw["TLIST(A1)"],  errors="coerce")
sia24_raw["VALUE"] = pd.to_numeric(sia24_raw["VALUE"], errors="coerce")
sia24_raw = sia24_raw.dropna(subset=["VALUE"])
sia24_raw = sia24_raw[sia24_raw["STATISTIC"].isin([ATRISK_STAT, CONSIST_STAT])].copy()
sia24_raw["Statistic"] = sia24_raw["STATISTIC"].replace({
    ATRISK_STAT:  "At Risk of Poverty Rate",
    CONSIST_STAT: "Consistent Poverty Rate",
})
sia24 = sia24_raw[["Year", "Statistic", "VALUE"]].sort_values(["Year", "Statistic"])

# ── 2020-2024: CSO SILC published headline figures ──────────────────────────
silc_rows = [
    # Year, At Risk of Poverty Rate, Consistent Poverty Rate
    (2020, 13.9, 5.5),
    (2021, 11.6, 4.3),
    (2022, 12.8, 5.5),
    (2023, 12.1, 4.2),
    (2024, 12.7, 4.0),
]
silc_records = []
for year, atrisk, consist in silc_rows:
    silc_records.append({"Year": year, "Statistic": "At Risk of Poverty Rate",  "VALUE": atrisk})
    silc_records.append({"Year": year, "Statistic": "Consistent Poverty Rate",  "VALUE": consist})
silc_2020 = pd.DataFrame(silc_records)

# ── Combine ──────────────────────────────────────────────────────────────────
poverty = (
    pd.concat([sia24, silc_2020], ignore_index=True)
    .drop_duplicates(subset=["Year", "Statistic"], keep="first")
    .sort_values(["Year", "Statistic"])
    .reset_index(drop=True)
)

poverty.to_csv(DATA_DIR / "Poverty_Rates_2004_2024.csv", index=False)
print(f"  Saved: {len(poverty)} rows  ({poverty['Year'].min()}-{poverty['Year'].max()})")
print(poverty.pivot(index="Year", columns="Statistic", values="VALUE").to_string())


# ---------------------------------------------------------------------------
# SIA52: Household Disposable Income by Tenure Type
# Annual, 2004 onwards. Owner-occupied vs renter split.
# Useful for splitting energy burden by housing tenure.
# ---------------------------------------------------------------------------

print("\nFetching SIA52: Disposable Income by Tenure...")
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

sia52.to_csv(DATA_DIR / "Disposable_Income_by_Tenure_2007_2023.csv", index=False)
print(f"  Saved: {len(sia52)} rows")
print(sia52.head())

print("\nAll done. Files saved to 'Cleaned Data/'.")
