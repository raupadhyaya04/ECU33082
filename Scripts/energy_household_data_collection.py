import requests
import pandas as pd
from itertools import product
from pathlib import Path

ROOT     = Path(__file__).resolve().parent.parent
DATA_DIR = ROOT / "Cleaned Data"
DATA_DIR.mkdir(exist_ok=True)

YEAR_MIN, YEAR_MAX = 2007, 2023

# ---------------------------------------------------------------------------
# Helper: fetch a CSO PxStat table : tidy DataFrame
# ---------------------------------------------------------------------------

def fetch_cso_table(table_id: str) -> pd.DataFrame:
    url = (
        "https://ws.cso.ie/public/api.restful/"
        f"PxStat.Data.Cube_API.ReadDataset/{table_id}/JSON-stat/2.0/en"
    )
    r = requests.get(url, timeout=60)
    r.raise_for_status()
    js = r.json()
    dims = js["dimension"]
    values = js["value"]
    dim_ids = list(dims.keys())
    cat_lists = [list(dims[d]["category"]["label"].items()) for d in dim_ids]
    rows = []
    for i, combo in enumerate(product(*cat_lists)):
        row = {dim_ids[j]: combo[j][1] for j in range(len(dim_ids))}
        row["VALUE"] = values[i]
        rows.append(row)
    return pd.DataFrame(rows)


def save(df: pd.DataFrame, filename: str):
    path = DATA_DIR / filename
    df.to_csv(path, index=False)
    print(f"  ✓ Saved {len(df)} rows : {path}")


# ============================================================
# ENERGY VARIABLES
# ============================================================

# ------------------------------------------------------------
# 1. Residential Gas Prices  (TMEGB01)
#    Semi-annual, 2015 onwards, by consumption band & tax treatment
# ------------------------------------------------------------
print("\n[1/6] Residential Gas Prices (TMEGB01)...")
gas = fetch_cso_table("TMEGB01")
gas = gas.rename(columns={
    "STATISTIC":      "Statistic",
    "TLIST(A1)":      "Year",
    "C04064V04826":   "Half",
    "C04065V04827":   "Product",
    "C04063V04825":   "Tax_Treatment",
    "C04061V04823":   "Consumption_Band",
})
gas["Year"] = pd.to_numeric(gas["Year"], errors="coerce")
gas = gas[(gas["Year"] >= YEAR_MIN) & (gas["Year"] <= YEAR_MAX)]
# Keep all-in price (including all taxes) for the typical band (D2: 20-199 GJ)
gas = gas[
    gas["Tax_Treatment"].str.contains("including all taxes", case=False) &
    gas["Consumption_Band"].str.contains("band D2", case=False)
]
gas["VALUE"] = pd.to_numeric(gas["VALUE"], errors="coerce")
gas = gas.dropna(subset=["VALUE"])
gas = gas[["Year", "Half", "Consumption_Band", "Tax_Treatment", "VALUE"]].rename(
    columns={"VALUE": "Price_EUR_per_GJ"}
)
gas = gas.sort_values(["Year", "Half"]).reset_index(drop=True)
save(gas, "Gas_Prices_Residential_2015_2023.csv")
print("  Note: TMEGB01 starts 2015: no CSO residential gas price data before that.")

# ------------------------------------------------------------
# 2. Residential Electricity Prices  (TMEGB03)
#    Semi-annual, 2015 onwards, by consumption band & tax treatment
# ------------------------------------------------------------
print("\n[2/6] Residential Electricity Prices (TMEGB03)...")
elec = fetch_cso_table("TMEGB03")
elec = elec.rename(columns={
    "STATISTIC":      "Statistic",
    "TLIST(A1)":      "Year",
    "C04064V04826":   "Half",
    "C04065V04827":   "Product",
    "C04063V04825":   "Tax_Treatment",
    "C04062V04824":   "Consumption_Band",
})
elec["Year"] = pd.to_numeric(elec["Year"], errors="coerce")
elec = elec[(elec["Year"] >= YEAR_MIN) & (elec["Year"] <= YEAR_MAX)]
# Keep all-in price for the mid-range band (DC: 2500–5000 kWh)
elec = elec[
    elec["Tax_Treatment"].str.contains("including all taxes", case=False) &
    elec["Consumption_Band"].str.contains("band DC", case=False)
]
elec["VALUE"] = pd.to_numeric(elec["VALUE"], errors="coerce")
elec = elec.dropna(subset=["VALUE"])
elec = elec[["Year", "Half", "Consumption_Band", "Tax_Treatment", "VALUE"]].rename(
    columns={"VALUE": "Price_EUR_per_kWh"}
)
elec = elec.sort_values(["Year", "Half"]).reset_index(drop=True)
save(elec, "Electricity_Prices_Residential_2015_2023.csv")
print("  Note: TMEGB03 starts 2015: no CSO residential electricity price data before that.")

# ------------------------------------------------------------
# 3. CPI for Energy Products  (EIIEEA04)
#    Annual index covering electricity, gas, petrol, diesel, solid fuels
#    Goes back to 2000: fills the pre-2015 gap above
# ------------------------------------------------------------
print("\n[3/6] CPI for Energy Products (EIIEEA04)...")
energy_cpi = fetch_cso_table("EIIEEA04")
energy_cpi = energy_cpi.rename(columns={
    "STATISTIC":      "Statistic",
    "TLIST(A1)":      "Year",
    "C03732V04480":   "Energy_Product",
})
energy_cpi["Year"] = pd.to_numeric(energy_cpi["Year"], errors="coerce")
energy_cpi = energy_cpi[
    (energy_cpi["Year"] >= YEAR_MIN) & (energy_cpi["Year"] <= YEAR_MAX)
]
energy_cpi["VALUE"] = pd.to_numeric(energy_cpi["VALUE"], errors="coerce")
energy_cpi = energy_cpi.dropna(subset=["VALUE"])
energy_cpi = energy_cpi[["Year", "Energy_Product", "VALUE"]].rename(
    columns={"VALUE": "CPI_Index"}
)
energy_cpi = energy_cpi.sort_values(["Year", "Energy_Product"]).reset_index(drop=True)
save(energy_cpi, "CPI_Energy_Products_2007_2023.csv")
print(f"  Products covered: {sorted(energy_cpi['Energy_Product'].unique())}")

# ============================================================
# HOUSEHOLD VARIABLES
# ============================================================

# ------------------------------------------------------------
# 4a. Quarterly ILO Participation & Unemployment Rates (QLF02)
#     Covers 1998Q1 - 2025Q4  (replaces QNQ20 which ended 2017Q2)
# ------------------------------------------------------------
print("\n[4/6] Quarterly Employment / Unemployment Rates (QLF02 + LRM03)...")
qnq = fetch_cso_table("QLF02")
qnq = qnq.rename(columns={
    "STATISTIC":      "Statistic",
    "TLIST(Q1)":      "Quarter",
    "C02199V02655":   "Sex",
})
# QLF02 has a typo in one stat name ("LO" instead of "ILO") — fix it
qnq["Statistic"] = qnq["Statistic"].replace({
    "LO Unemployment Rates (15 - 74 years)":
        "ILO Unemployment Rates (15 - 74 years)",
    "LO Unemployment Rates (15 - 74 years) (Seasonally Adjusted)":
        "ILO Unemployment Rates (15 - 74 years) (Seasonally Adjusted)",
})
qnq = qnq[qnq["Sex"] == "Both sexes"].copy()
qnq["Year"] = qnq["Quarter"].str[:4].astype(int)
qnq = qnq[(qnq["Year"] >= YEAR_MIN) & (qnq["Year"] <= YEAR_MAX)]
qnq["VALUE"] = pd.to_numeric(qnq["VALUE"], errors="coerce")
qnq = qnq.dropna(subset=["VALUE"])
qnq = qnq[["Quarter", "Statistic", "VALUE"]].rename(columns={"VALUE": "Rate_Pct"})
qnq = qnq.sort_values(["Quarter", "Statistic"]).reset_index(drop=True)

# ------------------------------------------------------------
# 4b. Seasonally Adjusted Monthly Unemployment Rate (LRM03)
#     Extends coverage past 2017 to present
# ------------------------------------------------------------
lrm = fetch_cso_table("LRM03")
lrm = lrm.rename(columns={
    "STATISTIC":    "Statistic",
    "TLIST(M1)":    "Month",
})
lrm["Year"] = lrm["Month"].str[:4].astype(int)
lrm = lrm[(lrm["Year"] >= YEAR_MIN) & (lrm["Year"] <= YEAR_MAX)]
lrm["VALUE"] = pd.to_numeric(lrm["VALUE"], errors="coerce")
lrm = lrm.dropna(subset=["VALUE"])
lrm = lrm[["Month", "Statistic", "VALUE"]].rename(
    columns={"VALUE": "Unemployment_Rate_Pct"}
)
lrm = lrm.sort_values("Month").reset_index(drop=True)

qnq.to_csv(DATA_DIR / "Employment_Rates_Quarterly_ILO_2007_2023.csv", index=False)
lrm.to_csv(DATA_DIR / "Unemployment_Rate_Monthly_2007_2023.csv", index=False)
print(f"  ✓ Saved {len(qnq)} rows : Employment_Rates_Quarterly_ILO_2007_2023.csv")
print(f"  ✓ Saved {len(lrm)} rows : Unemployment_Rate_Monthly_2007_2023.csv")

# ------------------------------------------------------------
# 5. Income: Median Household Disposable Income by Age (SIA13)
#    Annual, 2004 onwards: already explored, now clean fully
# ------------------------------------------------------------
print("\n[5/6] Income: Median Disposable Income by Age Group (SIA13)...")
sia13 = fetch_cso_table("SIA13")
sia13 = sia13.rename(columns={
    "STATISTIC":      "Statistic",
    "TLIST(A1)":      "Year",
    "C02076V02508":   "Age_Group",
})
sia13["Year"] = pd.to_numeric(sia13["Year"], errors="coerce")
sia13 = sia13[(sia13["Year"] >= YEAR_MIN) & (sia13["Year"] <= YEAR_MAX)]
# Keep Median Real Household Disposable Income (inflation-adjusted)
sia13 = sia13[
    sia13["Statistic"] == "Median Real Household Disposable Income"
].copy()
sia13["VALUE"] = pd.to_numeric(sia13["VALUE"], errors="coerce")
sia13 = sia13.dropna(subset=["VALUE"])
sia13 = sia13[["Year", "Age_Group", "VALUE"]].rename(
    columns={"VALUE": "Median_Real_Disposable_Income_EUR"}
)
sia13 = sia13.sort_values(["Year", "Age_Group"]).reset_index(drop=True)
save(sia13, "Income_Median_by_AgeGroup_2007_2023.csv")

# Income volatility: compute year-on-year % change per age group
sia13_wide = sia13.pivot(index="Year", columns="Age_Group",
                         values="Median_Real_Disposable_Income_EUR")
sia13_vol = sia13_wide.pct_change() * 100
sia13_vol.columns = [f"YoY_Pct_{c}" for c in sia13_vol.columns]
sia13_vol = sia13_vol.dropna(how="all").reset_index()
sia13_vol.to_csv(DATA_DIR / "Income_Volatility_YoY_by_AgeGroup_2007_2023.csv",
                 index=False)
print(f"  ✓ Also saved income volatility (YoY %) : "
      f"Income_Volatility_YoY_by_AgeGroup_2007_2023.csv")

# ------------------------------------------------------------
# 6. Inflation: Monthly CPI All Items + Housing component (CPM13)
#    Already have housing CPI from inflation_data_cleaning.py.
#    Here we add the broader All Items CPI for context.
# ------------------------------------------------------------
print("\n[6/6] Monthly CPI: All Items & Energy sub-index (CPM13)...")
cpm13 = fetch_cso_table("CPM13")
cpm13 = cpm13.rename(columns={
    "C02439V03429":   "Commodity_Group",
    "TLIST(M1)":      "Month_Code",
    "STATISTIC":      "Statistic",
})

# Keep YoY % change for All Items and Housing & Energy groups
keep_groups = [
    "All items",
    "Housing, water, electricity, gas and other fuels",
    "Energy",
]
keep_stat = "Percentage Change over 12 months for Consumer Price Index"
cpm13 = cpm13[
    cpm13["Commodity_Group"].isin(keep_groups) &
    (cpm13["Statistic"] == keep_stat)
].copy()

# Parse month code "2007M01" : datetime
cpm13["Month_dt"] = pd.to_datetime(cpm13["Month_Code"], format="%YM%m")
cpm13["Year"] = cpm13["Month_dt"].dt.year
cpm13 = cpm13[(cpm13["Year"] >= YEAR_MIN) & (cpm13["Year"] <= YEAR_MAX)]
cpm13["VALUE"] = pd.to_numeric(cpm13["VALUE"], errors="coerce")
cpm13 = cpm13.dropna(subset=["VALUE"])
cpm13 = cpm13[["Month_dt", "Commodity_Group", "VALUE"]].rename(
    columns={"VALUE": "YoY_Inflation_Pct"}
)
cpm13 = cpm13.sort_values(["Month_dt", "Commodity_Group"]).reset_index(drop=True)
save(cpm13, "CPI_AllItems_and_Housing_Monthly_2007_2023.csv")
print(f"  Groups: {cpm13['Commodity_Group'].unique()}")

print("\n✅  All done. Files saved to 'Cleaned Data/'.")
