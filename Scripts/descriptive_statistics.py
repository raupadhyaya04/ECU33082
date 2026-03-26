"""
descriptive_statistics.py
=============================================================
Project : Global Energy Poverty & Interest Rate Transitions
Stats   : Central Tendency · Dispersion · Position · Frequency
"""

import pandas as pd
import numpy as np
from pathlib import Path
from scipy import stats as scipy_stats

try:
    import openpyxl
except ImportError:
    import subprocess
    import sys
    subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl"])
    import openpyxl

from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

ROOT     = Path(__file__).resolve().parent.parent
DATA_DIR = ROOT / "Cleaned Data"
OUT_DIR  = ROOT / "Output"
OUT_DIR.mkdir(exist_ok=True)


def _full_stats(s: pd.Series, label: str) -> dict:
    s = s.dropna()
    n = len(s)
    if n == 0:
        return {"Variable": label, "N": 0}

    q = [0.05, 0.10, 0.25, 0.75, 0.90, 0.95]
    p5, p10, p25, p75, p90, p95 = s.quantile(q).values

    out = {
        "Variable": label,
        "N": n,
        "Mean": float(s.mean()),
        "Trimmed Mean (5%)": float(scipy_stats.trim_mean(s.values, 0.05)),
        "Median": float(s.median()),
        "Std Dev": float(s.std()),
        "Range": float(s.max() - s.min()),
        "IQR": float(p75 - p25),
        "Min": float(s.min()),
        "Max": float(s.max()),
        "P5": float(p5),
        "P25": float(p25),
        "P75": float(p75),
        "P95": float(p95),
        "Skewness": float(s.skew()),
        "Kurtosis": float(s.kurt()),
    }
    return out


def run_descriptive_stats():
    infile = DATA_DIR / "Global_Interest_Energy_Poverty_2007_2023.csv"
    if not infile.exists():
        print(f"Data not found: {infile}")
        return

    df = pd.read_csv(infile)
    stats_list = []
    
    # Process numeric columns
    numeric_cols = [
        "Interest_Rate", 
        "Electricity_Access", 
        "Clean_Cooking_Access", 
        "Energy_Poverty_Elec", 
        "Energy_Poverty_Cook", 
        "Composite_Energy_Poverty"
    ]
    
    for col in numeric_cols:
        if col in df.columns:
            stats_list.append(_full_stats(df[col], col))
            
    stats_df = pd.DataFrame(stats_list)

    # State frequencies
    freq_data = []
    
    # Interest Rate States
    if 'Interest_Rate_State' in df.columns:
        ir_counts = df['Interest_Rate_State'].value_counts(normalize=True).mul(100).round(2)
        for state, pct in ir_counts.items():
            freq_data.append({"Variable": f"State: Interest Rate [{state}]", "Percentage (%)": pct})
            
    # Energy Poverty States
    if 'Energy_Poverty_State' in df.columns:
        ep_counts = df['Energy_Poverty_State'].value_counts(normalize=True).mul(100).round(2)
        for state, pct in ep_counts.items():
            freq_data.append({"Variable": f"State: Energy Poverty [{state}]", "Percentage (%)": pct})

    freq_df = pd.DataFrame(freq_data)

    # Save to Excel
    out_path = OUT_DIR / "Global_Descriptive_Statistics.xlsx"
    with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
        if not stats_df.empty:
            stats_df.to_excel(writer, sheet_name="Numeric Variables", index=False)
        if not freq_df.empty:
            freq_df.to_excel(writer, sheet_name="State Frequencies", index=False)
            
    print(f"✅ Descriptive statistics exported to: {out_path}")


if __name__ == "__main__":
    print("Computing Descriptive Statistics...")
    run_descriptive_stats()
