"""
compute_transition_matrices.py
=================================
Compute transition matrices for Energy Poverty states using the panel dataset.

Outputs (CSV):
  Output/Transitions/Transitions_{yr}_{yr+1}_counts.csv   (raw counts)
  Output/Transitions/Transitions_{yr}_{yr+1}_rowpct.csv    (row percentages)
  Output/Transitions/Transitions_POOLED_counts.csv
  Output/Transitions/Transitions_POOLED_rowpct.csv

Assumes the panel is at Country_Code-Year resolution (Country_Code, Year, Energy_Poverty_State).
"""

from pathlib import Path
import pandas as pd
import numpy as np


ROOT = Path(__file__).resolve().parent.parent
PANEL = ROOT / "Cleaned Data" / "Global_Interest_Energy_Poverty_2007_2023.csv"
OUT_DIR = ROOT / "Output" / "Transitions"
OUT_DIR.mkdir(parents=True, exist_ok=True)


def _safe_sort_states(series):
    """Return ordered list of unique states with a stable ordering.
    Ensures consistent row/column ordering across saved matrices."""
    # Preferred ordering (visual): joint states
    preferred = ["lL", "lM", "lH", "mL", "mM", "mH", "hL", "hM", "hH"]
    uniques = list(pd.Index(series.dropna().unique()))
    ordered = [s for s in preferred if s in uniques]
    # Append any other unseen states (defensive)
    ordered += [s for s in uniques if s not in ordered]
    return ordered


def compute_transitions(panel_path: Path, out_dir: Path):
    df = pd.read_csv(panel_path)
    if not {'Year', 'Country_Code', 'Energy_Poverty_State'}.issubset(df.columns):
        raise ValueError("Panel must contain Year, Country_Code, Energy_Poverty_State columns")

    def get_joint_state(row):
        # Map Interest Rate state to lower case, Energy Poverty state to upper case
        ir_map = {'Low': 'l', 'Medium': 'm', 'High': 'h'}
        ep_map = {'Low': 'L', 'Medium': 'M', 'High': 'H'}
        # Handle NA or weird cases
        ir = ir_map.get(row['Interest_Rate_State'], '?')
        ep = ep_map.get(row['Energy_Poverty_State'], '?')
        return ir + ep

    df['Joint_State'] = df.apply(get_joint_state, axis=1)

    obs = df.copy()
    years = sorted(obs['Year'].unique())
    if len(years) < 2:
        print("Not enough observed years to compute transitions.")
        return

    pooled_counts = None
    pair_count = 0

    for y in years:
        if (y + 1) not in years:
            continue
        t0 = obs[obs['Year'] == y][['Country_Code', 'Joint_State']].copy()
        t1 = obs[obs['Year'] == (y + 1)][['Country_Code', 'Joint_State']].copy()

        # Ensure ID is unique per year
        t0 = t0.drop_duplicates(subset=['Country_Code'], keep='last')
        t1 = t1.drop_duplicates(subset=['Country_Code'], keep='last')

        merged = pd.merge(t0, t1, on='Country_Code', how='inner', suffixes=('_t', '_t1'))
        # merged now has Country_Code, Joint_State_t, Joint_State_t1
        if merged.empty:
            print(f"No matched countries between {y} and {y+1}; skipping")
            continue

        pair_count += 1
        states = sorted(list(set(merged['Joint_State_t']).union(set(merged['Joint_State_t1']))))
        # use safe ordering for presentation
        ordered = _safe_sort_states(pd.Series(states))

        ct = pd.crosstab(merged['Joint_State_t1'], merged['Joint_State_t']).reindex(index=ordered, columns=ordered, fill_value=0)

        # Save counts and column percentages (left-stochastic standard)
        counts_fp = out_dir / f"Transitions_{y}_{y+1}_counts.csv"
        colpct_fp  = out_dir / f"Transitions_{y}_{y+1}_colpct.csv"
        ct.to_csv(counts_fp)
        (ct.div(ct.sum(axis=0).replace(0, np.nan), axis=1).fillna(0) * 100).round(1).to_csv(colpct_fp)

        print(f"Wrote: {counts_fp}  (n_pairs={len(merged)})")
        print(f"Wrote: {colpct_fp}")

        if pooled_counts is None:
            pooled_counts = ct.copy()
        else:
            # align and add
            pooled_counts = pooled_counts.reindex(index=ct.index.union(pooled_counts.index), columns=ct.columns.union(pooled_counts.columns), fill_value=0)
            pooled_counts = pooled_counts.add(ct, fill_value=0)

    # Save pooled matrices
    if pooled_counts is None or pooled_counts.values.sum() == 0:
        print("No transitions were computed (no overlapping observed years).")
        return

    pooled_counts = pooled_counts.astype(int)
    pooled_counts_fp = out_dir / "Transitions_POOLED_counts.csv"
    pooled_colpct_fp = out_dir / "Transitions_POOLED_colpct.csv"
    pooled_counts.to_csv(pooled_counts_fp)
    (pooled_counts.div(pooled_counts.sum(axis=0).replace(0, np.nan), axis=1).fillna(0) * 100).round(1).to_csv(pooled_colpct_fp)

    print(f"\nPooled transitions across {pair_count} year-pairs written:")
    print(f"  {pooled_counts_fp}")
    print(f"  {pooled_colpct_fp}")


if __name__ == '__main__':
    print("Computing transition matrices from panel...")
    compute_transitions(PANEL, OUT_DIR)
    print("Done.")
