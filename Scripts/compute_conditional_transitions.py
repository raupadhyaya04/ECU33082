import pandas as pd
import numpy as np
from pathlib import Path

# Paths
ROOT = Path(__file__).resolve().parent.parent
PANEL_FILE = ROOT / "Cleaned Data" / "Global_Interest_Energy_Poverty_2007_2023.csv"
OUT_DIR = ROOT / "Output" / "Transitions" / "Conditional"
OUT_DIR.mkdir(parents=True, exist_ok=True)

def _safe_sort_states(series):
    preferred = ["lL", "lM", "lH", "mL", "mM", "mH", "hL", "hM", "hH"]
    uniques = list(pd.Index(series.dropna().unique()))
    ordered = [s for s in preferred if s in uniques]
    ordered += [s for s in uniques if s not in ordered]
    return ordered

def compute_conditional_transitions():
    if not PANEL_FILE.exists():
        print(f"Panel file not found: {PANEL_FILE}")
        return
        
    df = pd.read_csv(PANEL_FILE)
    
    def get_joint_state(row):
        ir_map = {'Low': 'l', 'Medium': 'm', 'High': 'h'}
        ep_map = {'Low': 'L', 'Medium': 'M', 'High': 'H'}
        ir = ir_map.get(row['Interest_Rate_State'], '?')
        ep = ep_map.get(row['Energy_Poverty_State'], '?')
        return ir + ep

    df['Joint_State'] = df.apply(get_joint_state, axis=1)
    df = df.sort_values(["Country_Code", "Year"]).reset_index(drop=True)
    
    states = _safe_sort_states(df['Joint_State'])
    
    # Standard Split
    pre_counts = pd.DataFrame(0, index=states, columns=states)
    post_counts = pd.DataFrame(0, index=states, columns=states)
    
    # Crisis vs Baseline Split
    crisis_counts = pd.DataFrame(0, index=states, columns=states)
    baseline_counts = pd.DataFrame(0, index=states, columns=states)
    
    # Pre/Post Split (EXCLUDING CRISIS YEARS for robustness check)
    pre_stable_counts = pd.DataFrame(0, index=states, columns=states)
    post_stable_counts = pd.DataFrame(0, index=states, columns=states)

    crisis_years = {2008, 2009, 2020, 2021, 2022}
    
    # Iterate countries
    for code, grp in df.groupby("Country_Code"):
        grp = grp.sort_values("Year")
        
        for i in range(len(grp) - 1):
            y1 = grp.iloc[i]["Year"]
            y2 = grp.iloc[i+1]["Year"]
            
            if y2 == y1 + 1:
                st1 = grp.iloc[i]["Joint_State"]
                st2 = grp.iloc[i+1]["Joint_State"]
                
                # 1. Standard Pre vs Post 2015
                if y1 < 2015:
                    pre_counts.loc[st1, st2] += 1
                else:
                    post_counts.loc[st1, st2] += 1
                
                # 2. Crisis vs Baseline
                is_crisis = y1 in crisis_years or y2 in crisis_years
                if is_crisis:
                    crisis_counts.loc[st1, st2] += 1
                else:
                    baseline_counts.loc[st1, st2] += 1
                    
                    # 3. Stable Pre vs Post 2015 (only non-crisis years)
                    if y1 < 2015:
                        pre_stable_counts.loc[st1, st2] += 1
                    else:
                        post_stable_counts.loc[st1, st2] += 1
                    
    def save_matrix(counts_df, regime_name):
        counts_df_t = counts_df.T
        counts_fp = OUT_DIR / f"Transitions_{regime_name}_runs_counts.csv"
        colpct_fp = OUT_DIR / f"Transitions_{regime_name}_runs_colpct.csv"
        
        counts_df_t.to_csv(counts_fp)
        
        col_sums = counts_df_t.sum(axis=0).replace(0, np.nan)
        colpct_df = counts_df_t.div(col_sums, axis=1).fillna(0) * 100
        colpct_df.round(1).to_csv(colpct_fp)
        
    save_matrix(pre_counts, "Pre_2015")
    save_matrix(post_counts, "Post_2015")
    save_matrix(crisis_counts, "Crisis")
    save_matrix(baseline_counts, "Baseline")
    save_matrix(pre_stable_counts, "Pre_2015_Stable")
    save_matrix(post_stable_counts, "Post_2015_Stable")
    print("✅ All conditional matrices generated seamlessly.")

if __name__ == "__main__":
    compute_conditional_transitions()
