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
    
    # We will split conditional transitions Pre-2015 (Pre-SDGs) vs Post-2015 (Post-SDGs)
    pre_counts = pd.DataFrame(0, index=states, columns=states)
    post_counts = pd.DataFrame(0, index=states, columns=states)
    
    pre_pairs = 0
    post_pairs = 0
    
    # Iterate countries
    for code, grp in df.groupby("Country_Code"):
        grp = grp.sort_values("Year")
        
        for i in range(len(grp) - 1):
            y1 = grp.iloc[i]["Year"]
            y2 = grp.iloc[i+1]["Year"]
            
            # Only strictly consecutive years
            if y2 == y1 + 1:
                st1 = grp.iloc[i]["Joint_State"]
                st2 = grp.iloc[i+1]["Joint_State"]
                
                # Check regime of year 1 (the starting state year indicates the regime)
                if y1 < 2015:
                    pre_counts.loc[st1, st2] += 1
                    pre_pairs += 1
                else:
                    post_counts.loc[st1, st2] += 1
                    post_pairs += 1
                    
    # Function to save
    def save_matrix(counts_df, regime_name, pairs_count):
        # We need Left-Stochastic (X-axis = t, Y-axis = t+1)
        # However, the loop above built it as counts_df.loc[st1, st2], meaning row=t, col=t+1.
        # Transpose to strictly left-stochastic before saving
        counts_df_t = counts_df.T
        
        counts_fp = OUT_DIR / f"Transitions_{regime_name}_runs_counts.csv"
        colpct_fp = OUT_DIR / f"Transitions_{regime_name}_runs_colpct.csv"
        
        counts_df_t.to_csv(counts_fp)
        
        # Column %
        col_sums = counts_df_t.sum(axis=0).replace(0, np.nan)
        colpct_df = counts_df_t.div(col_sums, axis=1).fillna(0) * 100
        colpct_df.round(1).to_csv(colpct_fp)
        
        print(f"\n[{regime_name.upper()}] Transition Matrix ({pairs_count} year-pairs):")
        print(f"Saved to {OUT_DIR}")
        
    save_matrix(pre_counts, "Pre_2015", pre_pairs)
    save_matrix(post_counts, "Post_2015", post_pairs)

if __name__ == "__main__":
    compute_conditional_transitions()
