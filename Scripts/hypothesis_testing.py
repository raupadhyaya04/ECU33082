import pandas as pd
import numpy as np
from pathlib import Path
from scipy.stats import chi2_contingency

ROOT = Path(__file__).resolve().parent.parent
COND_DIR = ROOT / "Output" / "Transitions" / "Conditional"

def run_homogeneity_test(group1_counts: pd.DataFrame, group2_counts: pd.DataFrame, label1: str, label2: str):
    """
    Runs the Anderson-Goodman homogeneity test (Pearson Chi-Square) across corresponding columns
    (initial states) for two temporal regimes.
    """
    total_chi2 = 0.0
    total_dof = 0
    results = []

    # Get the union of columns
    cols = sorted(list(set(group1_counts.columns) | set(group2_counts.columns)))
    
    for col in cols:
        obs1 = group1_counts[col].fillna(0).values if col in group1_counts else np.zeros(len(group1_counts))
        obs2 = group2_counts[col].fillna(0).values if col in group2_counts else np.zeros(len(group2_counts))
        
        # We need at least some transitions to test
        if obs1.sum() == 0 and obs2.sum() == 0:
            continue
            
        # Contingency table for this specific starting state
        contingency_table = np.array([obs1, obs2])
        
        # Remove zero columns (destinations that never happened) 
        # to avoid division by zero in expected frequencies
        active_cols = contingency_table.sum(axis=0) > 0
        contingency_table = contingency_table[:, active_cols]

        if contingency_table.shape[1] > 1:
            chi2, p_val, dof, expected = chi2_contingency(contingency_table)
            total_chi2 += chi2
            total_dof += dof
            results.append({
                "State": col, "Chi2": chi2, "DoF": dof, "P-Value": p_val
            })

    # Global significance
    from scipy.stats import chi2
    global_p_val = chi2.sf(total_chi2, total_dof)

    print(f"\n{'='*50}")
    print(f"HYPOTHESIS TEST: {label1} vs {label2}")
    print(f"{'='*50}")
    print(f"Total Chi-Square Stat : {total_chi2:.3f}")
    print(f"Global Degrees of Free: {total_dof}")
    print(f"Global p-value        : {global_p_val:.4e}")
    if global_p_val < 0.05:
        print("Conclusion            : WE REJECT THE NULL HYPOTHESIS. Structural break confirmed.")
    else:
        print("Conclusion            : WE FAIL TO REJECT THE NULL. No statistically significant difference.")
    
    # Save detailed results
    res_df = pd.DataFrame(results)
    res_df.to_csv(COND_DIR / f"Hypothesis_Test_{label1}_vs_{label2}.csv", index=False)

def main():
    # 1. Standard UN SDG Split (Includes shocks)
    pre = pd.read_csv(COND_DIR / "Transitions_Pre_2015_runs_counts.csv", index_col=0)
    post = pd.read_csv(COND_DIR / "Transitions_Post_2015_runs_counts.csv", index_col=0)
    run_homogeneity_test(pre, post, "Pre-2015", "Post-2015")
    
    # 2. Strict UN SDG Split (Excluding Shocks: 2008, 09, 20-22)
    pre_stable = pd.read_csv(COND_DIR / "Transitions_Pre_2015_Stable_runs_counts.csv", index_col=0)
    post_stable = pd.read_csv(COND_DIR / "Transitions_Post_2015_Stable_runs_counts.csv", index_col=0)
    run_homogeneity_test(pre_stable, post_stable, "Pre-2015_Stable", "Post-2015_Stable")
    
    # 3. Baseline vs Crisis
    baseline = pd.read_csv(COND_DIR / "Transitions_Baseline_runs_counts.csv", index_col=0)
    crisis = pd.read_csv(COND_DIR / "Transitions_Crisis_runs_counts.csv", index_col=0)
    run_homogeneity_test(baseline, crisis, "Baseline", "Crisis-Years")

if __name__ == "__main__":
    main()
