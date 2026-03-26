import pandas as pd
import numpy as np
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
TRANS_DIR = ROOT / "Output" / "Transitions"

def compute_stationary_and_mobility(matrix_path, name):
    if not matrix_path.exists():
        return None
    
    # Load left-stochastic matrix (columns sum to 100)
    df = pd.read_csv(matrix_path, index_col=0)
    P = df.values / 100.0  # Normalize to 0-1
    
    # 1. Stationary Distribution
    # For carefully structured left-stochastic matrix P, we want P * v = 1 * v
    eigenvalues, eigenvectors = np.linalg.eig(P)
    
    # Find the index of the eigenvalue closest to 1.0 (it's guaranteed to have one)
    idx = np.argmin(np.abs(eigenvalues - 1.0))
    
    # Extract corresponding eigenvector and take the real part
    stationary = np.real(eigenvectors[:, idx])
    
    # Normalize probabilities so they sum to 100%
    stationary = (stationary / np.sum(stationary)) * 100
    
    # 2. Shorrocks Mobility Index
    # M(P) = [N - trace(P)] / [N - 1]
    # Measures how 'mobile' the system is. 0 = total rigidly / no movement, 1 = perfect mobility.
    N = P.shape[0]
    trace = np.trace(P)
    shorrocks = (N - trace) / (N - 1)
    
    out = {
        "Model": name,
        "Shorrocks_Mobility_Index": shorrocks
    }
    
    for state, prob in zip(df.index, stationary):
        out[f"State_{state}_(%)"] = prob
        
    return out

def main():
    results = []
    
    # Pooled
    pooled_path = TRANS_DIR / "Transitions_POOLED_colpct.csv"
    res = compute_stationary_and_mobility(pooled_path, "Pooled (Whole Period)")
    if res: results.append(res)
    
    # Pre-2015
    pre_path = TRANS_DIR / "Conditional" / "Transitions_Pre_2015_runs_colpct.csv"
    res = compute_stationary_and_mobility(pre_path, "Pre-2015")
    if res: results.append(res)
    
    # Post-2015
    post_path = TRANS_DIR / "Conditional" / "Transitions_Post_2015_runs_colpct.csv"
    res = compute_stationary_and_mobility(post_path, "Post-2015")
    if res: results.append(res)
    
    out_df = pd.DataFrame(results)
    out_df.set_index("Model", inplace=True)
    
    print("\n=======================================================================")
    print("                LONG-RUN (STATIONARY) DISTRIBUTIONS (%)                ")
    print("=======================================================================")
    print("If countries follow these transition rules infinitely, this is the \npercentage of the world that eventually settles into each state:\n")
    dist_cols = [c for c in out_df.columns if "State" in c]
    print(out_df[dist_cols].round(2))
    
    print("\n=======================================================================")
    print("                       SHORROCKS MOBILITY INDEX                        ")
    print("=======================================================================")
    print("(0 = Perfect Immobility/Trapped,  1 = Perfect Mobility)\n")
    print(out_df[["Shorrocks_Mobility_Index"]].round(4))
    print("\n=======================================================================\n")
    
    out_csv = TRANS_DIR / "Stationary_Distributions_Analytics.csv"
    out_df.round(4).to_csv(out_csv)
    print(f"✅ Saved complete analytics to: {out_csv}")

if __name__ == "__main__":
    main()
