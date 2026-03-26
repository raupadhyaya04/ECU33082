import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from pathlib import Path

def plot_matrix(csv_path, out_path, title):
    if not csv_path.exists():
        print(f"File not found: {csv_path}")
        return
        
    df = pd.read_csv(csv_path, index_col=0)
    
    fig, ax = plt.subplots(figsize=(10, 8))
    sns.heatmap(df, annot=True, cmap="Blues", fmt=".1f", 
                cbar_kws={'label': 'Transition Probability (%)'},
                linewidths=.5, ax=ax)
    
    plt.yticks(rotation=0, fontsize=11)
    plt.xticks(fontsize=11)
    
    plt.title(title, pad=20, size=16, fontweight='bold')
    plt.ylabel("State at Year t+1", size=13, fontweight='medium', labelpad=10)
    plt.xlabel("State at Year t", size=13, fontweight='medium', labelpad=10)
    
    legend_text = (
        "State Notation (e.g., 'mH'):\n"
        "• 1st Letter (Interest Rate):   l = Low,  m = Medium,  h = High\n"
        "• 2nd Letter (Energy Poverty):  L = Low,  M = Medium,  H = High"
    )
    props = dict(boxstyle='round,pad=1.2', facecolor='#f8f9fa', edgecolor='#ced4da', alpha=1.0)
    fig.text(0.5, -0.05, legend_text, ha='center', va='top', fontsize=11, 
             bbox=props, fontfamily='monospace', linespacing=1.6)
    
    fig.tight_layout()
    plt.savefig(out_path, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"Saved PNG to: {out_path}")

def main():
    root = Path(__file__).resolve().parent.parent
    trans_dir = root / "Output" / "Transitions"
    plots_dir = root / "Output" / "plots"
    plots_dir.mkdir(parents=True, exist_ok=True)
    
    # 1. Pooled Matrix
    plot_matrix(
        trans_dir / "Transitions_POOLED_colpct.csv",
        plots_dir / "Global_Transitions_POOLED_matrix.png",
        "Global Joint State Transition Matrix (Pooled)"
    )
    
    # 2. Pre-2015 Matrix
    plot_matrix(
        trans_dir / "Conditional" / "Transitions_Pre_2015_runs_colpct.csv",
        plots_dir / "Global_Transitions_Pre_2015_matrix.png",
        "Global Joint State Matrix: PRE-2015 (Pre-SDGs)"
    )
    
    # 3. Post-2015 Matrix
    plot_matrix(
        trans_dir / "Conditional" / "Transitions_Post_2015_runs_colpct.csv",
        plots_dir / "Global_Transitions_Post_2015_matrix.png",
        "Global Joint State Matrix: POST-2015 (Post-SDGs)"
    )

if __name__ == "__main__":
    main()
