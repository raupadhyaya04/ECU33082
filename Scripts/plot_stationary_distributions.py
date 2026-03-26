import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from pathlib import Path

def main():
    root = Path(__file__).resolve().parent.parent
    trans_dir = root / "Output" / "Transitions"
    plots_dir = root / "Output" / "plots"
    plots_dir.mkdir(parents=True, exist_ok=True)
    
    data_file = trans_dir / "Stationary_Distributions_Analytics.csv"
    if not data_file.exists():
        print(f"Data not found: {data_file}")
        return
        
    df = pd.read_csv(data_file, index_col=0)
    
    # === 1. Plot Stationary Distributions (Grouped Bar Chart) ===
    state_cols = [c for c in df.columns if "State" in c]
    dist_df = df[state_cols].copy()
    
    # Clean up column names for the plot so X-axis just says "lL", "mH", etc.
    dist_df.columns = [c.replace("State_", "").replace("_(%)", "") for c in dist_df.columns]
    
    # Transpose so states are on X-axis and models are hues
    dist_df = dist_df.T
    
    fig, ax = plt.subplots(figsize=(12, 6))
    colors = ['#2c3e50', '#e74c3c', '#3498db']
    dist_df.plot(kind='bar', ax=ax, width=0.8, color=colors[:len(df)])
    
    plt.title("Long-Run (Stationary) Distributions by Regime", pad=20, size=16, fontweight='bold')
    plt.ylabel("Long-Run Probability (%)", size=13, fontweight='medium')
    plt.xlabel("Joint State", size=13, fontweight='medium')
    plt.xticks(rotation=0, fontsize=11)
    plt.yticks(fontsize=11)
    plt.legend(title="Regime Model", title_fontsize='12', fontsize='11')
    plt.grid(axis='y', linestyle='--', alpha=0.7)
    
    legend_text = (
        "State Notation: 1st Letter (Interest Rate): l=Low, m=Medium, h=High | "
        "2nd Letter (Energy Poverty): L=Low, M=Medium, H=High"
    )
    plt.figtext(0.5, 0.01, legend_text, ha='center', fontsize=10, 
                bbox=dict(boxstyle='round,pad=0.5', facecolor='#f8f9fa', edgecolor='#ced4da'))
    
    plt.tight_layout(rect=[0, 0.04, 1, 1])
    out_path1 = plots_dir / "Stationary_Distributions_BarChart.png"
    plt.savefig(out_path1, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"Saved PNG to: {out_path1}")

    # === 2. Plot Shorrocks Mobility Index ===
    mob_df = df[["Shorrocks_Mobility_Index"]].copy()
    fig, ax = plt.subplots(figsize=(8, 5))
    
    bars = ax.bar(mob_df.index, mob_df["Shorrocks_Mobility_Index"], color=['#2c3e50', '#e74c3c', '#3498db'][:len(df)])
    
    plt.title("Shorrocks Mobility Index by Regime", pad=20, size=16, fontweight='bold')
    plt.ylabel("Mobility Index (0 = Rigid, 1 = Perfect Mobile)", size=13, fontweight='medium')
    plt.ylim(0, max(mob_df["Shorrocks_Mobility_Index"]) * 1.4)
    plt.xticks(fontsize=12)
    
    # Annotate bar heights
    for bar in bars:
        yval = bar.get_height()
        ax.text(bar.get_x() + bar.get_width()/2, yval + 0.005, 
                f"{yval:.4f}", ha='center', va='bottom', fontsize=11, fontweight='bold')
                
    plt.tight_layout()
    out_path2 = plots_dir / "Mobility_Index_BarChart.png"
    plt.savefig(out_path2, dpi=300)
    plt.close()
    print(f"Saved PNG to: {out_path2}")

if __name__ == "__main__":
    main()
