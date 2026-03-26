import pandas as pd
import numpy as np
from pathlib import Path


def _compute_counts(panel: pd.DataFrame, year_t: int, year_t1: int) -> pd.DataFrame:
    """Small helper mirroring Scripts/compute_transition_matrices.py logic."""
    obs = panel.copy()
    t0 = obs[obs["Year"] == year_t][["Country_Code", "Joint_State"]].drop_duplicates("Country_Code", keep="last")
    t1 = obs[obs["Year"] == year_t1][["Country_Code", "Joint_State"]].drop_duplicates("Country_Code", keep="last")

    merged = pd.merge(t0, t1, on="Country_Code", how="inner", suffixes=("_t", "_t1"))
    # Remember it's now X-axis = t, Y-axis = t+1 (left-stochastic)
    ct = pd.crosstab(merged["Joint_State_t1"], merged["Joint_State_t"])
    return ct


def test_synthetic_transition_col_sums_100():
    panel = pd.DataFrame(
        {
            "Year": [2017, 2017, 2018, 2018],
            "Country_Code": ["USA", "GBR", "USA", "GBR"],
            "Joint_State": ["lL", "hH", "mM", "hH"]
        }
    )

    ct = _compute_counts(panel, 2017, 2018)
    # Counts should match 2 countries
    assert int(ct.values.sum()) == 2

    colpct = (ct.div(ct.sum(axis=0), axis=1) * 100).fillna(0)
    # Each non-empty column should sum to 100
    for col_name, col_data in colpct.items():
        if ct[col_name].sum() > 0:
            assert abs(col_data.sum() - 100) < 1e-9


def test_real_pooled_transition_files_exist_and_valid():
    root = Path(__file__).resolve().parent.parent
    counts_fp = root / "Output" / "Transitions" / "Transitions_POOLED_counts.csv"
    colpct_fp = root / "Output" / "Transitions" / "Transitions_POOLED_colpct.csv"

    assert counts_fp.exists(), "Missing pooled counts CSV."
    assert colpct_fp.exists(), "Missing pooled col% CSV."

    tc = pd.read_csv(counts_fp, index_col=0)
    tr = pd.read_csv(colpct_fp, index_col=0)

    # same shape and labels
    assert list(tc.index) == list(tr.index)
    assert list(tc.columns) == list(tr.columns)

    # counts are non-negative integers
    assert (tc.fillna(0).values >= 0).all()
    assert np.allclose(tc.fillna(0).values, np.round(tc.fillna(0).values))

    # col% columns sum to ~100 unless col is all zeros
    for c in tr.columns:
        col_sum_counts = float(tc[c].sum())
        col_sum_pct = float(tr[c].sum())
        if col_sum_counts == 0:
            assert abs(col_sum_pct) < 1e-9
        else:
            assert abs(col_sum_pct - 100) <= 0.2  # rounding to 1dp
