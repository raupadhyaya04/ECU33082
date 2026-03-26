# ECU33082 Mathematical Economics Project

## Collaborators:

- Raj Aryan Upadhyaya
- Maria Faro Santana
- Emma Falchetti
- Victor Grönlund

## Transition matrices and plots

Transition matrices (decile → decile energy-poverty state transitions) are produced from the panel dataset and saved under:

- `Output/Transitions/` — per-year counts and row-% CSVs, plus pooled counts/row-% and a summary CSV.
- `Output/Transitions/plots/` — simple visualizations (heatmap and summary bars) for the pooled transition matrix.

To re-generate these artifacts locally run:

```bash
python3 Scripts/compute_transition_matrices.py
python3 Scripts/transitions_summary_and_plots.py
```

If you want the pooled transitions included inside the document-ready Excel export (`Output/Descriptive_Statistics_Document.xlsx`), run `Scripts/export_stats_document.py`. Note: this script uses `openpyxl` to write the workbook — install it first if necessary:

```bash
python3 -m pip install --user openpyxl
python3 Scripts/export_stats_document.py
```
