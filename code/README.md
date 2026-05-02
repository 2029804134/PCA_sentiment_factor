# Code Overview

This directory contains only the scripts used in the PCA sentiment analysis section of the thesis.

## Scripts

- `Get_PCAMS.py`
  Standardizes the core sentiment proxies, runs PCA, and writes `PCAMS` and `exPCAMS`.
- `Baseline.py`
  Runs the CSI 300 baseline regressions and exports regression tables.
- `test1.py`
  Runs the CSI All Share baseline regressions and exports regression tables.
- `Stats.py`
  Produces descriptive statistics, charts, and selected statistical tests.
- `heterogeneity.py`
  Runs heterogeneity regressions based on quantile grid search.
- `mediation_bootstrap.py`
  Runs the dual-mediator bootstrap analysis.

## Suggested Order

1. `Get_PCAMS.py`
2. `Baseline.py`
3. `test1.py`
4. `Stats.py`
5. `heterogeneity.py`
6. `mediation_bootstrap.py`

You can also run the same subset workflow through [scripts/run_pipeline.py](../scripts/run_pipeline.py).
