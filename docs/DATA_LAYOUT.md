# Data Layout

This trimmed repository is built around the PCA sentiment analysis section only.

## Files Expected Locally In `aggregateData/`

- `TS.xlsx`
- `TS_CSI_300_PCA.xlsx`
- `TS_CSI_ALL_PCA.xlsx`

These Excel files are required for local reproduction but are not included in
the public repository because of copyright and data licensing restrictions.

The scripts in `code/` read from and write to this folder. If you rerun the
analysis locally, additional output files such as `*.docx`, `*.csv`, and `*.png`
may be generated here. Only outputs that do not contain restricted source data
should be committed.
