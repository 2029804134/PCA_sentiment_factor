# PCA Sentiment Analysis Repository

This repository is a trimmed public release for the PCA-based market sentiment analysis section of the thesis. It includes only the code and data needed for:

- building `PCAMS` and `exPCAMS`;
- running CSI 300 and CSI All Share baseline regressions;
- producing descriptive statistics and figures;
- running heterogeneity regressions;
- running the dual-mediator bootstrap analysis.

The repository is published under the MIT License.

## Included Code

- `code/Get_PCAMS.py`
  Standardizes the core sentiment proxies, runs PCA, and writes `PCAMS` and `exPCAMS`.
- `code/Baseline.py`
  Runs the CSI 300 baseline regressions and exports regression tables.
- `code/test1.py`
  Runs the CSI All Share baseline regressions and exports regression tables.
- `code/Stats.py`
  Produces descriptive statistics, charts, and selected statistical tests.
- `code/heterogeneity.py`
  Runs heterogeneity regressions based on quantile grid search.
- `code/mediation_bootstrap.py`
  Runs the dual-mediator bootstrap analysis.
- `scripts/run_pipeline.py`
  Runs the subset workflow from PCA construction through the later analysis scripts.

## Data Availability

The following datasets are required to reproduce the full analysis, but they are
not included in this public repository because of copyright and data licensing
restrictions:

- `aggregateData/TS.xlsx`
  Local input dataset for `Get_PCAMS.py`.
- `aggregateData/TS_CSI_300_PCA.xlsx`
  Local PCA dataset for the CSI 300 sample.
- `aggregateData/TS_CSI_ALL_PCA.xlsx`
  Local PCA dataset for the CSI All Share sample.

To run the scripts, place these files under `aggregateData/` in your local
working copy. Publicly shareable derived outputs, such as selected CSV summaries
and figures, may be kept in `aggregateData/` when they do not contain restricted
source data.

## Repository Structure

```text
.
|-- code/              Analysis scripts for the PCA sentiment section
|-- scripts/           Optional workflow entry point
|-- aggregateData/     Input and derived data for this section
|-- docs/              Short publication notes
|-- requirements.txt   Python dependencies
|-- LICENSE            MIT License
`-- README.md          Project overview
```

## Quick Start

Install dependencies:

```bash
pip install -r requirements.txt
```

Run the subset workflow from the repository root:

```bash
python scripts/run_pipeline.py
```

Or run the scripts individually in this order:

1. `python code/Get_PCAMS.py`
2. `python code/Baseline.py`
3. `python code/test1.py`
4. `python code/Stats.py`
5. `python code/heterogeneity.py`
6. `python code/mediation_bootstrap.py`

## Notes

- This repository no longer includes the earlier raw-data construction scripts.
- Word files generated during analysis are not tracked by default.
- If you want the GitHub repository to display your real name in the license, replace the placeholder line in [LICENSE](LICENSE).

## License

This project is released under the [MIT License](LICENSE).
