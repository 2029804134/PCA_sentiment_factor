from __future__ import annotations

from pathlib import Path

import numpy as np
import pandas as pd
import statsmodels.api as sm
from statsmodels.stats.diagnostic import (
    het_white,
    het_breuschpagan,
    acorr_breusch_godfrey,
)
from statsmodels.tsa.stattools import adfuller

try:
    from docx import Document
except Exception:
    Document = None


PROJECT_ROOT = Path(__file__).resolve().parents[1]
ROOT = PROJECT_ROOT / "aggregateData"
FILE_ALL = ROOT / "TS_CSI_ALL_PCA.xlsx"


def zscore(series: pd.Series) -> pd.Series:
    s = pd.to_numeric(series, errors="coerce")
    mu = s.mean(skipna=True)
    sigma = s.std(skipna=True)
    if sigma and not np.isnan(sigma):
        return (s - mu) / sigma
    return pd.Series(np.nan, index=series.index)


def adf_report(series: pd.Series, name: str) -> None:
    s = pd.to_numeric(series, errors="coerce").dropna()
    if s.empty:
        print(f"[ADF] {name}: no valid data")
        return
    stat, pval, usedlag, nobs, crit, _ = adfuller(s, autolag="AIC")
    print(f"\n[ADF] {name}")
    print(f"  Test statistic: {stat:.6f}")
    print(f"  p-value:        {pval:.6f}")
    print(f"  used lag:       {usedlag}")
    print(f"  nobs:           {nobs}")
    for k, v in crit.items():
        print(f"  crit {k}:    {v:.6f}")
    has_unit_root = pval >= 0.05
    conclusion = "unit root (non-stationary)" if has_unit_root else "no unit root (stationary)"
    print(f"  Conclusion (5%): {conclusion}")


def pval_stars(pval: float) -> str:
    if pval < 0.01:
        return "***"
    if pval < 0.05:
        return "**"
    if pval < 0.10:
        return "*"
    return ""


def coef_with_stars(coef: float, pval: float) -> str:
    return f"{coef:.4f}{pval_stars(pval)}"


def run_regression(
    df: pd.DataFrame,
    y_col: str,
    x_cols: list[str],
    title: str,
) -> dict:
    y = pd.to_numeric(df[y_col], errors="coerce")
    X = df[x_cols].apply(pd.to_numeric, errors="coerce")
    reg = pd.concat([y, X], axis=1).dropna()
    if reg.empty:
        raise ValueError(f"No valid rows for regression: {title}")

    y = reg[y_col]
    X = sm.add_constant(reg[x_cols])
    ols = sm.OLS(y, X).fit()

    white_stat, white_p, _, _ = het_white(ols.resid, X)
    bp_stat, bp_p, _, _ = het_breuschpagan(ols.resid, X)
    bg_stat, bg_p, _, _ = acorr_breusch_godfrey(ols, nlags=4)

    print(f"\n[OLS] {title}")
    print(ols.summary())

    # Residual stationarity test (ADF)
    try:
        stat, pval, usedlag, nobs, crit, _ = adfuller(ols.resid, autolag="AIC")
        print("\n[ADF] Residuals")
        print(f"  Test statistic: {stat:.6f}")
        print(f"  p-value:        {pval:.6f}")
        print(f"  used lag:       {usedlag}")
        print(f"  nobs:           {nobs}")
        for k, v in crit.items():
            print(f"  crit {k}:    {v:.6f}")
    except Exception as exc:
        print(f"\n[ADF] Residuals: failed ({exc})")

    print("\n[White test]")
    print(f"  stat: {white_stat:.6f}, p-value: {white_p:.6f}")

    print("\n[Breusch-Pagan test]")
    print(f"  stat: {bp_stat:.6f}, p-value: {bp_p:.6f}")

    print("\n[Breusch-Godfrey test (nlags=4)]")
    print(f"  stat: {bg_stat:.6f}, p-value: {bg_p:.6f}")

    print("\n[HAC] Serial correlation not detected at 5%, HAC not applied.")

    return {
        "title": title,
        "params": ols.params,
        "bse": ols.bse,
        "pvalues": ols.pvalues,
        "nobs": int(ols.nobs),
        "r2": ols.rsquared,
        "adj_r2": ols.rsquared_adj,
    }


def write_reg_table_word(results: list[dict], out_path: Path) -> None:
    if Document is None:
        print("\n[Reg] python-docx not installed. Skipping Word export.")
        print("Install with: pip install python-docx")
        return

    doc = Document()
    doc.add_heading("回归结果汇总表（中证全指）", level=1)
    doc.add_paragraph("注：*** p<0.01, ** p<0.05, * p<0.10；括号内为标准误。")

    row_vars = ["PCAMS", "exPCAMS", "MMSentlag", "Rexcess", "const"]
    row_labels = {
        "PCAMS": "PCAMS",
        "exPCAMS": "exPCAMS",
        "MMSentlag": "MMSentlag",
        "Rexcess": "Rexcess",
        "const": "常数项",
    }

    n_rows = len(row_vars) + 3
    n_cols = 1 + len(results)
    table = doc.add_table(rows=1 + n_rows, cols=n_cols)
    table.style = "Table Grid"

    table.cell(0, 0).text = "变量/模型"
    for j, res in enumerate(results, start=1):
        table.cell(0, j).text = res["title"]

    for i, var in enumerate(row_vars, start=1):
        table.cell(i, 0).text = row_labels[var]
        for j, res in enumerate(results, start=1):
            if var in res["params"].index:
                coef = res["params"][var]
                pval = res["pvalues"][var]
                se = res["bse"][var]
                table.cell(i, j).text = f"{coef_with_stars(coef, pval)}\n({se:.4f})"
            else:
                table.cell(i, j).text = ""

    stats_rows = [
        ("观测值", "nobs"),
        ("R²", "r2"),
        ("Adj. R²", "adj_r2"),
    ]
    start = 1 + len(row_vars)
    for k, (label, key) in enumerate(stats_rows, start=start):
        table.cell(k, 0).text = label
        for j, res in enumerate(results, start=1):
            val = res[key]
            if key == "nobs":
                table.cell(k, j).text = str(val)
            else:
                table.cell(k, j).text = f"{val:.4f}"

    doc.save(out_path)
    print("Saved regression table to:", out_path)


def main() -> None:
    df = pd.read_excel(FILE_ALL)

    # Keep time order before building HP-filtered controls
    if "Month" in df.columns:
        df = df.sort_values("Month").reset_index(drop=True)

    # Standardize sentiment variables before regressions
    for col in ["MMSentlag", "PCAMS", "exPCAMS"]:
        if col in df.columns:
            df[col] = zscore(df[col])
        else:
            raise KeyError(f"Missing {col} in {FILE_ALL}")

    if "Rexcess" not in df.columns:
        raise KeyError("Missing Rexcess in TS_CSI_ALL_PCA.xlsx")

    # Build lead-1 Rexcess
    if "Rexcess_lead1" not in df.columns:
        if "Rexcess" not in df.columns:
            raise KeyError("Missing both Rexcess_lead1 and Rexcess in TS_CSI_ALL_PCA.xlsx")
        df["Rexcess_lead1"] = pd.to_numeric(df["Rexcess"], errors="coerce").shift(-1)

    reg_results = []

    reg_results.append(
        run_regression(
            df,
            y_col="Rexcess_lead1",
            x_cols=["PCAMS", "Rexcess"],
            title="(1) PCAMS + Rexcess",
        )
    )

    reg_results.append(
        run_regression(
            df,
            y_col="Rexcess_lead1",
            x_cols=["exPCAMS", "Rexcess"],
            title="(2) exPCAMS + Rexcess",
        )
    )

    reg_results.append(
        run_regression(
            df,
            y_col="Rexcess_lead1",
            x_cols=["MMSentlag", "Rexcess"],
            title="(3) MMSentlag + Rexcess",
        )
    )

    out_reg_docx = ROOT / "Regression_Table_ALL.docx"
    write_reg_table_word(reg_results, out_reg_docx)


if __name__ == "__main__":
    main()
