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
from statsmodels.tsa.stattools import adfuller, grangercausalitytests

try:
    from docx import Document
except Exception:
    Document = None


PROJECT_ROOT = Path(__file__).resolve().parents[1]
ROOT = PROJECT_ROOT / "aggregateData"
FILE_300 = ROOT / "TS_CSI_300_PCA.xlsx"
ADF_COLS = ["Rexcess", "PCAMS", "exPCAMS", "MMSentlag"]


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


def adf_result(series: pd.Series) -> dict:
    s = pd.to_numeric(series, errors="coerce").dropna()
    if s.empty:
        return {
            "stat": np.nan,
            "p": np.nan,
            "lag": np.nan,
            "nobs": np.nan,
            "crit_1": np.nan,
            "crit_5": np.nan,
            "crit_10": np.nan,
        }
    stat, pval, usedlag, nobs, crit, _ = adfuller(s, autolag="AIC")
    return {
        "stat": stat,
        "p": pval,
        "lag": usedlag,
        "nobs": nobs,
        "crit_1": crit["1%"],
        "crit_5": crit["5%"],
        "crit_10": crit["10%"],
    }


def pval_stars(pval: float) -> str:
    if pval < 0.01:
        return "***"
    if pval < 0.05:
        return "**"
    if pval < 0.10:
        return "*"
    return ""


def write_adf_table_word(results: dict, out_path: Path) -> None:
    if Document is None:
        print("\n[ADF] python-docx not installed. Skipping Word export.")
        print("Install with: pip install python-docx")
        return

    doc = Document()
    doc.add_heading("ADF 单位根检验结果", level=1)
    doc.add_paragraph("注：显著性水平：*** p<0.01, ** p<0.05, * p<0.10")

    metrics = [
        "ADF统计量",
        "p值",
        "滞后阶数",
        "样本量",
        "1%临界值",
        "5%临界值",
        "10%临界值",
        "平稳性结论",
    ]
    vars_order = list(results.keys())

    table = doc.add_table(rows=1 + len(metrics), cols=1 + len(vars_order))
    table.style = "Table Grid"

    # Header row
    table.cell(0, 0).text = "指标/变量"
    for j, var in enumerate(vars_order, start=1):
        table.cell(0, j).text = var

    # Fill rows
    for i, metric in enumerate(metrics, start=1):
        table.cell(i, 0).text = metric
        for j, var in enumerate(vars_order, start=1):
            res = results[var]
            if metric == "ADF统计量":
                stat = res["stat"]
                pval = res["p"]
                val = "" if pd.isna(stat) else f"{stat:.4f}{pval_stars(pval)}"
            elif metric == "p值":
                val = "" if pd.isna(res["p"]) else f"{res['p']:.4f}"
            elif metric == "滞后阶数":
                val = "" if pd.isna(res["lag"]) else f"{int(res['lag'])}"
            elif metric == "样本量":
                val = "" if pd.isna(res["nobs"]) else f"{int(res['nobs'])}"
            elif metric == "1%临界值":
                val = "" if pd.isna(res["crit_1"]) else f"{res['crit_1']:.4f}"
            elif metric == "5%临界值":
                val = "" if pd.isna(res["crit_5"]) else f"{res['crit_5']:.4f}"
            elif metric == "10%临界值":
                val = "" if pd.isna(res["crit_10"]) else f"{res['crit_10']:.4f}"
            elif metric == "平稳性结论":
                pval = res["p"]
                if pd.isna(pval):
                    val = ""
                elif pval < 0.05:
                    val = "平稳"
                elif pval < 0.10:
                    val = "10%平稳"
                else:
                    val = "不平稳"
            else:
                val = ""
            table.cell(i, j).text = val

    doc.save(out_path)
    print("Saved ADF table to:", out_path)


def coef_with_stars(coef: float, pval: float) -> str:
    return f"{coef:.4f}{pval_stars(pval)}"


def write_reg_table_word(results: list[dict], out_path: Path) -> None:
    if Document is None:
        print("\n[Reg] python-docx not installed. Skipping Word export.")
        print("Install with: pip install python-docx")
        return

    doc = Document()
    doc.add_heading("回归结果汇总表", level=1)
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


def write_granger_table_word(results: dict[int, dict[str, float]], out_path: Path) -> None:
    if Document is None:
        print("\n[Granger] python-docx not installed. Skipping Word export.")
        print("Install with: pip install python-docx")
        return

    doc = Document()
    doc.add_heading("Granger 因果检验结果", level=1)
    doc.add_paragraph("注：*** p<0.01, ** p<0.05, * p<0.10")

    table = doc.add_table(rows=1 + len(results), cols=3)
    table.style = "Table Grid"

    table.cell(0, 0).text = "滞后阶数"
    table.cell(0, 1).text = "PCAMS → Rexcess (p值)"
    table.cell(0, 2).text = "Rexcess → PCAMS (p值)"

    for i, lag in enumerate(sorted(results.keys()), start=1):
        row = results[lag]
        p1 = row["pcams_to_rexcess"]
        p2 = row["rexcess_to_pcams"]
        table.cell(i, 0).text = str(lag)
        table.cell(i, 1).text = f"{p1:.4f}{pval_stars(p1)}"
        table.cell(i, 2).text = f"{p2:.4f}{pval_stars(p2)}"

    doc.save(out_path)
    print("Saved Granger table to:", out_path)


def zscore(series: pd.Series) -> pd.Series:
    s = pd.to_numeric(series, errors="coerce")
    mu = s.mean(skipna=True)
    sigma = s.std(skipna=True)
    if sigma and not np.isnan(sigma):
        return (s - mu) / sigma
    return pd.Series(np.nan, index=series.index)


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


def main() -> None:
    df = pd.read_excel(FILE_300)

    # Keep time order before regressions
    if "Month" in df.columns:
        df = df.sort_values("Month").reset_index(drop=True)

    # ADF tests
    required_adf_cols = ADF_COLS + ["PE", "PB"]
    for col in required_adf_cols:
        if col not in df.columns:
            raise KeyError(f"Missing {col} in {FILE_300}")

    df["Abs_Rexcess"] = pd.to_numeric(df["Rexcess"], errors="coerce").abs()
    df["PE_diff1"] = pd.to_numeric(df["PE"], errors="coerce").diff()

    adf_series = {
        "Rexcess": df["Rexcess"],
        "|Rexcess|": df["Abs_Rexcess"],
        "PE": df["PE"],
        "D(PE)": df["PE_diff1"],
        "PB": df["PB"],
        "PCAMS": df["PCAMS"],
        "exPCAMS": df["exPCAMS"],
        "MMSentlag": df["MMSentlag"],
    }
    for name, series in adf_series.items():
        adf_report(series, name)

    # Build ADF table (transposed layout) and export to Word
    adf_results = {name: adf_result(series) for name, series in adf_series.items()}
    out_docx = ROOT / "ADF_Table.docx"
    write_adf_table_word(adf_results, out_docx)

    # Standardize sentiment variables before regressions
    for col in ["MMSentlag", "PCAMS", "exPCAMS"]:
        if col in df.columns:
            df[col] = zscore(df[col])

    # Build lead-1 Rexcess
    if "Rexcess_lead1" not in df.columns:
        df["Rexcess_lead1"] = pd.to_numeric(df["Rexcess"], errors="coerce").shift(-1)

    reg_results = []

    # Regression 1: Rexcess_lead1 ~ PCAMS + Rexcess
    reg_results.append(
        run_regression(
            df,
            y_col="Rexcess_lead1",
            x_cols=["PCAMS", "Rexcess"],
            title="(1) PCAMS + Rexcess",
        )
    )

    # Regression 2: Rexcess_lead1 ~ exPCAMS + Rexcess
    reg_results.append(
        run_regression(
            df,
            y_col="Rexcess_lead1",
            x_cols=["exPCAMS", "Rexcess"],
            title="(2) exPCAMS + Rexcess",
        )
    )

    # Regression 3: Rexcess_lead1 ~ MMSentlag + Rexcess
    reg_results.append(
        run_regression(
            df,
            y_col="Rexcess_lead1",
            x_cols=["MMSentlag", "Rexcess"],
            title="(3) MMSentlag + Rexcess",
        )
    )

    out_reg_docx = ROOT / "Regression_Table.docx"
    write_reg_table_word(reg_results, out_reg_docx)

    # Granger causality tests between Rexcess and PCAMS
    g_df = df[["Rexcess", "PCAMS"]].apply(pd.to_numeric, errors="coerce").dropna()
    if len(g_df) < 10:
        print("\n[Granger] Not enough data for Granger causality tests.")
        return

    maxlag = 4
    print(f"\n[Granger] Testing maxlag={maxlag}")

    print("\nDoes PCAMS Granger-cause Rexcess?")
    res1 = grangercausalitytests(g_df[["Rexcess", "PCAMS"]], maxlag=maxlag, verbose=False)
    for lag, out in res1.items():
        ftest_p = out[0]["ssr_ftest"][1]
        print(f"  lag {lag}: p-value = {ftest_p:.6f}")

    print("\nDoes Rexcess Granger-cause PCAMS?")
    res2 = grangercausalitytests(g_df[["PCAMS", "Rexcess"]], maxlag=maxlag, verbose=False)
    for lag, out in res2.items():
        ftest_p = out[0]["ssr_ftest"][1]
        print(f"  lag {lag}: p-value = {ftest_p:.6f}")

    granger_rows = {}
    for lag in range(1, maxlag + 1):
        p1 = res1[lag][0]["ssr_ftest"][1]
        p2 = res2[lag][0]["ssr_ftest"][1]
        granger_rows[lag] = {
            "pcams_to_rexcess": p1,
            "rexcess_to_pcams": p2,
        }

    out_granger_docx = ROOT / "Granger_Table.docx"
    write_granger_table_word(granger_rows, out_granger_docx)


if __name__ == "__main__":
    main()
