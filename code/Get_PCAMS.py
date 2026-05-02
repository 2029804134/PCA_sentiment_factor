from __future__ import annotations

from pathlib import Path

import numpy as np
import pandas as pd
from sklearn.decomposition import PCA
from sklearn.preprocessing import StandardScaler


PROJECT_ROOT = Path(__file__).resolve().parents[1]
ROOT = PROJECT_ROOT / "aggregateData"
TS_FILE = ROOT / "TS.xlsx"
OUT_FILE_300 = ROOT / "TS_CSI_300_PCA.xlsx"
OUT_FILE_985 = ROOT / "TS_CSI_ALL_PCA.xlsx"

MONTH_COL = "Month"
INDEX_COL = "Indexcd"

X_COLS_FULL = ["Turnover_diff", "NIA", "NIPO", "MMSentlag", "CEFD"]
X_COLS_EX = ["Turnover_diff", "NIA", "NIPO", "CEFD"]
Y_COL = "Rexcess_lead1"
VALUATION_COLS = ["PB", "PE"]


def normalize_month(series: pd.Series) -> pd.Series:
    if pd.api.types.is_datetime64_any_dtype(series):
        return series.dt.to_period("M").astype("string")

    s = series.astype("string").str.strip()
    s = s.str.replace(r"\.0$", "", regex=True)
    s = (
        s.str.replace("/", "-", regex=False)
        .str.replace("年", "-", regex=False)
        .str.replace("月", "", regex=False)
    )

    out = pd.Series(pd.NaT, index=s.index, dtype="datetime64[ns]")

    mask = s.str.fullmatch(r"\d{6}")
    if mask.any():
        out.loc[mask] = pd.to_datetime(s[mask], format="%Y%m", errors="coerce")

    mask = s.str.fullmatch(r"\d{4}-\d{1,2}")
    if mask.any():
        parts = s[mask].str.split("-", expand=True)
        tmp = parts[0] + "-" + parts[1].str.zfill(2)
        out.loc[mask] = pd.to_datetime(tmp, format="%Y-%m", errors="coerce")

    mask = s.str.fullmatch(r"\d{8}")
    if mask.any():
        out.loc[mask] = pd.to_datetime(s[mask], format="%Y%m%d", errors="coerce")

    mask = s.str.fullmatch(r"\d{4}-\d{2}-\d{2}")
    if mask.any():
        out.loc[mask] = pd.to_datetime(s[mask], format="%Y-%m-%d", errors="coerce")

    return out.dt.to_period("M").astype("string")


def run_for_index(df: pd.DataFrame, index_code: str, out_file: Path) -> None:
    # Filter by Indexcd
    sub = df[df[INDEX_COL] == index_code].copy()
    sub = sub.sort_values(MONTH_COL).reset_index(drop=True)
    if sub.empty:
        raise ValueError(f"No rows for Indexcd={index_code}")

    # Turnover first difference and insert after Turnover column
    if "Turnover" not in sub.columns:
        raise KeyError("Missing Turnover column in TS.xlsx")
    sub["Turnover_diff"] = pd.to_numeric(sub["Turnover"], errors="coerce").diff()

    cols = sub.columns.tolist()
    if "Turnover_diff" in cols:
        cols.remove("Turnover_diff")
    turn_idx = cols.index("Turnover")
    cols.insert(turn_idx + 1, "Turnover_diff")
    sub = sub[cols]

    # Lead-1 Rexcess as Y
    if "Rexcess" not in sub.columns:
        raise KeyError("Missing Rexcess column in TS.xlsx")
    sub[Y_COL] = pd.to_numeric(sub["Rexcess"], errors="coerce").shift(-1)

    # Keep valuation columns in output files and normalize as numeric
    for col in VALUATION_COLS:
        if col not in sub.columns:
            raise KeyError(f"Missing {col} column in TS.xlsx")
        sub[col] = pd.to_numeric(sub[col], errors="coerce")

    # Build modeling frame
    for col in ["NIA", "NIPO", "MMSentlag", "CEFD"]:
        if col not in sub.columns:
            raise KeyError(f"Missing {col} column in TS.xlsx")
        sub[col] = pd.to_numeric(sub[col], errors="coerce")

    model_full = sub[[MONTH_COL, *X_COLS_FULL, Y_COL]].copy()
    model_full = model_full.dropna(subset=X_COLS_FULL + [Y_COL])
    model_ex = sub[[MONTH_COL, *X_COLS_EX, Y_COL]].copy()
    model_ex = model_ex.dropna(subset=X_COLS_EX + [Y_COL])
    if model_full.empty or model_ex.empty:
        raise ValueError(f"No valid rows after diff/lead for Indexcd={index_code}.")

    # Standardize X and Y, then replace original values
    scaler_x = StandardScaler()
    scaler_y = StandardScaler()
    X_scaled_full = scaler_x.fit_transform(model_full[X_COLS_FULL])
    y_scaled = scaler_y.fit_transform(model_full[[Y_COL]]).ravel()

    sub.loc[model_full.index, X_COLS_FULL] = X_scaled_full
    sub.loc[model_full.index, Y_COL] = y_scaled

    # PCA with 1 component
    pca_full = PCA(n_components=1)
    pca_scores_full = pca_full.fit_transform(X_scaled_full)[:, 0]

    # Build ex set from the already-standardized columns
    ex_idx = [X_COLS_FULL.index(c) for c in X_COLS_EX]
    X_scaled_ex = X_scaled_full[:, ex_idx]
    pca_ex = PCA(n_components=1)
    pca_scores_ex = pca_ex.fit_transform(X_scaled_ex)[:, 0]

    # Align sign with Rexcess (if negative correlation, flip)
    rexcess_aligned = pd.to_numeric(sub.loc[model_full.index, "Rexcess"], errors="coerce")
    corr_full = pd.Series(pca_scores_full, index=model_full.index).corr(rexcess_aligned)
    if pd.notna(corr_full) and corr_full < 0:
        pca_scores_full = -pca_scores_full
        pca_full.components_[0] = -pca_full.components_[0]
        print(f"PCAMS flipped to make correlation with Rexcess positive. corr={corr_full:.6f}")
    else:
        print(f"PCAMS correlation with Rexcess: {corr_full:.6f}")

    rexcess_ex = pd.to_numeric(sub.loc[model_ex.index, "Rexcess"], errors="coerce")
    corr_ex = pd.Series(pca_scores_ex, index=model_ex.index).corr(rexcess_ex)
    if pd.notna(corr_ex) and corr_ex < 0:
        pca_scores_ex = -pca_scores_ex
        pca_ex.components_[0] = -pca_ex.components_[0]
        print(f"exPCAMS flipped to make correlation with Rexcess positive. corr={corr_ex:.6f}")
    else:
        print(f"exPCAMS correlation with Rexcess: {corr_ex:.6f}")

    sub["PCAMS"] = np.nan
    sub.loc[model_full.index, "PCAMS"] = pca_scores_full
    sub["exPCAMS"] = np.nan
    sub.loc[model_ex.index, "exPCAMS"] = pca_scores_ex

    # Ensure PCAMS and exPCAMS are the last columns
    cols = sub.columns.tolist()
    for col in ["PCAMS", "exPCAMS"]:
        if col in cols:
            cols.remove(col)
    cols.extend(["PCAMS", "exPCAMS"])
    sub = sub[cols]

    try:
        sub.to_excel(out_file, index=False)
        print("Saved:", out_file)
    except PermissionError:
        alt = out_file.with_name(out_file.stem + "_new" + out_file.suffix)
        sub.to_excel(alt, index=False)
        print("File is open/locked. Saved to:", alt)

    weights = pca_full.components_[0]
    print(f"\nPCA weights (PCAMS, Indexcd={index_code}):")
    for name, w in zip(X_COLS_FULL, weights):
        print(f"{name}: {w:.6f}")
    var_ratio = pca_full.explained_variance_ratio_[0]
    formula = " + ".join(
        [f"({w:.6f} * {name})" for name, w in zip(X_COLS_FULL, weights)]
    )
    print("PCAMS = " + formula)
    print(f"Explained variance ratio (PC1): {var_ratio:.6f}")

    weights_ex = pca_ex.components_[0]
    print(f"\nPCA weights (exPCAMS, Indexcd={index_code}):")
    for name, w in zip(X_COLS_EX, weights_ex):
        print(f"{name}: {w:.6f}")
    var_ratio_ex = pca_ex.explained_variance_ratio_[0]
    formula_ex = " + ".join(
        [f"({w:.6f} * {name})" for name, w in zip(X_COLS_EX, weights_ex)]
    )
    print("exPCAMS = " + formula_ex)
    print(f"Explained variance ratio (PC1): {var_ratio_ex:.6f}")
    print("(All X variables are standardized; PCAMS/exPCAMS are PC1 scores.)")


def main() -> None:
    df = pd.read_excel(TS_FILE)
    if MONTH_COL not in df.columns:
        raise KeyError(f"Missing {MONTH_COL} in {TS_FILE}")
    if INDEX_COL not in df.columns:
        raise KeyError(f"Missing {INDEX_COL} in {TS_FILE}")

    # Normalize month and index code
    df[MONTH_COL] = normalize_month(df[MONTH_COL])
    df[INDEX_COL] = df[INDEX_COL].astype("string").str.replace(r"\.0$", "", regex=True)

    # Use Indexcd as index, then sort by Month (global)
    df = df.set_index(INDEX_COL, drop=False).sort_values(MONTH_COL)

    run_for_index(df, "000300", OUT_FILE_300)
    run_for_index(df, "000985", OUT_FILE_985)


if __name__ == "__main__":
    main()
