import pandas as pd
import numpy as np
from pathlib import Path

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
PROJECT_ROOT = Path(__file__).resolve().parent
DATA_DIR = PROJECT_ROOT / "data"
RESULTS_DIR = PROJECT_ROOT / "results"

BUDGET_FILE = DATA_DIR / "2021-2030 Budget Data.xlsx"
CAPITAL_BUDGET_FILE = (
    DATA_DIR / "Tables_A.1_A.2_Teekay_13_25_year_capital_budgeting_plan.xls"
)

# Default discount rate if cost of capital is missing (7%)
DISCOUNT_RATE = 0.07

# Results files (all written inside results/ directory)
OUTPUT_TEEKAY_XLSX = RESULTS_DIR / "teekay_capital_budgeting_stats_and_npv.xlsx"
OUTPUT_DOD_XLSX = RESULTS_DIR / "dod_budget_category_stats.xlsx"
OUTPUT_XLSX = RESULTS_DIR / "capital_budgeting_analysis.xlsx"

YEAR0_COL = 2  # column index where year-0 data begins in Teekay sheets


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def _find_row(df: pd.DataFrame, label: str) -> int:
    """Find the first row whose column-0 value matches *label* (stripped)."""
    for i in range(df.shape[0]):
        val = df.iloc[i, 0]
        if isinstance(val, str) and val.strip() == label:
            return i
    raise KeyError(f"Row label '{label}' not found in sheet")


def _extract_yearly_series(df: pd.DataFrame, row: int,
                           start_col: int, end_col: int) -> np.ndarray:
    """Pull a numeric row slice from a raw (header=None) DataFrame."""
    values = df.iloc[row, start_col:end_col + 1].values
    return pd.to_numeric(values, errors="coerce").astype(float)


# ---------------------------------------------------------------------------
# Data loading
# ---------------------------------------------------------------------------
def load_budget_data() -> dict[str, pd.DataFrame]:
    return pd.read_excel(BUDGET_FILE, sheet_name=None, header=None,
                         engine="openpyxl")


def load_teekay(sheet: str) -> pd.DataFrame:
    return pd.read_excel(CAPITAL_BUDGET_FILE, sheet_name=sheet, header=None)


# ---------------------------------------------------------------------------
# Teekay capital-budgeting analysis
# ---------------------------------------------------------------------------
def analyse_teekay(sheet_name: str):
    df = load_teekay(sheet_name)

    # Locate key rows by label
    row_year = _find_row(df, "Year")
    # Cost of capital / discount rate (sheet-specific, fall back to default)
    try:
        row_cost = _find_row(df, "Cost of Capital =")
        # In both Teekay sheets the numeric rate is in the 3rd column (index 2)
        rate_val = pd.to_numeric(df.iloc[row_cost, 2], errors="coerce")
        rate = float(rate_val) if not np.isnan(rate_val) else DISCOUNT_RATE
    except KeyError:
        rate = DISCOUNT_RATE

    row_revenue = _find_row(df, "Charter Revenue")
    row_opex = _find_row(df, "Total Operating Costs")
    row_net_income = _find_row(df, "Project After Tax Op Income")
    row_ncf = _find_row(df, "Net Cash Flow")

    # Last year column from the Year row
    year_vals = df.iloc[row_year, :]
    last_col = year_vals.apply(pd.to_numeric, errors="coerce").last_valid_index()

    first_year_col = YEAR0_COL + 1  # skip year-0 for income-statement items

    revenue = _extract_yearly_series(df, row_revenue, first_year_col, last_col)
    opex = _extract_yearly_series(df, row_opex, first_year_col, last_col)
    net_income = _extract_yearly_series(df, row_net_income, first_year_col, last_col)
    cashflows = _extract_yearly_series(df, row_ncf, YEAR0_COL, last_col)

    metrics = {
        "Charter Revenue": revenue,
        "Total Operating Costs": opex,
        "After-Tax Operating Income": net_income,
    }
    rows = []
    for name, arr in metrics.items():
        clean = arr[~np.isnan(arr)]
        rows.append({
            "Metric": name,
            "Average": np.mean(clean),
            "StdDev": np.std(clean, ddof=1),
            "Min": np.min(clean),
            "Max": np.max(clean),
            "N": len(clean),
        })
    stats_df = pd.DataFrame(rows).set_index("Metric")

    return {
        "revenue": revenue,
        "opex": opex,
        "net_income": net_income,
        "cashflows": cashflows,
        "rate": rate,
        "stats": stats_df,
    }


# ---------------------------------------------------------------------------
# Debug helper: print raw yearly series for a Teekay sheet
# ---------------------------------------------------------------------------
def print_teekay_series(sheet_name: str) -> None:
    """
    Print the yearly sequences for Revenue, Operating Expenses,
    Net Income, and Net Cash Flow for a given Teekay sheet.
    """
    df = load_teekay(sheet_name)

    row_year = _find_row(df, "Year")
    year_vals = df.iloc[row_year, :].apply(pd.to_numeric, errors="coerce")
    last_col = year_vals.last_valid_index()

    # Years 1..n (skip year 0, which is col YEAR0_COL)
    years = year_vals.iloc[YEAR0_COL + 1 : last_col + 1].astype(int).tolist()

    res = analyse_teekay(sheet_name)
    rev = res["revenue"][~np.isnan(res["revenue"])]
    opex = res["opex"][~np.isnan(res["opex"])]
    ni = res["net_income"][~np.isnan(res["net_income"])]
    ncf = res["cashflows"][~np.isnan(res["cashflows"])]

    print(f"\nRaw series for {sheet_name}:")
    print(f"  Years (1..n): {years}")
    print(f"  Revenue:      {[round(x, 2) for x in rev.tolist()]}")
    print(f"  OpEx:         {[round(x, 2) for x in opex.tolist()]}")
    print(f"  Net Income:   {[round(x, 2) for x in ni.tolist()]}")
    print(f"  Cash Flows:   {[round(x, 2) for x in ncf.tolist()]}")


# ---------------------------------------------------------------------------
# DoD budget-data summary (fiscal-year averages by category)
# ---------------------------------------------------------------------------
def analyse_budget():
    sheets = load_budget_data()
    rows = []
    for sheet_name, df in sheets.items():
        if sheet_name == "DATA SOURCE":
            continue

        year_row_idx = None
        for i in range(min(10, len(df))):
            numeric = pd.to_numeric(df.iloc[i, :], errors="coerce")
            if (numeric > 2000).sum() >= 5:
                year_row_idx = i
                break
        if year_row_idx is None:
            continue

        year_cols = []
        for c in range(df.shape[1]):
            v = pd.to_numeric(df.iloc[year_row_idx, c], errors="coerce")
            if 2000 < v < 2100:
                year_cols.append(c)

        for i in range(year_row_idx + 1, len(df)):
            label = df.iloc[i, 0]
            if pd.isna(label) or str(label).strip() == "":
                continue
            vals = pd.to_numeric(
                df.iloc[i, year_cols].values, errors="coerce"
            ).astype(float)
            clean = vals[~np.isnan(vals)]
            if len(clean) == 0:
                continue
            rows.append({
                "Sheet": sheet_name,
                "Category": str(label).strip(),
                "Average ($M)": np.mean(clean),
                "StdDev ($M)": np.std(clean, ddof=1),
                "N_Years": len(clean),
            })

    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# NPV
# ---------------------------------------------------------------------------
def calculate_npv(cashflows: np.ndarray, rate: float) -> float:
    """NPV = sum( CF_t / (1 + r)^t ),  t = 0..n"""
    cf = cashflows[~np.isnan(cashflows)]
    t = np.arange(len(cf))
    return float(np.sum(cf / (1.0 + rate) ** t))


# ---------------------------------------------------------------------------
# Export
# ---------------------------------------------------------------------------
def export_results(
    teekay_results: dict,
    budget_stats: pd.DataFrame,
):
    # Ensure results directory exists
    RESULTS_DIR.mkdir(parents=True, exist_ok=True)

    # Teekay capital-budgeting stats + NPV
    teekay_parts = []
    for label, res in teekay_results.items():
        s = res["stats"].reset_index()
        s.insert(0, "Source", label)
        s["DiscountRate"] = res["rate"]
        s["NPV"] = res["npv"]
        teekay_parts.append(s)
    teekay_combined = pd.concat(teekay_parts, ignore_index=True)

    # Budget stats renamed to align columns for the CSV
    budget_csv = budget_stats.rename(
        columns={
            "Sheet": "Source",
            "Category": "Metric",
            "Average ($M)": "Average",
            "StdDev ($M)": "StdDev",
        }
    )

    # Standalone Excel workbooks for each result set
    with pd.ExcelWriter(OUTPUT_TEEKAY_XLSX, engine="openpyxl") as w_teekay:
        teekay_combined.to_excel(w_teekay, sheet_name="Teekay_Stats", index=False)
        from copy import copy as _copy_font

        sheet_t = w_teekay.book["Teekay_Stats"]
        for cell in sheet_t[1]:
            nf = _copy_font(cell.font)
            nf.bold = True
            cell.font = nf

    with pd.ExcelWriter(OUTPUT_DOD_XLSX, engine="openpyxl") as w_dod:
        budget_csv.to_excel(w_dod, sheet_name="Budget_Stats", index=False)
        from copy import copy as _copy_font2

        sheet_b = w_dod.book["Budget_Stats"]
        for cell in sheet_b[1]:
            nf = _copy_font2(cell.font)
            nf.bold = True
            cell.font = nf

    # Combined Excel with separate sheets for easy reading and bold headers
    with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as w:
        teekay_combined.to_excel(w, sheet_name="Teekay_Stats", index=False)
        budget_stats.to_excel(w, sheet_name="Budget_Stats", index=False)

        # Make header row bold in each sheet (preserve existing font settings)
        workbook = w.book
        from copy import copy

        for sheet_name in ["Teekay_Stats", "Budget_Stats"]:
            sheet = workbook[sheet_name]
            for cell in sheet[1]:
                new_font = copy(cell.font)
                new_font.bold = True
                cell.font = new_font

    print(
        f"\nResults exported to '{OUTPUT_TEEKAY_XLSX.name}', "
        f"'{OUTPUT_DOD_XLSX.name}' and '{OUTPUT_XLSX.name}'."
    )


# ---------------------------------------------------------------------------
# Data quality checks
# ---------------------------------------------------------------------------
TEEKAY_REQUIRED_LABELS = [
    "Year",
    "Charter Revenue",
    "Total Operating Costs",
    "Project After Tax Op Income",
    "Net Cash Flow",
]


def validate_data():
    """Run data-quality checks on both input files before calculations."""
    sep = "-" * 65
    issues = []
    warnings = []

    print(sep)
    print("  DATA QUALITY CHECK")
    print(sep)

    # --- File existence ---
    for path, name in [
        (BUDGET_FILE, "Budget Data (.xlsx)"),
        (CAPITAL_BUDGET_FILE, "Teekay Capital Budget (.xls)"),
    ]:
        if not path.exists():
            issues.append(f"File not found: {path}")
            print(f"  [FAIL] {name}: file missing")
        else:
            size_kb = path.stat().st_size / 1024
            print(f"  [OK]   {name}: found ({size_kb:.0f} KB)")

    if issues:
        return issues, warnings

    # --- Teekay sheets ---
    for sheet in ("Teekay_13", "Teekay_25"):
        df = load_teekay(sheet)
        print(f"\n  --- {sheet} ---")
        print(f"  Shape: {df.shape[0]} rows x {df.shape[1]} columns")

        for label in TEEKAY_REQUIRED_LABELS:
            try:
                _find_row(df, label)
                print(f"  [OK]   Row label '{label}' found")
            except KeyError:
                issues.append(f"{sheet}: missing row '{label}'")
                print(f"  [FAIL] Row label '{label}' NOT found")

        # Cost of capital row is useful but not strictly required
        try:
            r_cost = _find_row(df, "Cost of Capital =")
            rate_val = pd.to_numeric(df.iloc[r_cost, 2], errors="coerce")
            if np.isnan(rate_val):
                warnings.append(f"{sheet}: Cost of Capital row has non-numeric value")
                print(f"  [WARN] Cost of Capital = non-numeric, will use default {DISCOUNT_RATE:.0%}")
            else:
                print(f"  [OK]   Cost of Capital = {float(rate_val):.2%}")
        except KeyError:
            warnings.append(f"{sheet}: missing Cost of Capital = row, will use default {DISCOUNT_RATE:.0%}")
            print(f"  [WARN] Cost of Capital = row NOT found (using default {DISCOUNT_RATE:.0%})")

        row_year = _find_row(df, "Year")
        year_vals = df.iloc[row_year, :].apply(
            pd.to_numeric, errors="coerce"
        )
        n_years = year_vals.notna().sum()
        print(f"  [OK]   Year columns detected: {n_years}")

        for label in ["Charter Revenue", "Total Operating Costs",
                       "Project After Tax Op Income", "Net Cash Flow"]:
            try:
                r = _find_row(df, label)
            except KeyError:
                continue
            last_col = year_vals.last_valid_index()
            start = YEAR0_COL + (0 if label == "Net Cash Flow" else 1)
            series = _extract_yearly_series(df, r, start, last_col)
            total = len(series)
            nan_ct = int(np.isnan(series).sum())
            valid = total - nan_ct
            if nan_ct > 0:
                warnings.append(
                    f"{sheet}/{label}: {nan_ct} NaN out of {total} values"
                )
                print(f"  [WARN] {label}: {valid} valid, {nan_ct} NaN")
            else:
                print(f"  [OK]   {label}: {valid} valid values, 0 NaN")

    # --- Budget workbook ---
    sheets = load_budget_data()
    print(f"\n  --- Budget Data ---")
    print(f"  Sheets found: {list(sheets.keys())}")
    for sname, df in sheets.items():
        if sname == "DATA SOURCE":
            continue
        numeric_cells = df.apply(
            pd.to_numeric, errors="coerce"
        ).notna().sum().sum()
        total_cells = df.shape[0] * df.shape[1]
        pct = numeric_cells / total_cells * 100 if total_cells else 0
        print(
            f"  [OK]   {sname}: {df.shape[0]} rows x {df.shape[1]} cols, "
            f"{numeric_cells} numeric cells ({pct:.0f}%)"
        )

    # --- Summary ---
    print(f"\n{sep}")
    if issues:
        print(f"  DATA CHECK: {len(issues)} ISSUE(S) FOUND -- aborting.")
        for i in issues:
            print(f"    - {i}")
    elif warnings:
        print(
            f"  DATA CHECK: PASSED with {len(warnings)} warning(s) -- "
            "proceeding."
        )
        for w in warnings:
            print(f"    - {w}")
    else:
        print("  DATA CHECK: ALL CLEAN -- proceeding.")
    print(sep)

    return issues, warnings


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    sep = "=" * 65

    issues, _ = validate_data()
    if issues:
        raise SystemExit("Aborting due to data quality issues.")

    print()
    teekay_results = {}
    for sheet in ("Teekay_13", "Teekay_25"):
        res = analyse_teekay(sheet)
        rate = res["rate"]
        npv = calculate_npv(res["cashflows"], rate)
        res["npv"] = npv
        teekay_results[sheet] = res

        print(sep)
        print(f"  {sheet}: Capital Budgeting Analysis")
        print(sep)
        print(f"\n  Average & Std Dev (Years 1-n)\n")
        for metric in res["stats"].index:
            avg = res["stats"].loc[metric, "Average"]
            std = res["stats"].loc[metric, "StdDev"]
            print(f"    {metric:35s}  Avg = {avg:>12,.2f}   Std = {std:>10,.2f}")
        print(f"\n  Net Present Value (discount rate = {rate:.0%})")
        print(f"    NPV = {npv:,.2f}\n")

    budget_stats = analyse_budget()
    print(sep)
    print("  2021-2030 Defense Budget - Average by Category ($M)")
    print(sep)
    print()
    print(budget_stats.to_string(index=False))
    print()

    export_results(teekay_results, budget_stats)


if __name__ == "__main__":
    main()
