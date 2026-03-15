# Capital Budgeting Analysis

A Python project that processes capital budgeting Excel data and calculates key financial metrics: **Average**, **Standard Deviation**, and **Net Present Value (NPV)**.

## Project Structure

```
capital_budgeting_project/
│
├── data/
│   ├── 2021-2030 Budget Data.xlsx
│   └── Tables_A.1_A.2_Teekay_13_25_year_capital_budgeting_plan.xls
│
├── main.py                  # Full capital budgeting + DoD analysis
├── teekay_irr_scenarios.py  # Standalone IRR scenarios for Teekay 13-year project
├── requirements.txt         # Dependencies for the project
├── .gitignore               # To exclude unnecessary files
└── README.md                # Project description and usage instructions
```

## Setup

### 1. Create and activate the virtual environment

**Windows (PowerShell)**

```powershell
cd E:\capital_budgeting_project
python -m venv venv
.\venv\Scripts\Activate.ps1
```

**macOS / Linux**

```bash
cd capital_budgeting_project
python3 -m venv venv
source venv/bin/activate
```

### 2. Install dependencies

```bash
pip install -r requirements.txt
```

### 3. Add Excel data files

Place the following files inside the `data/` folder:

- `2021-2030 Budget Data.xlsx`
- `Tables_A.1_A.2_Teekay_13_25_year_capital_budgeting_plan.xls`

## Usage

### 1. Full analysis (`main.py`)

Run the full analysis from the project root:

```bash
python main.py
```

### What `main.py` does

1. **Loads** both Excel files from the `data/` directory.
2. **Validates** the Teekay 13-year and 25-year sheets (row labels, year columns, numeric values).
3. **Calculates** the average and sample standard deviation of:
   - Charter Revenue
   - Total Operating Costs
   - Project After-Tax Operating Income
   for both the 13-year and 25-year Teekay plans.
4. **Reads** the cost of capital (%), if present, from the Teekay sheets (falls back to **7%** if not found).
5. **Calculates** the Net Present Value (NPV) of the Teekay 13-year and 25-year net cash flows at the detected cost of capital.
6. **Summarises** the 2021-2030 DoD budget data (average and standard deviation across fiscal years for every budget category in each sheet).
7. **Exports** the results to the `results/` folder as Excel workbooks:
   - `teekay_capital_budgeting_stats_and_npv.xlsx` (Teekay stats with bold column headers)
   - `dod_budget_category_stats.xlsx` (DoD budget stats with bold column headers)
   - `capital_budgeting_analysis.xlsx` (combined workbook with separate sheets for Teekay stats and DoD budget stats, headers bold)

> **Note:** If a CSV or the Excel file is open in Excel while you run `python main.py`, Windows may raise a *PermissionError*. Close the file and run the script again.

### 2. Standalone IRR scenarios (`teekay_irr_scenarios.py`)

Run only the Teekay 13-year IRR scenarios analysis:

```bash
python teekay_irr_scenarios.py
```

What this script does:

1. **Loads** the Teekay 13-year sheet from `Tables_A.1_A.2_Teekay_13_25_year_capital_budgeting_plan.xls`.
2. **Extracts** the `Net Cash Flow` row, drops any empty / non-numeric cells, and uses the resulting cash-flow series as the 13-year project cash flows.
3. **Computes** the internal rate of return (**IRR**) once, holding all cash flows constant.
4. **Loops** over three cost-of-capital scenarios using a conditional structure:
   - Scenario #1: Cost of capital **7%**
   - Scenario #2: Cost of capital **5%**
   - Scenario #3: Cost of capital **2%**
5. **Calculates** the NPV of the same cash flows at each scenario’s cost of capital.
6. **Prints** a clean table to the console with:
   - `Scenario`
   - `CostOfCapital`
   - `IRR`
   - `NPV_at_CostOfCapital`
7. **Exports** the IRR scenarios table to the `results/` folder as a formatted Excel file:
   - `teekay_irr_scenarios.xlsx`

## Results Summary

## Dependencies

| Package  | Purpose                                  |
|----------|------------------------------------------|
| pandas   | Data manipulation and analysis           |
| numpy    | Mathematical calculations                |
| openpyxl | Reading / writing `.xlsx` files          |
| xlrd     | Reading legacy `.xls` files              |
| scipy    | Advanced calculations (if needed)        |


