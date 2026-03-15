# Capital Budgeting Analysis

A Python project that processes capital budgeting Excel data and calculates key financial metrics such as **Average**, **Standard Deviation**, and **Net Present Value (NPV)**.

---

# Project Structure

```
capital_budgeting_project/
│
├── data/
│   ├── 2021-2030 Budget Data.xlsx
│   └── Tables_A.1_A.2_Teekay_13_25_year_capital_budgeting_plan.xls
│
├── main.py
├── teekay_irr_scenarios.py
├── requirements.txt
├── .gitignore
└── README.md
```

---

# Setup

## 1. Create and Activate Virtual Environment

### Windows (PowerShell)

```
cd E:\capital_budgeting_project
python -m venv venv
.\venv\Scripts\Activate.ps1
```

### macOS / Linux

```
cd capital_budgeting_project
python3 -m venv venv
source venv/bin/activate
```

---

## 2. Install Dependencies

```
pip install -r requirements.txt
```

---

## 3. Add Excel Data Files

Place the following files inside the **data/** folder:

* `2021-2030 Budget Data.xlsx`
* `Tables_A.1_A.2_Teekay_13_25_year_capital_budgeting_plan.xls`

---

# Usage

## 1. Full Analysis

Run the full analysis from the project root:

```
python main.py
```

### What `main.py` Does

* Loads Excel files from the **data/** directory
* Validates the Teekay 13-year and 25-year sheets
* Calculates **Average** and **Standard Deviation** for:

  * Charter Revenue
  * Total Operating Costs
  * Project After-Tax Operating Income
* Detects the **Cost of Capital (%)** (default 7% if not found)
* Calculates **Net Present Value (NPV)** for both projects
* Processes **2021-2030 budget data**
* Generates statistical summaries
* Exports results to Excel files in the **results/** folder

### Output Files

* `teekay_capital_budgeting_stats_and_npv.xlsx`
* `dod_budget_category_stats.xlsx`
* `capital_budgeting_analysis.xlsx`

**Note:** If any Excel file is open while running the script, Windows may raise a **PermissionError**. Close the file and run the script again.

---

# IRR Scenario Analysis

Run the standalone IRR scenarios script:

```
python teekay_irr_scenarios.py
```

### What This Script Does

* Loads the **Teekay 13-year sheet**
* Extracts **Net Cash Flow**
* Calculates **Internal Rate of Return (IRR)**
* Runs multiple **cost-of-capital scenarios**

### Scenarios

| Scenario   | Cost of Capital |
| ---------- | --------------- |
| Scenario 1 | 7%              |
| Scenario 2 | 5%              |
| Scenario 3 | 2%              |

For each scenario the script calculates:

* IRR
* NPV at that cost of capital

### Output

Exports a formatted Excel file:

```
results/teekay_irr_scenarios.xlsx
```

---

# Dependencies

| Package  | Purpose                           |
| -------- | --------------------------------- |
| pandas   | Data manipulation and analysis    |
| numpy    | Mathematical calculations         |
| openpyxl | Reading and writing `.xlsx` files |
| xlrd     | Reading legacy `.xls` files       |
| scipy    | Advanced numerical calculations   |
