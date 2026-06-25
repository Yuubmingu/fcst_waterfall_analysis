# FCST Waterfall 변동 분석 자동화

매주 업데이트되는 FCST 데이터를 기반으로 고객·모델·제품별 계획 수량을 Waterfall 형태로 펼쳐 보고, 전주 대비 수요 변동을 자동 계산하여 Excel 리포트로 생성하는 Python 기반 분석 프로젝트입니다.

## Overview

This project automates weekly FCST Waterfall analysis using Python, pandas, and openpyxl.

The tool reads rolling forecast data, converts it into a horizontal Waterfall view, calculates week-over-week demand changes, and generates a formatted Excel report for supply chain, procurement, and production planning teams.

The main purpose of this project is to quickly identify demand changes, volatile products, and urgent planning risks from weekly forecast updates.

## Business Context

In many supply chain environments, customers update their forecasts every week.

For example:

* The forecast issued in `2026-W22` may include plans from `W22` to `W34`.
* The forecast issued in `2026-W23` may include plans from `W23` to `W34`.
* The forecast issued in `2026-W24` may include plans from `W24` to `W34`.

As new forecast versions are received, the same production week may change multiple times.

This project helps users answer questions such as:

* Which product demand increased compared to the previous forecast?
* Which product demand decreased sharply?
* Which customer or model has the highest forecast volatility?
* What changed in the latest forecast compared to the previous one?
* Which items require urgent review from SCM, procurement, or production planning teams?

## Input File

The input Excel file should be located at:

```text
/kaggle/working/fcst_waterfall_w22_w34_sample.xlsx
```

Required columns:

| Column            | Description                        |
| ----------------- | ---------------------------------- |
| `fcst_issue_week` | Week when the forecast was issued  |
| `plan_week`       | Production or demand planning week |
| `customer`        | Customer name                      |
| `model`           | Model name                         |
| `product_code`    | Product code                       |
| `qty`             | Forecast quantity                  |

Example:

| fcst_issue_week | plan_week | customer   | model       | product_code | qty |
| --------------- | --------- | ---------- | ----------- | ------------ | --: |
| 2026-W22        | 2026-W22  | Customer A | Model Alpha | PROD-A       | 300 |
| 2026-W22        | 2026-W23  | Customer A | Model Alpha | PROD-A       | 320 |
| 2026-W23        | 2026-W23  | Customer A | Model Alpha | PROD-A       | 350 |

## Output File

The final Excel report is saved as:

```text
/kaggle/working/fcst_waterfall_analysis.xlsx
```

## Output Sheets

The generated Excel report contains the following sheets:

| Sheet Name               | Description                                                                            |
| ------------------------ | -------------------------------------------------------------------------------------- |
| `fcst_waterfall_view`    | Horizontal Waterfall view of forecast quantities by issue week and plan week           |
| `change_waterfall_view`  | Horizontal Waterfall view of week-over-week forecast quantity changes                  |
| `fcst_change_detail`     | Detailed comparison table between current and previous forecast quantities             |
| `top_volatility_items`   | Ranking of products with the highest forecast volatility                               |
| `latest_vs_previous`     | Comparison between the latest forecast issue week and the previous forecast issue week |
| `customer_model_summary` | Customer and model level summary of latest forecast changes                            |
| `parameter_info`         | Execution parameters, input/output paths, thresholds, and analysis metadata            |

## Key Features

### 1. Weekly Forecast Waterfall View

The `fcst_waterfall_view` sheet converts rolling forecast data into a horizontal structure.

Example:

| customer   | model       | product_code | fcst_issue_week | W22 | W23 | W24 | W25 |
| ---------- | ----------- | ------------ | --------------- | --: | --: | --: | --: |
| Customer A | Model Alpha | PROD-A       | 2026-W22        | 300 | 320 | 310 | 340 |
| Customer A | Model Alpha | PROD-A       | 2026-W23        |     | 350 | 330 | 360 |
| Customer A | Model Alpha | PROD-A       | 2026-W24        |     |     | 370 | 390 |

This layout helps users visually track how forecast plans shift over time.

### 2. Week-over-Week Change Analysis

The tool compares forecast quantities using the same:

* Customer
* Model
* Product code
* Plan week

Calculation logic:

```text
change_qty = current_qty - previous_qty
```

```text
change_rate = change_qty / previous_qty
```

If there is no previous forecast value, the row is treated as `New`.

### 3. Change Direction Classification

Each change is classified as one of the following:

| Direction   | Rule                        |
| ----------- | --------------------------- |
| `Increase`  | `change_qty > 0`            |
| `Decrease`  | `change_qty < 0`            |
| `No Change` | `change_qty == 0`           |
| `New`       | No previous forecast exists |

### 4. Change Flag Logic

A change is flagged when at least one of the following conditions is met:

```python
abs(change_rate) >= 0.20
```

or

```python
abs(change_qty) >= 100
```

Default thresholds:

```python
CHANGE_RATE_THRESHOLD = 0.20
CHANGE_QTY_THRESHOLD = 100
```

These values can be changed at the top of the code.

### 5. Volatility Ranking

The `top_volatility_items` sheet summarizes forecast volatility by customer, model, and product.

It includes:

* Change count
* Increase count
* Decrease count
* Total absolute change quantity
* Average absolute change rate
* Maximum absolute change rate
* Volatility rank

Sorting priority:

1. `change_count` descending
2. `total_abs_change_qty` descending
3. `max_abs_change_rate` descending

### 6. Latest vs Previous Forecast Comparison

The `latest_vs_previous` sheet compares only the latest forecast issue week with the immediately previous forecast issue week.

This sheet is intentionally separated from `change_waterfall_view`.

* `change_waterfall_view` shows the full forecast change history.
* `latest_vs_previous` focuses only on the newest forecast update.

This prevents duplicated analysis and keeps each sheet’s purpose clear.

### 7. Excel Formatting

The output Excel file is formatted using openpyxl.

Applied formatting includes:

* Bold header row
* Filter on the first row
* Frozen header row
* Auto-adjusted column widths
* Thousands separator for quantity columns
* Percentage format for rate columns
* Light blue fill for forecast quantity cells
* Red fill for demand increases
* Blue fill for demand decreases
* Yellow highlight for significant changes
* Highlighting for top volatility items

## Project Structure

```text
fcst-waterfall-analysis/
│
├── fcst_waterfall_analysis.py
├── fcst_waterfall_w22_w34_sample.xlsx
├── fcst_waterfall_analysis.xlsx
└── README.md
```

## How to Run

### 1. Prepare the input file

Place the input file in the Kaggle working directory:

```text
/kaggle/working/fcst_waterfall_w22_w34_sample.xlsx
```

### 2. Run the Python script

Execute the analysis code in Kaggle Notebook or another Python environment.

Required libraries:

```python
import pandas as pd
import numpy as np
import openpyxl
```

### 3. Check the output file

After execution, the Excel report will be created at:

```text
/kaggle/working/fcst_waterfall_analysis.xlsx
```

## Sample Data Logic

The sample data is generated using rolling forecast issue weeks from `W22` to `W34`.

For each forecast issue week, the sample creates forecast rows from the issue week through `W34`.

Example:

| fcst_issue_week | Plan Week Coverage |
| --------------- | ------------------ |
| 2026-W22        | W22 ~ W34          |
| 2026-W23        | W23 ~ W34          |
| 2026-W24        | W24 ~ W34          |
| 2026-W34        | W34 only           |

This structure reflects a rolling forecast where the remaining planning horizon becomes shorter as the issue week moves forward.

## Validation Logic

The code includes validation checks for:

### Required Columns

The script checks whether all required columns exist.

If a required column is missing, it raises an error:

```text
필수 컬럼이 없습니다: missing_column_name
```

### Week Parsing

The code supports multiple week formats:

* `2026-W22`
* `2026-W02`
* `W22`
* `W2`
* `22W`
* `202622`

The parsing logic converts week values into:

* Year
* Week number
* Sort key
* Display label

### Waterfall Structure

The code verifies that forecast values do not appear before the forecast issue week.

For example:

* `2026-W22` starts from `W22`
* `2026-W23` starts from `W23`
* `2026-W24` starts from `W24`

### Change Calculation

The script samples several rows and validates:

```text
change_qty = current_qty - previous_qty
```

```text
change_rate = change_qty / previous_qty
```

## Example Use Cases

This project can be used for:

* Weekly S&OP forecast review
* Demand volatility monitoring
* Customer forecast change tracking
* Procurement priority setting
* Production planning risk review
* Supply chain reporting automation
* SCM portfolio demonstration

## Why This Project Matters

Forecast changes directly affect procurement, inventory, production planning, and supplier communication.

Manual Excel comparison is time-consuming and error-prone, especially when forecast versions are updated every week.

This project reduces repetitive manual work by automatically:

* Reshaping forecast data
* Comparing forecast versions
* Detecting significant changes
* Ranking volatile items
* Creating a formatted Excel report

## Tech Stack

| Tool            | Purpose                            |
| --------------- | ---------------------------------- |
| Python          | Main programming language          |
| pandas          | Data processing and transformation |
| numpy           | Numeric calculation                |
| openpyxl        | Excel formatting                   |
| Excel           | Final report output                |
| Kaggle Notebook | Execution environment              |

## Limitations

This project currently assumes:

* Forecast data is provided in weekly format.
* Forecast issue weeks and plan weeks are expressed as week-based labels.
* The same customer, model, product, and plan week combination is used for comparison.
* The latest forecast comparison only compares the latest issue week with the immediately previous issue week.
* The default week display can be configured as either the actual data range or `W01~W52`.

## Future Improvements

Possible improvements include:

* Adding a dashboard summary sheet
* Adding charts for top volatility products
* Creating customer-level forecast stability scores
* Supporting monthly forecast formats
* Adding supplier and material impact analysis
* Connecting forecast changes to BOM-based material demand
* Exporting a Power BI-ready dataset
* Creating an automated email report

## Portfolio Value

This project demonstrates practical SCM automation skills, including:

* Rolling forecast data handling
* Waterfall-style demand visualization
* Forecast revision analysis
* Excel report automation
* Business-oriented data transformation
* Procurement and supply chain decision support

It is especially relevant for roles in:

* Supply Chain Management
* Procurement
* Production Planning
* S&OP
* Business Operations
* Demand Planning
* SCM Data Analytics

## License

This project is intended for personal portfolio and educational use.

## Author

Created as a practical SCM automation project using Python, pandas, and Excel.
