# Students Performance Dashboard

A small, reproducible pipeline that builds a two‑sheet Excel dashboard from `StudentsPerformance.csv` using pandas, matplotlib, and numpy (and openpyxl for Excel output with optional embedded charts).

## What it produces

- `Students_Performance_Dashboard.xlsx`
  - Sheet `Cleaned_Data`
    - Original data with normalized column names
    - Derived `Average` (mean of Math, Reading, Writing) and `Performance_Band`
  - Sheet `Summary`
    - Data quality: duplicate count, per‑column missing values
    - Aggregations: Average scores by Gender, Race/Ethnicity, Test Prep
    - Top lists: Top 20 by Average, Top 20 by Math
    - Charts: Average by Gender, Average by Race/Ethnicity, Scores by Test Prep, Performance Band Distribution
      - Embedded if Pillow is available; otherwise image files are saved under `plots/`

## Input

- `StudentsPerformance.csv`
  - Expected columns (case/spacing normalized by the script):
    - `gender` → `Gender`
    - `race/ethnicity` → `Race_Ethnicity`
    - `parental level of education` → `Parental_Education`
    - `lunch` → `Lunch`
    - `test preparation course` → `Test_Prep`
    - `math score` → `Math`
    - `reading score` → `Reading`
    - `writing score` → `Writing`

## How to run

1. (Optional) Create and activate a virtual environment.
2. Install dependencies:
   - `pip install pandas numpy matplotlib openpyxl pillow`
     - `pillow` is only needed if you want charts embedded into the Excel workbook.
3. Execute the build script:
   - `python build_students_dashboard.py`
4. The workbook will be written to:
   - `Students_Performance_Dashboard.xlsx`
5. Charts are saved to the `plots/` directory and embedded into the Summary sheet when possible.

## Data processing steps

- Null checks and filling:
  - Numeric (`Math`, `Reading`, `Writing`): fill with column mean (only if nulls exist)
  - Categorical: fill with column mode (only if nulls exist)
- Derived metrics:
  - `Average` = mean of the three scores (rounded to 2 decimals)
  - `Performance_Band` buckets: `Low` (<50), `Fair` (50–70), `Good` (70–85), `Excellent` (≥85)
- Filtering & sorting examples:
  - Top 20 by `Average` and by `Math`
- Aggregations:
  - Average scores by `Gender`, `Race_Ethnicity`, and `Test_Prep`

## Customization

You can adjust behavior in `build_students_dashboard.py`:
- Change performance bands (`bins` and `labels`)
- Modify which “Top N” listings are exported
- Add or remove groupings and charts

## Troubleshooting

- If charts do not appear in the Excel `Summary` sheet, ensure `pillow` is installed and that Excel supports embedded images. Images are always saved in `plots/` regardless.
- If the script cannot find the CSV, confirm that `StudentsPerformance.csv` is in the same folder as `build_students_dashboard.py`.

## Files in this folder

- `build_students_dashboard.py` — main script
- `StudentsPerformance.csv` — input data
- `Students_Performance_Dashboard.xlsx` — generated dashboard (two sheets)
- `plots/` — PNG charts used by the Summary sheet
