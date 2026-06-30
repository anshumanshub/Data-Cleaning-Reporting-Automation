# Data Cleaning & Reporting Automation

A Python tool that automates the messy parts of preparing raw data: deduplication,
missing-value handling, date/phone standardization, outlier detection, and a
polished multi-sheet Excel report with charts.

## Project structure

```
data_cleaning_project/
├── cleaner.py                 # Core pipeline (DataCleaner class + CLI)
├── generate_sample_data.py    # Creates a messy demo CSV
├── requirements.txt
├── input/
│   └── sales_data.csv         # Demo input (generated)
└── output/
    └── cleaning_report.xlsx   # Generated report
```

## What it does

1. **Ingestion** – reads CSV, TSV, or Excel files.
2. **Duplicate removal** – drops exact duplicate rows.
3. **Missing values** – numeric columns filled with median, categorical
   columns filled with mode; every fill is logged.
4. **Date standardization** – auto-detects date-like columns (by name) and
   normalizes mixed formats (`MM/DD/YYYY`, `DD-Mon-YYYY`, etc.) to `YYYY-MM-DD`.
5. **Phone standardization** – auto-detects phone-like columns and reformats
   10/11-digit numbers to `+1 (XXX) XXX-XXXX`.
6. **Outlier detection** – flags numeric outliers using the IQR method
   (1.5× interquartile range) and adds a `<column>_outlier` flag.
7. **Data quality scoring** – a 0–100 composite score (60% completeness,
   40% uniqueness) computed before and after cleaning.
8. **Excel report** – a workbook with 4 sheets:
   - **Summary** – key metrics, quality score, imputation detail
   - **Cleaned Data** – the full cleaned dataset
   - **Outliers & Flags** – only the rows that were flagged, highlighted
   - **Charts** – before/after quality bar chart + issue-breakdown pie chart

## Usage

```bash
pip install -r requirements.txt

# (optional) generate a messy demo dataset
python generate_sample_data.py

# run the pipeline on any CSV/TSV/Excel file
python cleaner.py input/sales_data.csv --output output/cleaning_report.xlsx
```

### Programmatic use

```python
from cleaner import DataCleaner

dc = DataCleaner("input/sales_data.csv")
dc.load()
dc.run_pipeline()
dc.export_report("output/cleaning_report.xlsx")

print(dc.log.quality_score_before, dc.log.quality_score_after)
```

## Extending it

- Swap the imputation strategy: `dc.handle_missing_values(numeric_strategy="mean")`
- Adjust outlier sensitivity: `dc.flag_outliers_iqr(k=3.0)` (more lenient)
- Add new column-detection hints by editing `DATE_COL_HINTS` / `PHONE_COL_HINTS`
  in `cleaner.py`
- For scheduled automation, wrap `cleaner.py` in a cron job or Task Scheduler
  entry pointed at a watched input folder

## Notes

- Date/phone column detection is name-based (looks for words like `date`,
  `created`, `phone`, `mobile`, etc.) — rename ambiguous columns if detection
  misses them.
- The quality score is a heuristic, not a statistical guarantee; tune the
  weighting in `DataCleaner._quality_score` if your use case needs a
  different definition of "quality."
