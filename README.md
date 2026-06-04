# Hotel Early Bird Intelligence

A Python project skeleton for ingesting daily hotel KPI Excel files, combining them into a historical SQLite database, and later building dashboard and report capabilities.

## Project Goal

- Collect daily hotel KPI files named like `Waterloo Early Bird DD.MM.YY.xlsx`
- Store historical metrics in a SQLite database
- Provide a foundation for a Streamlit dashboard and PowerPoint/PDF report generator

## Structure

- `app/` - core application modules
  - `config.py` - application settings and configuration
  - `extractor.py` - placeholder for Excel extraction logic
  - `database.py` - placeholder for SQLite database management
  - `validation.py` - placeholder for data validation rules
  - `analytics.py` - placeholder for KPI analytics functions
  - `report_generator.py` - placeholder for report generation logic
- `dashboard/` - dashboard application entrypoint
  - `streamlit_app.py` - Streamlit dashboard starter
- `data/` - data directories
  - `raw/` - raw daily Excel files
  - `processed/` - processed datasets
  - `database/` - SQLite database file
- `reports/` - generated report outputs
  - `powerpoint/`
  - `pdf/`
- `tests/` - tests and validation utilities

## Python Requirements

- Python 3.11+
- pandas
- openpyxl
- streamlit
- plotly
- python-pptx
- pydantic
- pytest

## Usage

This is an initial skeleton. Extraction, database ingestion, analytics, and report generation are not yet implemented.
