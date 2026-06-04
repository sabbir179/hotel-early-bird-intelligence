# Hotel Early Bird Intelligence

A local-first analytics dashboard for combining daily hotel Early Bird Excel reports into a historical KPI database, dashboard, validation report, and PowerPoint summary.

## Problem statement

Daily Early Bird reports are delivered as individual Excel files. Management can review one day at a time, but there is no easy way to view historical performance or compare consecutive reports.

## Solution

This app combines multiple Excel files into a single master dataset and exposes the results through a Streamlit dashboard. It supports data validation, archive storage in SQLite, and PowerPoint summary generation for management review.

## Features

- Upload Early Bird Excel files through the dashboard
- Read Excel files from a configurable folder path
- Extract daily KPIs from each report
- Build a master CSV dataset
- Build a SQLite database for historical analysis
- Validate data quality and report warning/failure counts
- Streamlit dashboard with charts and KPI cards
- Management summary section for quick insights
- PowerPoint report generation and download
- Pytest tests for core functionality

## Data privacy

- Raw Excel files are not committed to GitHub
- Processed CSV files are not committed
- SQLite database files are not committed
- Generated reports are not committed
- The app is designed to run locally so data stays under user control

## Project structure

- `app/` - core application logic
- `dashboard/` - Streamlit user interface
- `data/` - local data storage
  - `raw/` - raw Excel source files
  - `processed/` - generated CSV outputs
  - `database/` - SQLite database files
- `reports/` - generated report outputs
  - `powerpoint/`
  - `pdf/`
- `tests/` - automated tests

## Setup

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Run pipeline

```bash
python -m app.extractor
```

## Run dashboard

```bash
streamlit run dashboard/streamlit_app.py
```

## Run tests

```bash
python -m pytest
```

## GitHub safety note

This repository uses `.gitignore` to protect sensitive local data and report files. The following directories are intended to stay local and not be committed:

- `data/raw`
- `data/processed`
- `data/database`
- `reports`

## Roadmap

- Better UI design and dashboard polish
- PDF export support
- Optional AI-generated executive summary using aggregated KPI data only
- Scheduled refresh for recurring ingestion
- Authentication for production deployment
