import pandas as pd
import pytest

from app.validation import validate_records


def test_validate_records_basic():
    # Build a simple valid row
    df = pd.DataFrame([
        {
            "business_date": pd.to_datetime("2025-02-04"),
            "file_name": "Waterloo Early Bird 04.02.25.xlsx",
            "daily_rooms_sold": 10,
            "daily_occupancy": 0.75,
            "daily_average_rate": 120.0,
            "revenue_mtd_total_revenue": 5000.0,
            "revenue_stats_rooms_revenue": 1200.0,
        }
    ])

    results = validate_records(df)
    assert isinstance(results, pd.DataFrame)

    expected_cols = {"business_date", "file_name", "check_name", "status", "message"}
    assert expected_cols.issubset(set(results.columns))

    # There should be at least one PASS
    assert "PASS" in set(results["status"].values)
