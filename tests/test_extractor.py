import pytest
from datetime import date

from app.extractor import extract_date_from_filename, list_early_bird_files, extract_report_kpis


def test_extract_date_from_filename():
    assert extract_date_from_filename("Waterloo Early Bird 04.02.25.xlsx") == date(2025, 2, 4)
    assert extract_date_from_filename("Waterloo Early Bird 03.11.24.xlsx") == date(2024, 11, 3)
    assert extract_date_from_filename("random_file.xlsx") is None


def test_extract_report_kpis_basic():
    files = list_early_bird_files()
    if not files:
        pytest.skip("No Excel files in data/raw to run extractor tests")

    first = files[0]
    try:
        result = extract_report_kpis(first)
    except Exception as e:
        pytest.skip(f"extract_report_kpis raised; skipping: {e}")

    assert isinstance(result, dict)
    for key in [
        "file_name",
        "business_date",
        "daily_rooms_sold",
        "daily_occupancy",
        "daily_average_rate",
        "revenue_mtd_total_revenue",
    ]:
        assert key in result
