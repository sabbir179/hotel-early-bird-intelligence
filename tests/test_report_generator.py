import pandas as pd

from app.report_generator import generate_powerpoint_report


def test_generate_powerpoint_report_creates_pptx_file(tmp_path):
    df = pd.DataFrame([
        {
            "business_date": "2025-07-14",
            "daily_rooms_sold": 120,
            "daily_occupancy": 0.75,
            "daily_occupancy_pct": 75.0,
            "daily_average_rate": 150.0,
            "revenue_mtd_total_revenue": 18000.0,
            "revenue_stats_rooms_revenue": 12000.0,
            "revenue_mtd_bar": 3000.0,
        },
        {
            "business_date": "2025-07-13",
            "daily_rooms_sold": 110,
            "daily_occupancy": 0.70,
            "daily_occupancy_pct": 70.0,
            "daily_average_rate": 145.0,
            "revenue_mtd_total_revenue": 16000.0,
            "revenue_stats_rooms_revenue": 11000.0,
            "revenue_mtd_bar": 2800.0,
        },
    ])

    output_file = tmp_path / "test_report.pptx"
    result_path = generate_powerpoint_report(
        df,
        validation_counts={"PASS": 10},
        output_path=output_file,
    )

    assert result_path.exists()
    assert result_path.suffix == ".pptx"
    assert result_path.stat().st_size > 0
