from pathlib import Path
import sys

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import streamlit as st
import pandas as pd
import plotly.express as px

from app.config import RAW_DATA_DIR
from app.database import load_daily_kpis
from app.extractor import run_pipeline


def format_currency(v):
    try:
        return f"£{float(v):,.2f}"
    except Exception:
        return "-"


def format_percent(v, decimals: int = 1):
    try:
        return f"{float(v):.{decimals}f}%"
    except Exception:
        return "-"


def main() -> None:
    st.set_page_config(page_title="Hotel Early Bird Intelligence", layout="wide")

    st.title("Hotel Early Bird Intelligence")
    st.write("Historical KPI view from daily Early Bird Excel reports.")

    # Sidebar data controls
    with st.sidebar:
        st.header("Data controls")
        
        # Excel source folder path
        folder_path = st.text_input(
            "Excel source folder path",
            value=str(RAW_DATA_DIR),
            help="Path to folder containing Excel files"
        )
        
        uploaded_files = st.file_uploader(
            "Upload Early Bird Excel files",
            type=["xlsx"],
            accept_multiple_files=True,
        )
        if uploaded_files:
            RAW_DATA_DIR.mkdir(parents=True, exist_ok=True)
            for uploaded_file in uploaded_files:
                target_path = RAW_DATA_DIR / uploaded_file.name
                target_path.write_bytes(uploaded_file.read())
            st.success(f"Uploaded {len(uploaded_files)} file(s) to {RAW_DATA_DIR}")
            st.info("After uploading, click Refresh data from Excel files.")

        if st.button("Refresh data from Excel files"):
            # Validate folder exists
            selected_folder = Path(folder_path)
            if not selected_folder.exists():
                st.error(f"Folder does not exist: {folder_path}")
            elif not selected_folder.is_dir():
                st.error(f"Path is not a directory: {folder_path}")
            else:
                with st.spinner("Refreshing data from Excel files..."):
                    try:
                        summary = run_pipeline(raw_dir=selected_folder)
                        st.success(f"Processed {summary['records_processed']} records")
                        st.write("Validation status counts:")
                        st.write(summary.get("validation_status_counts", {}))
                    except Exception as e:
                        st.error(f"Refresh failed: {e}")

    # Load data
    try:
        df = load_daily_kpis()
    except Exception:
        st.warning("No database found. Run `python -m app.extractor` to create the master data.")
        return

    if df is None or df.empty:
        st.warning("No data available. Run `python -m app.extractor` to populate the database.")
        return

    # Prepare data
    if "business_date" in df.columns:
        df["business_date"] = pd.to_datetime(df["business_date"])
    df = df.sort_values("business_date").reset_index(drop=True)

    # derived columns
    if "daily_occupancy" in df.columns:
        df["daily_occupancy_pct"] = pd.to_numeric(df["daily_occupancy"], errors="coerce") * 100

    # Header
    st.header("Hotel Early Bird Intelligence Dashboard")
    st.markdown("_Historical KPI view from daily Early Bird Excel reports._")

    # KPI cards
    latest = df.iloc[-1]
    total_reports = len(df)
    date_min = df['business_date'].min()
    date_max = df['business_date'].max()
    date_range_display = f"Date range: {date_min.strftime('%Y-%m-%d')} → {date_max.strftime('%Y-%m-%d')}"

    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric("Total reports", total_reports)
    # Use markdown for date range to avoid truncation issues with st.metric
    col2.markdown(date_range_display)
    col3.metric("Latest rooms sold", latest.get("daily_rooms_sold", "-"))

    occ = latest.get("daily_occupancy_pct") if "daily_occupancy_pct" in latest.index else None
    occ_display = format_percent(occ) if occ is not None and pd.notna(occ) else "-"
    col4.metric("Latest occupancy", occ_display)

    adr_display = format_currency(latest.get("daily_average_rate"))
    col5.metric("Latest ADR", adr_display)

    # Latest report table (friendly labels)
    st.subheader("Latest report details")
    latest_row = df.loc[df["business_date"] == df["business_date"].max()].tail(1)
    if not latest_row.empty:
        row = latest_row.iloc[0]
        friendly = {
            "Business Date": row.get("business_date").strftime("%Y-%m-%d") if pd.notna(row.get("business_date")) else "-",
            "Rooms Sold": int(row.get("daily_rooms_sold")) if pd.notna(row.get("daily_rooms_sold")) else "-",
            "Occupancy %": format_percent(row.get("daily_occupancy") * 100) if pd.notna(row.get("daily_occupancy")) else "-",
            "ADR £": format_currency(row.get("daily_average_rate")),
            "Total Revenue £": format_currency(row.get("revenue_mtd_total_revenue")),
            "Rooms Revenue £": format_currency(row.get("revenue_stats_rooms_revenue")),
            "Bar Revenue £": format_currency(row.get("revenue_mtd_bar")),
        }
        st.table(pd.DataFrame([friendly]))
    else:
        st.write("No latest report available")

    # Trend charts
    st.subheader("Trends")
    charts_col1, charts_col2 = st.columns(2)

    if "daily_rooms_sold" in df.columns:
        fig_rooms = px.line(df, x="business_date", y="daily_rooms_sold", title="Daily Rooms Sold")
        fig_rooms.update_layout(yaxis_title="Rooms Sold")
        charts_col1.plotly_chart(fig_rooms, use_container_width=True)

    if "daily_occupancy_pct" in df.columns:
        fig_occ = px.line(df, x="business_date", y="daily_occupancy_pct", title="Daily Occupancy")
        fig_occ.update_layout(yaxis_title="Occupancy (%)")
        charts_col1.plotly_chart(fig_occ, use_container_width=True)

    if "daily_average_rate" in df.columns:
        fig_adr = px.line(df, x="business_date", y="daily_average_rate", title="Daily Average Rate")
        fig_adr.update_layout(yaxis_title="ADR (£)")
        charts_col2.plotly_chart(fig_adr, use_container_width=True)

    if "revenue_mtd_total_revenue" in df.columns:
        fig_rev = px.line(df, x="business_date", y="revenue_mtd_total_revenue", title="MTD Total Revenue")
        fig_rev.update_layout(yaxis_title="Revenue (£)")
        charts_col2.plotly_chart(fig_rev, use_container_width=True)

    # Data Validation section
    st.subheader("Data Validation")
    val_path = Path("data") / "processed" / "early_bird_validation.csv"
    if val_path.exists():
        vdf = pd.read_csv(val_path)
        status_counts = vdf["status"].value_counts().to_dict()
        st.write("Validation status counts:")
        st.write(status_counts)

        problems = vdf[vdf["status"].isin(["WARN", "FAIL"])]
        if not problems.empty:
            st.markdown("**Warnings and Failures:**")
            st.dataframe(problems[["business_date", "file_name", "check_name", "status", "message"]])
        else:
            st.success("No WARN or FAIL validations found.")
    else:
        st.info("Validation file not found. Run `python -m app.extractor` to produce validation results.")

    # Full data table
    st.subheader("Full data")
    st.dataframe(df)


if __name__ == "__main__":
    main()
