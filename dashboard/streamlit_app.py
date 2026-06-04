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
from app.report_generator import generate_powerpoint_report


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


def format_delta(value, prefix: str = "", suffix: str = "", decimals: int = 1):
    try:
        delta = float(value)
    except Exception:
        return "-"

    if delta > 0:
        return f"▲ {prefix}{abs(delta):.{decimals}f}{suffix}"
    if delta < 0:
        return f"▼ {prefix}{abs(delta):.{decimals}f}{suffix}"
    return f"– {prefix}{abs(delta):.{decimals}f}{suffix}"


def format_delta_html(value, prefix: str = "", suffix: str = "", invert: bool = False, decimals: int = 1):
    try:
        delta = float(value)
    except Exception:
        return "-"

    if invert:
        delta = -delta

    if delta > 0:
        return f"<span class='delta-positive'>▲ {prefix}{abs(delta):.{decimals}f}{suffix}</span>"
    if delta < 0:
        return f"<span class='delta-negative'>▼ {prefix}{abs(delta):.{decimals}f}{suffix}</span>"
    return f"<span class='delta-neutral'>– {prefix}{abs(delta):.{decimals}f}{suffix}</span>"


def format_card_value(value):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return "-"
    return str(value)


def main() -> None:
    st.set_page_config(page_title="Hotel Early Bird Intelligence", layout="wide")

    st.markdown(
        """
        <style>
            .page-title h1 { margin: 0; color: #0f172a; font-size: 2.6rem; }
            .page-title p { margin: 0.4rem 0 1.6rem; color: #475569; font-size: 1.05rem; }
            .page-note { color: #64748b; margin-bottom: 24px; }
            .kpi-row { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 16px; margin-bottom: 24px; }
            .kpi-card { background: #ffffff; border: 1px solid #e2e8f0; border-radius: 18px; padding: 20px 18px; box-shadow: 0 10px 30px rgba(15, 23, 42, 0.05); }
            .kpi-label { color: #64748b; font-size: 0.90rem; margin-bottom: 8px; }
            .kpi-value { color: #0f172a; font-size: 1.6rem; font-weight: 700; }
            .section-card { background: #ffffff; border: 1px solid #e2e8f0; border-radius: 18px; padding: 24px; box-shadow: 0 10px 30px rgba(15, 23, 42, 0.06); margin-bottom: 24px; }
            .section-title { color: #0f172a; font-size: 1.3rem; font-weight: 700; margin-bottom: 16px; }
            .bullet-list { padding-left: 20px; color: #334155; }
            .bullet-list li { margin-bottom: 10px; line-height: 1.6; }
            .status-badge { display: inline-flex; align-items: center; margin-bottom: 12px; padding: 8px 14px; border-radius: 999px; font-size: 0.95rem; font-weight: 600; }
            .status-success { background: #ecfdf5; color: #0f766e; }
            .status-warning { background: #ffedd5; color: #92400e; }
            .summary-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 16px; margin-bottom: 24px; }
            .summary-card { background: #ffffff; border: 1px solid #e2e8f0; border-radius: 18px; padding: 20px; }
            .summary-card-title { color: #0f172a; font-size: 1rem; font-weight: 700; margin-bottom: 12px; }
            .summary-row { display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px; font-size: 0.95rem; }
            .summary-label { color: #475569; }
            .summary-value { color: #0f172a; font-weight: 700; }
            .delta-positive { color: #16a34a; font-weight: 700; }
            .delta-negative { color: #dc2626; font-weight: 700; }
            .delta-neutral { color: #475569; font-weight: 700; }
            .sidebar hr { margin: 18px 0; }
            .sidebar .stTextInput > div, .sidebar .stButton > button { width: 100%; }
            .sidebar .stButton > button { margin-top: 10px; }
        </style>
        <div class="page-title">
            <h1>Hotel Early Bird Intelligence</h1>
            <p>Historical KPI view from daily Early Bird Excel reports.</p>
        </div>
        <div class="page-note">Use the sidebar to upload files, select a source folder, refresh the dataset, or generate a PowerPoint report.</div>
        """,
        unsafe_allow_html=True,
    )

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

        if st.button("Generate PowerPoint report"):
            try:
                # Load current KPI data
                df = load_daily_kpis()
                if df is None or df.empty:
                    st.error("No data available. Please refresh data first.")
                else:
                    # Prepare data with derived columns
                    if "business_date" in df.columns:
                        df["business_date"] = pd.to_datetime(df["business_date"])
                    df = df.sort_values("business_date").reset_index(drop=True)
                    
                    if "daily_occupancy" in df.columns:
                        df["daily_occupancy_pct"] = pd.to_numeric(df["daily_occupancy"], errors="coerce") * 100
                    
                    # Load validation counts if available
                    validation_counts = None
                    val_path = Path("data") / "processed" / "early_bird_validation.csv"
                    if val_path.exists():
                        vdf = pd.read_csv(val_path)
                        validation_counts = vdf["status"].value_counts().to_dict()
                    
                    # Generate PowerPoint
                    with st.spinner("Generating PowerPoint report..."):
                        report_path = generate_powerpoint_report(df, validation_counts=validation_counts)
                    
                    st.success(f"PowerPoint report created: {report_path.name}")
                    st.info(f"Saved to: {report_path}")

                    try:
                        with open(report_path, "rb") as pptx_file:
                            pptx_bytes = pptx_file.read()
                        st.download_button(
                            label="Download PowerPoint report",
                            data=pptx_bytes,
                            file_name=report_path.name,
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        )
                    except Exception as e:
                        st.error(f"Unable to prepare download file: {e}")
            except Exception as e:
                st.error(f"Failed to generate report: {e}")

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

    latest = df.iloc[-1]
    total_reports = len(df)
    date_min = df["business_date"].min()
    date_max = df["business_date"].max()
    date_range_display = f"{date_min.strftime('%Y-%m-%d')} → {date_max.strftime('%Y-%m-%d')}"

    latest_rooms = int(latest.get("daily_rooms_sold")) if pd.notna(latest.get("daily_rooms_sold")) else "-"
    latest_occ = latest.get("daily_occupancy_pct") if "daily_occupancy_pct" in latest.index else None
    latest_occ_display = format_percent(latest_occ) if latest_occ is not None and pd.notna(latest_occ) else "-"
    latest_adr_display = format_currency(latest.get("daily_average_rate"))
    latest_total_revenue = format_currency(latest.get("revenue_mtd_total_revenue"))

    latest_early_dm = latest.get("early_dm_name") if "early_dm_name" in latest.index else None
    latest_late_dm = latest.get("late_dm_name") if "late_dm_name" in latest.index else None
    latest_night_dm = latest.get("night_dm_name") if "night_dm_name" in latest.index else None
    early_dm_display = format_card_value(latest_early_dm)
    late_dm_display = format_card_value(latest_late_dm)
    night_dm_display = format_card_value(latest_night_dm)

    previous_row = df.iloc[-2] if len(df) > 1 else None
    rooms_change = None
    occupancy_change = None
    adr_change = None
    revenue_change = None
    if previous_row is not None:
        prev_rooms = previous_row.get("daily_rooms_sold")
        if pd.notna(prev_rooms) and pd.notna(latest.get("daily_rooms_sold")):
            rooms_change = int(latest.get("daily_rooms_sold")) - int(prev_rooms)

        prev_occ = previous_row.get("daily_occupancy_pct") if "daily_occupancy_pct" in previous_row.index else None
        if latest_occ is not None and pd.notna(latest_occ) and prev_occ is not None and pd.notna(prev_occ):
            occupancy_change = latest_occ - prev_occ

        prev_adr = previous_row.get("daily_average_rate")
        if pd.notna(prev_adr) and pd.notna(latest.get("daily_average_rate")):
            adr_change = latest.get("daily_average_rate") - prev_adr

        prev_rev = previous_row.get("revenue_mtd_total_revenue")
        if pd.notna(prev_rev) and pd.notna(latest.get("revenue_mtd_total_revenue")):
            revenue_change = latest.get("revenue_mtd_total_revenue") - prev_rev

    val_path = Path("data") / "processed" / "early_bird_validation.csv"
    validation_counts = None
    pass_count = 0
    validation_status_text = "No validation data"
    fail_count = 0
    warn_count = 0
    if val_path.exists():
        vdf = pd.read_csv(val_path)
        validation_counts = vdf["status"].value_counts().to_dict()
        fail_count = int(validation_counts.get("FAIL", 0))
        warn_count = int(validation_counts.get("WARN", 0))
        pass_count = int(validation_counts.get("PASS", 0))
        if fail_count == 0 and warn_count == 0:
            validation_status_text = "All checks passed"
        else:
            validation_status_text = f"{fail_count} failures, {warn_count} warnings"

    overview_tab, trends_tab, validation_tab, data_tab = st.tabs([
        "Overview",
        "Trends",
        "Validation",
        "Data",
    ])

    with overview_tab:
        st.markdown(
            f"""
            <div class="kpi-row">
                <div class="kpi-card">
                    <div class="kpi-label">Total reports</div>
                    <div class="kpi-value">{total_reports}</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-label">Date range</div>
                    <div class="kpi-value">{date_range_display}</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-label">Duty Managers</div>
                    <div class="kpi-value">Early: {early_dm_display}<br>Late: {late_dm_display}<br>Night: {night_dm_display}</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-label">Latest rooms sold</div>
                    <div class="kpi-value">{latest_rooms}</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-label">Latest occupancy</div>
                    <div class="kpi-value">{latest_occ_display}</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-label">Latest ADR</div>
                    <div class="kpi-value">{latest_adr_display}</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-label">Latest total revenue</div>
                    <div class="kpi-value">{latest_total_revenue}</div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        operations_items = [
            ("📅 Business date", latest.get("business_date").strftime("%Y-%m-%d") if pd.notna(latest.get("business_date")) else "-"),
            ("🛏️ Rooms sold", latest_rooms),
            ("🔼 Rooms sold delta", format_delta_html(rooms_change, decimals=0)),
            ("👤 Early DM", early_dm_display),
            ("👤 Late DM", late_dm_display),
            ("👤 Night DM", night_dm_display),
        ]

        commercial_items = [
            ("📈 Occupancy", latest_occ_display),
            ("🔼 Occupancy delta", format_delta_html(occupancy_change, suffix=" pp", decimals=1)),
            ("💷 ADR", latest_adr_display),
            ("🔼 ADR delta", format_delta_html(adr_change, prefix="£", decimals=2)),
        ]

        revenue_items = [
            ("💰 Total revenue", latest_total_revenue),
            ("🔼 Revenue delta", format_delta_html(revenue_change, prefix="£", decimals=2)),
            ("✅ Validation status", validation_status_text),
            ("📝 PASS count", str(pass_count)),
        ]

        summary_html = "<div class='summary-grid'>"
        for title, items in [
            ("Operations", operations_items),
            ("Commercial", commercial_items),
            ("Revenue & Validation", revenue_items),
        ]:
            summary_html += "<div class='summary-card'>"
            summary_html += f"<div class='summary-card-title'>{title}</div>"
            for label, value in items:
                summary_html += (
                    f"<div class='summary-row'><span class='summary-label'>{label}</span>"
                    f"<span class='summary-value'>{value}</span></div>"
                )
            summary_html += "</div>"
        summary_html += "</div>"
        st.markdown(summary_html, unsafe_allow_html=True)

        st.markdown("<div class='section-card'><div class='section-title'>Latest report details</div></div>", unsafe_allow_html=True)
        latest_details = {
            "Business Date": latest.get("business_date").strftime("%Y-%m-%d") if pd.notna(latest.get("business_date")) else "-",
            "Early DM": early_dm_display,
            "Late DM": late_dm_display,
            "Night DM": night_dm_display,
            "Rooms Sold": latest_rooms,
            "Occupancy %": format_percent(latest.get("daily_occupancy") * 100) if pd.notna(latest.get("daily_occupancy")) else "-",
            "ADR £": latest_adr_display,
            "Total Revenue £": latest_total_revenue,
            "Rooms Revenue £": format_currency(latest.get("revenue_stats_rooms_revenue")),
            "Bar Revenue £": format_currency(latest.get("revenue_mtd_bar")),
        }
        st.table(pd.DataFrame([latest_details]))

    with trends_tab:
        st.markdown("<div class='section-card'><div class='section-title'>Trends</div></div>", unsafe_allow_html=True)
        chart_cols = st.columns(2)

        if "daily_rooms_sold" in df.columns:
            fig_rooms = px.line(df, x="business_date", y="daily_rooms_sold", title="Daily Rooms Sold")
            fig_rooms.update_layout(yaxis_title="Rooms Sold")
            chart_cols[0].plotly_chart(fig_rooms, use_container_width=True)

        if "daily_occupancy_pct" in df.columns:
            fig_occ = px.line(df, x="business_date", y="daily_occupancy_pct", title="Daily Occupancy")
            fig_occ.update_layout(yaxis_title="Occupancy (%)")
            chart_cols[0].plotly_chart(fig_occ, use_container_width=True)

        if "daily_average_rate" in df.columns:
            fig_adr = px.line(df, x="business_date", y="daily_average_rate", title="Daily Average Rate")
            fig_adr.update_layout(yaxis_title="ADR (£)")
            chart_cols[1].plotly_chart(fig_adr, use_container_width=True)

        if "revenue_mtd_total_revenue" in df.columns:
            fig_rev = px.line(df, x="business_date", y="revenue_mtd_total_revenue", title="MTD Total Revenue")
            fig_rev.update_layout(yaxis_title="Revenue (£)")
            chart_cols[1].plotly_chart(fig_rev, use_container_width=True)

    with validation_tab:
        st.markdown("<div class='section-card'><div class='section-title'>Validation Summary</div></div>", unsafe_allow_html=True)
        if validation_counts is not None:
            status_html = "<div class='summary-grid'>"
            status_html += "<div class='summary-card'>"
            status_html += "<div class='summary-card-title'>Overall quality</div>"
            status_html += f"<div class='summary-row'><span class='summary-label'>Validation status</span><span class='summary-value'>{validation_status_text}</span></div>"
            status_html += "<div class='summary-row'><span class='summary-label'>PASS</span><span class='summary-value'>" + str(pass_count) + "</span></div>"
            status_html += "<div class='summary-row'><span class='summary-label'>WARN</span><span class='summary-value'>" + str(warn_count) + "</span></div>"
            status_html += "<div class='summary-row'><span class='summary-label'>FAIL</span><span class='summary-value'>" + str(fail_count) + "</span></div>"
            status_html += "</div></div>"
            st.markdown(status_html, unsafe_allow_html=True)

            if fail_count == 0 and warn_count == 0:
                st.success("All validation checks passed. No WARN or FAIL issues found.")
            else:
                st.markdown("<div class='section-card'><div class='section-title'>Warnings and Failures</div></div>", unsafe_allow_html=True)
                problems = vdf[vdf["status"].isin(["WARN", "FAIL"])]
                if not problems.empty:
                    st.dataframe(problems[["business_date", "file_name", "check_name", "status", "message"]])
                else:
                    st.info("No WARN or FAIL issues found, but validation data is present.")
        else:
            st.info("Validation file not found. Run `python -m app.extractor` to produce validation results.")

    with data_tab:
        st.markdown("<div class='section-card'><div class='section-title'>Full data</div></div>", unsafe_allow_html=True)
        preferred_columns = [
            "business_date",
            "early_dm_name",
            "late_dm_name",
            "night_dm_name",
            "daily_rooms_sold",
            "daily_occupancy",
            "daily_average_rate",
            "revenue_mtd_total_revenue",
            "revenue_stats_rooms_revenue",
            "revenue_mtd_bar",
            "file_name",
        ]
        remaining_columns = [c for c in df.columns if c not in preferred_columns]
        display_df = df[[c for c in preferred_columns if c in df.columns] + remaining_columns]
        st.dataframe(display_df)


if __name__ == "__main__":
    main()
