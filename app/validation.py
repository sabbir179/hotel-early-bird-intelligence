import pandas as pd
from typing import List


def validate_records(df: pd.DataFrame) -> pd.DataFrame:
    """Run validation checks on the master KPI DataFrame.

    Returns a DataFrame with columns: business_date, file_name, check_name, status, message
    """
    results: List[dict] = []

    # Helper to append a result
    def add(row, check_name: str, status: str, message: str = ""):
        results.append(
            {
                "business_date": row.get("business_date"),
                "file_name": row.get("file_name"),
                "check_name": check_name,
                "status": status,
                "message": message,
            }
        )

    # Per-row checks
    for _, row in df.iterrows():
        # a) business_date is not missing
        if pd.isna(row.get("business_date")):
            add(row, "business_date_present", "FAIL", "business_date is missing")
        else:
            add(row, "business_date_present", "PASS", "")

        # b) daily_rooms_sold is not missing and >= 0
        drs = row.get("daily_rooms_sold")
        if pd.isna(drs):
            add(row, "daily_rooms_sold_nonmissing", "FAIL", "daily_rooms_sold is missing")
        else:
            try:
                if float(drs) >= 0:
                    add(row, "daily_rooms_sold_nonmissing", "PASS", "")
                else:
                    add(row, "daily_rooms_sold_nonmissing", "FAIL", "daily_rooms_sold < 0")
            except Exception:
                add(row, "daily_rooms_sold_nonmissing", "WARN", "daily_rooms_sold not numeric")

        # c) daily_occupancy is not missing and between 0 and 1.05
        occ = row.get("daily_occupancy")
        if pd.isna(occ):
            add(row, "daily_occupancy_range", "FAIL", "daily_occupancy is missing")
        else:
            try:
                occ_v = float(occ)
                if 0 <= occ_v <= 1.05:
                    add(row, "daily_occupancy_range", "PASS", "")
                else:
                    add(row, "daily_occupancy_range", "FAIL", f"daily_occupancy out of range: {occ_v}")
            except Exception:
                add(row, "daily_occupancy_range", "WARN", "daily_occupancy not numeric")

        # d) daily_average_rate is not missing and > 0
        adr = row.get("daily_average_rate")
        if pd.isna(adr):
            add(row, "daily_average_rate_positive", "FAIL", "daily_average_rate is missing")
        else:
            try:
                if float(adr) > 0:
                    add(row, "daily_average_rate_positive", "PASS", "")
                else:
                    add(row, "daily_average_rate_positive", "FAIL", "daily_average_rate <= 0")
            except Exception:
                add(row, "daily_average_rate_positive", "WARN", "daily_average_rate not numeric")

        # e) revenue_mtd_total_revenue is not missing
        rev = row.get("revenue_mtd_total_revenue")
        if pd.isna(rev):
            add(row, "revenue_mtd_total_revenue_present", "FAIL", "revenue_mtd_total_revenue is missing")
        else:
            add(row, "revenue_mtd_total_revenue_present", "PASS", "")

        # ADR consistency check (4)
        rooms_rev = row.get("revenue_stats_rooms_revenue")
        drs = row.get("daily_rooms_sold")
        adr_val = row.get("daily_average_rate")
        if pd.isna(rooms_rev) or pd.isna(drs) or pd.isna(adr_val):
            add(row, "adr_consistency", "SKIP", "missing values for ADR consistency check")
        else:
            try:
                expected_adr = float(rooms_rev) / float(drs) if float(drs) != 0 else None
                if expected_adr is None:
                    add(row, "adr_consistency", "SKIP", "daily_rooms_sold is zero")
                else:
                    if abs(expected_adr - float(adr_val)) <= 1.0:
                        add(row, "adr_consistency", "PASS", "")
                    else:
                        add(row, "adr_consistency", "WARN", f"expected {expected_adr:.2f} vs adr {adr_val}")
            except Exception:
                add(row, "adr_consistency", "WARN", "error computing ADR consistency")

    # f) duplicate business_date check across full DataFrame
    if "business_date" in df.columns:
        dup_mask = df["business_date"].duplicated(keep=False)
        for idx, is_dup in dup_mask.items():
            row = df.iloc[idx]
            if pd.isna(row.get("business_date")):
                # already reported missing above
                continue
            if is_dup:
                results.append(
                    {
                        "business_date": row.get("business_date"),
                        "file_name": row.get("file_name"),
                        "check_name": "duplicate_business_date",
                        "status": "FAIL",
                        "message": "duplicate business_date",
                    }
                )
            else:
                results.append(
                    {
                        "business_date": row.get("business_date"),
                        "file_name": row.get("file_name"),
                        "check_name": "duplicate_business_date",
                        "status": "PASS",
                        "message": "",
                    }
                )

    return pd.DataFrame(results)
