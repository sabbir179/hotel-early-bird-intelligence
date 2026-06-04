import pandas as pd
from pathlib import Path
from typing import List, Dict, Optional


def records_to_dataframe(records: List[Dict]) -> pd.DataFrame:
    """Convert list of KPI dictionaries to a pandas DataFrame.

    Converts `business_date` to datetime, sorts by it, and resets the index.
    """
    df = pd.DataFrame(records)
    if "business_date" in df.columns:
        df["business_date"] = pd.to_datetime(df["business_date"])
        df = df.sort_values("business_date")
    df = df.reset_index(drop=True)
    return df


def save_master_csv(df: pd.DataFrame, output_path: Optional[Path] = None) -> Path:
    """Save the master DataFrame to CSV and return the output path.

    Default path: `data/processed/early_bird_master.csv`.
    """
    if output_path is None:
        output_path = Path("data/processed/early_bird_master.csv")

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(output_path, index=False)
    return output_path


def save_validation_csv(validation_df: pd.DataFrame, output_path: Optional[Path] = None) -> Path:
    """Save the validation results DataFrame to CSV.

    Default path: `data/processed/early_bird_validation.csv`.
    """
    if output_path is None:
        output_path = Path("data/processed/early_bird_validation.csv")

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    validation_df.to_csv(output_path, index=False)
    return output_path


class AnalyticsEngine:
    """Compute KPI analytics and summaries."""

    def calculate_daily_summary(self, dataframe: pd.DataFrame) -> pd.DataFrame:
        """Return a summary of daily metrics."""
        return dataframe

    def calculate_trends(self, dataframe: pd.DataFrame) -> pd.DataFrame:
        """Return trend analytics for the dataset."""
        return dataframe
