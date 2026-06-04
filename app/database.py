import sqlite3
from pathlib import Path
from typing import Optional

import pandas as pd


DEFAULT_DB_PATH = Path("data") / "database" / "early_bird.db"


def get_connection(db_path: Optional[Path] = None) -> sqlite3.Connection:
    """Return a sqlite3 connection, creating parent folders if needed."""
    db_path = Path(db_path) if db_path is not None else DEFAULT_DB_PATH
    db_path.parent.mkdir(parents=True, exist_ok=True)
    return sqlite3.connect(db_path)


def save_daily_kpis(df: pd.DataFrame, db_path: Optional[Path] = None) -> Path:
    """Save the master KPI DataFrame to the `daily_kpis` table in SQLite.

    Uses `if_exists='replace'` for now and returns the database path used.
    The table schema is recreated from the DataFrame columns, including new fields
    such as `early_dm_name`, `late_dm_name`, and `night_dm_name` when present.
    """
    db_path = Path(db_path) if db_path is not None else DEFAULT_DB_PATH
    db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(db_path)
    try:
        df.to_sql("daily_kpis", conn, if_exists="replace", index=False)
    finally:
        conn.close()
    return db_path


def load_daily_kpis(db_path: Optional[Path] = None) -> pd.DataFrame:
    """Load the `daily_kpis` table from SQLite into a pandas DataFrame."""
    db_path = Path(db_path) if db_path is not None else DEFAULT_DB_PATH
    conn = sqlite3.connect(db_path)
    try:
        return pd.read_sql_query("SELECT * FROM daily_kpis", conn)
    finally:
        conn.close()
