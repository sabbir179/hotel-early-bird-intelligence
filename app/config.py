from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = PROJECT_ROOT / "data"
RAW_DATA_DIR = DATA_DIR / "raw"
PROCESSED_DATA_DIR = DATA_DIR / "processed"
DATABASE_DIR = DATA_DIR / "database"
REPORTS_DIR = PROJECT_ROOT / "reports"
POWERPOINT_REPORTS_DIR = REPORTS_DIR / "powerpoint"
PDF_REPORTS_DIR = REPORTS_DIR / "pdf"
