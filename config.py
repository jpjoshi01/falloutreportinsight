from __future__ import annotations

from pathlib import Path


APP_NAME = "Migration Fallout Report Intelligence Dashboard"
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
ATTACHMENT_DIR = DATA_DIR / "attachments"
DB_PATH = DATA_DIR / "migration_fallout.db"

IMPORTANT_SHEETS = {
    "summary by rules",
    "sdb summary(with dependencies)",
    "customers by validation",
}

EXCEL_EXTENSIONS = {".xlsx", ".xlsm", ".xls"}

DEFAULT_EMAIL_KEYWORDS = [
    "fallout report",
    "prd dump validation",
    "artemis strategic accounts",
    "increment fallout",
]


def ensure_data_dirs() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    ATTACHMENT_DIR.mkdir(parents=True, exist_ok=True)
