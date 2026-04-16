import os
from pathlib import Path

try:
    from dotenv import load_dotenv
except Exception:
    load_dotenv = None


BASE_DIR = Path(__file__).resolve().parent

if load_dotenv is not None:
    load_dotenv(BASE_DIR / ".env")


def _as_bool(value: str | None, default: bool) -> bool:
    if value is None:
        return default
    return value.strip().lower() in {"1", "true", "yes", "y", "on"}


def _require_env(key: str) -> str:
    value = os.getenv(key)
    if not value:
        raise ValueError(f"Environment variable '{key}' is required but not set.")
    return value


# REQUIRED
OUTPUT_BASE = _require_env("OUTPUT_BASE")
BASE_DIR_LEGACY = _require_env("BASE_DIR_LEGACY")
BASE_DIR_PARQUET = _require_env("BASE_DIR_PARQUET")
TABLE_REFERENCE_PATH = _require_env("TABLE_REFERENCE_PATH")
PROJECT_REFERENCE_SOURCE_XLSX = _require_env("PROJECT_REFERENCE_SOURCE_XLSX")
RT_REFERENCE_PATH = _require_env("RT_REFERENCE_PATH")


# OPTIONAL
PROJECT_REFERENCE_SHEET = os.getenv("PROJECT_REFERENCE_SHEET", "ACC & SHIPPER GROUPING")

PROJECT_REFERENCE_TARGET_CSV = os.getenv(
    "PROJECT_REFERENCE_TARGET_CSV",
    str(BASE_DIR / "data" / "project_reference.csv"),
)

MANUALS_PATH = os.getenv(
    "MANUALS_PATH",
    str(BASE_DIR / "data" / "manuals.csv"),
)

WA_USER_DATA_DIR = os.getenv("WA_USER_DATA_DIR", str(BASE_DIR / "selenium_profile"))
WA_PROFILE_DIRECTORY = os.getenv("WA_PROFILE_DIRECTORY", "Default")
WA_HEADLESS = _as_bool(os.getenv("WA_HEADLESS"), default=True)