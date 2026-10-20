"""Runtime configuration loaded from environment / .env."""

from __future__ import annotations

import os
from datetime import date
from pathlib import Path

from dotenv import load_dotenv

# Project root = parent of src/
ROOT_DIR = Path(__file__).resolve().parent.parent
load_dotenv(ROOT_DIR / ".env")

DPCI_LIBRARY_XLSX = Path(
    os.getenv("DPCI_LIBRARY_XLSX", ROOT_DIR / "DPCI LIBRARY.xlsx")
)
DEFECTS_LIBRARY_XLSX = Path(
    os.getenv("DEFECTS_LIBRARY_XLSX", ROOT_DIR / "DEFECTS LIST.xlsx")
)
OUTPUT_DIR = Path(os.getenv("OUTPUT_DIR", ROOT_DIR / "output"))
REVIEWS_RANGE_XLSX = OUTPUT_DIR / "reviews_range.xlsx"
PRODUCTS_AND_REVIEWS_XLSX = OUTPUT_DIR / f"{date.today():%B} Target Data.xlsx"
RUN_LOG = OUTPUT_DIR / "run.log"

REQUEST_DELAY_SECONDS = float(os.getenv("REQUEST_DELAY_SECONDS", "1.2"))
REVIEW_PAGE_SIZE = int(os.getenv("REVIEW_PAGE_SIZE", "50"))
REVIEW_LIMIT = int(os.getenv("REVIEW_LIMIT", "50"))
MAX_API_REVIEW_LIMIT = 51 * min(max(REVIEW_PAGE_SIZE, 1), 50)
REQUEST_TIMEOUT = int(os.getenv("REQUEST_TIMEOUT", "30"))
MAX_RETRIES = int(os.getenv("MAX_RETRIES", "3"))

STORE_ID = os.getenv("STORE_ID", "3991")
ZIP = os.getenv("ZIP", "10001")
STATE = os.getenv("STATE", "NY")
LATITUDE = os.getenv("LATITUDE", "40.75")
LONGITUDE = os.getenv("LONGITUDE", "-73.99")

# Fallback public web keys (also bootstrapped from Target HTML when possible).
DEFAULT_REDSKY_KEY = "9f36aeafbe60771e321a7cc95a78140772ab3e96"
DEFAULT_NOVA_KEY = "c6b68aaef0eac4df4931aae70500b7056531cb37"

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/122.0.0.0 Safari/537.36"
)

PRODUCT_COLUMNS = [
    "DPCI",
    "TCIN",
    "Item Description",
    "Item Type",
    "Brand",
    "Avg. Rating",
    "Positive %",
    "Negative %",
    "Neutral %",
    "Average Satisfaction %",
    "Top Remark",
    "Program",
    "Status",
    "Website Status",
    # "Run Date",
]

REVIEW_COLUMNS = [
    "DPCI",
    # "TCIN",
    # "Review ID",
    "Review Rating",
    "Review Title",
    "Review Text",
    "Clean Text",
    "Date Posted",
    "Image URLs",
    "Remark",
    "Reason",
    "Defect",
    "Main Category",
    "Sentiment",
    "Sentiment Score",
    "Satisfaction Percentage",
    # "Run Date",
]

DPCI_PATTERN = r"^\d{3}-\d{2}-\d{4}$"