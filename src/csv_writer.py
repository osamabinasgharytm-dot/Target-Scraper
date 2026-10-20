"""Write combined product/review workbooks and optional review-range Excel files."""

from __future__ import annotations

import logging
from pathlib import Path

import pandas as pd

from src.config import PRODUCT_COLUMNS, REVIEW_COLUMNS
from src.dashboard import write_dashboard

logger = logging.getLogger("scrapper")


# def write_products_csv(rows: list[dict], path: Path) -> Path:
    # path.parent.mkdir(parents=True, exist_ok=True)
    # df = pd.DataFrame(rows, columns=PRODUCT_COLUMNS)
    # df.to_csv(path, index=False, encoding="utf-8-sig")
    # logger.info("Wrote %d product rows -> %s", len(df), path)
    # return path


# def write_reviews_csv(rows: list[dict], path: Path) -> Path:
    # path.parent.mkdir(parents=True, exist_ok=True)
    # df = pd.DataFrame(rows, columns=REVIEW_COLUMNS)
    # df.to_csv(path, index=False, encoding="utf-8-sig")
    # logger.info("Wrote %d review rows -> %s", len(df), path)
    # return path


def write_reviews_excel(rows: list[dict], path: Path) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    df = pd.DataFrame(rows, columns=REVIEW_COLUMNS)
    df.to_excel(path, index=False)
    logger.info("Wrote %d review rows -> %s", len(df), path)
    return path


def write_products_and_reviews_excel(
    product_rows: list[dict],
    review_rows: list[dict],
    path: Path,
) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    products_df = pd.DataFrame(product_rows, columns=PRODUCT_COLUMNS)
    reviews_df = pd.DataFrame(review_rows, columns=REVIEW_COLUMNS)
    with pd.ExcelWriter(path) as writer:
        products_df.to_excel(writer, sheet_name="Products", index=False)
        reviews_df.to_excel(writer, sheet_name="Reviews", index=False)
    write_dashboard(
        products_df.to_dict(orient="records"),
        reviews_df.to_dict(orient="records"),
        path.with_name(f"{path.stem} Dashboard.html"),
    )
    logger.info(
        "Wrote %d product rows and %d review rows -> %s",
        len(products_df),
        len(reviews_df),
        path,
    )
    return path
