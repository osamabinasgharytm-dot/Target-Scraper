"""Write combined product/review workbooks and optional review-range Excel files."""

from __future__ import annotations

import logging
from datetime import date, datetime
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

from src.config import PRODUCT_COLUMNS, REVIEW_COLUMNS
from src.dashboard import write_dashboard, write_dashboard_from_excel

logger = logging.getLogger("scrapper")


def _review_identity(row: dict) -> tuple[str, ...] | None:
    """Build a stable review key, with a content fallback for older exports."""
    dpci = str(row.get("DPCI") or "").strip()
    review_id = str(row.get("Review ID") or "").strip()
    if review_id:
        return ("id", dpci, review_id)

    def normalized(value: object) -> str:
        if value is None or pd.isna(value):
            return ""
        if isinstance(value, (date, datetime, pd.Timestamp)):
            return value.strftime("%Y-%m-%d")
        return " ".join(str(value).split()).casefold()

    date_value = row.get("Date Posted")
    parsed_date = pd.to_datetime(date_value, dayfirst=True, errors="coerce")
    signature = (
        normalized(row.get("DPCI")),
        "" if pd.isna(parsed_date) else parsed_date.strftime("%Y-%m-%d"),
        normalized(row.get("Review Rating")),
        normalized(row.get("Review Title")),
        normalized(row.get("Review Text")),
    )
    if not any(signature[1:]):
        return None
    return ("content", *signature)


def load_website_error_dpcis(path: Path) -> list[str]:
    """Find DPCIs whose latest Products-sheet entry has Website Status Error."""
    if not path.is_file():
        raise FileNotFoundError(f"Excel workbook not found: {path}")

    with pd.ExcelFile(path) as workbook:
        if "Products" not in workbook.sheet_names:
            raise ValueError(f"Excel workbook has no Products sheet: {path}")
        products = pd.read_excel(
            workbook,
            sheet_name="Products",
            dtype={"DPCI": str},
        )

    missing_columns = {"DPCI", "Website Status"}.difference(products.columns)
    if missing_columns:
        missing = ", ".join(sorted(missing_columns))
        raise ValueError(
            f"Products sheet must contain DPCI and Website Status columns; "
            f"missing: {missing}"
        )

    latest_status: dict[str, str] = {}
    for dpci, status in zip(products["DPCI"], products["Website Status"]):
        if pd.isna(dpci) or pd.isna(status):
            continue
        dpci_value = str(dpci).strip()
        if dpci_value:
            latest_status[dpci_value] = str(status).strip()

    return [
        dpci
        for dpci, status in latest_status.items()
        if status.casefold() == "error"
    ]


def append_products_and_reviews_excel(
    product_rows: list[dict],
    review_rows: list[dict],
    path: Path,
) -> Path:
    """Append product and review rows to an existing combined workbook."""
    if not path.is_file():
        raise FileNotFoundError(f"Excel workbook not found: {path}")

    workbook = load_workbook(path)
    required_sheets = {"Products", "Reviews"}
    missing_sheets = required_sheets.difference(workbook.sheetnames)
    if missing_sheets:
        missing = ", ".join(sorted(missing_sheets))
        raise ValueError(
            f"Excel workbook must contain Products and Reviews sheets; "
            f"missing: {missing}"
        )

    appended_review_count = 0
    for sheet_name, rows, expected_columns in (
        ("Products", product_rows, PRODUCT_COLUMNS),
        ("Reviews", review_rows, REVIEW_COLUMNS),
    ):
        sheet = workbook[sheet_name]
        headers = {
            str(sheet.cell(row=1, column=column).value).strip(): column
            for column in range(1, sheet.max_column + 1)
            if sheet.cell(row=1, column=column).value is not None
        }
        for column in expected_columns:
            if column not in headers:
                new_column = sheet.max_column + 1
                sheet.cell(row=1, column=new_column, value=column)
                headers[column] = new_column

        seen_review_ids: set[tuple[str, ...]] = set()
        legacy_content_keys: set[tuple[str, ...]] = set()
        if sheet_name == "Reviews":
            for row_number in range(2, sheet.max_row + 1):
                existing_row = {
                    header: sheet.cell(row=row_number, column=column).value
                    for header, column in headers.items()
                }
                identity = _review_identity(existing_row)
                if identity is None:
                    continue
                if identity[0] == "id":
                    seen_review_ids.add(identity)
                else:
                    legacy_content_keys.add(identity)

        for row in rows:
            if sheet_name == "Reviews":
                identity = _review_identity(row)
                if identity is not None and (
                    identity in seen_review_ids
                    or (
                        identity[0] == "id"
                        and _review_identity(
                            {
                                **row,
                                "Review ID": "",
                            }
                        )
                        in legacy_content_keys
                    )
                ):
                    continue
            values = [None] * sheet.max_column
            for column, value in row.items():
                if column not in headers:
                    continue
                if value is None or pd.isna(value):
                    value = None
                elif hasattr(value, "item"):
                    value = value.item()
                values[headers[column] - 1] = value
            sheet.append(values)
            if sheet_name == "Reviews" and identity is not None:
                if identity[0] == "id":
                    seen_review_ids.add(identity)
                else:
                    legacy_content_keys.add(identity)
                appended_review_count += 1

    product_sheet = workbook["Products"]
    product_headers = {
        str(product_sheet.cell(row=1, column=column).value).strip(): column
        for column in range(1, product_sheet.max_column + 1)
        if product_sheet.cell(row=1, column=column).value is not None
    }
    if "DPCI" not in product_headers or "Website Status" not in product_headers:
        raise ValueError("Products sheet must contain DPCI and Website Status columns")

    successfully_scraped = {
        str(row["DPCI"]).strip()
        for row in product_rows
        if row.get("DPCI")
        and str(row.get("Website Status", "")).strip().casefold() == "found"
    }
    for row_number in range(2, product_sheet.max_row + 1):
        dpci = product_sheet.cell(
            row=row_number,
            column=product_headers["DPCI"],
        ).value
        website_status = product_sheet.cell(
            row=row_number,
            column=product_headers["Website Status"],
        )
        if (
            dpci is not None
            and str(dpci).strip() in successfully_scraped
            and str(website_status.value or "").strip().casefold() == "error"
        ):
            website_status.value = "Resolved"

    workbook.save(path)
    dashboard_path = write_dashboard_from_excel(path)
    logger.info(
        "Appended %d product rows and %d new review rows to %s; dashboard -> %s",
        len(product_rows),
        appended_review_count,
        path,
        dashboard_path,
    )
    return path


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
