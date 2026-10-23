"""CLI entry point for the Target product scraper pipeline."""

from __future__ import annotations

import argparse
import logging
import sys
from datetime import date, datetime, timezone
from pathlib import Path

from src import config
from src.aggregator import aggregate_product_row, empty_error_row
from src.csv_writer import (
    append_products_and_reviews_excel,
    load_website_error_dpcis,
    write_products_and_reviews_excel,
    write_reviews_excel,
)
from src.dashboard import write_dashboard_from_excel
from src.defect_library import load_defects_library, match_defects
from src.dpci_library import load_dpci_library
from src.dpci_loader import load_dpcis
from src.logger_setup import setup_logging
from src.product_parser import parse_product
from src.review_pipeline import process_reviews
from src.target_client import NotFoundError, TargetAPIError, TargetClient

logger = logging.getLogger("scrapper")


def _is_review_in_month_range(
    review: dict,
    month_range: tuple[tuple[int, int], tuple[int, int]],
) -> bool:
    posted_date = review.get("Date Posted")
    if isinstance(posted_date, datetime):
        parsed_date = posted_date
    elif isinstance(posted_date, date):
        parsed_date = posted_date
    elif isinstance(posted_date, str):
        try:
            posted_date = posted_date.strip()
            try:
                parsed_date = datetime.strptime(posted_date, "%d/%m/%Y")
            except ValueError:
                parsed_date = datetime.fromisoformat(posted_date.replace("Z", "+00:00"))
        except ValueError:
            return False
    else:
        return False

    posted_month = (parsed_date.year, parsed_date.month)
    return month_range[0] <= posted_month <= month_range[1]


def _parse_month(value: str) -> tuple[int, int]:
    if len(value) != 7 or value[4] != "-":
        raise argparse.ArgumentTypeError("month must use YYYY-MM format")
    try:
        parsed = datetime.strptime(value, "%Y-%m")
    except ValueError as exc:
        raise argparse.ArgumentTypeError("month must use YYYY-MM format") from exc
    return parsed.year, parsed.month


def _previous_calendar_month(today: date) -> tuple[int, int]:
    if today.month == 1:
        return today.year - 1, 12
    return today.year, today.month - 1


def _default_workbook_path(
    month_range: tuple[tuple[int, int], tuple[int, int]],
) -> Path:
    start_year, start_month = month_range[0]
    end_year, end_month = month_range[1]
    today = date.today()
    if month_range == ((today.year, today.month), (today.year, today.month)):
        return config.PRODUCTS_AND_REVIEWS_XLSX
    start_date = date(start_year, start_month, 1)
    end_date = date(end_year, end_month, 1)
    if month_range[0] == month_range[1]:
        filename = f"{start_date:%B} Target Data.xlsx"
    elif start_year == end_year:
        filename = f"{start_date:%B}-{end_date:%B} {start_year} Target Data.xlsx"
    else:
        filename = (
            f"{start_date:%B} {start_year}-{end_date:%B} {end_year} Target Data.xlsx"
        )
    return config.OUTPUT_DIR / filename


def _parse_dpcis(value: str) -> list[str]:
    dpcis = [dpci.strip() for dpci in value.split(",")]
    if not dpcis or any(not dpci for dpci in dpcis):
        raise argparse.ArgumentTypeError("provide comma-separated DPCIs without empty values")
    return dpcis


def _positive_int(value: str) -> int:
    try:
        number = int(value)
    except ValueError as exc:
        raise argparse.ArgumentTypeError("must be a positive integer") from exc
    if number < 1:
        raise argparse.ArgumentTypeError("must be a positive integer")
    return number


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Scrape Target.com products and reviews for DPCIs, run sentiment, "
            "and write Excel workbooks."
        )
    )
    parser.add_argument(
        "--dpci-library",
        type=str,
        default=str(config.DPCI_LIBRARY_XLSX),
        help="Excel workbook used as the DPCI source (default: DPCI LIBRARY.xlsx)",
    )
    parser.add_argument(
        "--limit",
        type=int,
        default=None,
        help="Process only the first N DPCIs (after validate/dedupe)",
    )
    parser.add_argument(
        "--review-limit",
        type=_positive_int,
        default=None,
        metavar="N",
        help=(
            "Maximum reviews to fetch per DPCI "
            f"(default: {config.REVIEW_LIMIT}; current/last-month modes default "
            f"to the API maximum of {config.MAX_API_REVIEW_LIMIT})"
        ),
    )
    parser.add_argument(
        "--dpci",
        type=str,
        default=None,
        help="Process a single DPCI (e.g. 097-02-2055)",
    )
    parser.add_argument(
        "--dpcis",
        type=_parse_dpcis,
        default=None,
        help="Process specific comma-separated DPCIs (e.g. 097-02-2055,060-15-0108)",
    )
    parser.add_argument(
        "--from-month",
        type=_parse_month,
        default=None,
        metavar="YYYY-MM",
        help="First review month to include (requires --to-month)",
    )
    parser.add_argument(
        "--to-month",
        type=_parse_month,
        default=None,
        metavar="YYYY-MM",
        help="Last review month to include, inclusive",
    )
    parser.add_argument(
        "--excel-output",
        type=str,
        default=None,
        help="Excel output path for a month-range export",
    )
    parser.add_argument(
        "--workbook-output",
        type=str,
        default=None,
        help="Excel workbook path for combined Products and Reviews sheets",
    )
    parser.add_argument(
        "--save-reviews",
        action="store_true",
        help="Deprecated; per-review data is included in the combined workbook",
    )
    parser.add_argument(
        "--current-month",
        action="store_true",
        help=(
            "Export current-month reviews; use --dpcis for selected DPCIs "
            "(defaults to the API review maximum per DPCI)"
        ),
    )
    parser.add_argument(
        "--all-current-month",
        action="store_true",
        help=(
            "Export current-month reviews for every DPCI in the DPCI library "
            "(defaults to the API review maximum per DPCI)"
        ),
    )
    parser.add_argument(
        "--last-month",
        action="store_true",
        help=(
            "Export previous calendar month's reviews for every DPCI in the "
            "DPCI library (defaults to the API maximum per DPCI)"
        ),
    )
    parser.add_argument(
        "--retry-errors",
        action="store_true",
        help=(
            "Re-scrape DPCIs whose latest Website Status is Error, update product "
            "rows, and append new reviews"
        ),
    )
    parser.add_argument(
        "--update-reviews",
        action="store_true",
        help=(
            "Scrape current-month reviews for all library DPCIs and append only "
            "new reviews to the existing combined workbook"
        ),
    )
    parser.add_argument(
        "--create-dashboard",
        action="store_true",
        help="Generate a dashboard from this month's existing combined Excel workbook",
    )
    parser.add_argument(
        "--dashboard-workbook",
        type=str,
        default=None,
        help=(
            "Excel workbook to use with --create-dashboard "
            "(default: this month's workbook in the output folder)"
        ),
    )
    return parser.parse_args(argv)


def process_one(
    client: TargetClient,
    dpci: str,
    *,
    run_date: str,
    save_reviews: bool,
    review_month_range: tuple[tuple[int, int], tuple[int, int]] | None = None,
    library_fields: dict[str, str] | None = None,
    defects_library: list[dict[str, str]] | None = None,
    review_limit: int = config.REVIEW_LIMIT,
) -> tuple[dict, list[dict], str]:
    """Scrape one DPCI; return its product row, review rows, and scrape outcome."""
    if review_month_range is None:
        today = date.today()
        current_month = (today.year, today.month)
        review_month_range = (current_month, current_month)

    try:
        bundle = client.scrape_dpci(dpci, review_limit=review_limit)
    except NotFoundError as exc:
        logger.warning("%s -> not_found (%s)", dpci, exc)
        product_row = empty_error_row(dpci, "not_found", run_date=run_date)
        _apply_library_fields(product_row, library_fields)
        return product_row, [], "not_found"
    except TargetAPIError as exc:
        logger.error("%s -> error (%s)", dpci, exc)
        product_row = empty_error_row(dpci, f"error:{exc}", run_date=run_date)
        _apply_library_fields(product_row, library_fields)
        return product_row, [], f"error:{exc}"
    except Exception as exc:  # noqa: BLE001 — keep batch running
        logger.exception("%s -> unexpected error", dpci)
        product_row = empty_error_row(dpci, f"error:{exc}", run_date=run_date)
        _apply_library_fields(product_row, library_fields)
        return product_row, [], f"error:{exc}"

    tcin = bundle["tcin"]
    product_fields = parse_product(dpci, tcin, bundle["product"])
    processed = process_reviews(bundle["reviews"])
    product_row = aggregate_product_row(
        product_fields,
        processed,
        status="ok",
        run_date=run_date,
    )
    _apply_library_fields(product_row, library_fields)

    review_rows: list[dict] = []
    if save_reviews:
        for rev in processed:
            if not _is_review_in_month_range(rev, review_month_range):
                continue
            sentiment = str(rev.get("Sentiment", "")).strip().casefold()
            if sentiment == "negative":
                defect_fields = match_defects(
                    rev.get("Text", ""),
                    defects_library or [],
                )
            else:
                defect_fields = {
                    "Reason": "-",
                    "Defect": "-",
                    "Main Category": "-",
                }
            review_rows.append(
                {
                    "DPCI": dpci,
                    "TCIN": tcin,
                    "Review ID": rev.get("Review ID", ""),
                    "Review Rating": rev.get("Rating", ""),
                    "Review Title": rev.get("Title", ""),
                    "Review Text": rev.get("Text", ""),
                    **defect_fields,
                    "Date Posted": rev.get("Date Posted", ""),
                    "Image URLs": rev.get("Image URLs", ""),
                    "Clean Text": rev.get("Clean Text", ""),
                    "Sentiment": rev.get("Sentiment", ""),
                    "Sentiment Score": rev.get("Sentiment Score", ""),
                    "Satisfaction Percentage": rev.get(
                        "Satisfaction Percentage", ""
                    ),
                    "Remark": rev.get("Remark", ""),
                    "Run Date": run_date,
                }
            )

    logger.info(
        "%s -> ok TCIN=%s reviews=%d avg=%s top=%s",
        dpci,
        tcin,
        len(processed),
        product_row["Avg. Rating"],
        product_row["Top Remark"] or "-",
    )
    return product_row, review_rows, "ok"


def _apply_library_fields(
    product_row: dict,
    library_fields: dict[str, str] | None,
) -> None:
    for field in ("Program", "Status"):
        product_row[field] = (library_fields or {}).get(field, "")


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    config.OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    setup_logging(config.RUN_LOG)

    if args.dashboard_workbook and not args.create_dashboard:
        logger.error("--dashboard-workbook requires --create-dashboard")
        return 1
    if args.create_dashboard:
        if (
            args.dpci
            or args.dpcis
            or args.from_month is not None
            or args.to_month is not None
            or args.current_month
            or args.all_current_month
            or args.last_month
            or args.retry_errors
            or args.update_reviews
            or args.limit is not None
            or args.review_limit is not None
            or args.excel_output
        ):
            logger.error("--create-dashboard cannot be combined with scrape options")
            return 1
        workbook_path = Path(
            args.dashboard_workbook or config.PRODUCTS_AND_REVIEWS_XLSX
        )
        try:
            dashboard_path = write_dashboard_from_excel(workbook_path)
        except (FileNotFoundError, ValueError, OSError) as exc:
            logger.error("%s", exc)
            return 1
        logger.info("Dashboard created from existing workbook -> %s", dashboard_path)
        return 0

    has_month_range = args.from_month is not None or args.to_month is not None
    is_current_month = (
        args.current_month
        or args.all_current_month
        or args.retry_errors
        or args.update_reviews
    )
    if is_current_month and has_month_range:
        logger.error("Current-month mode cannot be combined with month-range flags")
        return 1
    if args.last_month and (is_current_month or has_month_range):
        logger.error("--last-month cannot be combined with other month modes")
        return 1
    if (args.from_month is None) != (args.to_month is None):
        logger.error("--from-month and --to-month must be provided together")
        return 1
    if args.dpci and args.dpcis:
        logger.error("Use either --dpci or --dpcis, not both")
        return 1
    if args.all_current_month and (args.dpci or args.dpcis):
        logger.error("--all-current-month cannot be combined with --dpci or --dpcis")
        return 1
    if args.last_month and (args.dpci or args.dpcis):
        logger.error("--last-month cannot be combined with --dpci or --dpcis")
        return 1
    if args.retry_errors and (
        args.dpci
        or args.dpcis
        or args.current_month
        or args.all_current_month
        or has_month_range
    ):
        logger.error(
            "--retry-errors cannot be combined with DPCI selection or month flags"
        )
        return 1
    if args.update_reviews and (
        args.dpci
        or args.dpcis
        or args.current_month
        or args.all_current_month
        or args.retry_errors
        or has_month_range
    ):
        logger.error(
            "--update-reviews cannot be combined with DPCI selection or other month modes"
        )
        return 1
    if args.from_month and args.to_month and args.from_month > args.to_month:
        logger.error("--from-month must be earlier than or equal to --to-month")
        return 1
    if args.excel_output and not (has_month_range or args.last_month):
        logger.error(
            "--excel-output requires --from-month/--to-month or --last-month"
        )
        return 1

    if args.last_month:
        last_month = _previous_calendar_month(date.today())
        review_month_range = (last_month, last_month)
    elif has_month_range:
        review_month_range = (args.from_month, args.to_month)
    else:
        today = date.today()
        current_month = (today.year, today.month)
        review_month_range = (current_month, current_month)
    workbook_path = (
        Path(args.workbook_output)
        if args.workbook_output
        else _default_workbook_path(review_month_range)
    )

    requested_dpcis: str | list[str] | None = (
        None if args.all_current_month or args.last_month else args.dpci or args.dpcis
    )
    try:
        dpci_library = load_dpci_library(Path(args.dpci_library))
        if args.update_reviews:
            if not workbook_path.is_file():
                raise FileNotFoundError(
                    f"Combined workbook not found for review update: {workbook_path}"
                )
            dpcis = load_dpcis(limit=args.limit, only=list(dpci_library))
        elif args.retry_errors:
            retry_dpcis = load_website_error_dpcis(workbook_path)
            if not retry_dpcis:
                logger.info(
                    "No DPCIs with latest Website Status Error in %s",
                    workbook_path,
                )
                return 0
            dpcis = load_dpcis(limit=args.limit, only=retry_dpcis)
        else:
            dpcis = load_dpcis(
                limit=args.limit,
                only=requested_dpcis or list(dpci_library),
            )
        defects_library = load_defects_library(config.DEFECTS_LIBRARY_XLSX)
    except (FileNotFoundError, ValueError) as exc:
        logger.error("%s", exc)
        return 1

    if not dpcis:
        logger.error("No DPCIs to process")
        return 1
    review_limit = args.review_limit or (
        config.MAX_API_REVIEW_LIMIT
        if is_current_month or args.last_month
        else config.REVIEW_LIMIT
    )

    run_date = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    logger.info("Starting run: %d DPCIs | run_date=%s", len(dpcis), run_date)

    client = TargetClient()
    product_rows: list[dict] = []
    review_rows: list[dict] = []
    ok = not_found = errors = 0

    total = len(dpcis)
    for index, dpci in enumerate(dpcis, start=1):
        logger.info("[%d/%d] Processing %s", index, total, dpci)
        library_fields = dpci_library.get(dpci)
        if library_fields is None:
            logger.warning("%s -> no matching DPCI row in library workbook", dpci)
        row, reviews, scrape_status = process_one(
            client,
            dpci,
            run_date=run_date,
            save_reviews=True,
            review_month_range=review_month_range,
            library_fields=library_fields,
            defects_library=defects_library,
            review_limit=review_limit,
        )
        product_rows.append(row)
        review_rows.extend(reviews)

        if scrape_status == "ok":
            ok += 1
        elif scrape_status == "not_found":
            not_found += 1
        else:
            errors += 1

    if args.retry_errors or args.update_reviews:
        append_products_and_reviews_excel(
            product_rows,
            review_rows,
            workbook_path,
        )
    else:
        write_products_and_reviews_excel(
            product_rows,
            review_rows,
            workbook_path,
        )
    if has_month_range or args.last_month:
        excel_path = Path(args.excel_output or config.REVIEWS_RANGE_XLSX)
        write_reviews_excel(review_rows, excel_path)

    logger.info(
        "Done. total=%d ok=%d not_found=%d errors=%d -> %s",
        total,
        ok,
        not_found,
        errors,
        workbook_path,
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
