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
    write_products_and_reviews_excel,
    write_reviews_excel,
)
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
        description="Scrape Target.com product + reviews for DPCIs, run sentiment, write CSV."
    )
    parser.add_argument(
        "--csv",
        type=str,
        default=str(config.DPCI_CSV),
        help="Path to DPCI input CSV (default: dpcis.csv)",
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
        default=config.REVIEW_LIMIT,
        metavar="N",
        help=f"Maximum reviews to fetch per DPCI (default: {config.REVIEW_LIMIT})",
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
        default=str(config.PRODUCTS_AND_REVIEWS_XLSX),
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
        help="Export current-month reviews for DPCIs from the input CSV",
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

    has_month_range = args.from_month is not None or args.to_month is not None
    if args.current_month and (args.dpci or args.dpcis or has_month_range):
        logger.error("--current-month uses the input CSV and cannot be combined with DPCI or month-range flags")
        return 1
    if (args.from_month is None) != (args.to_month is None):
        logger.error("--from-month and --to-month must be provided together")
        return 1
    if args.dpci and args.dpcis:
        logger.error("Use either --dpci or --dpcis, not both")
        return 1
    if args.from_month and args.to_month and args.from_month > args.to_month:
        logger.error("--from-month must be earlier than or equal to --to-month")
        return 1
    if args.excel_output and not has_month_range:
        logger.error("--excel-output requires --from-month and --to-month")
        return 1

    csv_path = Path(args.csv)
    requested_dpcis: str | list[str] | None = args.dpci or args.dpcis
    try:
        dpcis = load_dpcis(csv_path, limit=args.limit, only=requested_dpcis)
        dpci_library = load_dpci_library(config.DPCI_LIBRARY_XLSX)
        defects_library = load_defects_library(config.DEFECTS_LIBRARY_XLSX)
    except (FileNotFoundError, ValueError) as exc:
        logger.error("%s", exc)
        return 1

    if not dpcis:
        logger.error("No DPCIs to process")
        return 1

    run_date = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    if has_month_range:
        review_month_range = (args.from_month, args.to_month)
    else:
        today = date.today()
        current_month = (today.year, today.month)
        review_month_range = (current_month, current_month)
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
            review_limit=args.review_limit,
        )
        product_rows.append(row)
        review_rows.extend(reviews)

        if scrape_status == "ok":
            ok += 1
        elif scrape_status == "not_found":
            not_found += 1
        else:
            errors += 1

    write_products_and_reviews_excel(
        product_rows,
        review_rows,
        Path(args.workbook_output),
    )
    if has_month_range:
        excel_path = Path(args.excel_output or config.REVIEWS_RANGE_XLSX)
        write_reviews_excel(review_rows, excel_path)

    logger.info(
        "Done. total=%d ok=%d not_found=%d errors=%d -> %s",
        total,
        ok,
        not_found,
        errors,
        args.workbook_output,
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
