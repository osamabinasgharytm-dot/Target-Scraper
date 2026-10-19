import unittest
from datetime import date, datetime, timedelta
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch

import pandas as pd

from src.aggregator import aggregate_product_row
from src.config import PRODUCTS_AND_REVIEWS_XLSX, PRODUCT_COLUMNS
from src.csv_writer import write_products_and_reviews_excel, write_reviews_excel
from src.defect_library import match_defects
from src.dpci_loader import load_dpcis
from src.main import main, parse_args, process_one
from src.review_pipeline import process_reviews
from src.target_client import TargetClient


class ReviewPipelineTests(unittest.TestCase):
    def setUp(self):
        self.sentiment_patcher = patch(
            "src.review_pipeline._sentiment_score",
            return_value=0.8,
        )
        self.sentiment_patcher.start()
        self.addCleanup(self.sentiment_patcher.stop)

    def test_combined_workbook_default_name_uses_current_month(self):
        self.assertEqual(
            PRODUCTS_AND_REVIEWS_XLSX.name,
            f"{datetime.now():%B} Target Data.xlsx",
        )

    def test_process_one_exports_reviews_from_inclusive_month_range(self):
        class FakeClient:
            def scrape_dpci(self, _dpci, *, review_limit):
                self.review_limit = review_limit
                return {
                    "tcin": "123",
                    "product": {},
                    "reviews": [
                        {
                            "id": "current",
                            "text": "Current month review",
                            "submitted_at": "2026-10-14T12:00:00Z",
                        },
                        {
                            "id": "previous",
                            "text": "Previous month review",
                            "submitted_at": "2026-07-31T23:59:59Z",
                        },
                        {
                            "id": "start-month",
                            "rating": 4,
                            "title": "Start month title",
                            "text": "Start month review",
                            "submitted_at": "2026-08-01T00:00:00Z",
                        },
                        {
                            "id": "end-month",
                            "rating": 2,
                            "title": "End month title",
                            "text": "End month review",
                            "submitted_at": "2026-09-30T23:59:59Z",
                        },
                        {
                            "id": "after-range",
                            "text": "After range review",
                            "submitted_at": "2026-10-01T00:00:00Z",
                        },
                        {
                            "id": "unknown-date",
                            "text": "Review without a date",
                        },
                    ],
                }

        with patch("src.main.parse_product", return_value={}), patch(
            "src.main.aggregate_product_row",
            side_effect=lambda _product, reviews, **_kwargs: {
                "Avg. Rating": None,
                "Top Remark": "",
            },
        ):
            client = FakeClient()
            product_row, exported_reviews, scrape_status = process_one(
                client,
                "060-15-0108",
                run_date="2026-10-17T00:00:00Z",
                save_reviews=True,
                review_month_range=((2026, 8), (2026, 9)),
                library_fields={
                    "Program": "TARGET PILLOWFORT COTTON SHEETS",
                    "Status": "CF",
                },
            )

        self.assertEqual(scrape_status, "ok")
        self.assertEqual(client.review_limit, 50)
        self.assertNotIn("Design", product_row)
        self.assertNotIn("Class", product_row)
        self.assertNotIn("Reviews Count", product_row)
        self.assertNotIn("Design", PRODUCT_COLUMNS)
        self.assertNotIn("Class", PRODUCT_COLUMNS)
        self.assertNotIn("Reviews Count", PRODUCT_COLUMNS)
        self.assertEqual(product_row["Program"], "TARGET PILLOWFORT COTTON SHEETS")
        self.assertEqual(product_row["Status"], "CF")
        self.assertEqual(
            [review["Review Text"] for review in exported_reviews],
            ["Start month review", "End month review"],
        )
        self.assertEqual(exported_reviews[0]["Review Rating"], 4)
        self.assertEqual(exported_reviews[0]["Review Title"], "Start month title")
        self.assertEqual(exported_reviews[0]["Review Text"], "Start month review")
        self.assertEqual(exported_reviews[0]["Date Posted"], "01/08/2026")
        self.assertEqual(exported_reviews[0]["Sentiment"], "Positive")
        self.assertEqual(exported_reviews[0]["Sentiment Score"], 0.8)
        self.assertEqual(
            exported_reviews[0]["Satisfaction Percentage"],
            90.0,
        )

    def test_process_one_exports_current_month_for_selected_dpcis(self):
        today = date.today()
        previous_month_end = today.replace(day=1) - timedelta(days=1)

        class FakeClient:
            def scrape_dpci(self, _dpci, *, review_limit):
                return {
                    "tcin": "123",
                    "product": {},
                    "reviews": [
                        {
                            "id": "this-month",
                            "title": "Current month title",
                            "text": "Current month review",
                            "submitted_at": f"{today.isoformat()}T12:00:00Z",
                        },
                        {
                            "id": "previous-month",
                            "title": "Old title",
                            "text": "Previous month review",
                            "submitted_at": f"{previous_month_end.isoformat()}T12:00:00Z",
                        },
                    ],
                }

        with patch("src.main.parse_product", return_value={}), patch(
            "src.main.aggregate_product_row",
            side_effect=lambda _product, _reviews, **_kwargs: {
                "DPCI": "060-15-0108",
                "Avg. Rating": None,
                "Top Remark": "",
            },
        ):
            product_row, reviews, status = process_one(
                FakeClient(),
                "060-15-0108",
                run_date="2026-10-17T00:00:00Z",
                save_reviews=True,
            )

        self.assertEqual(status, "ok")
        self.assertEqual(product_row["DPCI"], "060-15-0108")
        self.assertEqual(len(reviews), 1)
        self.assertEqual(reviews[0]["Review Text"], "Current month review")
        self.assertEqual(
            reviews[0]["Date Posted"],
            today.strftime("%d/%m/%Y"),
        )

    def test_process_one_adds_matching_defect_fields_to_review(self):
        today = date.today()

        class FakeClient:
            def scrape_dpci(self, _dpci, *, review_limit):
                return {
                    "tcin": "123",
                    "product": {},
                    "reviews": [
                        {
                            "title": "Product feedback",
                            "text": "The sheets are not soft; seams unraveling.",
                            "submitted_at": f"{today.isoformat()}T12:00:00Z",
                        }
                    ],
                }

        with patch("src.main.parse_product", return_value={}), patch(
            "src.main.aggregate_product_row",
            side_effect=lambda _product, _reviews, **_kwargs: {
                "Avg. Rating": None,
                "Top Remark": "",
            },
        ), patch("src.review_pipeline._sentiment_score", return_value=-0.8):
            _product, reviews, status = process_one(
                FakeClient(),
                "060-15-0108",
                run_date="2026-10-17T00:00:00Z",
                save_reviews=True,
                defects_library=[
                    {
                        "Reason": "NOT SOFT",
                        "Sub Category": "POOR HANDFEEL",
                        "Main Category": "FABRIC QUALITY",
                    },
                    {
                        "Reason": "SEAMS UNRAVELING",
                        "Sub Category": "POOR STITCHING",
                        "Main Category": "WORKMANSHIP",
                    },
                ],
            )

        self.assertEqual(status, "ok")
        self.assertEqual(reviews[0]["Reason"], "NOT SOFT | SEAMS UNRAVELING")
        self.assertEqual(
            reviews[0]["Defect"],
            "POOR HANDFEEL | POOR STITCHING",
        )
        self.assertEqual(
            reviews[0]["Main Category"],
            "FABRIC QUALITY | WORKMANSHIP",
        )

    def test_process_one_sets_defect_fields_to_dash_for_positive_review(self):
        today = date.today()

        class FakeClient:
            def scrape_dpci(self, _dpci, *, review_limit):
                return {
                    "tcin": "123",
                    "product": {},
                    "reviews": [
                        {
                            "text": "The sheets are not soft.",
                            "submitted_at": f"{today.isoformat()}T12:00:00Z",
                        }
                    ],
                }

        with patch("src.main.parse_product", return_value={}), patch(
            "src.main.aggregate_product_row",
            side_effect=lambda _product, _reviews, **_kwargs: {
                "Avg. Rating": None,
                "Top Remark": "",
            },
        ), patch("src.review_pipeline._sentiment_score", return_value=0.8), patch(
            "src.main.match_defects"
        ) as match_defects:
            _product, reviews, status = process_one(
                FakeClient(),
                "060-15-0108",
                run_date="2026-10-17T00:00:00Z",
                save_reviews=True,
                defects_library=[
                    {
                        "Reason": "NOT SOFT",
                        "Sub Category": "POOR HANDFEEL",
                        "Main Category": "FABRIC QUALITY",
                    }
                ],
            )

        self.assertEqual(status, "ok")
        self.assertEqual(reviews[0]["Sentiment"], "Positive")
        self.assertEqual(reviews[0]["Reason"], "-")
        self.assertEqual(reviews[0]["Defect"], "-")
        self.assertEqual(reviews[0]["Main Category"], "-")
        match_defects.assert_not_called()

    def test_process_one_leaves_defect_fields_blank_for_neutral_review(self):
        today = date.today()

        class FakeClient:
            def scrape_dpci(self, _dpci, *, review_limit):
                return {
                    "tcin": "123",
                    "product": {},
                    "reviews": [
                        {
                            "text": "The sheets are not soft.",
                            "submitted_at": f"{today.isoformat()}T12:00:00Z",
                        }
                    ],
                }

        with patch("src.main.parse_product", return_value={}), patch(
            "src.main.aggregate_product_row",
            side_effect=lambda _product, _reviews, **_kwargs: {
                "Avg. Rating": None,
                "Top Remark": "",
            },
        ), patch("src.review_pipeline._sentiment_score", return_value=0.0), patch(
            "src.main.match_defects"
        ) as match_defects:
            _product, reviews, status = process_one(
                FakeClient(),
                "060-15-0108",
                run_date="2026-10-17T00:00:00Z",
                save_reviews=True,
                defects_library=[
                    {
                        "Reason": "NOT SOFT",
                        "Sub Category": "POOR HANDFEEL",
                        "Main Category": "FABRIC QUALITY",
                    }
                ],
            )

        self.assertEqual(status, "ok")
        self.assertEqual(reviews[0]["Sentiment"], "Neutral")
        self.assertEqual(reviews[0]["Reason"], "-")
        self.assertEqual(reviews[0]["Defect"], "-")
        self.assertEqual(reviews[0]["Main Category"], "-")
        match_defects.assert_not_called()

    def test_defect_matching_requires_whole_words_and_matches_phrases(self):
        defects = [
            {
                "Reason": "STAIN",
                "Sub Category": "STAIN",
                "Main Category": "PRODUCT QUALITY",
            },
            {
                "Reason": "NOT SOFT",
                "Sub Category": "POOR HANDFEEL",
                "Main Category": "FABRIC QUALITY",
            },
        ]

        self.assertEqual(
            match_defects("This fabric is not soft, but is stainless.", defects),
            {
                "Reason": "NOT SOFT",
                "Defect": "POOR HANDFEEL",
                "Main Category": "FABRIC QUALITY",
            },
        )

    def test_cli_parses_multiple_dpcis_and_inclusive_month_range(self):
        args = parse_args(
            [
                "--dpcis",
                "060-15-0108,097-02-2055",
                "--from-month",
                "2026-08",
                "--to-month",
                "2026-09",
            ]
        )

        self.assertEqual(args.dpcis, ["060-15-0108", "097-02-2055"])
        self.assertEqual(args.from_month, (2026, 8))
        self.assertEqual(args.to_month, (2026, 9))

    def test_cli_review_limit_is_independent_from_dpci_limit(self):
        args = parse_args(["--limit", "3", "--review-limit", "120"])

        self.assertEqual(args.limit, 3)
        self.assertEqual(args.review_limit, 120)

    def test_cli_parses_current_month_csv_mode(self):
        args = parse_args(["--current-month"])

        self.assertTrue(args.current_month)
        self.assertIsNone(args.dpci)
        self.assertIsNone(args.dpcis)
        self.assertIsNone(args.from_month)
        self.assertIsNone(args.to_month)

    def test_cli_rejects_non_positive_review_limit(self):
        with self.assertRaises(SystemExit):
            parse_args(["--review-limit", "0"])

    def test_selected_dpcis_are_validated_and_deduplicated(self):
        self.assertEqual(
            load_dpcis(
                Path("unused.csv"),
                only=["060-15-0108", "097-02-2055", "060-15-0108"],
            ),
            ["060-15-0108", "097-02-2055"],
        )
        with self.assertRaises(ValueError):
            load_dpcis(Path("unused.csv"), only=["invalid"])

    def test_write_reviews_excel_creates_xlsx_with_review_columns(self):
        with TemporaryDirectory() as directory:
            output_path = Path(directory) / "reviews.xlsx"
            write_reviews_excel(
                [
                    {
                        "DPCI": "060-15-0108",
                        "Review Rating": 5,
                        "Review Title": "Great",
                        "Review Text": "Works very well",
                    }
                ],
                output_path,
            )

            result = pd.read_excel(output_path)

        self.assertEqual(result.loc[0, "DPCI"], "060-15-0108")
        self.assertEqual(result.loc[0, "Review Rating"], 5)
        self.assertEqual(result.loc[0, "Review Title"], "Great")
        self.assertEqual(result.loc[0, "Review Text"], "Works very well")
        self.assertIn("Sentiment Score", result.columns)
        self.assertIn("Satisfaction Percentage", result.columns)

    def test_write_combined_excel_creates_products_and_reviews_sheets(self):
        with TemporaryDirectory() as directory:
            output_path = Path(directory) / "combined.xlsx"
            write_products_and_reviews_excel(
                [
                    {
                        "DPCI": "060-15-0108",
                        "Program": "BEDDING",
                        "Average Satisfaction %": 90.0,
                    }
                ],
                [
                    {
                        "DPCI": "060-15-0108",
                        "Review Text": "Comfortable",
                        "Reason": "NOT SOFT",
                        "Defect": "POOR HANDFEEL",
                        "Main Category": "FABRIC QUALITY",
                        "Sentiment Score": 0.8,
                        "Satisfaction Percentage": 90.0,
                    }
                ],
                output_path,
            )

            with pd.ExcelFile(output_path) as workbook:
                products = pd.read_excel(workbook, sheet_name="Products")
                reviews = pd.read_excel(workbook, sheet_name="Reviews")

        self.assertEqual(workbook.sheet_names, ["Products", "Reviews"])
        self.assertEqual(products.loc[0, "DPCI"], "060-15-0108")
        self.assertEqual(products.loc[0, "Average Satisfaction %"], 90.0)
        self.assertEqual(reviews.loc[0, "Review Text"], "Comfortable")
        self.assertEqual(reviews.loc[0, "Reason"], "NOT SOFT")
        self.assertEqual(reviews.loc[0, "Defect"], "POOR HANDFEEL")
        self.assertEqual(reviews.loc[0, "Main Category"], "FABRIC QUALITY")
        review_columns = list(reviews.columns)
        self.assertEqual(
            review_columns[review_columns.index("Review Text") + 1 :][:3],
            ["Reason", "Defect", "Main Category"],
        )
        self.assertEqual(reviews.loc[0, "Sentiment Score"], 0.8)
        self.assertEqual(reviews.loc[0, "Satisfaction Percentage"], 90.0)

    def test_main_writes_combined_workbook_without_separate_csv_files(self):
        with TemporaryDirectory() as directory:
            output_dir = Path(directory)
            product_row = {"DPCI": "060-15-0108", "Status": "CF"}
            review_row = {
                "DPCI": "060-15-0108",
                "Review Text": "Comfortable",
                "Sentiment Score": 0.8,
                "Satisfaction Percentage": 90.0,
            }
            with (
                patch("src.main.config.OUTPUT_DIR", output_dir),
                patch("src.main.config.RUN_LOG", output_dir / "run.log"),
                patch("src.main.setup_logging"),
                patch("src.main.load_dpcis", return_value=["060-15-0108"]),
                patch(
                    "src.main.load_dpci_library",
                    return_value={"060-15-0108": {"Program": "BEDDING", "Status": "CF"}},
                ),
                patch("src.main.load_defects_library", return_value=[]),
                patch("src.main.TargetClient"),
                patch(
                    "src.main.process_one",
                    return_value=(product_row, [review_row], "ok"),
                ) as process,
            ):
                self.assertEqual(
                    main(["--workbook-output", str(output_dir / "products_and_reviews.xlsx")]),
                    0,
                )

            output_path = output_dir / "products_and_reviews.xlsx"
            with pd.ExcelFile(output_path) as workbook:
                products = pd.read_excel(workbook, sheet_name="Products")
                reviews = pd.read_excel(workbook, sheet_name="Reviews")

        self.assertEqual(workbook.sheet_names, ["Products", "Reviews"])
        self.assertEqual(products.loc[0, "DPCI"], "060-15-0108")
        self.assertIn("Average Satisfaction %", products.columns)
        self.assertEqual(reviews.loc[0, "Review Text"], "Comfortable")
        self.assertEqual(reviews.loc[0, "Sentiment Score"], 0.8)
        self.assertEqual(reviews.loc[0, "Satisfaction Percentage"], 90.0)
        self.assertTrue(process.call_args.kwargs["save_reviews"])
        self.assertEqual(list(output_dir.glob("*.csv")), [])

    def test_main_current_month_mode_loads_csv_without_writing_separate_csv(self):
        with TemporaryDirectory() as directory:
            output_dir = Path(directory)
            product_row = {"DPCI": "060-15-0108", "Status": "CF"}
            review_row = {
                "DPCI": "060-15-0108",
                "Review Text": "Current month review",
            }
            with (
                patch("src.main.config.OUTPUT_DIR", output_dir),
                patch("src.main.config.RUN_LOG", output_dir / "run.log"),
                patch(
                    "src.main.config.PRODUCTS_AND_REVIEWS_XLSX",
                    output_dir / "combined.xlsx",
                ),
                patch("src.main.setup_logging"),
                patch("src.main.load_dpcis", return_value=["060-15-0108"]) as load_dpcis_mock,
                patch(
                    "src.main.load_dpci_library",
                    return_value={"060-15-0108": {"Program": "BEDDING", "Status": "CF"}},
                ),
                patch("src.main.load_defects_library", return_value=[]),
                patch("src.main.TargetClient"),
                patch(
                    "src.main.process_one",
                    return_value=(product_row, [review_row], "ok"),
                ) as process_mock,
                patch("src.main.write_products_and_reviews_excel"),
            ):
                self.assertEqual(main(["--current-month"]), 0)

        self.assertIsNone(load_dpcis_mock.call_args.kwargs["only"])
        month = date.today().year, date.today().month
        self.assertEqual(process_mock.call_args.kwargs["review_month_range"], (month, month))
        self.assertTrue(process_mock.call_args.kwargs["save_reviews"])
        self.assertEqual(list(output_dir.glob("*.csv")), [])

    def test_process_reviews_includes_posted_date_and_image_urls(self):
        reviews = process_reviews(
            [
                {
                    "id": "review-1",
                    "title": "Great product",
                    "text": "Works very well",
                    "submitted_at": "2026-10-14T12:00:00Z",
                    "photos": [
                        {"sizes": [{"url": "https://example.com/review.jpg"}]},
                        {"url": "//example.com/another.jpg"},
                    ],
                }
            ]
        )

        self.assertEqual(len(reviews), 1)
        self.assertEqual(reviews[0]["Date Posted"], "14/10/2026")
        self.assertEqual(
            reviews[0]["Image URLs"],
            "https://example.com/review.jpg | https://example.com/another.jpg",
        )
        self.assertEqual(reviews[0]["Rating"], None)
        self.assertEqual(reviews[0]["Title"], "Great product")
        self.assertEqual(reviews[0]["Text"], "Works very well")

    def test_process_reviews_reads_target_api_review_fields(self):
        reviews = process_reviews(
            [
                {
                    "id": "target-review",
                    "rating": 4,
                    "title": "Nice sheets",
                    "text": "Soft and comfortable",
                    "submitted_at": "2026-08-10T12:00:00Z",
                }
            ]
        )

        self.assertEqual(len(reviews), 1)
        self.assertEqual(reviews[0]["Rating"], 4)
        self.assertEqual(reviews[0]["Title"], "Nice sheets")
        self.assertEqual(reviews[0]["Text"], "Soft and comfortable")
        self.assertEqual(reviews[0]["Date Posted"], "10/08/2026")

    def test_process_reviews_formats_date_objects_without_time(self):
        reviews = process_reviews(
            [
                {
                    "text": "Review text",
                    "submitted_at": datetime(2026, 8, 10, 14, 30),
                }
            ]
        )

        self.assertEqual(reviews[0]["Date Posted"], "10/08/2026")

    def test_process_reviews_applies_neutral_and_five_star_overrides(self):
        with patch(
            "src.review_pipeline._sentiment_score",
            side_effect=[-0.1, -0.9],
        ):
            reviews = process_reviews(
                [
                    {"rating": 2, "text": "A review with a neutral score"},
                    {"rating": 5, "text": "A review with a negative score"},
                ]
            )

        self.assertEqual(reviews[0]["Sentiment"], "Neutral")
        self.assertEqual(reviews[0]["Sentiment Score"], -0.1)
        self.assertEqual(
            reviews[0]["Satisfaction Percentage"],
            45.0,
        )
        self.assertEqual(reviews[1]["Sentiment"], "Positive")
        self.assertEqual(reviews[1]["Sentiment Score"], 0.99)
        self.assertEqual(
            reviews[1]["Satisfaction Percentage"],
            99.5,
        )

    def test_product_aggregation_includes_average_satisfaction(self):
        row = aggregate_product_row(
            {},
            [
                {"Sentiment": "Positive", "Satisfaction Percentage": 90.0},
                {"Sentiment": "Negative", "Satisfaction Percentage": 20.0},
            ],
        )

        self.assertEqual(row["Average Satisfaction %"], 55.0)

    def test_fetch_reviews_matches_separate_photo_records(self):
        client = TargetClient.__new__(TargetClient)
        client.nova_key = "test-key"
        client._throttle = lambda: None
        client._mark_request = lambda: None
        client._get_json = lambda *_args, **_kwargs: {
            "reviews": {
                "results": [{"id": "review-1", "title": "Great", "text": "Works well"}],
                "last_page": True,
                "total_pages": 1,
            },
            "reviewsWithPhotos": {
                "results": [
                    {
                        "reviewId": "review-1",
                        "photos": [{"url": "https://example.com/review.jpg"}],
                    }
                ]
            },
        }

        reviews, _statistics = client.fetch_reviews("123")

        self.assertEqual(
            reviews[0]["photos"][0]["url"],
            "https://example.com/review.jpg",
        )

    def test_fetch_reviews_uses_available_pages_and_warns_when_api_caps_results(self):
        client = TargetClient.__new__(TargetClient)
        client.nova_key = "test-key"
        client._throttle = lambda: None
        client._mark_request = lambda: None
        requested_pages = []
        requested_sizes = set()

        def get_page(_url, params):
            page = int(params["page"])
            requested_pages.append(page)
            requested_sizes.add(params["size"])
            return {
                "reviews": {
                    "results": [{"id": f"review-{page}"}],
                    "last_page": False,
                    "total_pages": 60,
                    "total_results": 60,
                }
            }

        client._get_json = get_page

        with patch("src.target_client.config.REVIEW_PAGE_SIZE", 50):
            with self.assertLogs("scrapper", level="WARNING") as logs:
                reviews, _statistics = client.fetch_reviews("123", review_limit=3000)

        self.assertEqual(len(reviews), 51)
        self.assertEqual(requested_pages, list(range(51)))
        self.assertEqual(requested_sizes, {"50"})
        self.assertIn("Fetched 51 of 60 reviews", logs.output[0])

    def test_fetch_reviews_stops_at_requested_total_review_limit(self):
        client = TargetClient.__new__(TargetClient)
        client.nova_key = "test-key"
        client._throttle = lambda: None
        client._mark_request = lambda: None
        requested_pages = []

        def get_page(_url, params):
            page = int(params["page"])
            requested_pages.append(page)
            return {
                "reviews": {
                    "results": [{"id": f"review-{page}-{index}"} for index in range(50)],
                    "last_page": False,
                    "total_pages": 100,
                    "total_results": 4807,
                }
            }

        client._get_json = get_page

        reviews, _statistics = client.fetch_reviews("123", review_limit=75)

        self.assertEqual(len(reviews), 75)
        self.assertEqual(requested_pages, [0, 1])

    def test_fetch_reviews_default_limit_requests_only_one_page(self):
        client = TargetClient.__new__(TargetClient)
        client.nova_key = "test-key"
        client._throttle = lambda: None
        client._mark_request = lambda: None
        requested_pages = []

        def get_page(_url, params):
            requested_pages.append(int(params["page"]))
            return {
                "reviews": {
                    "results": [{"id": str(index)} for index in range(50)],
                    "last_page": False,
                    "total_pages": 100,
                    "total_results": 4807,
                }
            }

        client._get_json = get_page

        reviews, _statistics = client.fetch_reviews("123")

        self.assertEqual(len(reviews), 50)
        self.assertEqual(requested_pages, [0])


if __name__ == "__main__":
    unittest.main()
