import unittest
from datetime import date
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch

from src import config
from src.main import _default_workbook_path, _previous_calendar_month, main, parse_args


class MainTests(unittest.TestCase):
    def test_previous_calendar_month_handles_year_boundary(self):
        self.assertEqual(_previous_calendar_month(date(2026, 1, 1)), (2025, 12))

    def test_previous_calendar_month_uses_previous_month(self):
        self.assertEqual(_previous_calendar_month(date(2026, 10, 23)), (2026, 9))

    def test_multi_month_export_names_workbook_for_month_range(self):
        with patch("src.main.config.OUTPUT_DIR", Path("output")):
            workbook = _default_workbook_path(((2026, 8), (2026, 9)))

        self.assertEqual(
            workbook,
            Path("output") / "August-September 2026 Target Data.xlsx",
        )

    def test_last_month_flag_is_available(self):
        args = parse_args(["--last-month"])

        self.assertTrue(args.last_month)

    def test_last_month_scrapes_every_library_dpci_and_exports_previous_month(self):
        library = {
            "060-15-0108": {"Program": "A", "Status": "CF"},
            "097-02-2055": {"Program": "B", "Status": "NCF"},
        }
        with TemporaryDirectory() as directory:
            output_directory = Path(directory)
            with (
                patch("src.main.config.OUTPUT_DIR", output_directory),
                patch("src.main.setup_logging"),
                patch("src.main.load_dpci_library", return_value=library),
                patch("src.main.load_dpcis", return_value=list(library)) as load_dpcis,
                patch("src.main.load_defects_library", return_value=[]),
                patch("src.main.TargetClient"),
                patch("src.main.process_one", return_value=({}, [], "ok")) as process_one,
                patch(
                    "src.main.write_products_and_reviews_excel"
                ) as write_combined,
                patch("src.main.write_reviews_excel") as write_reviews,
            ):
                result = main(["--last-month"])

        today = date.today()
        previous_month = _previous_calendar_month(today)
        self.assertEqual(result, 0)
        load_dpcis.assert_called_once_with(limit=None, only=list(library))
        self.assertEqual(process_one.call_count, len(library))
        for scrape_call in process_one.call_args_list:
            self.assertEqual(
                scrape_call.kwargs["review_month_range"],
                (previous_month, previous_month),
            )
            self.assertEqual(
                scrape_call.kwargs["review_limit"],
                config.MAX_API_REVIEW_LIMIT,
            )
        expected_workbook = (
            output_directory
            / f"{date(previous_month[0], previous_month[1], 1):%B} Target Data.xlsx"
        )
        write_combined.assert_called_once_with([{}, {}], [], expected_workbook)
        write_reviews.assert_called_once_with([], config.REVIEWS_RANGE_XLSX)


if __name__ == "__main__":
    unittest.main()
