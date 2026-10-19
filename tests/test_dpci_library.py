import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

import pandas as pd

from src.dpci_library import load_dpci_library


class DpciLibraryTests(unittest.TestCase):
    def test_loads_program_and_status_by_dpci(self):
        with TemporaryDirectory() as directory:
            workbook = Path(directory) / "library.xlsx"
            pd.DataFrame(
                [
                    {
                        "DPCI": "060-15-0108",
                        "PROGRAM": "BEDDING",
                        "STATUS": "CF",
                    }
                ]
            ).to_excel(workbook, index=False)

            result = load_dpci_library(workbook)

        self.assertEqual(
            result["060-15-0108"],
            {"Program": "BEDDING", "Status": "CF"},
        )

    def test_missing_status_column_is_rejected(self):
        with TemporaryDirectory() as directory:
            workbook = Path(directory) / "library.xlsx"
            pd.DataFrame(
                [{"DPCI": "060-15-0108", "PROGRAM": "BEDDING"}]
            ).to_excel(workbook, index=False)

            with self.assertRaisesRegex(ValueError, "STATUS"):
                load_dpci_library(workbook)

    def test_duplicate_dpci_rows_are_rejected(self):
        with TemporaryDirectory() as directory:
            workbook = Path(directory) / "library.xlsx"
            pd.DataFrame(
                [
                    {"DPCI": "060-15-0108", "PROGRAM": "A", "STATUS": "CF"},
                    {"DPCI": "060-15-0108", "PROGRAM": "B", "STATUS": "NCF"},
                ]
            ).to_excel(workbook, index=False)

            with self.assertRaisesRegex(ValueError, "Duplicate DPCI"):
                load_dpci_library(workbook)


if __name__ == "__main__":
    unittest.main()
