import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

import pandas as pd

from src.defect_library import load_defects_library


class DefectLibraryTests(unittest.TestCase):
    def test_loads_required_fields_and_deduplicates_identical_reasons(self):
        with TemporaryDirectory() as directory:
            workbook = Path(directory) / "defects.xlsx"
            pd.DataFrame(
                [
                    {
                        " reason ": "NOT SOFT",
                        "SUB CATEGORY": "POOR HANDFEEL",
                        "Main Category": "FABRIC QUALITY",
                    },
                    {
                        " reason ": " not soft ",
                        "SUB CATEGORY": "POOR HANDFEEL",
                        "Main Category": "FABRIC QUALITY",
                    },
                ]
            ).to_excel(workbook, index=False)

            defects = load_defects_library(workbook)

        self.assertEqual(
            defects,
            [
                {
                    "Reason": "NOT SOFT",
                    "Sub Category": "POOR HANDFEEL",
                    "Main Category": "FABRIC QUALITY",
                }
            ],
        )

    def test_rejects_conflicting_duplicate_reason_mappings(self):
        with TemporaryDirectory() as directory:
            workbook = Path(directory) / "defects.xlsx"
            pd.DataFrame(
                [
                    {
                        "Reason": "STAIN",
                        "Sub Category": "STAIN",
                        "Main Category": "PRODUCT QUALITY",
                    },
                    {
                        "Reason": " stain ",
                        "Sub Category": "SPOTS",
                        "Main Category": "FABRIC QUALITY",
                    },
                ]
            ).to_excel(workbook, index=False)

            with self.assertRaisesRegex(ValueError, "Conflicting defect library"):
                load_defects_library(workbook)

    def test_rejects_missing_required_columns(self):
        with TemporaryDirectory() as directory:
            workbook = Path(directory) / "defects.xlsx"
            pd.DataFrame([{"Reason": "STAIN"}]).to_excel(workbook, index=False)

            with self.assertRaisesRegex(ValueError, "Sub Category"):
                load_defects_library(workbook)


if __name__ == "__main__":
    unittest.main()
