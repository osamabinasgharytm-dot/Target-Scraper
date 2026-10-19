"""Load product metadata mapped by DPCI from the DPCI library workbook."""

from __future__ import annotations

import logging
import re
from pathlib import Path

import pandas as pd

from src.config import DPCI_PATTERN

logger = logging.getLogger("scrapper")

_DPCI_COLUMNS = ("DPCI", "PROGRAM", "STATUS")
_DPCI_PATTERN = re.compile(DPCI_PATTERN)


def load_dpci_library(path: Path) -> dict[str, dict[str, str]]:
    """Return Program and Status values keyed by DPCI."""
    if not path.exists():
        raise FileNotFoundError(f"DPCI library workbook not found: {path}")

    frame = pd.read_excel(path, dtype=str)
    columns = {
        str(column).strip().upper(): column
        for column in frame.columns
    }
    missing_required = [column for column in _DPCI_COLUMNS if column not in columns]
    if missing_required:
        raise ValueError(
            "DPCI library must contain columns: DPCI, PROGRAM, and STATUS. "
            f"Missing: {', '.join(missing_required)}"
        )

    library: dict[str, dict[str, str]] = {}
    for _, row in frame.iterrows():
        dpci = str(row[columns["DPCI"]] or "").strip()
        if not dpci or dpci.lower() == "nan":
            continue
        if not _DPCI_PATTERN.fullmatch(dpci):
            logger.warning("Skipping invalid DPCI in library: %s", dpci)
            continue
        if dpci in library:
            raise ValueError(f"Duplicate DPCI in library workbook: {dpci}")

        values: dict[str, str] = {}
        for field in _DPCI_COLUMNS[1:]:
            column = columns.get(field)
            value = row[column] if column is not None else ""
            values[field.title()] = "" if pd.isna(value) else str(value).strip()
        library[dpci] = values

    logger.info("Loaded %d DPCI metadata rows from %s", len(library), path)
    return library
