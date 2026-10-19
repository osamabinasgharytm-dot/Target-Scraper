"""Load and match review text against the defects library workbook."""

from __future__ import annotations

import logging
import re
from pathlib import Path
from typing import Any

import pandas as pd

logger = logging.getLogger("scrapper")

_REQUIRED_COLUMNS = ("Reason", "Sub Category", "Main Category")
_WORD_RE = re.compile(r"[a-z0-9]+")


def _tokens(value: Any) -> tuple[str, ...]:
    return tuple(_WORD_RE.findall(str(value).casefold()))


def load_defects_library(path: Path) -> list[dict[str, str]]:
    """Load unique reason-to-category mappings from the defects workbook."""
    if not path.exists():
        raise FileNotFoundError(f"Defects library workbook not found: {path}")

    frame = pd.read_excel(path, dtype=str)
    columns = {str(column).strip().casefold(): column for column in frame.columns}
    missing = [field for field in _REQUIRED_COLUMNS if field.casefold() not in columns]
    if missing:
        raise ValueError(
            "Defects library must contain columns: Reason, Sub Category, and "
            f"Main Category. Missing: {', '.join(missing)}"
        )

    library: list[dict[str, str]] = []
    seen: dict[tuple[str, ...], tuple[str, str, str]] = {}
    for _, row in frame.iterrows():
        values = tuple(
            ""
            if pd.isna(row[columns[field.casefold()]])
            else str(row[columns[field.casefold()]]).strip()
            for field in _REQUIRED_COLUMNS
        )
        reason, sub_category, main_category = values
        if not reason:
            continue
        reason_tokens = _tokens(reason)
        if not reason_tokens:
            continue
        previous = seen.get(reason_tokens)
        if previous is not None:
            if tuple(value.casefold() for value in previous) != tuple(
                value.casefold() for value in values
            ):
                raise ValueError(
                    f"Conflicting defect library mappings for reason: {reason}"
                )
            continue
        seen[reason_tokens] = values
        library.append(
            {
                "Reason": reason,
                "Sub Category": sub_category,
                "Main Category": main_category,
            }
        )

    logger.info("Loaded %d defect reasons from %s", len(library), path)
    return library


def match_defects(
    review_text: str,
    defects_library: list[dict[str, str]],
) -> dict[str, str]:
    """Find whole-word or contiguous-phrase defect reasons in review text."""
    review_tokens = _tokens(review_text)
    matches: list[dict[str, str]] = []
    for defect in defects_library:
        reason_tokens = _tokens(defect.get("Reason", ""))
        if not reason_tokens:
            continue
        width = len(reason_tokens)
        if any(
            review_tokens[index : index + width] == reason_tokens
            for index in range(len(review_tokens) - width + 1)
        ):
            matches.append(defect)

    def join_unique(field: str) -> str:
        values = dict.fromkeys(
            defect[field].strip()
            for defect in matches
            if defect.get(field, "").strip()
        )
        return " | ".join(values)

    return {
        "Reason": join_unique("Reason"),
        "Defect": join_unique("Sub Category"),
        "Main Category": join_unique("Main Category"),
    }
