"""Validate and deduplicate DPCI identifiers, or load them from a legacy CSV."""

from __future__ import annotations

import logging
import re
from pathlib import Path

import pandas as pd

from src.config import DPCI_PATTERN

logger = logging.getLogger("scrapper")
_DPCI_RE = re.compile(DPCI_PATTERN)


def load_dpcis(
    csv_path: Path | None = None,
    *,
    limit: int | None = None,
    only: str | list[str] | None = None,
) -> list[str]:
    """Return a clean, deduped list of DPCIs.

    Args:
        csv_path: Optional legacy CSV path with a DPCI column. Omit when `only` is set.
        limit: Optional max count after validation/dedupe (for test runs).
        only: If set, return only these DPCIs (each must be valid format).
    """
    if only is not None:
        requested = [only] if isinstance(only, str) else only
        selected: list[str] = []
        seen: set[str] = set()
        for value in requested:
            dpci = value.strip()
            if not _DPCI_RE.match(dpci):
                raise ValueError(f"Invalid DPCI format: {value!r} (expected ###-##-####)")
            if dpci not in seen:
                selected.append(dpci)
                seen.add(dpci)
        if limit is not None:
            if limit < 1:
                raise ValueError("--limit must be >= 1")
            selected = selected[:limit]
        return selected

    if csv_path is None:
        raise ValueError("Provide a CSV path or an explicit DPCI list")
    if not csv_path.exists():
        raise FileNotFoundError(f"DPCI CSV not found: {csv_path}")

    df = pd.read_csv(csv_path, dtype=str)
    if "DPCI" not in df.columns:
        raise ValueError(f"CSV must contain a 'DPCI' column. Found: {list(df.columns)}")

    raw = [str(x).strip() for x in df["DPCI"].tolist() if str(x).strip() and str(x).strip().lower() != "nan"]
    valid: list[str] = []
    invalid = 0
    for value in raw:
        if _DPCI_RE.match(value):
            valid.append(value)
        else:
            invalid += 1
            logger.warning("Skipping invalid DPCI: %s", value)

    # Preserve first-seen order while deduping
    seen: set[str] = set()
    deduped: list[str] = []
    for value in valid:
        if value not in seen:
            seen.add(value)
            deduped.append(value)

    duplicates_removed = len(valid) - len(deduped)
    logger.info(
        "Loaded DPCIs: raw=%d valid=%d duplicates_removed=%d invalid=%d unique=%d",
        len(raw),
        len(valid),
        duplicates_removed,
        invalid,
        len(deduped),
    )

    if limit is not None:
        if limit < 1:
            raise ValueError("--limit must be >= 1")
        deduped = deduped[:limit]
        logger.info("Applying --limit %d -> %d DPCIs", limit, len(deduped))

    return deduped


def class_from_dpci(dpci: str) -> str:
    """Middle segment of DPCI (dept-class-item)."""
    parts = dpci.split("-")
    return parts[1] if len(parts) == 3 else ""
