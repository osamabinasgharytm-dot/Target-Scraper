"""Aggregate review-level NLP into per-DPCI product rows."""

from __future__ import annotations

from collections import Counter
from datetime import datetime, timezone
from typing import Any

from src.config import PRODUCT_COLUMNS


def _pct(count: int, total: int) -> float:
    if total <= 0:
        return 0.0
    return round(100.0 * count / total, 2)


def _avg_rating(reviews: list[dict[str, Any]], fallback: Any) -> float | None:
    ratings: list[float] = []
    for rev in reviews:
        value = rev.get("Rating")
        if value is None:
            continue
        try:
            ratings.append(float(value))
        except (TypeError, ValueError):
            continue
    if ratings:
        return round(sum(ratings) / len(ratings), 2)
    if fallback is None or fallback == "":
        return None
    try:
        return round(float(fallback), 2)
    except (TypeError, ValueError):
        return None


def aggregate_product_row(
    product_fields: dict[str, Any],
    processed_reviews: list[dict[str, Any]],
    *,
    status: str = "ok",
    run_date: str | None = None,
) -> dict[str, Any]:
    """Build one final product summary row."""
    run_date = run_date or datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    total = len(processed_reviews)
    sentiments = Counter(r.get("Sentiment", "Neutral") for r in processed_reviews)
    remarks = [r["Remark"] for r in processed_reviews if r.get("Remark")]
    top_remark = Counter(remarks).most_common(1)[0][0] if remarks else ""
    satisfaction_scores = [
        float(r["Satisfaction Percentage"])
        for r in processed_reviews
        if r.get("Satisfaction Percentage") is not None
    ]

    row = {
        "DPCI": product_fields.get("DPCI", ""),
        "Item Description": product_fields.get("Item Description", ""),
        "Program": product_fields.get("Program", ""),
        "Item Type": product_fields.get("Item Type", ""),
        "Brand": product_fields.get("Brand", ""),
        "Avg. Rating": _avg_rating(processed_reviews, product_fields.get("api_avg_rating")),
        "Positive %": _pct(sentiments.get("Positive", 0), total),
        "Negative %": _pct(sentiments.get("Negative", 0), total),
        "Neutral %": _pct(sentiments.get("Neutral", 0), total),
        "Average Satisfaction %": (
            round(sum(satisfaction_scores) / len(satisfaction_scores), 2)
            if satisfaction_scores
            else None
        ),
        "Top Remark": top_remark,
        "Run Date": run_date,
        "Status": product_fields.get("Status", status),
        "TCIN": product_fields.get("TCIN", ""),
    }
    # Ensure all expected columns exist
    return {col: row.get(col, "") for col in PRODUCT_COLUMNS}


def empty_error_row(dpci: str, status: str, run_date: str | None = None) -> dict[str, Any]:
    """Row for not_found / error cases."""
    return aggregate_product_row(
        {
            "DPCI": dpci,
            "Item Description": "",
            "Program": "",
            "Item Type": "",
            "Brand": "",
            "TCIN": "",
            "api_avg_rating": None,
        },
        [],
        status=status,
        run_date=run_date,
    )
