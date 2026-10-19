"""Map Target PDP JSON into flat product fields."""

from __future__ import annotations

import html
from typing import Any

from src.dpci_loader import class_from_dpci


def parse_product(dpci: str, tcin: str, product: dict[str, Any]) -> dict[str, Any]:
    """Extract schema fields from a Redsky PDP product object."""
    item = product.get("item") or {}
    desc = item.get("product_description") or {}
    brand_obj = item.get("primary_brand") or {}
    classification = item.get("product_classification") or {}
    item_type_obj = classification.get("item_type") or {}

    category = product.get("category") or {}
    breadcrumbs = category.get("breadcrumbs") or []
    # Prefer last non-root breadcrumb as Item Type fallback
    breadcrumb_names = [
        b.get("name", "").strip()
        for b in breadcrumbs
        if isinstance(b, dict) and b.get("name") and b.get("name", "").lower() != "target"
    ]

    item_type = (
        item_type_obj.get("name")
        or (breadcrumb_names[-1] if breadcrumb_names else "")
        or classification.get("product_type_name")
        or ""
    )

    title = html.unescape(str(desc.get("title") or "")).strip()
    brand = html.unescape(str(brand_obj.get("name") or "")).strip()

    # Prefer API rating summary if present
    ratings = (product.get("ratings_and_reviews") or {}).get("statistics") or {}
    rating_block = ratings.get("rating") or {}
    avg_rating = rating_block.get("average")
    review_count = ratings.get("review_count")

    # Confirm DPCI from payload when available
    payload_dpci = item.get("dpci") or dpci

    return {
        "DPCI": payload_dpci or dpci,
        "Item Description": title,
        "Program": "",  # not available on public PDP; reserved for future mapping
        "Design": "",
        "Class": class_from_dpci(payload_dpci or dpci),
        "Item Type": str(item_type).strip(),
        "Brand": brand,
        "TCIN": tcin,
        "api_avg_rating": avg_rating,
        "api_review_count": review_count,
        "breadcrumbs": breadcrumb_names,
    }
