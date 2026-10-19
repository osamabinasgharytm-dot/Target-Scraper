"""HTTP client for Target Redsky (catalog) and r2d2 (reviews) APIs."""

from __future__ import annotations

import logging
import re
import time
import uuid
from typing import Any

import requests
from tenacity import (
    retry,
    retry_if_exception_type,
    stop_after_attempt,
    wait_exponential,
)

from src import config

logger = logging.getLogger("scrapper")

_KEY_RE = re.compile(r"(?:key=|\"apiKey\"\s*:\s*\"|\"key\"\s*:\s*\")([a-f0-9]{40})")


class TargetAPIError(Exception):
    """Raised for retryable HTTP/API failures."""

    def __init__(self, message: str, status_code: int | None = None):
        super().__init__(message)
        self.status_code = status_code


class NotFoundError(Exception):
    """Product or DPCI could not be resolved."""


class TargetClient:
    def __init__(self) -> None:
        self.session = requests.Session()
        self.session.headers.update(
            {
                "User-Agent": config.USER_AGENT,
                "Accept": "application/json,text/html,*/*",
                "Accept-Language": "en-US,en;q=0.9",
                "Referer": "https://www.target.com/",
                "Origin": "https://www.target.com",
            }
        )
        self.redsky_key = config.DEFAULT_REDSKY_KEY
        self.nova_key = config.DEFAULT_NOVA_KEY
        self.visitor_id = uuid.uuid4().hex.upper()
        self._last_request_at = 0.0
        self._bootstrap()

    def _throttle(self) -> None:
        elapsed = time.monotonic() - self._last_request_at
        wait_for = config.REQUEST_DELAY_SECONDS - elapsed
        if wait_for > 0:
            time.sleep(wait_for)

    def _mark_request(self) -> None:
        self._last_request_at = time.monotonic()

    def _bootstrap(self) -> None:
        """Warm cookies and try to discover live API keys / visitor id."""
        try:
            self._throttle()
            resp = self.session.get("https://www.target.com/", timeout=config.REQUEST_TIMEOUT)
            self._mark_request()
            if resp.status_code == 200:
                keys = _KEY_RE.findall(resp.text)
                if keys:
                    # Prefer the well-known Redsky key if present; else first key found.
                    if config.DEFAULT_REDSKY_KEY in keys:
                        self.redsky_key = config.DEFAULT_REDSKY_KEY
                    else:
                        self.redsky_key = keys[0]
                    logger.info("Bootstrapped Redsky key from Target homepage")

                vid = self.session.cookies.get("visitorId")
                if vid:
                    self.visitor_id = vid
                    logger.info("Using visitorId cookie: %s...", vid[:12])

                # Nova (r2d2) key sometimes appears in page bundles
                if config.DEFAULT_NOVA_KEY in resp.text:
                    self.nova_key = config.DEFAULT_NOVA_KEY
            else:
                logger.warning("Homepage bootstrap returned HTTP %s", resp.status_code)
        except requests.RequestException as exc:
            logger.warning("Homepage bootstrap failed (%s); using default keys", exc)

    @retry(
        retry=retry_if_exception_type(TargetAPIError),
        stop=stop_after_attempt(config.MAX_RETRIES),
        wait=wait_exponential(multiplier=1, min=2, max=20),
        reraise=True,
    )
    def _get_json(self, url: str, params: dict[str, Any] | None = None) -> dict[str, Any]:
        self._throttle()
        try:
            resp = self.session.get(url, params=params, timeout=config.REQUEST_TIMEOUT)
        finally:
            self._mark_request()

        if resp.status_code in (429, 500, 502, 503, 504):
            raise TargetAPIError(f"Retryable HTTP {resp.status_code} for {url}", resp.status_code)
        if resp.status_code == 404:
            raise NotFoundError(f"HTTP 404 for {url}")
        if resp.status_code >= 400:
            raise TargetAPIError(
                f"HTTP {resp.status_code}: {resp.text[:200]}",
                resp.status_code,
            )
        try:
            return resp.json()
        except ValueError as exc:
            raise TargetAPIError(f"Invalid JSON from {url}") from exc

    def resolve_dpci(self, dpci: str) -> str:
        """Resolve DPCI → TCIN via Redsky search; require exact DPCI match."""
        candidates = [dpci, dpci.replace("-", "")]
        for keyword in candidates:
            products = self._search_products(keyword)
            for product in products:
                item = product.get("item") or {}
                if item.get("dpci") == dpci:
                    tcin = str(product.get("tcin") or item.get("tcin") or "")
                    if tcin:
                        logger.debug("Resolved %s -> TCIN %s", dpci, tcin)
                        return tcin
                # Some payloads nest child variants
                for child in item.get("children") or []:
                    if child.get("dpci") == dpci:
                        tcin = str(child.get("tcin") or "")
                        if tcin:
                            return tcin
        raise NotFoundError(f"No Target product found for DPCI {dpci}")

    def _search_products(self, keyword: str) -> list[dict[str, Any]]:
        params = {
            "key": self.redsky_key,
            "channel": "WEB",
            "count": "24",
            "default_purchasability_filter": "false",
            "include_sponsored": "false",
            "offset": "0",
            "page": f"/s/{keyword}",
            "platform": "desktop",
            "pricing_store_id": config.STORE_ID,
            "keyword": keyword,
            "visitor_id": self.visitor_id,
        }
        try:
            data = self._get_json(
                "https://redsky.target.com/redsky_aggregations/v1/web/plp_search_v2",
                params=params,
            )
        except NotFoundError:
            return []
        search = (data.get("data") or {}).get("search") or {}
        products = search.get("products")
        if products is None:
            # Some responses nest under search_response
            products = (search.get("search_response") or {}).get("items") or []
        return products if isinstance(products, list) else []

    def fetch_product(self, tcin: str) -> dict[str, Any]:
        """Fetch full PDP payload for a TCIN."""
        params = {
            "key": self.redsky_key,
            "tcin": tcin,
            "store_id": config.STORE_ID,
            "pricing_store_id": config.STORE_ID,
            "channel": "WEB",
            "page": f"/p/A-{tcin}",
            "zip": config.ZIP,
            "state": config.STATE,
            "latitude": config.LATITUDE,
            "longitude": config.LONGITUDE,
        }
        data = self._get_json(
            "https://redsky.target.com/redsky_aggregations/v1/web/pdp_client_v1",
            params=params,
        )
        product = (data.get("data") or {}).get("product")
        if not product:
            raise NotFoundError(f"Empty product payload for TCIN {tcin}")
        return product

    def fetch_reviews(
        self,
        tcin: str,
        review_limit: int | None = None,
    ) -> tuple[list[dict[str, Any]], dict[str, Any]]:
        """Fetch paginated reviews + statistics for a TCIN.

        review_limit caps the total number fetched, not the number per page.
        Returns (review_dicts, statistics_dict).
        """
        review_limit = config.REVIEW_LIMIT if review_limit is None else review_limit
        if review_limit < 1:
            raise ValueError("review_limit must be >= 1")

        all_reviews: list[dict[str, Any]] = []
        statistics: dict[str, Any] = {}
        total_pages = 1
        total_results = 0
        page_size = min(max(config.REVIEW_PAGE_SIZE, 1), 50, review_limit)
        max_pages = min((review_limit + page_size - 1) // page_size, 51)

        # The API accepts page indexes 0 through 50 and rejects higher pages.
        for page in range(max_pages):
            params = {
                "key": self.nova_key,
                "hasOnlyPhotos": "false",
                "includes": "reviews,reviewsWithPhotos,entities,metadata,statistics",
                "page": str(page),
                "entity": "",
                "reviewedId": tcin,
                "reviewType": "PRODUCT",
                "size": str(page_size),
                "sortBy": "most_recent",
                "verifiedOnly": "false",
            }
            data = self._get_json(
                "https://r2d2.target.com/ratings_reviews_api/v1/summary",
                params=params,
            )

            if page == 0:
                statistics = data.get("statistics") or {}

            reviews_block = data.get("reviews") or {}
            if isinstance(reviews_block, dict):
                batch = reviews_block.get("results") or []
                last_page = bool(reviews_block.get("last_page"))
                total_pages = max(1, int(reviews_block.get("total_pages") or 1))
                total_results = int(reviews_block.get("total_results") or 0)
            else:
                batch = reviews_block if isinstance(reviews_block, list) else []
                last_page = True
                total_pages = 1

            photos_block = data.get("reviewsWithPhotos") or {}
            if isinstance(photos_block, dict):
                photo_reviews = photos_block.get("results") or []
            else:
                photo_reviews = photos_block if isinstance(photos_block, list) else []
            if not isinstance(photo_reviews, list):
                photo_reviews = []

            photos_by_review_id: dict[str, dict[str, Any]] = {}
            for photo_review in photo_reviews:
                if not isinstance(photo_review, dict):
                    continue
                photo_review_id = (
                    photo_review.get("reviewId")
                    or photo_review.get("review_id")
                    or photo_review.get("id")
                )
                if photo_review_id:
                    photos_by_review_id[str(photo_review_id)] = photo_review

            batch = batch[: review_limit - len(all_reviews)]
            for review in batch:
                review_id = str(
                    review.get("id")
                    or review.get("external_id")
                    or review.get("reviewId")
                    or ""
                )
                photo_review = photos_by_review_id.get(review_id)
                if photo_review:
                    review_photos = (
                        photo_review.get("photos")
                        or photo_review.get("images")
                        or photo_review.get("reviewImages")
                    )
                    if review_photos and not any(
                        review.get(key)
                        for key in ("photos", "images", "reviewImages", "review_images")
                    ):
                        review["photos"] = review_photos

            all_reviews.extend(batch)
            if (
                len(all_reviews) >= review_limit
                or last_page
                or page >= total_pages - 1
                or not batch
            ):
                break

        if len(all_reviews) < total_results and max_pages == 51 and review_limit > 51 * page_size:
            logger.warning(
                "Fetched %d of %d reviews for TCIN %s; Target's API page limit is 0-50 "
                "(%d reviews per page).",
                len(all_reviews),
                total_results,
                tcin,
                page_size,
            )

        logger.info(
            "Fetched %d reviews for TCIN %s (requested review limit: %d)",
            len(all_reviews),
            tcin,
            review_limit,
        )
        return all_reviews, statistics

    def scrape_dpci(
        self,
        dpci: str,
        review_limit: int | None = None,
    ) -> dict[str, Any]:
        """Resolve DPCI and return product + reviews bundle."""
        tcin = self.resolve_dpci(dpci)
        product = self.fetch_product(tcin)
        reviews, statistics = self.fetch_reviews(tcin, review_limit=review_limit)
        return {
            "dpci": dpci,
            "tcin": tcin,
            "product": product,
            "reviews": reviews,
            "statistics": statistics,
        }
