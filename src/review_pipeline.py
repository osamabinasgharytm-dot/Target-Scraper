"""Clean reviews, run VADER sentiment, and extract one-word remarks."""

from __future__ import annotations

import re
from collections import Counter
from datetime import date, datetime
from typing import Any

from vaderSentiment.vaderSentiment import SentimentIntensityAnalyzer

NEUTRAL_THRESHOLD = 0.2
POSITIVE_RATING_THRESHOLD = 4.5

_analyzer = SentimentIntensityAnalyzer()

_HTML_RE = re.compile(r"<[^>]+>")
_NON_WORD_RE = re.compile(r"[^a-z0-9\s\-']")
_WS_RE = re.compile(r"\s+")

STOPWORDS = frozenset(
    {
        "a",
        "an",
        "the",
        "and",
        "or",
        "but",
        "if",
        "in",
        "on",
        "at",
        "to",
        "for",
        "of",
        "as",
        "by",
        "with",
        "from",
        "is",
        "are",
        "was",
        "were",
        "be",
        "been",
        "being",
        "it",
        "this",
        "that",
        "these",
        "those",
        "i",
        "me",
        "my",
        "we",
        "our",
        "you",
        "your",
        "he",
        "she",
        "they",
        "them",
        "their",
        "his",
        "her",
        "not",
        "no",
        "yes",
        "so",
        "too",
        "very",
        "just",
        "than",
        "then",
        "there",
        "here",
        "what",
        "which",
        "who",
        "whom",
        "when",
        "where",
        "why",
        "how",
        "all",
        "any",
        "both",
        "each",
        "few",
        "more",
        "most",
        "other",
        "some",
        "such",
        "only",
        "own",
        "same",
        "can",
        "will",
        "don",
        "should",
        "now",
        "have",
        "has",
        "had",
        "do",
        "does",
        "did",
        "would",
        "could",
        "about",
        "into",
        "over",
        "after",
        "before",
        "up",
        "down",
        "out",
        "off",
        "again",
        "further",
        "once",
        "also",
        "really",
        "get",
        "got",
        "one",
        "two",
        "like",
        "love",
        "loved",
        "great",
        "good",
        "bad",
        "nice",
        "well",
        "much",
        "many",
        "lot",
        "lots",
        "product",
        "item",
        "buy",
        "bought",
        "purchase",
        "purchased",
        "review",
        "reviews",
        "target",
        "amazon",
        "im",
        "ive",
        "id",
        "dont",
        "doesnt",
        "didnt",
        "wasnt",
        "werent",
        "cant",
        "wont",
        "its",
        "thats",
        "size",
        "color",
        "colour",
    }
)


def clean_review_text(text: str) -> str:
    if not text:
        return ""
    cleaned = _HTML_RE.sub(" ", text)
    cleaned = cleaned.replace("\xa0", " ")
    cleaned = _NON_WORD_RE.sub(" ", cleaned.lower())
    cleaned = _WS_RE.sub(" ", cleaned).strip()
    return cleaned


def sentiment_label(text: str) -> str:
    """Return the VADER sentiment label using the project's neutral threshold."""
    return _sentiment_label_from_score(_sentiment_score(text))


def _sentiment_label_from_score(normalized_score: float) -> str:
    if abs(normalized_score) <= NEUTRAL_THRESHOLD:
        return "Neutral"
    return "Positive" if normalized_score > 0 else "Negative"


def _sentiment_score(text: str) -> float:
    if not text:
        return 0.0
    return _analyzer.polarity_scores(text)["compound"]


def one_word_remark(cleaned_text: str) -> str:
    """Pick the most frequent non-stopword token from a review."""
    if not cleaned_text:
        return ""
    tokens = [
        t.strip("-'")
        for t in cleaned_text.split()
        if len(t.strip("-'")) >= 3 and t.strip("-'") not in STOPWORDS
    ]
    if not tokens:
        return ""
    return Counter(tokens).most_common(1)[0][0]


def _image_urls(value: Any) -> list[str]:
    """Extract unique image URLs from Target review photo objects."""
    urls: list[str] = []

    def collect(item: Any) -> None:
        if isinstance(item, str):
            url = item.strip()
            if url.startswith("//"):
                url = f"https:{url}"
            if url.startswith(("https://", "http://")) and url not in urls:
                urls.append(url)
        elif isinstance(item, dict):
            for nested in item.values():
                collect(nested)
        elif isinstance(item, list):
            for nested in item:
                collect(nested)

    for key in ("photos", "images", "reviewImages", "review_images"):
        collect(value.get(key))
    return urls


def _date_only(value: Any) -> str:
    if isinstance(value, datetime):
        return value.strftime("%d/%m/%Y")
    if isinstance(value, date):
        return value.strftime("%d/%m/%Y")
    if not isinstance(value, str):
        return ""

    value = value.strip()
    if not value:
        return ""
    try:
        parsed_date = datetime.fromisoformat(value.replace("Z", "+00:00")).date()
        return parsed_date.strftime("%d/%m/%Y")
    except ValueError:
        match = re.match(r"^(\d{4}-\d{2}-\d{2})(?:\b|T)", value)
        if match:
            try:
                parsed_date = date.fromisoformat(match.group(1))
                return parsed_date.strftime("%d/%m/%Y")
            except ValueError:
                pass
        return value


def process_reviews(raw_reviews: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Normalize Target review objects into cleaned/sentiment rows."""
    prepared: list[dict[str, Any]] = []
    for rev in raw_reviews:
        text = str(
            rev.get("text")
            or rev.get("Text")
            or rev.get("review text")
            or rev.get("Review Text")
            or ""
        ).strip()
        title = str(
            rev.get("title")
            or rev.get("Title")
            or rev.get("review title")
            or rev.get("Review Title")
            or ""
        ).strip()
        combined = f"{title} {text}".strip()
        cleaned = clean_review_text(combined)
        if not cleaned:
            continue

        rating = next(
            (
                rev[key]
                for key in ("rating", "Rating", "review rating", "Review Rating")
                if rev.get(key) is not None
            ),
            None,
        )

        posted_date = next(
            (
                rev[key]
                for key in (
                    "submitted_at",
                    "submittedAt",
                    "submissionTime",
                    "submission_time",
                    "postedDate",
                    "posted_date",
                    "reviewDate",
                    "review_date",
                    "createdAt",
                    "created_at",
                )
                if rev.get(key) is not None
            ),
            "",
        )
        posted_date = _date_only(posted_date)

        prepared.append(
            {
                "Review ID": str(rev.get("id") or rev.get("external_id") or ""),
                "Rating": rating,
                "Title": title,
                "Text": text,
                "Date Posted": posted_date,
                "Image URLs": " | ".join(_image_urls(rev)),
                "Clean Text": cleaned,
                "Remark": one_word_remark(cleaned),
            }
        )

    if not prepared:
        return []

    processed: list[dict[str, Any]] = []
    for review in prepared:
        normalized_score = _sentiment_score(review["Clean Text"])
        sentiment = _sentiment_label_from_score(normalized_score)

        rating = review["Rating"]
        rating_text = str(rating).strip()
        try:
            rating_value = float(rating)
        except (TypeError, ValueError):
            rating_value = None
        if rating_text.startswith("5") or (
            rating_value is not None and rating_value >= POSITIVE_RATING_THRESHOLD
        ):
            sentiment = "Positive"
            normalized_score = 0.99

        review["Sentiment"] = sentiment
        review["Sentiment Score"] = normalized_score
        review["Satisfaction Percentage"] = (
            (normalized_score + 1) / 2
        ) * 100
        processed.append(review)
    return processed
