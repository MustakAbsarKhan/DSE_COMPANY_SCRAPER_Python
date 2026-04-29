"""Parser architecture preview.

The private parser handles inconsistent DSE HTML by combining table scanning,
DOM traversal, defensive normalization, and fallback extraction. Detailed
selectors, index assumptions, and field-level extraction rules are intentionally
excluded from this public repository.
"""

from __future__ import annotations

from typing import Any


def parse_html(html: str) -> Any:
    """Convert raw HTML into a parser object in the private implementation."""
    raise NotImplementedError("Private implementation omitted from public preview.")


def extract_sectors(parsed_html: Any, ignored_sectors: list[str]) -> list[dict[str, str]]:
    """Return tradable sector names and relative URLs."""
    raise NotImplementedError("Private implementation omitted from public preview.")


def extract_company_urls(parsed_html: Any) -> list[str]:
    """Return company detail-page URLs discovered on a sector page."""
    raise NotImplementedError("Private implementation omitted from public preview.")


def extract_company_profile(parsed_html: Any, sector: str) -> dict[str, object]:
    """Return one normalized company row for workbook export."""
    raise NotImplementedError("Private implementation omitted from public preview.")


NORMALIZED_COMPANY_PROFILE_SCHEMA = [
    "Company Name",
    "Trading Code",
    "Scrip Code",
    "Sector",
    "Market Date",
    "LTP",
    "YCP",
    "Day Volume",
    "Market Cap (mn)",
    "Free Float Cap (mn)",
    "Paid-up Capital (mn)",
    "Latest EPS Used",
    "Latest P/E Used",
    "Latest NAVPS Used",
    "Dividend Yield %",
    "Sponsor/Director Holding %",
    "Institute Holding %",
    "Foreign Holding %",
    "Public Holding %",
    "Data Quality Notes",
]
