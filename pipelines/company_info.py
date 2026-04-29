"""Company profile pipeline preview."""

from __future__ import annotations


async def fetch_company_profiles(
    company_urls: list[str],
    sector: str,
) -> list[dict[str, object]]:
    """Fetch, parse, normalize, and validate company rows privately."""
    raise NotImplementedError("Private implementation omitted from public preview.")
