"""Public architecture preview for the DSE company data pipeline.

The production implementation is intentionally not included in this public
repository. This file shows the orchestration shape used by the private system:
sector discovery, company discovery, company-profile extraction, and workbook
export.
"""

from __future__ import annotations

import asyncio

from config import IGNORED_SECTORS, MAIN_URL
from export.excel import export_company_rows_to_excel
from pipelines.companies import fetch_company_urls_for_sector
from pipelines.company_info import fetch_company_profiles
from pipelines.sectors import fetch_tradable_sector_links


async def process_sector(sector: dict[str, str]) -> dict[str, object]:
    """Preview one sector-level unit of work."""
    company_urls = await fetch_company_urls_for_sector(sector["url"])
    company_rows = await fetch_company_profiles(company_urls, sector["name"])

    return {
        "sector": sector["name"],
        "found": len(company_urls),
        "scraped": len(company_rows),
        "data": company_rows,
    }


async def main() -> None:
    """Preview the private scraper's high-level async control flow."""
    sectors = await fetch_tradable_sector_links(MAIN_URL, IGNORED_SECTORS)

    semaphore = asyncio.Semaphore(3)

    async def bounded_sector_task(sector: dict[str, str]) -> dict[str, object]:
        async with semaphore:
            return await process_sector(sector)

    results = await asyncio.gather(*(bounded_sector_task(sector) for sector in sectors))
    rows = [row for result in results for row in result["data"]]

    export_company_rows_to_excel(rows)


if __name__ == "__main__":
    raise SystemExit(
        "This public repository is an architecture preview only. "
        "The runnable production scraper is kept private."
    )
