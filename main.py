"""
DSE Company Data Scraper
Author: Mohammad Mustak Absar Khan
GitHub: https://github.com/MustakAbsarKhan
"""

import asyncio

# Central configuration values for the DSE domain and sector filtering.
from config import DOMAIN, MAIN_URL, IGNORED_SECTORS

# Pipeline functions keep each scraping stage small and readable.
from pipelines.sectors import get_sectors
from pipelines.companies import get_companies
from pipelines.company_info import get_company_infos

# Final output writer.
from export.excel import save_to_excel

# Optional market holiday checker. It is currently disabled in main(), but kept
# here so the scraper can skip closed market days when needed.
from core.holidays import holiday_checker


# =========================
# PROCESS SINGLE SECTOR
# =========================
async def process_sector(sector):
    """Fetch all company links for one sector, then scrape those companies."""
    print(f"\n🔹 Processing Sector: {sector['name']}")

    # Step 1: open the sector page and collect all company detail-page links.
    companies = await get_companies(DOMAIN, sector["url"])
    company_count = len(companies)

    # print(f"   ➤ Companies Found: {company_count}")

    # Step 2: fetch and parse every company page found inside this sector.
    sector_data = await get_company_infos(
        DOMAIN,
        companies,
        sector["name"]
    )

    scraped_count = len(sector_data)

    print(f"   ✔ Companies Scraped: {scraped_count}")

    # Return both counts and data so main() can build a final summary.
    return {
        "sector": sector["name"],
        "found": company_count,
        "scraped": scraped_count,
        "data": sector_data
    }


# =========================
# MAIN
# =========================
async def main():
    """Main async orchestrator for the full scraping run."""
    # # 🔥 HOLIDAY CHECK - FIRST PRIORITY
    # # Uncomment this block if you want the scraper to stop automatically on
    # # Friday/Saturday or an official DSE holiday.
    # is_holiday = await holiday_checker.check_and_exit_if_holiday()
    # if is_holiday:
    #     return

    # Flat list of all company rows that will later be exported to Excel.
    all_data = []

    # Counters used only for the final terminal summary.
    total_sectors_found = 0
    total_sectors_scraped = 0
    total_companies_found = 0
    total_companies_scraped = 0

    # =========================
    # STEP 1: GET SECTORS
    # =========================
    # Fetch the DSE industry listing page and remove ignored sectors.
    sectors = await get_sectors(MAIN_URL, IGNORED_SECTORS)

    total_sectors_found = len(sectors)

    print("\n" + "=" * 60)
    print(f"TOTAL SECTORS FOUND: {total_sectors_found}")
    print("=" * 60)

    # =========================
    # 🔥 CONCURRENT SECTOR PROCESSING
    # =========================
    # Only process a few sectors at the same time. Inside each sector, company
    # pages are also fetched concurrently through the shared adaptive client.
    semaphore = asyncio.Semaphore(3)  # 🔥 CONTROL PARALLEL SECTORS

    async def sem_task(sector):
        """Wrap process_sector() with a sector-level concurrency limit."""
        async with semaphore:
            return await process_sector(sector)

    # Create one async task per sector.
    tasks = [sem_task(sector) for sector in sectors]

    # 🔥 PRESERVES ORDER
    # gather() waits for all sector tasks and returns results in task-list order.
    results = await asyncio.gather(*tasks)

    # =========================
    # COLLECT RESULTS
    # =========================
    for result in results:
        # Each result contains one sector's discovered links and scraped rows.
        total_sectors_scraped += 1
        total_companies_found += result["found"]
        total_companies_scraped += result["scraped"]

        all_data.extend(result["data"])

    # =========================
    # SAVE DATA
    # =========================
    # Convert the list of dictionaries into an Excel file.
    save_to_excel(all_data)

    # =========================
    # FINAL SUMMARY
    # =========================
    print("\n" + "=" * 60)
    print("FINAL SCRAPING SUMMARY")
    print("=" * 60)

    print(f"Total Sectors Found     : {total_sectors_found}")
    print(f"Total Sectors Scraped   : {total_sectors_scraped}")
    print(f"Total Companies Found   : {total_companies_found}")
    print(f"Total Companies Scraped : {total_companies_scraped}")

    print("=" * 60)


if __name__ == "__main__":
    # Start the async event loop when this file is run directly.
    asyncio.run(main())
