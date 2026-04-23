import asyncio

from config import DOMAIN, MAIN_URL, IGNORED_SECTORS

from pipelines.sectors import get_sectors
from pipelines.companies import get_companies
from pipelines.company_info import get_company_infos

from export.excel import save_to_excel


# =========================
# PROCESS SINGLE SECTOR
# =========================
async def process_sector(sector):
    print(f"\n🔹 Processing Sector: {sector['name']}")

    # STEP 1: GET COMPANIES
    companies = await get_companies(DOMAIN, sector["url"])
    company_count = len(companies)

    # print(f"   ➤ Companies Found: {company_count}")

    # STEP 2: SCRAPE COMPANY DATA
    sector_data = await get_company_infos(
        DOMAIN,
        companies,
        sector["name"]
    )

    scraped_count = len(sector_data)

    print(f"   ✔ Companies Scraped: {scraped_count}")

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
    all_data = []

    total_sectors_found = 0
    total_sectors_scraped = 0
    total_companies_found = 0
    total_companies_scraped = 0

    # =========================
    # STEP 1: GET SECTORS
    # =========================
    sectors = await get_sectors(MAIN_URL, IGNORED_SECTORS)

    total_sectors_found = len(sectors)

    print("\n" + "=" * 60)
    print(f"TOTAL SECTORS FOUND: {total_sectors_found}")
    print("=" * 60)

    # =========================
    # 🔥 CONCURRENT SECTOR PROCESSING
    # =========================
    semaphore = asyncio.Semaphore(3)  # 🔥 CONTROL PARALLEL SECTORS

    async def sem_task(sector):
        async with semaphore:
            return await process_sector(sector)

    tasks = [sem_task(sector) for sector in sectors]

    # 🔥 PRESERVES ORDER
    results = await asyncio.gather(*tasks)

    # =========================
    # COLLECT RESULTS
    # =========================
    for result in results:
        total_sectors_scraped += 1
        total_companies_found += result["found"]
        total_companies_scraped += result["scraped"]

        all_data.extend(result["data"])

    # =========================
    # SAVE DATA
    # =========================
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
    asyncio.run(main())