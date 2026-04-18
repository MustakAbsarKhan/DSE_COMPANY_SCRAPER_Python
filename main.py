import asyncio

from config import DOMAIN, MAIN_URL, IGNORED_SECTORS

from pipelines.sectors import get_sectors
from pipelines.companies import get_companies
from pipelines.company_info import get_company_infos

from export.excel import save_to_excel


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
    # LOOP THROUGH SECTORS
    # =========================
    for sector in sectors:
        print(f"\n🔹 Processing Sector: {sector['name']}")

        # =========================
        # STEP 2: GET COMPANIES
        # =========================
        companies = await get_companies(DOMAIN, sector["url"])

        sector_company_count = len(companies)
        total_companies_found += sector_company_count

        print(f"   ➤ Companies Found: {sector_company_count}")

        # =========================
        # STEP 3: SCRAPE COMPANIES
        # =========================
        sector_data = await get_company_infos(
            DOMAIN,
            companies,
            sector["name"]
        )

        scraped_count = len(sector_data)
        total_companies_scraped += scraped_count
        total_sectors_scraped += 1

        print(f"   ✔ Companies Scraped: {scraped_count}")

        all_data.extend(sector_data)

    # =========================
    # STEP 4: SAVE DATA
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