from config import MAIN_URL, DOMAIN, DEFAULT_IGNORED_SECTORS

from pipelines.sectors import get_sectors
from pipelines.companies import get_companies
from pipelines.company_info import get_company_info
from export.excel import save_to_excel


def main():
    data = []

    sectors = get_sectors(MAIN_URL, DEFAULT_IGNORED_SECTORS)

    for sector in sectors:
        print(f"\nProcessing sector: {sector['name']}")

        companies = get_companies(DOMAIN, sector['url'])

        for company in companies:
            company_name = company.split("name=")[1] if "name=" in company else company
            print(f"Scraping: {company_name}") 

            info = get_company_info(
                DOMAIN,
                company,
                sector['name']
            )

            data.append(info)

    save_to_excel(data)


if __name__ == "__main__":
    main()