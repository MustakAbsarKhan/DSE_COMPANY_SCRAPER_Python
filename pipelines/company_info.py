from core.client import global_client
from core.parser import parse_html, extract_company_info

async def get_company_infos(domain, company_urls, sector):
    """Fetch and parse all company pages for a single sector."""
    # # ✅ Temporary FILTER - To Debug Only
    # # Uncomment this block when you want to scrape only one company while
    # # debugging parser behavior.
    # company_urls = [
    #     url for url in company_urls
    #     if "BSCPLC" in url.upper()
    # ]

    # print("Filtered URLs:", company_urls)  # 🔍 DEBUG

    # Company URLs from DSE are relative paths. Convert them into absolute URLs
    # before sending them to the shared async client.
    full_urls = [domain + url for url in company_urls]

    html_list = await global_client.fetch_all(full_urls)

    data = []

    for html in html_list:
        # A failed request returns None. Those pages are skipped here after the
        # client has already retried them.
        if html:
            soup = parse_html(html)
            info = extract_company_info(soup, sector)

            # Only keep rows where the parser found a company name. This avoids
            # exporting empty rows from malformed or unexpected pages.
            if info["Company Name"]:
                data.append(info)

    return data
