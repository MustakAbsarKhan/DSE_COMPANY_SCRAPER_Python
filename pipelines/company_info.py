from core.client import AsyncClient
from core.parser import parse_html, extract_company_info

async def get_company_infos(domain, company_urls, sector):
    client = AsyncClient(concurrency=5)

    # # ✅ Temporary FILTER - To Debug Only
    # company_urls = [
    #     url for url in company_urls
    #     if "BSCPLC" in url.upper()
    # ]

    # print("Filtered URLs:", company_urls)  # 🔍 DEBUG

    full_urls = [domain + url for url in company_urls]

    html_list = await client.fetch_all(full_urls)

    data = []

    for html in html_list:
        if html:
            soup = parse_html(html)
            info = extract_company_info(soup, sector)

            if info["Company Name"]:
                data.append(info)

    return data