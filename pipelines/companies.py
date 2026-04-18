from core.client import AsyncClient
from core.parser import parse_html, extract_companies


async def get_companies(domain, sector_url):
    client = AsyncClient(concurrency=2)

    html = (await client.fetch_all([domain + sector_url]))[0]

    if not html:
        return []

    soup = parse_html(html)
    return extract_companies(soup)