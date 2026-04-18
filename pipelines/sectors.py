from core.client import AsyncClient
from core.parser import parse_html, extract_sectors


async def get_sectors(url, ignored):
    client = AsyncClient(concurrency=2)

    html = (await client.fetch_all([url]))[0]

    if not html:
        return []

    soup = parse_html(html)
    return extract_sectors(soup, ignored)