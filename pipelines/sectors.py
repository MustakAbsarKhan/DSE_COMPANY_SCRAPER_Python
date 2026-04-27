from core.client import global_client
from core.parser import parse_html, extract_sectors


async def fetch_tradable_sector_links(url, ignored_sectors):
    """Fetch the DSE industry listing page and return usable sector links."""
    # fetch_all() returns a list because it can fetch many URLs. Here we only
    # request one page, so the first item is the industry listing HTML.
    html = (await global_client.fetch_all([url]))[0]

    if not html:
        return []

    # Convert raw HTML into a BeautifulSoup object, then let the parser extract
    # sector names and URLs while applying the ignored-sector filter.
    soup = parse_html(html)
    return extract_sectors(soup, ignored_sectors)
