from core.client import global_client
from core.parser import parse_html, extract_companies


async def get_companies(domain, sector_url):
    """Fetch one sector page and return company detail-page links from it."""
    # Sector URLs from DSE are relative paths, so prepend the configured domain.
    html = (await global_client.fetch_all([domain + sector_url]))[0]

    if not html:
        return []

    # The parser searches the sector page for displayCompany.php links.
    soup = parse_html(html)
    return extract_companies(soup)
