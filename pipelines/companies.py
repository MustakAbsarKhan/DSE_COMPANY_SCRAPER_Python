from core.client import get
from core.parser import parse_html, extract_companies


def get_companies(domain, sector_url):
    html = get(domain + sector_url)
    soup = parse_html(html)

    return extract_companies(soup)