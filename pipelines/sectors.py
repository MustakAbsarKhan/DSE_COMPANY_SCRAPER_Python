from core.client import get
from core.parser import parse_html, extract_sectors


def get_sectors(url, ignored_sectors):
    html = get(url)
    soup = parse_html(html)

    return extract_sectors(soup, ignored_sectors)