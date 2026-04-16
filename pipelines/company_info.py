from core.client import get
from core.parser import parse_html, extract_company_info


def get_company_info(domain, company_url, sector_name):
    html = get(domain + company_url)
    soup = parse_html(html)

    return extract_company_info(soup, sector_name)