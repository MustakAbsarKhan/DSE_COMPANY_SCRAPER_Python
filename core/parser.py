import re
from bs4 import BeautifulSoup as bs


def parse_html(html):
    return bs(html, "lxml")


# -------- SECTORS --------
def extract_sectors(soup, ignored):
    sectors = []

    for td in soup.find_all("td", class_="text-left"):
        a = td.find("a", class_="ab1")

        if a:
            name = a.text.strip()
            if name not in ignored:
                sectors.append({
                    "name": name,
                    "url": a["href"]
                })

    return sectors


# -------- COMPANIES --------
def extract_companies(soup):
    links = soup.find_all("a", class_="ab1")

    company_links = []

    for link in links:
        href = link.get("href")

        # Only include valid company pages
        if href and "displayCompany.php" in href:
            company_links.append(href)

    return company_links


# -------- TABLE PARSER --------
def extract_table_data(soup):
    data = {}

    rows = soup.select("#company tr")

    for row in rows:
        ths = row.find_all("th")
        tds = row.find_all("td")

        for i in range(min(len(ths), len(tds))):
            key = ths[i].text.strip()
            val = tds[i].text.strip()
            data[key] = val

    return data


# -------- 🔥 FINAL CODE EXTRACTION (WORKS FOR DSE) --------
def extract_codes(soup):
    trading_code = None
    scrip_code = None

    try:
        # Look specifically in the "alt" row first (most reliable)
        row = soup.find("tr", class_="alt")

        if row:
            text = row.get_text(" ", strip=True)

            # Example text:
            # "Trading Code: ABBANK Scrip Code: 11101"

            trading_match = re.search(r'Trading Code[:\s]+([A-Z0-9]+)', text)
            scrip_match = re.search(r'Scrip Code[:\s]+([0-9]+)', text)

            if trading_match:
                trading_code = trading_match.group(1)

            if scrip_match:
                scrip_code = scrip_match.group(1)

        # Fallback: search entire page if above fails
        if not trading_code or not scrip_code:
            full_text = soup.get_text(" ", strip=True)

            if not trading_code:
                trading_match = re.search(r'Trading Code[:\s]+([A-Z0-9]+)', full_text)
                if trading_match:
                    trading_code = trading_match.group(1)

            if not scrip_code:
                scrip_match = re.search(r'Scrip Code[:\s]+([0-9]+)', full_text)
                if scrip_match:
                    scrip_code = scrip_match.group(1)

    except:
        pass

    return trading_code, scrip_code


# -------- COMPANY INFO --------
def extract_company_info(soup, sector):
    table = extract_table_data(soup)

    def to_float(x):
        try:
            return float(x.replace(",", ""))
        except:
            return 0.0

    # 52 Week Range
    low, high = 0.0, 0.0
    try:
        if "52 Weeks' Moving Range" in table:
            low, high = table["52 Weeks' Moving Range"].split(" - ")
    except:
        pass

    trading_code, scrip_code = extract_codes(soup)

    return {
        "Company Name": soup.find_all("i")[1].text.strip() if soup.find_all("i") else None,
        "Trading Code": trading_code,
        "Scrip Code": int(scrip_code) if scrip_code and scrip_code.isdigit() else None,
        "Sector": sector,
        "LTP": to_float(table.get("Last Trading Price", "0")),
        "52 Weeks Moving Range Lowest": to_float(low),
        "52 Weeks Moving Range Highest": to_float(high),
    }