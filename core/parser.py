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
    print(f"Total Sectors Found: {len(sectors)}", end="\r")

    return sectors


# -------- COMPANIES --------
def extract_companies(soup):
    links = soup.find_all("a", class_="ab1")
    l = [l["href"] for l in links][:-3]
    print(f"Total Companies Found: {len(l)}", end="\r")
    return l


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


# -------- COMPANY INFO --------
def extract_company_info(soup, sector):
    table = extract_table_data(soup)

    def to_float(x):
        try:
            return float(x.replace(",", ""))
        except:
            return 0.0

    low, high = 0.0, 0.0
    if "52 Weeks' Moving Range" in table:
        try:
            low, high = table["52 Weeks' Moving Range"].split(" - ")
        except:
            pass

    return {
        "Company Name": soup.find_all("i")[1].text.strip() if soup.find_all("i") else None,
        "Trading Code": table.get("Trading Code"),
        "Sector": sector,
        "LTP": to_float(table.get("Last Trading Price", "0")),
        "52 Weeks Moving Range Lowest": to_float(low),
        "52 Weeks Moving Range Highest": to_float(high),
    }