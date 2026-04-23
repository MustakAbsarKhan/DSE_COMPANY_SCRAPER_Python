import re
from bs4 import BeautifulSoup as bs


# =============================
# BASIC PARSER
# =============================
def parse_html(html):
    return bs(html, "lxml")


# =============================
# SAFE HELPERS
# =============================
def to_float(x):
    try:
        if not x:
            return None
        x = x.replace(",", "").strip()
        if x in ["", "-", "N/A"]:
            return None
        return float(x)
    except:
        return None


def to_int(x):
    try:
        if not x:
            return None
        x = x.replace(",", "").strip()
        if x in ["", "-", "N/A"]:
            return None
        return int(float(x))
    except:
        return None


def clean_str(x):
    if not x:
        return None
    x = x.strip()
    return x if x not in ["", "-"] else None


# =============================
# ✅ SECTOR EXTRACTION (FIXED)
# =============================
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


# =============================
# COMPANY LINKS
# =============================
def extract_companies(soup):
    links = soup.find_all("a", class_="ab1")
    company_links = []

    for link in links:
        href = link.get("href")
        if href and "displayCompany.php" in href:
            company_links.append(href)

    return company_links


# =============================
# TABLE PARSER
# =============================
def extract_table_data(soup):
    data = {}
    tables = soup.find_all("table", id="company")

    for table in tables:
        for row in table.find_all("tr"):
            ths = row.find_all("th", recursive=False)
            tds = row.find_all("td", recursive=False)

            for i in range(min(len(ths), len(tds))):
                key = ths[i].get_text(strip=True)
                val = tds[i].get_text(strip=True)

                if key:
                    data[key] = val

    return data


# =============================
# MARKET DATE
# =============================
def extract_market_date(soup):
    for h2 in soup.find_all("h2"):
        if "Market Information" in h2.get_text():
            i = h2.find("i")
            if i:
                return i.get_text(strip=True)
    return None


# =============================
# CODES
# =============================
def extract_codes(soup):
    trading_code = None
    scrip_code = None

    try:
        text = soup.get_text(" ", strip=True)

        t_match = re.search(r"Trading Code[:\s]+([A-Z0-9]+)", text)
        s_match = re.search(r"Scrip Code[:\s]+([0-9]+)", text)

        if t_match:
            trading_code = t_match.group(1)

        if s_match:
            scrip_code = int(s_match.group(1))

    except:
        pass

    return trading_code, scrip_code


# =============================
# CHANGE
# =============================
def extract_change(soup):
    change_value = None
    change_percent = None

    try:
        th = soup.find("th", string=lambda x: x and "Change" in x)

        if th:
            tr = th.find_parent("tr")
            tds = tr.select("td table tr td")

            if len(tds) >= 2:
                change_value = to_float(tds[0].get_text(strip=True))

                percent_text = tds[1].get_text(strip=True).replace("%", "")
                change_percent = to_float(percent_text)

    except:
        pass

    return change_value, change_percent


# =============================
# BASIC INFO
# =============================
def extract_basic_info(soup):
    data = {}

    try:
        header = soup.find("h2", string=lambda x: x and "Basic Information" in x)
        table = header.find_next("table")

        for row in table.find_all("tr"):
            ths = row.find_all("th", recursive=False)
            tds = row.find_all("td", recursive=False)

            for i in range(min(len(ths), len(tds))):
                data[ths[i].get_text(strip=True)] = tds[i].get_text(strip=True)

    except:
        pass

    return data


# =============================
# EXTRA FIELDS - AGM AND OCI DETAILS
# =============================
def extract_extra_fields(soup):
    data = {}
    text = soup.get_text("\n", strip=True)

    agm = re.search(r"Last AGM held on:\s*([0-9\-]+)", text)
    if agm:
        data["Last AGM held on"] = agm.group(1)

    year = re.search(r"For the year ended:\s*([A-Za-z0-9, ]+)", text)
    if year:
        data["For the year ended"] = year.group(1)

    target_fields = [
        "Cash Dividend",
        "Bonus Issue (Stock Dividend)",
        "Right Issue",
        "Year End",
        "Reserve & Surplus without OCI (mn)",
        "Other Comprehensive Income (OCI) (mn)",
    ]

    shrink = (soup.find_all(class_='shrink'))[-1]
    div_info = shrink.findAll('td')
    data["Last Div Year"] = int(div_info[0].text.strip())
    data["Last Div Yield %"] = to_float(div_info[-1].text.strip())
    

    for table in soup.find_all("table", id="company"):
        for row in table.find_all("tr"):
            th = row.find("th")
            td = row.find("td")

            if th and td:
                key = th.get_text(strip=True)

                if key in target_fields:
                    data[key] = td.get_text(strip=True)
    
    return data

# ================================
# OTHER INFORMATION OF THE COMPANY
# ================================
def other_company_info(soup):
    other_company_data = {}
    data_fields = [
        "Listing Year",
        "Market Category",
        "Electronic Share",
    ]

    for table in soup.find_all("table", id="company"):
        for row in table.find_all("tr"):
            tds = row.find_all("td")  

            if len(tds) >= 2:
                key = tds[0].get_text(strip=True)
                val = tds[1].get_text(strip=True)

                if key in data_fields:
                    other_company_data[key] = val

    return other_company_data
# ================================
# Shareholding INFO OF THE COMPANY
# ================================
def parse_shareholding_rows(soup):
    rows_data = {}

    # Only target rows that contain "Share Holding Percentage"
    for idx, row in enumerate(soup.find_all("tr")):
        tds = row.find_all("td", recursive=False)
        if len(tds) < 2:
            continue

        key = tds[0].get_text(" ", strip=True)

        if "Share Holding Percentage" in key:
            val_td = tds[1]

            # Normalize key and append index to avoid duplicates
            key = " ".join(key.split())
            if key in rows_data:
                key = f"{key} ({idx})"

            # Flatten nested table into one string
            if val_td.find("table"):
                inner_values = []
                for inner_td in val_td.find_all("td"):
                    text = inner_td.get_text(" ", strip=True)
                    if text:
                        inner_values.append(text)
                rows_data[key] = " | ".join(inner_values)
            else:
                rows_data[key] = val_td.get_text(" ", strip=True)

    return rows_data



# =============================
# MAIN FUNCTION
# =============================
def extract_company_info(soup, sector):
    table = extract_table_data(soup)
    basic_info = extract_basic_info(soup)

    name = None
    try:
        name = soup.find_all("i")[1].text.strip()
    except:
        pass

    market_date = extract_market_date(soup)

    low, high = None, None
    if "52 Weeks' Moving Range" in table:
        try:
            low, high = table["52 Weeks' Moving Range"].split(" - ")
        except:
            pass

    day_low, day_high = None, None
    if "Day's Range" in table:
        try:
            day_low, day_high = table["Day's Range"].split(" - ")
        except:
            pass

    change_value, change_percent = extract_change(soup)
    trading_code, scrip_code = extract_codes(soup)
    extra = extract_extra_fields(soup)
    other_company_data = other_company_info(soup)
    shareholding_data = parse_shareholding_rows(soup)
    
    

    result = {
        "Market Date": market_date,
        "Last Update": clean_str(table.get("Last Update")),
        "Sector": sector or clean_str(basic_info.get("Sector")),
        "Company Name": clean_str(name),
        "Trading Code": clean_str(trading_code),
        "Scrip Code": scrip_code,

        # PRICE
        "LTP": to_float(table.get("Last Trading Price")),
        "Opening Price": to_float(table.get("Opening Price")),
        "Closing Price": to_float(table.get("Closing Price")),
        "YCP": to_float(table.get("Yesterday's Closing Price")),
        "Adj Opening Price": to_float(table.get("Adjusted Opening Price")),

        # RANGE
        "Day Low": to_float(day_low),
        "Day High": to_float(day_high),
        "52W Low": to_float(low),
        "52W High": to_float(high),

        # MOMENTUM
        "Change Value": change_value,
        "Change %": change_percent,

        # LIQUIDITY
        "Day Trade No": to_int(table.get("Day's Trade (Nos.)")),
        "Day Volume": to_int(table.get("Day's Volume (Nos.)")),
        "Day Value (mn)": to_float(table.get("Day's Value (mn)")),

        # SIZE
        "Market Cap (mn)": to_float(table.get("Market Capitalization (mn)")),
        "Free Float Cap (mn)": to_float(table.get("Free Float Market Cap. (mn)")),

        # FUNDAMENTALS
        "Authorized Capital (mn)": to_float(basic_info.get("Authorized Capital (mn)")),
        "Paid-up Capital (mn)": to_float(basic_info.get("Paid-up Capital (mn)")),
        "Reserve & Surplus without OCI (mn)": None,
        "Other Comprehensive Income (OCI) (mn)": None,

        # STRUCTURE
        "Face Value": to_float(basic_info.get("Face/par Value")),
        "Market Lot": to_int(basic_info.get("Market Lot")),
        "Total Securities": to_int(basic_info.get("Total No. of Outstanding Securities")),

        # META
        "Instrument Type": clean_str(basic_info.get("Type of Instrument")),
        "Debut Trading Date": clean_str(basic_info.get("Debut Trading Date")),
        "Listing Year": None,
        "Market Category": None,
        "Electronic Share": None,

        # CORPORATE
        "Last AGM held on": None,
        "For the year ended": None,
        "Last Div Year": None,
        "Last Div Yield %": None,
        "Cash Dividend": None,
        "Bonus Issue (Stock Dividend)": None,
        "Right Issue": None,
        "Year End": None,
        
        # SHAREHOLDING INFORMATION
        "Share Holding Percentage [as on Jun 30, 2025 (year ended)]": None,
        "Share Holding Percentage [as on Feb 28, 2026]":None,
        "Share Holding Percentage [as on Mar 31, 2026]":None
    }

    # Merge extra
    result.update(extra)
    result.update(other_company_data)
    result["Listing Year"] = to_int(result.get("Listing Year"))
    result.update(shareholding_data)
    
    # Convert numeric extra fields
    for field in [
        "Reserve & Surplus without OCI (mn)",
        "Other Comprehensive Income (OCI) (mn)"
    ]:
        if field in result:
            result[field] = to_float(result[field])

    # =============================
    # PURPOSE BASED FINAL ORDER CONTROL
    # =============================
    ordered_keys = list(result.keys())  # already structured logically

    ordered_result = {k: result.get(k) for k in ordered_keys}

    return ordered_result
