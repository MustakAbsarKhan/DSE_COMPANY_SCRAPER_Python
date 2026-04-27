import re
from bs4 import BeautifulSoup as bs


# This file contains the HTML extraction layer. DSE company pages are not very
# consistent, so most functions are defensive: they return None or an empty
# dictionary/list instead of stopping the full scraper when one field is missing.


# =============================
# BASIC PARSER
# =============================
def parse_html(html):
    """Convert raw HTML text into a BeautifulSoup object using lxml."""
    return bs(html, "lxml")


# =============================
# SAFE HELPERS
# =============================
def parse_float(x):
    """Convert a DSE numeric string into float, returning None for blanks."""
    try:
        if not x:
            return None
        x = x.replace(",", "").strip()
        if x in ["", "-", "N/A"]:
            return None
        return float(x)
    except:
        return None


def parse_int(x):
    """Convert a DSE numeric string into int, returning None for blanks."""
    try:
        if not x:
            return None
        x = x.replace(",", "").strip()
        if x in ["", "-", "N/A"]:
            return None
        return int(float(x))
    except:
        return None


def clean_text(x):
    """Normalize text fields and turn empty placeholders into None."""
    if not x:
        return None
    x = x.strip()
    return x if x not in ["", "-"] else None


MARKET_TABLE_FIELDS_ALREADY_MODELED = {
    "Change*",
    "Last Update",
    "Last Trading Price",
    "Opening Price",
    "Closing Price",
    "Yesterday's Closing Price",
    "Adjusted Opening Price",
    "Day's Range",
    "52 Weeks' Moving Range",
    "Day's Trade (Nos.)",
    "Day's Volume (Nos.)",
    "Day's Value (mn)",
    "Market Capitalization (mn)",
    "Free Float Market Cap. (mn)",
    "Type of Instrument",
    "Face/par Value",
    "Total No. of Outstanding Securities",
}


BASIC_INFORMATION_FIELDS_ALREADY_MODELED = {
    "Sector",
    "Authorized Capital (mn)",
    "Paid-up Capital (mn)",
    "Face/par Value",
    "Market Lot",
    "Total No. of Outstanding Securities",
    "Type of Instrument",
    "Debut Trading Date",
}


EXCLUDED_SOURCE_FIELDS = {
    "Closing Price Graph:",
    "Closing Price Graph",
    "Total Trade Graph:",
    "Total Trade Graph",
    "Total Volume Graph:",
    "Total Volume Graph",
}


def preserve_unmodeled_source_fields(result, source_fields, modeled_fields):
    """Append scraped DSE fields that are not already represented elsewhere."""
    for source_key, source_value in source_fields.items():
        normalized_source_key = source_key.strip()
        if (
            normalized_source_key in modeled_fields
            or normalized_source_key in result
            or normalized_source_key in EXCLUDED_SOURCE_FIELDS
        ):
            continue

        result[normalized_source_key] = clean_text(source_value)


# =============================
# ✅ SECTOR EXTRACTION (FIXED)
# =============================
def extract_sectors(soup, ignored_sectors):
    """Extract sector names and relative URLs from the industry listing page."""
    sectors = []

    # On the industry page, sector links live inside left-aligned table cells.
    for td in soup.find_all("td", class_="text-left"):
        a = td.find("a", class_="ab1")

        if a:
            name = a.text.strip()

            if name not in ignored_sectors:
                sectors.append({
                    "name": name,
                    "url": a["href"]
                })

    return sectors


# =============================
# COMPANY LINKS
# =============================
def extract_company_urls(soup):
    """Extract company detail-page links from a sector page."""
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
def extract_company_table_key_values(soup):
    """Read key/value pairs from all DSE tables with id='company'."""
    data = {}
    tables = soup.find_all("table", id="company")

    for table in tables:
        for row in table.find_all("tr"):
            # recursive=False means only direct cells are used. This avoids
            # mixing nested table values into the wrong parent row.
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
    """Extract the market date shown under the Market Information heading."""
    for h2 in soup.find_all("h2"):
        if "Market Information" in h2.get_text():
            i = h2.find("i")
            if i:
                return i.get_text(strip=True)
    return None


# =============================
# CODES
# =============================
def extract_security_codes(soup):
    """Extract Trading Code and Scrip Code from page text using regex."""
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
def extract_price_change(soup):
    """Extract price change value and percentage from the nested change table."""
    change_value = None
    change_percent = None

    try:
        # The Change field is stored in a nested table, so regular key/value
        # table parsing does not capture it cleanly.
        th = soup.find("th", string=lambda x: x and "Change" in x)

        if th:
            tr = th.find_parent("tr")
            tds = tr.select("td table tr td")

            if len(tds) >= 2:
                change_value = parse_float(tds[0].get_text(strip=True))

                percent_text = tds[1].get_text(strip=True).replace("%", "")
                change_percent = parse_float(percent_text)

    except:
        pass

    return change_value, change_percent


# =============================
# BASIC INFO
# =============================
def extract_basic_information(soup):
    """Extract the Basic Information table into a dictionary."""
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
def extract_corporate_action_fields(soup):
    """Extract dividend, AGM, year-end, reserve, and OCI fields."""
    data = {}
    text = soup.get_text("\n", strip=True)

    # Some fields appear as plain page text instead of regular table cells.
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

    # The last 'shrink' block contains latest dividend year/yield information
    # on the current DSE layout.
    shrink = (soup.find_all(class_='shrink'))[-1]
    div_info = shrink.findAll('td')
    data["Last Div Year"] = int(div_info[0].text.strip())
    data["Last Div Yield %"] = parse_float(div_info[-1].text.strip())
    

    # Scan all company tables for named corporate/fundamental fields.
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
def extract_listing_metadata(soup):
    """Extract listing/category/share metadata from company information tables."""
    listing_metadata = {}
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
                    listing_metadata[key] = val

    return listing_metadata
# ================================
# Shareholding INFO OF THE COMPANY
# ================================
def extract_shareholding_percentages(soup):
    """Flatten shareholding-percentage rows into export-friendly strings."""
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

# Extract Audited EPS and Unaudited EPS for Continuing Operations
def extract_eps_metrics(soup):
    """Extract EPS values from the DSE EPS table.

    Important: this uses DSE's current table order. The parser intentionally
    keeps this index-based extraction because the page layout is complex and was
    tuned against the live DSE structure.
    """
    data = {}

    # Table index 4 currently contains EPS sections on DSE company pages.
    table = soup.find_all("table", id="company")[4]

    periods = ["Q1", "Q2", "HalfYearly", "Q3", "9Months", "Annual"]

    # --- First EPS section ---
    eps_header = table.find("td", string=lambda t: t and "Earnings Per Share (EPS)" in t)
    if eps_header:
        basic_row = eps_header.find_parent("tr").find_next_sibling("tr")
        if basic_row:
            values = [parse_float(td.get_text(strip=True)) for td in basic_row.find_all("td")[1:]]
            for label, val in zip(periods, values):
                data[f"{label}_EPS"] = val

    # --- Continuing operations EPS section ---
    eps_cop_header = table.find("td", string=lambda t: t and "Earnings Per Share (EPS) - continuing operations" in t)
    if eps_cop_header:
        basic_row = eps_cop_header.find_parent("tr").find_next_sibling("tr")
        if basic_row:
            values = [parse_float(td.get_text(strip=True)) for td in basic_row.find_all("td")[1:]]
            for label, val in zip(periods, values):
                data[f"{label}_EPS_COP"] = val

            # The row after Basic contains diluted EPS for continuing
            # operations. Example DSE label: "Diluted*".
            diluted_row = basic_row.find_next_sibling("tr")
            if diluted_row:
                first_cell = diluted_row.find("td")

                if first_cell and "Diluted" in first_cell.get_text(strip=True):
                    values = [parse_float(td.get_text(strip=True)) for td in diluted_row.find_all("td")[1:]]
                    for label, val in zip(periods, values):
                        data[f"{label}_Diluted_EPS_COP"] = val

    return data


def extract_pe_ratios_and_annual_performance(soup):
    """
    Extract unaudited and audited P/E ratios (Basic, Diluted, Trailing) 
    from the HTML tables with id="company".

    Returns:
        dict: Combined dictionary of all extracted values with 
              date-based keys and descriptive suffixes.
    """

    # -----------------------------
    # UNAUDITED TABLE (index 5)
    # -----------------------------
    # Table index 5 currently contains unaudited P/E values.
    table_unaudited = soup.find_all("table", id="company")[5]
    tds_unaudited = table_unaudited.find_all("td")

    # Extract date headers (first row, skip the first cell "Particulars")
    date_keys = [td.get_text(strip=True) for td in tds_unaudited[1:7]]

    # --- Basic EPS P/E values ---
    data_pe_basic = {}
    pe_basic_values = [td.get_text(strip=True) for td in tds_unaudited[8:14]]
    for date, val in zip(date_keys, pe_basic_values):
        key = f"{date}_PEwBasEPS"
        data_pe_basic[key] = parse_float(val)

    # --- Diluted EPS P/E values ---
    data_pe_diluted = {}
    pe_diluted_values = [td.get_text(strip=True) for td in tds_unaudited[15:21]]
    for date, val in zip(date_keys, pe_diluted_values):
        key = f"{date}_PEwDilutEPS"
        data_pe_diluted[key] = parse_float(val)

    # --- Trailing P/E Ratio ---
    data_petrail_ratio = {}
    pe_trailing_values = [td.get_text(strip=True) for td in tds_unaudited[22:28]]
    for date, val in zip(date_keys, pe_trailing_values):
        key = f"{date}_PEwTrailRatio"
        data_petrail_ratio[key] = parse_float(val)

    # -----------------------------
    # AUDITED TABLE (index 6)
    # -----------------------------
    # Table index 6 currently contains audited P/E values.
    table_audited = soup.find_all("table", id="company")[6]
    tds_audited = table_audited.find_all("td")

    # --- Audited Basic EPS P/E values ---
    data_audited_pe_basic = {}
    pe_audited_basic_values = [td.get_text(strip=True) for td in tds_audited[8:14]]
    for date, val in zip(date_keys, pe_audited_basic_values):
        key = f"{date}_PEwAuditBascEPS"
        data_audited_pe_basic[key] = parse_float(val)

    # -----------------------------
    # AUDITED TABLE (index 7)
    # -----------------------------
    # Table index 7 currently contains annual audited financial performance.
    table_financial_performance = soup.find_all("table", id="company")[7]
    data_audited_financial_performance = {}

    annual_fields = [
        ("Aud_EPS_COP_Basic", 4),
        ("Aud_EPS_COP_Diluted", 6),
        ("Aud_NAVPS", 7),
        ("Aud_PCO_mn", 10),
        ("Aud_Profit_mn", 11),
        ("Aud_TCI_mn", 12),
    ]

    for row in table_financial_performance.find_all("tr", class_="shrink"):
        tds = row.find_all("td")
        if len(tds) <= max(index for _, index in annual_fields):
            continue

        year_text = tds[0].get_text(strip=True)
        if not re.search(r"\d{4}", year_text):
            continue

        year_key = re.sub(r"\s+", "_", year_text)
        year_key = re.sub(r"[^A-Za-z0-9_/-]", "", year_key)

        for field_name, td_index in annual_fields:
            key = f"{field_name}_{year_key}"
            val = tds[td_index].get_text(strip=True)
            data_audited_financial_performance[key] = parse_float(val)

    # -----------------------------
    # MERGE ALL RESULTS
    # -----------------------------
    return {
        **data_pe_basic,
        **data_pe_diluted,
        **data_petrail_ratio,
        **data_audited_pe_basic,
        **data_audited_financial_performance
    }

    
# =============================
# MAIN FUNCTION
# =============================
def extract_company_profile(soup, sector):
    """Build one clean export row from a single DSE company page."""
    # Generic table extraction catches most key/value market fields.
    market_table_fields = extract_company_table_key_values(soup)

    # Basic info is parsed separately because it lives under a named heading.
    basic_information = extract_basic_information(soup)

    name = None
    try:
        # The second <i> tag currently contains the company name on DSE pages.
        name = soup.find_all("i")[1].text.strip()
    except:
        pass

    market_date = extract_market_date(soup)

    # Split range strings like "10.00 - 20.00" into separate low/high columns.
    low, high = None, None
    if "52 Weeks' Moving Range" in market_table_fields:
        try:
            low, high = market_table_fields["52 Weeks' Moving Range"].split(" - ")
        except:
            pass

    day_low, day_high = None, None
    if "Day's Range" in market_table_fields:
        try:
            day_low, day_high = market_table_fields["Day's Range"].split(" - ")
        except:
            pass

    change_value, change_percent = extract_price_change(soup)
    trading_code, scrip_code = extract_security_codes(soup)

    # Specialized extractors handle fields that are nested, repeated, or stored
    # outside ordinary key/value table rows.
    corporate_action_fields = extract_corporate_action_fields(soup)
    listing_metadata = extract_listing_metadata(soup)
    eps_metrics = extract_eps_metrics(soup)
    pe_and_annual_metrics = extract_pe_ratios_and_annual_performance(soup)
    shareholding_percentages = extract_shareholding_percentages(soup)
    
    # Start with a predictable base schema. Some fields are initialized as None
    # and filled later by more specialized extraction results.
    result = {
    # =============================
    # 🟢 IDENTIFICATION
    # =============================
    "Company Name": clean_text(name),
    "Trading Code": clean_text(trading_code),
    "Scrip Code": scrip_code,
    "Sector": sector or clean_text(basic_information.get("Sector")),

    "Market Date": market_date,
    "Last Update": clean_text(market_table_fields.get("Last Update")),


    # =============================
    # 🟢 PRICE (CORE MARKET DATA)
    # =============================
    "LTP": parse_float(market_table_fields.get("Last Trading Price")),
    "Opening Price": parse_float(market_table_fields.get("Opening Price")),
    "Closing Price": parse_float(market_table_fields.get("Closing Price")),
    "YCP": parse_float(market_table_fields.get("Yesterday's Closing Price")),
    "Adj Opening Price": parse_float(market_table_fields.get("Adjusted Opening Price")),


    # =============================
    # 🟢 RANGE
    # =============================
    "Day Low": parse_float(day_low),
    "Day High": parse_float(day_high),
    "52W Low": parse_float(low),
    "52W High": parse_float(high),


    # =============================
    # 🟢 MOMENTUM
    # =============================
    "Change Value": change_value,
    "Change %": change_percent,


    # =============================
    # 🔵 LIQUIDITY
    # =============================
    "Day Trade No": parse_int(market_table_fields.get("Day's Trade (Nos.)")),
    "Day Volume": parse_int(market_table_fields.get("Day's Volume (Nos.)")),
    "Day Value (mn)": parse_float(market_table_fields.get("Day's Value (mn)")),


    # =============================
    # 🔵 SIZE
    # =============================
    "Market Cap (mn)": parse_float(market_table_fields.get("Market Capitalization (mn)")),
    "Free Float Cap (mn)": parse_float(market_table_fields.get("Free Float Market Cap. (mn)")),


    # =============================
    # 🟡 FUNDAMENTALS
    # =============================
    "Authorized Capital (mn)": parse_float(basic_information.get("Authorized Capital (mn)")),
    "Paid-up Capital (mn)": parse_float(basic_information.get("Paid-up Capital (mn)")),
    "Reserve & Surplus without OCI (mn)": None,
    "Other Comprehensive Income (OCI) (mn)": None,


    # =============================
    # 🔴 VALUATION (EPS + P/E)
    # =============================
    **eps_metrics,
    **pe_and_annual_metrics,


    # =============================
    # 🟡 STRUCTURE
    # =============================
    "Face Value": parse_float(basic_information.get("Face/par Value")),
    "Market Lot": parse_int(basic_information.get("Market Lot")),
    "Total Securities": parse_int(basic_information.get("Total No. of Outstanding Securities")),


    # =============================
    # 🟣 CORPORATE ACTIONS
    # =============================
    "Last AGM held on": None,
    "For the year ended": None,
    "Last Div Year": None,
    "Last Div Yield %": None,
    "Cash Dividend": None,
    "Bonus Issue (Stock Dividend)": None,
    "Right Issue": None,
    "Year End": None,


    # =============================
    # ⚫ META
    # =============================
    "Instrument Type": clean_text(basic_information.get("Type of Instrument")),
    "Debut Trading Date": clean_text(basic_information.get("Debut Trading Date")),
    "Listing Year": None,
    "Market Category": None,
    "Electronic Share": None,
    }
    
    # Merge specialized data into the base schema. Later values override earlier
    # placeholder None values where available.
    result.update(corporate_action_fields)
    result.update(listing_metadata)

    # Listing Year is scraped as text, so convert it after merging.
    result["Listing Year"] = parse_int(result.get("Listing Year"))

    # EPS/P/E are already included in the base result, but these updates keep
    # the final row aligned with the latest extracted dictionaries.
    result.update(eps_metrics)
    result.update(pe_and_annual_metrics)
    result.update(shareholding_percentages)
   
    
    # Convert numeric extra fields
    for field in [
        "Reserve & Surplus without OCI (mn)",
        "Other Comprehensive Income (OCI) (mn)"
    ]:
        if field in result:
            result[field] = parse_float(result[field])

    preserve_unmodeled_source_fields(
        result,
        market_table_fields,
        MARKET_TABLE_FIELDS_ALREADY_MODELED
    )
    preserve_unmodeled_source_fields(
        result,
        basic_information,
        BASIC_INFORMATION_FIELDS_ALREADY_MODELED
    )

    # =============================
    # PURPOSE BASED FINAL ORDER CONTROL
    # =============================
    # Python dictionaries preserve insertion order. Rebuilding through
    # ordered_keys keeps the export column order explicit and stable.
    ordered_keys = list(result.keys())  # already structured logically

    ordered_result = {k: result.get(k) for k in ordered_keys}

    return ordered_result
