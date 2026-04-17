import logging

from bs4 import BeautifulSoup as bs


def parse_html(html):
    return bs(html, "lxml")


def extract_sectors(soup, ignored_sectors):
    td_elements = soup.find_all('td', class_='text-left')

    sectors = []
    for td in td_elements:
        a = td.find('a', class_='ab1')
        if a and a.text.strip() not in ignored_sectors:
            sectors.append({
                "name": a.text.strip(),
                "url": a['href']
            })
    return sectors


def extract_companies(soup):
    companies = soup.find_all('a', class_='ab1')
    return [c['href'] for c in companies][:-3]


def extract_company_info(soup, sector_name):
    info = {
        'Company Name': None,
        'Trading Code': None,
        'Sector': sector_name,
        'LTP': 0.0,
        '52 Weeks Moving Range Lowest': 0.0,
        '52 Weeks Moving Range Highest': 0.0,
        'Last Dividend Year': None,
        'Dividend %': None,
        'Dividend Yield': None
    }

    # Company Name
    try:
        info['Company Name'] = soup.find_all('i')[1].text.strip()
    except (AttributeError, ValueError) as e:
        logging.warning(f"Error extracting {'Company Name'}: {e}")
        
    # Trading Code
    try:
        info['Trading Code'] = soup.find('tr', {'class': 'alt'}).text.split("\n")[1].replace("Trading Code:", "").strip()
    except (AttributeError, ValueError) as e:
        logging.warning(f"Error extracting {'Trading Code'}: {e}")

    # Table Data
    try:
        table = soup.find_all('table', {'id': 'company'})[1]
        tds = table.find_all('td')
        print(tds)

        info['LTP'] = float(tds[0].text.replace(',', ''))

        range_text = tds[3].text.split(' - ')
        info['52 Weeks Moving Range Lowest'] = float(range_text[0].replace(",", ""))
        info['52 Weeks Moving Range Highest'] = float(range_text[1].replace(",", ""))

    except (AttributeError, ValueError) as e:
        logging.warning(f"Error extracting table data: {e}")

    # Dividend Info
    try:
        shrink = soup.find_all(class_='shrink')[-1]
    except (AttributeError, ValueError) as e:
        try:
            shrink = soup.find_all(class_='shrink alt')[-1]
        except (AttributeError, ValueError) as e:
            shrink = None

    if shrink:
        try:
            div_info = shrink.find_all('td')
            info['Last Dividend Year'] = div_info[0].text.strip()
            info['Dividend %'] = div_info[-2].text.strip()
            info['Dividend Yield'] = div_info[-1].text.strip()
        except (AttributeError, ValueError) as e:
            logging.warning(f"Error extracting dividend info: {e}")

    return info