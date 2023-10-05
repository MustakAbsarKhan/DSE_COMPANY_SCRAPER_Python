# Introduction of the Developer
print("\nWelcome to our DSE Company Data Scraper!")
print("Developed by Mohammad Mustak Absar Khan")
# Initially Contributed by [Akib Sadmanee]
print("Contact: mustak.absar.khan@gmail.com\n")

input("Press Enter to initiate the code...")
print("\nInitializing.... \n")

import requests
from bs4 import BeautifulSoup as bs
from time import sleep
from random import randint
import pandas as pd
import xlsxwriter

from requests.adapters import HTTPAdapter
from urllib3 import Retry
from tqdm import tqdm

sleep(3)

sector_letters = {
    'A': 'Bank',
    'B': 'Cement',
    'C': 'Ceramics Sector',
    'D': 'Engineering',
    'E': 'Financial Institutions',
    'F': 'Food & Allied',
    'G': 'Fuel & Power',
    'H': 'Insurance',
    'I': 'IT Sector',
    'J': 'Jute',
    'K': 'Miscellaneous',
    'L': 'Mutual Funds',
    'M': 'Paper & Printing',
    'N': 'Pharmaceuticals & Chemicals',
    'O': 'Services & Real Estate',
    'P': 'Tannery Industries',
    'Q': 'Telecommunication',
    'R': 'Textile',
    'S': 'Travel & Leisure',
    'T': 'Corporate Bond',
    'U': 'Debenture',
    'V': 'G-SEC (T.Bond)'
}

ignored_sectors_default = []

use_default = input("Do you want to go with the default settings? (y/n): ")

if use_default.lower() == 'y':
    #ignored_sectors_default = ['Corporate Bond', 'Debenture', 'G-SEC (T.Bond)']
    ignored_sectors_default = ['Bank',
                               'Cement',
                               'Ceramics Sector',
                               'Engineering',
                               'Financial Institutions',
                               'Food & Allied',
                               'Fuel & Power',
                               'Insurance',
                               'IT Sector',
                               'Jute',
                               'Miscellaneous',
                               'Mutual Funds',
                               'Paper & Printing',
                               'Pharmaceuticals & Chemicals',
                               'Services & Real Estate',
                               'Tannery Industries',
                               'Telecommunication',
                               'Textile',
                               'Travel & Leisure',
                               'Debenture']
else:
    print("Sector Letters:")
    for key, value in sector_letters.items():
        print(f"{key}: {value}")

    sector_keys = input("Enter the sector keys to ignore, separated by commas (e.g., 'A,B,C'): ")
    sector_keys = sector_keys.upper().split(',')

    ignored_sectors_default = [sector_letters[key] for key in sector_keys if key in sector_letters]

print("Ignored Sectors:", ignored_sectors_default)


def get_sectors(main_page_url):
    html = bs(requests.get(main_page_url).content, features="lxml")
    td_elements = html.find_all('td', class_='text-left')
    sectors = [td for td in td_elements if not (td.find('a', class_='ab1') and td.find('a', class_='ab1').text.strip() in ignored_sectors_default)]
    print("\nWorking on These Sectors:\n")
    for sector in sectors:
        print(f"{sector.text.strip()}")
    return [{'url': sector.find('a', class_='ab1')['href'], 'sector': sector.find('a', class_='ab1').text} for sector in sectors if sector.find('a', class_='ab1') is not None]

def get_companies(sector_url):
    html = bs(requests.get(sector_url).content, features="lxml")
    companies = html.findAll('a', {'class':'ab1'})
    return [company['href'] for company in companies][:-3]

def get_info(url, sector):
    info = {
        'Company Name' : '',
        'Trading Code': '',
        'Sector' : '',
        'LTP' : '',
        '52 Weeks Moving Range Lowest' : '',
        '52 Weeks Moving Range Highest' : '',
        'Last Dividend Year' : '',
        'Dividend %' : '',
        'Dividend Yield' : ''
    }
    session = requests.Session()
    retry_strategy = Retry(total=5, backoff_factor=1, status_forcelist=[429, 500, 502, 503, 504])
    adapter = HTTPAdapter(max_retries=retry_strategy)
    session.mount("http://", adapter)
    session.mount("https://", adapter)

    stat = 0
    while stat != 200:
        response = session.get(url)
        stat = response.status_code
        sleep(3)

    html = bs(response.content, features="lxml")

    table = html.findAll('table', {'id':'company'})[1]
    tds = table.findAll('td')

    info['Company Name'] = html.findAll('i')[1].text.strip()
    print("\n"+info['Company Name'])
    info['Trading Code'] = html.find('tr', {'class':'alt'}).text.split("\n")[1]\
                                            .replace("Trading Code:", "").strip()
    info['Sector'] = sector
    try:
        info['LTP'] = float(tds[0].text.replace(',', ''))
    except:
        info['LTP'] = float(0.00)

    try:
        info['52 Weeks Moving Range Lowest'] = float(tds[7].text.split(' - ')[0].replace(",", ""))
    except:
        info['52 Weeks Moving Range Lowest'] = float(0.00)

    try:
        info['52 Weeks Moving Range Highest'] = float(tds[7].text.split(' - ')[1].replace(",", ""))
    except:
        info['52 Weeks Moving Range Highest'] = float(0.00)

    try:
        shrink = (html.find_all(class_='shrink'))[-1]
        div_info = shrink.findAll('td')

        last_div_year = div_info[0].text.strip()
        last_div_percent = div_info[-2].text.strip()
        last_div_yield = div_info[-1].text.strip()

        info['Last Dividend Year'] = last_div_year
        info['Dividend %'] = last_div_percent
        info['Dividend Yield'] = last_div_yield
    except:
        shrink_alt = (html.find_all(class_='shrink alt'))[-1]
        div_info = shrink_alt.findAll('td')

        last_div_year = div_info[0].text.strip()
        last_div_percent = div_info[-2].text.strip()
        last_div_yield = div_info[-1].text.strip()

        info['Last Dividend Year'] = last_div_year
        info['Dividend %'] = last_div_percent
        info['Dividend Yield'] = last_div_yield

    return info

def main():
    data = []
    domain = 'https://www.dsebd.org/'
    main_page_url = domain + 'by_industrylisting.php'

    sectors = get_sectors(main_page_url)[:-1]
    total_sectors = len(sectors)

    # Create a list of ignored sectors
    ignored_sectors = [sector['sector'] for sector in sectors if sector['sector'] in ignored_sectors_default]

    # Deduct ignored sectors from the total count
    remaining_sectors = total_sectors - len(ignored_sectors)

    # Create a progress bar for sectors
    with tqdm(total=remaining_sectors, ncols=70, desc="\nProcessing Sectors") as sector_pbar:
        for sector in sectors:
            if sector['sector'] in ignored_sectors:
                continue  # Skip ignored sectors
            
            print("\n\nWorking on {} Sector:=>".format(sector['sector']))
            sectorURL = sector['url']
            companies = get_companies(domain + sectorURL)
            
            # Create a progress bar for companies within the sector
            with tqdm(total=len(companies), ncols=70, desc=f"\nProcessing {sector['sector']} Sector") as company_pbar:
                for i, company in enumerate(companies):
                    data.append(get_info(domain + company, sector['sector']))
                    sleep(randint(1, 4))
                    company_pbar.update(1)
            
            sector_pbar.update(1)
    
    return data


if __name__ == "__main__":
    data = main()

    ordered_list = ["Sector", "Company Name", "Trading Code", "LTP", "52 Weeks Moving Range Lowest", "52 Weeks Moving Range Highest", 'Last Dividend Year', 'Dividend %', 'Dividend Yield']
    workbook = xlsxwriter.Workbook("DSE_Company_Details.xlsx")
    worksheet = workbook.add_worksheet("Scrapped_Data")

    first_row = 0
    for header in ordered_list:
        col = ordered_list.index(header)
        worksheet.write(first_row, col, header)

    row = 1
    for company in data:
        for _key, _value in company.items():
            col = ordered_list.index(_key)
            worksheet.write(row, col, _value)
        row += 1

    workbook.close()
