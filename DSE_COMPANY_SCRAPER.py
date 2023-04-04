import requests
from bs4 import BeautifulSoup as bs
from time import sleep
from random import randint
from xlsxwriter import Workbook

def get_sectors(main_page_url):
    html = bs(requests.get(main_page_url).content, features="lxml")
    sectors = html.findAll('td', {'class':'text-left'})
    return [{'url': sector.find('a', {'class':'ab1'})['href'], 
    'sector': sector.find('a', {'class':'ab1'}).text} for sector in sectors]
    
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
        '52 Weeks Moving Range Highest' : ''
    }
    stat = 200
    response = requests.get(url)
    stat = response.status_code
    
    while stat != 200:
        response = requests.get(url)
        stat = response.status_code
        sleep(3)
    
    html = bs(response.content, features="lxml")
    
    table = html.findAll('table', {'id':'company'})[1]
    tds = table.findAll('td')
    
    info['Company Name'] = html.findAll('i')[1].text.strip()
    print(info['Company Name'])
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

    return info
    

def main():
    data = []
    domain = 'https://www.dsebd.org/'
    main_page_url = domain + 'by_industrylisting.php'

    sectors = get_sectors(main_page_url)[:-1]
    for sector in sectors:
        print("Working on {} sector".format(sector['sector']))
        sectorURL = sector['url']
        companies = get_companies(domain + sectorURL)
        for i, company in enumerate(companies):
            data.append(get_info(domain + company, sector['sector']))
            sleep(randint(1,4))
        
    return data

if __name__ == "__main__":
    data = main()

    ordered_list=["Sector", "Company Name", "Trading Code", "LTP", "52 Weeks Moving Range Lowest", "52 Weeks Moving Range Highest"]
    wb=Workbook("DSE_Comapany_Details.xlsx")
    ws=wb.add_worksheet("New Sheet")
    
    first_row=0
    for header in ordered_list:
        col=ordered_list.index(header) 
        ws.write(first_row,col,header) 
    
    row=1
    for company in data:
        for _key,_value in company.items():
            col=ordered_list.index(_key)
            ws.write(row,col,_value)
        row+=1
    wb.close()
    