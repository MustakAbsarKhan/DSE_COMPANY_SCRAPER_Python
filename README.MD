# DSE Company Data Scraper

Asynchronous Dhaka Stock Exchange (DSE) company data scraper that collects sector-wise listed company data, normalizes the fields, and exports an analysis-ready Excel workbook.

The scraper is built for DSE pages that are often inconsistent, nested, and table-heavy. It combines table parsing, targeted DOM traversal, regex extraction, defensive null handling, and ordered Excel export.

## Current Features

- Async scraping with `aiohttp` and `asyncio`
- Adaptive request throttling with retry, delay, and concurrency control
- Sector discovery from the DSE industry listing page
- Company URL discovery per sector
- Company profile extraction with normalized market, financial, corporate, and ownership data
- Optional holiday checker for DSE market holidays and Friday/Saturday weekends
- Excel export with finance-analysis-friendly column order
- GitHub Actions workflow for scheduled and manual scraper runs

## Data Coverage

The exported workbook includes these major groups:

- Identification: company name, trading code, scrip code, sector, instrument type, listing metadata
- Market price: LTP, YCP, open, adjusted open, close, day range, 52-week range, change value and percent
- Liquidity: trade count, volume, traded value
- Capitalization and structure: market cap, free float cap, authorized capital, paid-up capital, securities, face value, market lot
- EPS metrics: quarterly, nine-month, half-yearly, annual EPS, EPS continuing operations, and diluted EPS continuing operations
- Audited annual financial performance:
  - `Aud_EPS_COP_Basic_{Year}`
  - `Aud_EPS_COP_Diluted_{Year}`
  - `Aud_NAVPS_{Year}`
  - `Aud_PCO_mn_{Year}`
  - `Aud_Profit_mn_{Year}`
  - `Aud_TCI_mn_{Year}`
- P/E ratios: basic EPS, diluted EPS, trailing ratio, audited basic EPS
- Corporate actions: AGM, year ended, dividend year/yield, cash dividend, bonus issue, right issue, year end
- Shareholding: flattened shareholding percentage rows
- Additional useful scraped DSE fields that are not already normalized

Noise fields such as `Closing Price Graph:`, `Total Trade Graph:`, and `Total Volume Graph:` are intentionally excluded from the export.

## Excel Output Order

The Excel writer uses explicit column groups in `export/excel.py`:

1. Identity and listing metadata
2. Price and movement
3. Liquidity
4. Capitalization and structure
5. EPS metrics
6. Audited annual financial performance
7. P/E ratios
8. Corporate actions
9. Shareholding
10. Remaining preserved scraped fields

This keeps the workbook suitable for stock screening, financial analysis, and decision support while still preserving useful scraped data.

Output files are written to:

```bash
Export_Data/DSE_Data_{Market_Date}.xlsx
```

Logs are written to:

```bash
DSECompanyScraper.log
```

## Project Structure

```bash
DSE_COMPANY_SCRAPER_Python/
|
├── core/
│   ├── client.py        # Async HTTP client with adaptive throttle
│   ├── holidays.py      # DSE holiday/weekend checker
│   ├── logger.py        # Shared logger setup
│   └── parser.py        # HTML parsing and field normalization
|
├── pipelines/
│   ├── sectors.py       # fetch_tradable_sector_links()
│   ├── companies.py     # fetch_company_urls_for_sector()
│   └── company_info.py  # fetch_company_profiles()
|
├── export/
│   └── excel.py         # export_company_rows_to_excel()
|
├── .github/workflows/
│   └── dse_scraper.yml  # Scheduled/manual GitHub Actions run
|
├── config.py
├── main.py
├── requirements.txt
└── README.MD
```

## Main Code Flow

```bash
main.py
  -> fetch_tradable_sector_links()
  -> fetch_company_urls_for_sector()
  -> fetch_company_profiles()
  -> extract_company_profile()
  -> export_company_rows_to_excel()
```

`main.py` processes sectors concurrently with a semaphore. Company pages are fetched through the shared `global_client`, so adaptive throttling learns across the full run.

## Key Parser Functions

`core/parser.py` contains the extraction layer:

- `parse_html()` converts raw HTML into BeautifulSoup
- `parse_float()`, `parse_int()`, and `clean_text()` normalize raw values
- `extract_company_table_key_values()` reads general DSE table key/value fields
- `extract_security_codes()` extracts trading and scrip codes
- `extract_price_change()` parses the nested change table
- `extract_basic_information()` parses the Basic Information section
- `extract_corporate_action_fields()` extracts dividend, AGM, year-end, reserve, and OCI fields
- `extract_listing_metadata()` extracts listing year, market category, and electronic share
- `extract_shareholding_percentages()` flattens shareholding rows
- `extract_eps_metrics()` extracts EPS and continuing-operations EPS
- `extract_pe_ratios_and_annual_performance()` extracts P/E ratios and audited annual performance metrics
- `extract_company_profile()` builds one final export row

The parser also preserves useful unmodeled DSE fields, while avoiding duplicates for fields that are already normalized under clearer names.

## Temporary Single-Company Debugging

`pipelines/company_info.py` includes a commented debug block:

```python
# company_urls = [
#     url for url in company_urls
#     if "ABBANK" in url.upper()
# ]
#
# print("Filtered URLs:", company_urls)  # DEBUG
```

Uncomment and change the trading-code text when you want to test the parser against one company only. Keep it commented for full production runs.

## Installation

```bash
git clone https://github.com/MustakAbsarKhan/DSE_COMPANY_SCRAPER_Python.git
cd DSE_COMPANY_SCRAPER_Python
pip install -r requirements.txt
```

## Run Locally

```bash
python main.py
```

The scraper will:

1. Fetch DSE sectors
2. Skip ignored non-equity sectors from `config.py`
3. Fetch company links for each sector
4. Scrape company pages
5. Normalize and combine all rows
6. Export the Excel workbook

## Configuration

`config.py` defines:

- `DOMAIN`: DSE base URL
- `MAIN_URL`: DSE industry listing page
- `IGNORED_SECTORS`: sectors skipped because they do not follow normal listed-equity data structure

Current ignored sectors:

- Corporate Bond
- Debenture
- G-SEC (T.Bond)

## GitHub Actions

The workflow at `.github/workflows/dse_scraper.yml`:

- Runs on Ubuntu
- Uses Python 3.11
- Installs dependencies
- Runs `python main.py`
- Uploads generated Excel files from `Export_Data/*.xlsx`

Schedule:

```yaml
cron: "20 14 * * 0-4"
```

This is Sunday through Thursday at 14:20 UTC.

Manual runs are also supported through `workflow_dispatch`.

## Tech Stack

- Python
- asyncio
- aiohttp
- BeautifulSoup
- lxml
- pandas
- openpyxl

## Notes

- DSE HTML structure can change. Some financial tables are still index-based because the source layout is complex.
- The parser is defensive: missing fields should produce blank cells instead of stopping the entire scraper.
- Generated Excel files and logs are runtime output and may change on every scraper run.

## Author

Mohammad Mustak Absar Khan

Original creator and maintainer of this project.

Email: [mustak.absar.khan@gmail.com](mailto:mustak.absar.khan@gmail.com)

## License

This project is licensed under the MIT License.

All derivatives must retain original attribution to the author.
