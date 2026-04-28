# DSE Company Data Scraper

> Production-minded async data pipeline for collecting, normalizing, and exporting Dhaka Stock Exchange company fundamentals from inconsistent public web pages.

[![Python](https://img.shields.io/badge/Python-3.11-3776AB?logo=python&logoColor=white)](https://www.python.org/)
[![Async](https://img.shields.io/badge/Async-aiohttp-2F80ED)](https://docs.aiohttp.org/)
[![Data Export](https://img.shields.io/badge/Export-Excel-217346?logo=microsoftexcel&logoColor=white)](#excel-output)
[![Automation](https://img.shields.io/badge/Automation-GitHub%20Actions-2088FF?logo=githubactions&logoColor=white)](#automation)
[![License](https://img.shields.io/badge/License-All%20Rights%20Reserved-red)](#license-and-usage)

## Why This Project Exists

The Dhaka Stock Exchange website exposes valuable company data, but the pages are not shaped like a clean API. Tables are nested, labels vary, values can move between sections, and missing fields are common. This project turns that messy public HTML into a structured Excel dataset that can be used for screening, research, and decision support.

This is not a toy scraper. It is built as a small data engineering system with async IO, adaptive throttling, retries, modular parsing, normalized financial fields, ordered exports, and scheduled GitHub Actions execution.

![Execution flow](Execution%20Flow.jpg)

## Table Of Contents

- [Highlights](#highlights)
- [What It Collects](#what-it-collects)
- [Architecture](#architecture)
- [Data Quality Strategy](#data-quality-strategy)
- [Excel Output](#excel-output)
- [Automation](#automation)
- [Run Locally](#run-locally)
- [Security And IP Protection](#security-and-ip-protection)
- [Portfolio Value](#portfolio-value)
- [Project Structure](#project-structure)
- [Roadmap](#roadmap)
- [License And Usage](#license-and-usage)

## Highlights

| Area                  | Implementation                                                                            |
| --------------------- | ----------------------------------------------------------------------------------------- |
| Async pipeline        | `asyncio` orchestration with `aiohttp` network calls                                      |
| Adaptive throttling   | Dynamic request delay and concurrency reduction after failures                            |
| Retry strategy        | Request retries with jitter to reduce transient failures                                  |
| Sector discovery      | Automatically discovers tradable sector links from DSE                                    |
| Company discovery     | Collects company detail URLs per sector                                                   |
| Robust parsing        | Combines table extraction, DOM traversal, regex, and fallback logic                       |
| Normalization         | Converts raw strings into clean numeric and text fields                                   |
| Financial coverage    | Market data, capital structure, EPS, P/E, annual audited metrics, dividends, shareholding |
| Analysis-ready export | Explicit Excel column order designed for finance workflows                                |
| Automation            | Scheduled and manual GitHub Actions workflows                                             |

## What It Collects

The final workbook is designed around practical equity research workflows.

| Group                      | Example Fields                                                                                   |
| -------------------------- | ------------------------------------------------------------------------------------------------ |
| Identity                   | Company name, trading code, scrip code, sector, instrument type                                  |
| Listing metadata           | Listing year, market category, electronic share, debut trading date                              |
| Market price               | LTP, YCP, open, adjusted open, close, day range, 52-week range                                   |
| Movement                   | Change value and change percentage                                                               |
| Liquidity                  | Trade count, volume, traded value                                                                |
| Capital structure          | Market cap, free float cap, authorized capital, paid-up capital, securities, face value          |
| EPS                        | Quarterly, half-yearly, nine-month, annual, continuing operations, diluted continuing operations |
| Annual audited performance | EPS, diluted EPS, NAVPS, operating cash flow, profit, total comprehensive income                 |
| Valuation                  | Basic EPS P/E, diluted EPS P/E, trailing P/E, audited basic EPS P/E                              |
| Corporate actions          | AGM, year ended, dividend year, dividend yield, cash dividend, bonus issue, right issue          |
| Ownership                  | Flattened shareholding percentage rows                                                           |

Noise fields such as graph labels are intentionally excluded from the export, while useful unmodeled DSE fields are still preserved at the end of the workbook.

## Architecture

```mermaid
flowchart TD
    A[main.py] --> B[Fetch sector links]
    B --> C[Filter ignored sectors]
    C --> D[Fetch company URLs per sector]
    D --> E[Fetch company pages concurrently]
    E --> F[Parse inconsistent HTML]
    F --> G[Normalize fields]
    G --> H[Order analysis columns]
    H --> I[Export Excel workbook]
```

The scraper keeps the pipeline split into small modules so each concern is easy to inspect, test, and extend.

```text
main.py
  -> pipelines/sectors.py
  -> pipelines/companies.py
  -> pipelines/company_info.py
  -> core/client.py
  -> core/parser.py
  -> export/excel.py
```

## Data Quality Strategy

DSE pages are inconsistent, so the parser does not rely on a single brittle selector. It uses layered extraction:

- Table scanning for structured key/value fields
- Targeted DOM traversal for nested sections
- Regex extraction for values embedded in page text
- Defensive null handling for missing or shifted fields
- Duplicate prevention when a field is already normalized under a clearer name
- Stable output ordering so every workbook is easy to compare over time

## Excel Output

Generated files follow this pattern:

```bash
Export_Data/DSE_Data_{Market_Date}.xlsx
```

Column groups are ordered for financial analysis:

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

## Automation

The GitHub Actions workflows support both scheduled and manual runs.

```yaml
cron: "20 14 * * 0-4"
```

This schedule runs Sunday through Thursday at 14:20 UTC, matching Bangladesh market days after the close. Generated Excel files are uploaded as workflow artifacts.

## Run Locally

```bash
git clone https://github.com/MustakAbsarKhan/DSE_COMPANY_SCRAPER_Python.git
cd DSE_COMPANY_SCRAPER_Python
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python main.py
```

On Linux or macOS, activate the environment with:

```bash
source .venv/bin/activate
```

## Configuration

`config.py` contains the main project configuration:

| Name              | Purpose                                        |
| ----------------- | ---------------------------------------------- |
| `DOMAIN`          | DSE base URL                                   |
| `MAIN_URL`        | DSE industry listing page                      |
| `IGNORED_SECTORS` | Non-equity-like sectors skipped by the scraper |

Current ignored sectors:

- Corporate Bond
- Debenture
- G-SEC (T.Bond)

## Security And IP Protection

This repository is designed to showcase engineering skill without exposing secrets or unnecessary runtime artifacts.

- No API keys, tokens, passwords, or private credentials are required.
- Runtime logs and generated Excel exports are ignored by Git.
- `.env` and virtual environment folders are ignored.
- Network access is limited to public DSE pages.
- Generated artifacts should be shared intentionally through GitHub Actions, not committed as source.
- See [SECURITY.md](SECURITY.md) for reporting guidance.

Important: public repositories can always be viewed and copied technically. Protection comes from license terms, attribution, commit history, public proof of authorship, and limiting what sensitive or proprietary material is published.

## Portfolio Value

This project demonstrates practical skills that matter in freelance and professional data roles:

- Async Python system design
- Resilient scraping under unstable HTML structures
- Financial data normalization
- Defensive parsing and graceful degradation
- Excel reporting for business users
- GitHub Actions automation
- Modular code organization
- Clear documentation for maintainability

## Project Structure

```text
DSE_COMPANY_SCRAPER_Python/
|
├── .github/workflows/
│   ├── dse_scraper.yml      # Scheduled/manual production-style run
│   └── manual.yml           # Manual execution workflow
|
├── core/
│   ├── client.py            # Async HTTP client with adaptive throttle
│   ├── holidays.py          # DSE holiday/weekend checker
│   ├── logger.py            # Shared logger setup
│   └── parser.py            # HTML parsing and normalization
|
├── export/
│   └── excel.py             # Analysis-ready Excel export
|
├── pipelines/
│   ├── sectors.py           # Sector discovery
│   ├── companies.py         # Company URL discovery
│   └── company_info.py      # Company profile scraping
|
├── config.py
├── main.py
├── requirements.txt
├── SECURITY.md
└── README.MD
```

## Roadmap

- Add unit tests for parser helpers and Excel column ordering
- Add optional CSV and Parquet exports
- Add historical snapshot storage
- Add schema validation for exported rows
- Add a dashboard layer for screening and comparisons
- Add CI checks for formatting and tests

## Author

**Mohammad Mustak Absar Khan**

Original creator and maintainer.

- GitHub: [MustakAbsarKhan](https://github.com/MustakAbsarKhan)
- Email: [mustak.absar.khan@gmail.com](mailto:mustak.absar.khan@gmail.com)

## License And Usage

Copyright (c) Mohammad Mustak Absar Khan.

This project is published for portfolio review, educational inspection, and hiring evaluation. Commercial use, resale, redistribution, or claiming this work as your own is not permitted without written permission from the author.

See [LICENSE.md](LICENSE.md) for the full terms.
