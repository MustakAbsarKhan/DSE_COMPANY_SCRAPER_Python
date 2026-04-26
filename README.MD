
---
# 📊 DSE Company Data Scraper — Institutional-Grade Data Pipeline

![Python](https://img.shields.io/badge/Python-3.10%2B-blue)
![Async](https://img.shields.io/badge/Async-Enabled-success)
![Architecture](https://img.shields.io/badge/Architecture-Pipeline%20%7C%20Resilient-informational)
![Scraping](https://img.shields.io/badge/Scraping-Robust%20%26%20Fault--Tolerant-orange)
![Data Quality](https://img.shields.io/badge/Data-Structured%20%26%20Clean-brightgreen)
![License](https://img.shields.io/badge/License-MIT-lightgrey)

---

## 🚀 Overview

A **high-performance, fault-tolerant, asynchronous data pipeline** for extracting structured company-level financial data from the Dhaka Stock Exchange (DSE).

> ⚠️ DSE pages are **inconsistent, non-standardized, and structurally unstable**
> ✅ This project is engineered specifically to **handle real-world HTML unpredictability**

This is not a basic scraper — it is a **resilient data engineering system** designed to operate reliably under messy, changing web structures.

---

## 🧠 What Makes This Different

### 🛡️ Built for Unreliable HTML (Core Upgrade)

Unlike traditional scrapers that break when layouts change, this system:

* 🔍 Uses **multi-layer parsing strategies**
* 🧩 Extracts data across **multiple tables dynamically**
* 🔁 Applies **fallback logic instead of fixed structure dependency**
* 🧠 Handles **missing, shifted, or duplicated fields gracefully**

---

### ⚡ Performance + Reliability

* 🚀 Fully **asynchronous pipeline (aiohttp + asyncio)**
* 🔄 **Adaptive concurrency control**
* ⏱️ **Rate limiting + retry with jitter**
* 📉 Reduced failure rate on unstable endpoints

---

### 📊 Institutional-Level Data Structuring

Data is **not just scraped — it's normalized and analysis-ready**:

* Clean numeric conversion (`to_float`, `to_int`)
* Null-safe handling
* Flattened nested structures (shareholding tables)
* Time-aware metrics (EPS / P/E with date keys)
* Logically grouped financial data

---

## 🏗️ Architecture

```bash
main.py (Async Orchestrator)
   ↓
Semaphore-Controlled Concurrency
   ↓
sectors → companies → company_info
   ↓
robust parsing layer (multi-table extraction)
   ↓
data normalization + validation
   ↓
export (Excel structured dataset)
```

---

## 🧩 Core Engineering Concepts

### ✅ 1. Resilient Parsing Layer

* Avoids brittle indexing wherever possible
* Uses:

  * text-based matching
  * regex extraction
  * structural fallbacks

---

### ✅ 2. Data Normalization Layer

* Converts inconsistent raw values into:

  * floats
  * integers
  * clean strings
* Eliminates:

  * `"N/A"`, `"-"`, empty noise

---

### ✅ 3. Hybrid Extraction Strategy

| Method           | Use Case                |
| ---------------- | ----------------------- |
| Table Parsing    | Structured fields       |
| Regex Extraction | Hidden / scattered text |
| DOM Traversal    | Nested / broken layouts |

---

### ✅ 4. Handling DSE Structural Problems

| Problem                 | Solution              |
| ----------------------- | --------------------- |
| Different table layouts | Multi-table scanning  |
| Missing fields          | Safe fallback logic   |
| Nested tables           | Flattening strategy   |
| Duplicate keys          | Indexed normalization |
| Inconsistent EPS/P.E    | Period-based labeling |

---

## 📊 Data Coverage

### 🟢 Market Data

* LTP, Open, Close, YCP
* Intraday range
* 52-week range
* Price movement %

### 🔵 Liquidity & Volume

* Trade count
* Volume
* Value traded

### 🟡 Fundamentals

* Paid-up capital
* Authorized capital
* Reserves & OCI

### 🔴 Valuation (Advanced)

* EPS (Quarterly → Annual)
* EPS (Continuing Operations)
* P/E:

  * Basic
  * Diluted
  * Trailing
  * Audited vs Unaudited

### 🟣 Corporate Actions

* Dividends (Cash / Bonus / Right)
* AGM
* Year End

### ⚫ Ownership

* Shareholding breakdown (flattened + normalized)

---

## 📁 Project Structure

```bash
dse_scraper/
│
├── core/
│   ├── client.py        # Async HTTP client
│   ├── holidays.py      # Market calendar detection
│   ├── parser.py        # Robust extraction engine
│   ├── logger.py        # Logging system
│
├── pipelines/
│   ├── sectors.py
│   ├── companies.py
│   ├── company_info.py
│
├── export/
│   ├── excel.py
│
├── config.py
├── main.py
├── requirements.txt
└── README.md
```

---

## ⚙️ Tech Stack

* Python (asyncio)
* aiohttp
* BeautifulSoup (lxml parser)
* Pandas
* openpyxl

---

## 🧪 Execution

```bash
git clone https://github.com/yourusername/dse-scraper.git
cd dse-scraper
pip install -r requirements.txt
python main.py
```

---

## 📦 Output

* 📁 `Export_Data/DSE_Data_{Market_Date}.xlsx`
* 📄 `DSECompanyScraper.log`

✔ Clean
✔ Structured
✔ Analysis-ready

---

## 🤖 Automation

Includes GitHub Actions:

* ⏰ Runs Sunday–Thursday (Dhaka trading days)
* 🚫 Skips holidays automatically
* 📤 Uploads output as artifacts

---

## 🔄 Version Evolution

### 🔥 v3.0 — Resilient Parsing Upgrade (Current)

* Major redesign of parsing layer
* Handles **non-uniform DSE HTML structures**
* Multi-source extraction strategy
* Improved EPS & P/E extraction reliability
* Enhanced data cleaning & normalization

---

### ⚡ v2.0 — Async System Upgrade

* Async pipeline architecture
* Rate limiting + retries
* Logging system
* Modular design

---

## 🎯 Use Cases

* 📊 Financial modeling
* 📈 Equity research
* 🧠 Quant analysis
* 💼 Portfolio showcase
* 🤖 Automated data pipelines

---

## 🚧 Future Roadmap

* Database integration (PostgreSQL / MongoDB)
* REST API layer
* Stock ranking engine
* Valuation models (DCF / Relative)
* Dashboard (Streamlit / React)

---

## 👨‍💻 Author

**Mohammad Mustak Absar Khan**

***Original creator and maintainer of this project***

📧 [mustak.absar.khan@gmail.com](mailto:mustak.absar.khan@gmail.com)

---

## 📄 License

This project is licensed under the MIT License.

All derivatives must retain original attribution to the author.