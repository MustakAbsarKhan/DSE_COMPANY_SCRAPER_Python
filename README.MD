# 📊 DSE Company Data Scraper

![Python](https://img.shields.io/badge/Python-3.10%2B-blue)
![Status](https://img.shields.io/badge/Status-Active-success)
![Architecture](https://img.shields.io/badge/Architecture-Pipeline%20%7C%20Modular-informational)
![License](https://img.shields.io/badge/License-MIT-lightgrey)

---

## 🚀 Overview

A **production-ready, modular pipeline-based web scraper** designed to extract structured company data from the Dhaka Stock Exchange (DSE).

This project demonstrates **clean architecture, scalable design, and real-world scraping practices**, making it suitable for both **portfolio showcasing and freelancing work**.

---

## 🧠 Key Highlights

* 🔄 **Pipeline-Based Architecture** (step-by-step data flow)
* 🧩 **Modular Design** (clean separation of concerns)
* 🔁 **Retry Mechanism** for resilient scraping
* ⚡ **Efficient Data Extraction** using BeautifulSoup
* 📦 **Structured Data Output** (Excel via Pandas)
* 🛠️ **Maintainable & Scalable Codebase**

---

## 🏗️ Architecture

```bash
main.py
   ↓
sectors pipeline
   ↓
companies pipeline
   ↓
company info pipeline
   ↓
export module
```

---

## 📁 Project Structure

```bash
dse_scraper/
│
├── core/
│   ├── client.py        # HTTP session & retry logic
│   ├── parser.py        # HTML parsing logic
│
├── pipelines/
│   ├── sectors.py       # Fetch sectors
│   ├── companies.py     # Fetch companies per sector
│   ├── company_info.py  # Extract company data
│
├── export/
│   ├── excel.py         # Export data to Excel
│
├── config.py            # Configurations
├── main.py              # Entry point
├── requirements.txt
└── README.md
```

---

## 📊 Data Extracted

The scraper collects the following:

* Company Name
* Trading Code
* Sector
* Last Traded Price (LTP)
* 52 Week Low & High
* Last Dividend Year
* Dividend %
* Dividend Yield

---

## ⚙️ Tech Stack

* **Python**
* **Requests** (HTTP handling)
* **BeautifulSoup (bs4)** (HTML parsing)
* **Pandas** (data processing & export)
* **openpyxl** (Excel engine)
* **tqdm** (progress visualization)

---

## 🧪 How to Run

### 1. Clone the repository

```bash
git clone https://github.com/yourusername/dse-scraper.git
cd dse-scraper
```

### 2. Install dependencies

```bash
pip install -r requirements.txt
```

### 3. Run the scraper

```bash
python main.py
```

---

## 📦 Output

* Generates: `DSE_Company_Details.xlsx`
* Structured, clean dataset ready for analysis

---

## 🔄 Version History

### 🔹 v2.0 (Current)

* Refactored into **pipeline-based modular architecture**
* Improved maintainability and scalability
* Separated parsing, networking, and processing logic

### 🔹 v1.0

* Initial monolithic scraper implementation

---

## 🎯 Use Cases

* 📈 Stock market data collection
* 📊 Financial analysis datasets
* 🤖 Automation & scraping projects
* 💼 Freelancing portfolio showcase

---

## 🚧 Future Improvements

* Async scraping (aiohttp) for performance boost
* Proxy rotation support
* CLI arguments for customization
* Logging system for production debugging
* API-based data serving

---

## 👨‍💻 Author

**Mohammad Mustak Absar Khan**
📧 [mustak.absar.khan@gmail.com](mailto:mustak.absar.khan@gmail.com)

---

## ⭐ Why This Project Matters

This project demonstrates the transition from:

> ❌ Basic scripting → ✅ Scalable software design

It reflects **real-world engineering practices**, including:

* Refactoring
* Modular architecture
* Pipeline design

---

## 📜 License

This project is licensed under the MIT License.

---

## 🌟 Support

If you found this useful:

* ⭐ Star the repository
* 🍴 Fork it
* 🛠️ Contribute

---
