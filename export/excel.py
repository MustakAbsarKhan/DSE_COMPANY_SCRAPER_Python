"""Excel export architecture preview.

The private exporter creates a multi-sheet workbook containing raw scraped
rows, processed analysis, sector summaries, a watchlist, and data-quality
notes. The scoring rules, workbook formatting implementation, and derived
metric calculations are intentionally excluded from this public repository.
"""

from __future__ import annotations


WORKBOOK_SHEETS = [
    "Workbook_Guide",
    "Raw_Scraped_Data",
    "Processed_Analysis",
    "Sector_Summary",
    "Watchlist",
    "Data_Quality_Issues",
    "Dropped_Empty_Columns",
]


PROCESSED_ANALYSIS_COLUMNS = [
    "Company Name",
    "Trading Code",
    "Sector",
    "Latest EPS Used",
    "Latest NAVPS Used",
    "Latest P/E Used",
    "P/E vs Sector Median",
    "Debt to Market Cap",
    "Cash Flow to Profit",
    "Price Position in 52W Range",
    "Reliability Score",
    "Final Screening Score",
    "Valuation Label",
    "Positive Signals",
    "Negative Signals",
    "Key Risks",
    "What To Check Next",
]


def export_company_rows_to_excel(rows: list[dict[str, object]]) -> None:
    """Preview the public interface of the private workbook exporter."""
    raise NotImplementedError("Private implementation omitted from public preview.")
