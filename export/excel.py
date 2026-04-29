import pandas as pd
import os


IDENTITY_COLUMNS = [
    "Company Name",
    "Trading Code",
    "Scrip Code",
    "Sector",
    "Instrument Type",
    "Listing Year",
    "Market Category",
    "Electronic Share",
    "Debut Trading Date",
    "Present Operational Status",
    "Market Date",
    "Last Update",
]

PRICE_COLUMNS = [
    "LTP",
    "YCP",
    "Opening Price",
    "Adj Opening Price",
    "Closing Price",
    "Day Low",
    "Day High",
    "52W Low",
    "52W High",
    "Change Value",
    "Change %",
]

LIQUIDITY_COLUMNS = [
    "Day Trade No",
    "Day Volume",
    "Day Value (mn)",
]

MARKET_VALUE_COLUMNS = [
    "Market Cap (mn)",
    "Free Float Cap (mn)",
]

SHARE_CAPITAL_COLUMNS = [
    "Authorized Capital (mn)",
    "Paid-up Capital (mn)",
    "Total Securities",
    "Face Value",
    "Market Lot",
]

DEBT_AND_EQUITY_COLUMNS = [
    "Present Loan Status Date",
    "Short-term Loan (mn)",
    "Long-term Loan (mn)",
    "Total Loan (mn)",
    "Reserve & Surplus without OCI (mn)",
    "Other Comprehensive Income (OCI) (mn)",
]

EPS_PERIODS = [
    "Q1",
    "Q2",
    "HalfYearly",
    "Q3",
    "9Months",
    "Annual",
]

EPS_SUFFIXES = [
    "EPS",
    "EPS_COP",
    "Diluted_EPS_COP",
]

ANNUAL_PERFORMANCE_PREFIXES = [
    "Aud_EPS_COP_Basic",
    "Aud_EPS_COP_Diluted",
    "Aud_NAVPS",
    "Aud_PCO_mn",
    "Aud_Profit_mn",
    "Aud_TCI_mn",
]

PE_RATIO_SUFFIXES = [
    "PEwBasEPS",
    "PEwDilutEPS",
    "PEwTrailRatio",
    "PEwAuditBascEPS",
]

CORPORATE_ACTION_COLUMNS = [
    "Last Div Year",
    "Last Div Yield %",
    "Latest Dividend Status (%)",
    "Cash Dividend",
    "Bonus Issue (Stock Dividend)",
    "Right Issue",
    "Year End",
    "For the year ended",
    "Last AGM held on",
]


def existing_columns(columns, preferred_order):
    """Return preferred columns that exist in the scraped data."""
    return [column for column in preferred_order if column in columns]


def is_annual_performance_column(column):
    """Return True for audited annual performance columns."""
    return any(column.startswith(f"{prefix}_") for prefix in ANNUAL_PERFORMANCE_PREFIXES)


def annual_performance_sort_key(column):
    """Sort annual performance fields by year, then by metric."""
    metric_order = {
        prefix: index
        for index, prefix in enumerate(ANNUAL_PERFORMANCE_PREFIXES)
    }

    for prefix in ANNUAL_PERFORMANCE_PREFIXES:
        column_prefix = f"{prefix}_"
        if column.startswith(column_prefix):
            year_part = column[len(column_prefix):]
            year_tokens = [
                int(token)
                for token in year_part.replace("/", "_").replace("-", "_").split("_")
                if token.isdigit()
            ]
            sort_year = max(year_tokens) if year_tokens else 0
            return (-sort_year, metric_order[prefix], column)

    return (0, len(ANNUAL_PERFORMANCE_PREFIXES), column)


def is_eps_metric_column(column):
    """Return True for quarterly/annual EPS columns from the EPS table."""
    return any(
        column == f"{period}_{suffix}"
        for period in EPS_PERIODS
        for suffix in EPS_SUFFIXES
    )


def eps_metric_sort_key(column):
    """Sort EPS columns by period first, then EPS type."""
    for period_index, period in enumerate(EPS_PERIODS):
        for suffix_index, suffix in enumerate(EPS_SUFFIXES):
            if column == f"{period}_{suffix}":
                return (period_index, suffix_index, column)

    return (len(EPS_PERIODS), len(EPS_SUFFIXES), column)


def is_pe_ratio_column(column):
    """Return True for dynamic P/E ratio columns."""
    return any(column.endswith(f"_{suffix}") for suffix in PE_RATIO_SUFFIXES)


def pe_ratio_sort_key(column):
    """Sort P/E ratio columns by metric family, then by source date/header."""
    for suffix_index, suffix in enumerate(PE_RATIO_SUFFIXES):
        marker = f"_{suffix}"
        if column.endswith(marker):
            date_part = column[:-len(marker)]
            return (suffix_index, date_part, column)

    return (len(PE_RATIO_SUFFIXES), column, column)


def is_shareholding_column(column):
    """Return True for flattened shareholding percentage columns."""
    return "Share Holding Percentage" in column


def order_columns_for_analysis(columns):
    """Return a finance-analysis-friendly column order without dropping data."""
    original_columns = list(columns)
    ordered_columns = []

    fixed_groups = [
        IDENTITY_COLUMNS,
        PRICE_COLUMNS,
        LIQUIDITY_COLUMNS,
        MARKET_VALUE_COLUMNS,
        SHARE_CAPITAL_COLUMNS,
        DEBT_AND_EQUITY_COLUMNS,
    ]

    for group in fixed_groups:
        ordered_columns.extend(existing_columns(original_columns, group))

    eps_columns = sorted(
        [column for column in original_columns if is_eps_metric_column(column)],
        key=eps_metric_sort_key
    )
    ordered_columns.extend(eps_columns)

    audited_annual_columns = sorted(
        [column for column in columns if is_annual_performance_column(column)],
        key=annual_performance_sort_key
    )
    ordered_columns.extend(audited_annual_columns)

    pe_columns = sorted(
        [column for column in original_columns if is_pe_ratio_column(column)],
        key=pe_ratio_sort_key
    )
    ordered_columns.extend(pe_columns)

    ordered_columns.extend(existing_columns(original_columns, CORPORATE_ACTION_COLUMNS))

    shareholding_columns = [
        column for column in original_columns
        if is_shareholding_column(column)
    ]
    ordered_columns.extend(shareholding_columns)

    ordered_set = set(ordered_columns)
    remaining_columns = [
        column for column in original_columns
        if column not in ordered_set
    ]

    return ordered_columns + remaining_columns


def export_company_rows_to_excel(company_rows):
    """Save scraped company dictionaries into an Excel workbook."""
    # =============================
    # GET MARKET DATE (FROM SCRAPED DATA)
    # =============================
    market_date = None

    # The market date is the same for all rows, so the first company row is
    # enough to build a readable file name.
    if company_rows and isinstance(company_rows, list):
        market_date = company_rows[0].get("Market Date")

    # fallback if missing
    if not market_date:
        market_date = "Unknown_Date"

    # clean filename (Windows-safe)
    market_date = str(market_date).replace(",", "").replace(" ", "_").replace("/", "-")

    # =============================
    # FILE NAME
    # =============================
    folder = "Export_Data"
    os.makedirs(folder, exist_ok=True)  # Create folder if it doesn't exist
    file_name = f"DSE_Data_{market_date}.xlsx"
    file_path = os.path.join(folder, file_name)

    # =============================
    # SAVE EXCEL
    # =============================
    # Pandas turns a list of dictionaries into columns automatically. Missing
    # values become blank cells in Excel.
    df = pd.DataFrame(company_rows)
    df = df.reindex(columns=order_columns_for_analysis(list(df.columns)))
    df.to_excel(file_path, index=False)

    print(f"\n✅ Data saved to: {file_path}")
