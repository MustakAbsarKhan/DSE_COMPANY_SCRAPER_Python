import pandas as pd
import os


def save_to_excel(data):
    # =============================
    # GET MARKET DATE (FROM SCRAPED DATA)
    # =============================
    market_date = None

    if data and isinstance(data, list):
        market_date = data[0].get("Market Date")

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
    df = pd.DataFrame(data)
    df.to_excel(file_path, index=False)

    print(f"\n✅ Data saved to: {file_path}")