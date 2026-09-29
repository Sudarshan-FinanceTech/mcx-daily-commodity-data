import json
from pathlib import Path
from datetime import datetime

import pandas as pd
from bs4 import BeautifulSoup
from curl_cffi import requests


# =========================================================
# SETTINGS
# =========================================================

# Store the Excel output inside the project's "data" folder.
# This makes the project portable and GitHub-friendly.

OUTPUT_FILE = (
    Path(__file__).resolve().parent
    / "data"
    / "MCX_Commodity_Data.xlsx"
)

# Create the data folder automatically if it does not exist
OUTPUT_FILE.parent.mkdir(
    parents=True,
    exist_ok=True
)

# Official MCX Bhav Copy page
MCX_URL = "https://www.mcxindia.com/market-data/bhavcopy"


# =========================================================
# DOWNLOAD MCX PAGE
# =========================================================

def download_mcx_page():

    print("Connecting to MCX...")

    response = requests.get(
        MCX_URL,
        impersonate="chrome",
        timeout=60
    )

    response.raise_for_status()

    print("MCX connection successful.")

    return response.text


# =========================================================
# EXTRACT BHAV COPY JSON
# =========================================================

def extract_mcx_data(html):

    print("Extracting MCX Bhav Copy...")

    soup = BeautifulSoup(
        html,
        "html.parser"
    )

    data_container = soup.find(
        "div",
        id="bhavcopy-data"
    )

    if data_container is None:

        raise ValueError(
            "MCX Bhav Copy data was not found."
        )

    json_text = data_container.get_text(
        strip=True
    )

    mcx_data = json.loads(
        json_text
    )

    df = pd.DataFrame(
        mcx_data
    )

    print(
        f"Total MCX records found: {len(df)}"
    )

    return df


# =========================================================
# FILTER COMMODITY FUTURES
# =========================================================

def filter_commodity_futures(df):

    print("Filtering commodity futures...")

    commodity_df = df[
        df["InstrumentName"] == "FUTCOM"
    ].copy()

    if commodity_df.empty:

        raise ValueError(
            "No FUTCOM records found."
        )

    print(
        f"MCX commodity futures: {len(commodity_df)}"
    )

    return commodity_df


# =========================================================
# CLEAN DATA
# =========================================================

def clean_data(df):

    columns = [
        "Date",
        "Symbol",
        "ExpiryDate",
        "Open",
        "High",
        "Low",
        "Close",
        "PreviousClose",
        "Volume",
        "VolumeInThousands",
        "Value",
        "OpenInterest",
        "InstrumentName"
    ]

    # Check that all required columns exist
    missing_columns = [
        column
        for column in columns
        if column not in df.columns
    ]

    if missing_columns:

        raise ValueError(
            f"Missing MCX columns: {missing_columns}"
        )

    df = df[
        columns
    ].copy()

    # Convert MCX date into a proper date
    df["Date"] = pd.to_datetime(
        df["Date"],
        errors="coerce"
    ).dt.date

    # Remove records where MCX date could not be read
    df = df.dropna(
        subset=["Date"]
    )

    # Sort data
    df = df.sort_values(
        by=[
            "Date",
            "Symbol",
            "ExpiryDate"
        ]
    )

    return df


# =========================================================
# SAVE / APPEND TO EXCEL
# =========================================================

def save_to_excel(new_data):

    print("Preparing Excel file...")

    # -----------------------------------------------------
    # If Excel already exists
    # -----------------------------------------------------

    if OUTPUT_FILE.exists():

        print(
            f"Existing Excel file found: {OUTPUT_FILE}"
        )

        existing_data = pd.read_excel(
            OUTPUT_FILE,
            sheet_name="MCX Data"
        )

        # Convert existing dates
        existing_data["Date"] = pd.to_datetime(
            existing_data["Date"],
            errors="coerce"
        ).dt.date

        print(
            f"Existing records: {len(existing_data)}"
        )

        # Combine old + new
        combined_data = pd.concat(
            [
                existing_data,
                new_data
            ],
            ignore_index=True
        )

    else:

        print(
            "No previous Excel file found."
        )

        combined_data = new_data.copy()

    # -----------------------------------------------------
    # Remove duplicate contract/date records
    # -----------------------------------------------------

    combined_data = combined_data.drop_duplicates(
        subset=[
            "Date",
            "Symbol",
            "ExpiryDate"
        ],
        keep="last"
    )

    # Sort data
    combined_data = combined_data.sort_values(
        by=[
            "Date",
            "Symbol",
            "ExpiryDate"
        ]
    )

    # -----------------------------------------------------
    # Determine MCX data date
    # -----------------------------------------------------

    latest_date = new_data["Date"].max()

    # -----------------------------------------------------
    # Run information
    # -----------------------------------------------------

    run_time = datetime.now().strftime(
        "%d-%b-%Y %H:%M:%S"
    )

    run_log_entry = pd.DataFrame(
        [{
            "Run Timestamp": run_time,
            "MCX Data Date": latest_date,
            "Records Downloaded": len(new_data),
            "Total Records in Excel": len(combined_data),
            "Status": "SUCCESS"
        }]
    )

    # -----------------------------------------------------
    # Existing Run Log
    # -----------------------------------------------------

    if OUTPUT_FILE.exists():

        try:

            existing_log = pd.read_excel(
                OUTPUT_FILE,
                sheet_name="Run Log"
            )

            run_log = pd.concat(
                [
                    existing_log,
                    run_log_entry
                ],
                ignore_index=True
            )

        except Exception:

            run_log = run_log_entry

    else:

        run_log = run_log_entry

    # -----------------------------------------------------
    # README SHEET
    # -----------------------------------------------------

    readme = pd.DataFrame({

        "MCX Daily Commodity Data": [
            "Source",
            "Instrument",
            "Frequency",
            "Data Type",
            "Purpose",
            "Last Run"
        ],

        "Details": [
            "MCX Bhav Copy",
            "FUTCOM",
            "Daily",
            "End-of-Day",
            "Personal market research",
            run_time
        ]

    })

    # -----------------------------------------------------
    # WRITE EXCEL
    # -----------------------------------------------------

    with pd.ExcelWriter(
        OUTPUT_FILE,
        engine="openpyxl"
    ) as writer:

        combined_data.to_excel(
            writer,
            sheet_name="MCX Data",
            index=False
        )

        run_log.to_excel(
            writer,
            sheet_name="Run Log",
            index=False
        )

        readme.to_excel(
            writer,
            sheet_name="README",
            index=False
        )

    # -----------------------------------------------------
    # SUCCESS MESSAGE
    # -----------------------------------------------------

    print("")
    print("========================================")
    print("MCX DATA UPDATE SUCCESSFUL")
    print("========================================")
    print(
        f"MCX Data Date       : {latest_date}"
    )
    print(
        f"Records Downloaded  : {len(new_data)}"
    )
    print(
        f"Total Excel Records : {len(combined_data)}"
    )
    print(
        f"Excel Location      : {OUTPUT_FILE}"
    )
    print("========================================")


# =========================================================
# MAIN PROGRAM
# =========================================================

def main():

    try:

        html = download_mcx_page()

        df = extract_mcx_data(
            html
        )

        commodity_df = filter_commodity_futures(
            df
        )

        commodity_df = clean_data(
            commodity_df
        )

        save_to_excel(
            commodity_df
        )

    except Exception as e:

        print("")
        print("========================================")
        print("MCX DATA UPDATE FAILED")
        print("========================================")
        print(
            f"Error: {e}"
        )
        print("========================================")


# =========================================================
# RUN
# =========================================================

if __name__ == "__main__":

    main()
