# Production Report Sorter

A Streamlit application that automates the sorting of production rejection reports.

## Features

- Upload an Excel (.xls/.xlsx) production report.
- Automatically detects the header row.
- Cleans incomplete records.
- Converts and validates data types.
- Sorts records by **Month-Day** and **Shift (A → B)**.
- Displays the sorted report in the browser.
- Downloads the sorted report as an Excel file.

## Tech Stack

- Python
- Pandas
- Streamlit
- OpenPyXL

## How to Run

1. Install the required packages:

pip install streamlit pandas openpyxl xlrd

2. Start the application:

streamlit run app.py

3. Upload the production report and download the sorted Excel file.

## Output

The application generates a sorted Excel report with:
- Date displayed in `dd-mmm` format.
- Records sorted by Month-Day and Shift.
- Five blank rows reserved at the top for future report headers.
