import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Production Report Sorter", layout="wide")

st.title("Production Report Sorter")
st.write("Upload the production report to sort it by Date and Shift.")

uploaded_file = st.file_uploader(
    "Upload Excel Workbook",
    type=["xls", "xlsx"]
)


def find_header(sheet):
    for i, row in sheet.iterrows():
        vals = [str(v).strip().lower() for v in row.values]
        if "date" in vals and "shift" in vals:
            return i
    return 0


def process_sheet(uploaded_file):

    excel = pd.ExcelFile(uploaded_file)

    # Read first sheet
    df = pd.read_excel(
        excel,
        sheet_name=excel.sheet_names[0],
        header=None
    )

    # Set header
    header_row = find_header(df)

    df.columns = df.iloc[header_row]
    df.columns.name = None

    df = df.iloc[header_row + 1:].reset_index(drop=True)

    # Remove incomplete rows
    df.dropna(
        subset=["Component", "Customer", "Rej. Qty."],
        inplace=True
    )

    # Convert Date
    df["Date"] = pd.to_datetime(
        df["Date"],
        errors="coerce"
    )

    df.dropna(subset=["Date"], inplace=True)

    # Numeric columns
    numeric_cols = [
        "Ok Qty.",
        "Rej. Qty.",
        "Total Qty.",
        "Total Cavity",
        "Run. Cavity"
    ]

    df[numeric_cols] = (
        df[numeric_cols]
        .apply(pd.to_numeric, errors="coerce")
        .fillna(0)
        .astype(int)
    )

    # Helper column for sorting (Month-Day)
    df["Sort_Date"] = df["Date"].dt.strftime("%m-%d")

    # Sort
    df = (
        df.sort_values(
            by=["Sort_Date", "Shift"],
            ascending=[True, True]
        )
        .reset_index(drop=True)
    )

    # Convert back for display
    df["Date"] = df["Date"].dt.strftime("%d-%b")

    # Remove helper column
    df.drop(columns=["Sort_Date"], inplace=True)

    return df


if uploaded_file is not None:

    with st.spinner("Sorting report... Please wait."):

        sorted_df = process_sheet(uploaded_file)

        output = BytesIO()

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            sorted_df.to_excel(
                writer,
                sheet_name="Sorted Data",
                startrow=4,
                index=False
            )

        output.seek(0)

    st.success("Sorting completed successfully.")

    st.dataframe(
        sorted_df,
        use_container_width=True,
        hide_index=True
    )

    st.download_button(
        label="Download Sorted Report",
        data=output,
        file_name="Sorted_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
