import os
import shutil
import pandas as pd
import xlwings as xw
import re
import streamlit as st
import tempfile
from Preview import preview_csv

# Function to create a valid sheet name from the file name
def create_valid_sheet_name(file_name):
    name = os.path.splitext(file_name)[0]  # Remove extension
    # Remove any invalid characters from the sheet name
    name = re.sub(r'[\\/*?:\[\]]', '', name)  # Remove invalid characters
    return name[:31]  # Excel sheet name max length is 31 characters

# Function to populate data from CSV into a macro-enabled Excel (.xlsm) file
def process_csv_to_macro_xlsm(csv_file, xlsm_file_path, delimiter):
    try:
        # Read CSV into DataFrame, skipping the first 4 rows (header-level attributes)
        df = pd.read_csv(csv_file, delimiter=delimiter, skiprows=4)

        # Create a valid sheet name
        sheet_name = create_valid_sheet_name(csv_file.name)

        # Open the macro-enabled Excel file with xlwings
        app = xw.App(visible=False)  # Set visible=False to run in the background

        # Open the copied file
        wb = app.books.open(xlsm_file_path)

        # Check if the sheet already exists; if not, create it
        if sheet_name in [sheet.name for sheet in wb.sheets]:
            sheet = wb.sheets[sheet_name]

        # If the DataFrame is empty (i.e., no data), clear the data from the 5th row onward
        if df.empty:
            sheet.range("A5").expand().clear_contents()  # Clear contents starting from the 5th row
            return f"{csv_file.name} is blank. Data in {sheet_name} has been cleared."

        # If there is data, clear the contents from the 5th row onward and write new data
        sheet.range("A5").expand().clear_contents()  # Clear contents starting from the 5th row
        sheet.range("A5").value = df.values.tolist()  # Write the new data starting from cell A5

        # Save and close the workbook
        wb.save(xlsm_file_path)
        wb.close()
        app.quit()

        return f"Data from {csv_file.name} successfully added to {sheet_name} in {xlsm_file_path}"

    except Exception as e:
        return f"Error processing file {csv_file.name}: {str(e)}"

# Streamlit app for uploading files and calling the function
def convert_csv_to_xlsm():
    st.title("CSV to Macro-Enabled Excel (XLSM) Converter")
    st.sidebar.title("Configurations")

    # Upload multiple CSV files
    uploaded_csvs = st.sidebar.file_uploader("Choose CSV files", type=["csv"], accept_multiple_files=True)

    # Upload the existing Macro-Enabled Excel file
    xlsm_file = st.sidebar.file_uploader("Choose an existing Macro-Enabled Excel file (.xlsm)", type=["xlsm"])

    # Ask for the delimiter for CSV files
    delimiter = st.sidebar.selectbox(
        "Select the delimiter for your CSV files",
        options=[",", ";", "\t", "|"],
        index=0  # Default to comma
    )

    # Start Conversion button
    start_button = st.sidebar.button("Start Conversion")

    if uploaded_csvs and xlsm_file and start_button:
        # Save the uploaded .xlsm file to a temporary location with a writable path
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsm") as tmp_xlsm_file:
            tmp_xlsm_file.write(xlsm_file.getvalue())
            xlsm_file_path = tmp_xlsm_file.name  # Get the temporary file path
            original_xlsm_name = os.path.splitext(xlsm_file.name)[0]  # Get original name (without .xlsm extension)

        # Ensure the .xlsm file is not marked as read-only
        writable_xlsm_file_path = f"/tmp/{original_xlsm_name}.xlsm"
        shutil.copy(xlsm_file_path, writable_xlsm_file_path)

        # Create the progress bar
        progress = st.progress(0)
        num_files = len(uploaded_csvs)

        # Process each CSV and update the .xlsm file
        for idx, uploaded_csv in enumerate(uploaded_csvs):
            result = process_csv_to_macro_xlsm(uploaded_csv, writable_xlsm_file_path, delimiter)
            st.write(result)

            # Update the progress bar
            progress_percentage = (idx + 1) / num_files
            progress.progress(progress_percentage)

        # Provide a download button for the updated XLSM file with the same name as the original
        with open(writable_xlsm_file_path, "rb") as f:
            st.download_button(
                label="Download updated XLSM file",
                data=f,
                file_name=f"{original_xlsm_name}.xlsm",  # Save with the original .xlsm file name
                mime="application/vnd.ms-excel.sheet.macroenabled.12"
            )

    for uploaded_file in uploaded_csvs:
        st.subheader(f"Preview of {uploaded_file.name}")
        preview_csv(uploaded_file, delimiter)

# Run the app
if __name__ == "__main__":
    convert_csv_to_xlsm()
