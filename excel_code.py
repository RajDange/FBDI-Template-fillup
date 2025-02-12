import os
import shutil
import pandas as pd
import xlwings as xw
import re
import streamlit as st
import tempfile
from Preview import preview_excel  # Assuming you'll need a new preview function for Excel files.

# Function to create a valid sheet name from the file name
def create_valid_sheet_name(file_name):
    name = os.path.splitext(file_name)[0]  # Remove extension
    # Remove any invalid characters from the sheet name
    name = re.sub(r'[\\/*?:\[\]]', '', name)  # Remove invalid characters
    return name[:31]  # Excel sheet name max length is 31 characters

# Function to populate data from Excel into a macro-enabled Excel (.xlsm) file
def process_excel_to_macro_xlsm(excel_file, xlsm_file_path):
    try:
        # Read Excel file into DataFrame
        df = pd.read_excel(excel_file, sheet_name=None)  # Read all sheets into a dictionary of DataFrames

        # For this example, let's assume we use the first sheet
        sheet_name = list(df.keys())[0]
        df_sheet = df[sheet_name]

        # Create a valid sheet name
        sheet_name = create_valid_sheet_name(excel_file.name)

        # Open the macro-enabled Excel file with xlwings
        app = xw.App(visible=False)  # Set visible=False to run in the background

        # Open the copied file
        wb = app.books.open(xlsm_file_path)

        # Check if the sheet already exists; if not, create it
        if sheet_name in [sheet.name for sheet in wb.sheets]:
            sheet = wb.sheets[sheet_name]


        # # If the DataFrame is empty (i.e., no data), clear the data from the 5th row onward
        # if df_sheet.empty:
        #     sheet.range("A5").expand().clear_contents()  # Clear contents starting from the 5th row
        #     return f"{excel_file.name} is blank. Data in {sheet_name} has been cleared."

        # If there is data, clear the contents from the 5th row onward and write new data
        sheet.range("A5").expand().clear_contents()  # Clear contents starting from the 5th row
        sheet.range("A5").value = df_sheet.values.tolist()  # Write the new data starting from cell A5

        # Save and close the workbook
        wb.save(xlsm_file_path)
        wb.close()
        app.quit()

        return f"Data from {excel_file.name} successfully added to {sheet_name} in {xlsm_file_path}"

    except Exception as e:
        return f"Error processing file {excel_file.name}: {str(e)}"

# Streamlit app for uploading files and calling the function
def convert_excel_to_xlsm():
    st.title("Excel to Macro-Enabled Excel (XLSM) Converter")
    st.sidebar.title("Configurations")

    # Upload multiple Excel files
    uploaded_excels = st.sidebar.file_uploader("Choose Excel files", type=["xlsx"], accept_multiple_files=True)

    # Upload the existing Macro-Enabled Excel file
    xlsm_file = st.sidebar.file_uploader("Choose an existing Macro-Enabled Excel file (.xlsm)", type=["xlsm"])

    # Start Conversion button
    start_button = st.sidebar.button("Start Conversion")

    if uploaded_excels and xlsm_file and start_button:
        # Save the uploaded .xlsm file to a temporary location with a writable path
        # Use os.path.join to ensure the correct path separator for your platform
        temp_dir = tempfile.gettempdir()
        xlsm_file_path = os.path.join(temp_dir, xlsm_file.name)

        # Write the uploaded .xlsm file to the temporary file path
        with open(xlsm_file_path, 'wb') as tmp_file:
            tmp_file.write(xlsm_file.getvalue())
        
        original_xlsm_name = os.path.splitext(xlsm_file.name)[0]  # Get original name (without .xlsm extension)

        # Ensure the .xlsm file is not marked as read-only
        writable_xlsm_file_path = os.path.join(temp_dir, f"{original_xlsm_name}.xlsm")

        # Check if the source and destination are the same, and if so, skip the copy
        if xlsm_file_path != writable_xlsm_file_path:
            shutil.copy(xlsm_file_path, writable_xlsm_file_path)

        # Create the progress bar
        progress = st.progress(0)
        num_files = len(uploaded_excels)

        # Process each Excel file and update the .xlsm file
        for idx, uploaded_excel in enumerate(uploaded_excels):
            result = process_excel_to_macro_xlsm(uploaded_excel, writable_xlsm_file_path)
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

    # Preview the uploaded Excel files
    for uploaded_file in uploaded_excels:
        st.subheader(f"Preview of {uploaded_file.name}")
        preview_excel(uploaded_file)  # Assuming a similar preview function for Excel

# Run the app
if __name__ == "__main__":
    convert_excel_to_xlsm()
