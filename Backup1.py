import os
import pandas as pd
import xlwings as xw
import re
import streamlit as st
import sqlite3
import io  # for handling in-memory file-like objects

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

# Function to save the file into the SQLite database
def save_file_to_db(file_data, file_name):
    conn = sqlite3.connect('uploaded_files.db')
    cursor = conn.cursor()

    # Convert the file data into a binary format
    binary_data = file_data.read()

    # Insert the file into the database
    cursor.execute('''
        INSERT INTO files (file_name, file_data)
        VALUES (?, ?)
    ''', (file_name, binary_data))

    conn.commit()
    conn.close()

# Streamlit app for uploading files and calling the function
def convert_csv_to_xlsm():
    st.title("CSV to Macro-Enabled Excel (XLSM) Converter")
    st.sidebar.title("Configurations")

    uploaded_csvs = st.sidebar.file_uploader("Choose CSV files", type=["csv"], accept_multiple_files=True)
    xlsm_file = st.sidebar.file_uploader("Choose an existing Macro-Enabled Excel file (.xlsm)", type=["xlsm"])
    delimiter = st.sidebar.selectbox("Select the delimiter for your CSV files", options=[",", ";", "\t", "|"], index=0)
    start_button = st.sidebar.button("Start Conversion")

    # Create the SQLite database and table if they don't exist
    conn = sqlite3.connect('uploaded_files.db')
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS files (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            file_name TEXT NOT NULL,
            file_data BLOB NOT NULL,
            timestamp DATETIME DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    conn.commit()
    conn.close()

    if uploaded_csvs and xlsm_file and start_button:
        # Save the uploaded .xlsm file into SQLite DB
        save_file_to_db(xlsm_file, xlsm_file.name)
        st.write(f"File {xlsm_file.name} has been uploaded and saved to the database.")

        # Retrieve file from database (in binary format)
        conn = sqlite3.connect('uploaded_files.db')
        cursor = conn.cursor()
        cursor.execute('SELECT file_data FROM files WHERE file_name = ?', (xlsm_file.name,))
        file_data = cursor.fetchone()[0]  # Extract the binary data

        # Use the retrieved file_data to process the CSV files
        file_like_object = io.BytesIO(file_data)  # Convert binary data to file-like object in memory

        # Process each CSV and update the .xlsm file
        for uploaded_csv in uploaded_csvs:
            result = process_csv_to_macro_xlsm(uploaded_csv, file_like_object, delimiter)
            st.write(result)

        # Provide download button for the updated XLSM file
        st.download_button(
            label="Download updated XLSM file",
            data=file_like_object,
            file_name=f"{xlsm_file.name}",
            mime="application/vnd.ms-excel.sheet.macroenabled.12"
        )

# Run the app
if __name__ == "__main__":
    convert_csv_to_xlsm()
