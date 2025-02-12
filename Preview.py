import pandas as pd
import streamlit as st

def preview_excel(uploaded_file):
    """
    Function to preview the first few rows of an Excel (.xlsx) file before processing.
    :param uploaded_file: The uploaded Excel file.
    :return: None
    """
    try:
        # Read the Excel file
        df = pd.read_excel(uploaded_file, sheet_name=None)  # Read all sheets into a dictionary

        # Get the first sheet name (you can modify this logic as needed)
        sheet_name = list(df.keys())[0]
        df_sheet = df[sheet_name]

        # Show the first few rows of the sheet
        st.write(f"Preview of the first sheet: {sheet_name}")
        df_sheet = df_sheet.fillna("")  # Fill NaN values with empty strings
        df_sheet = df_sheet.apply(lambda col: col.apply(lambda x: str(x).replace(",", "") if isinstance(x, (int, float, str)) else x))

        # Display the preview
        st.dataframe(df_sheet.head())  # Preview the first 5 rows of the sheet

    except Exception as e:
        st.error(f"Error loading the file: {e}")
