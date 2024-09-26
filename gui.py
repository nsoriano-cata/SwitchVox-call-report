import streamlit as st
import pandas as pd
import xlrd
from zipfile import BadZipFile
from openpyxl.utils.exceptions import InvalidFileException

def read_excel_file(file):
    try:
        # For .xlsx files
        return pd.read_excel(file, engine='openpyxl')
    except (BadZipFile, InvalidFileException):
        try:
            # For .xls files
            return pd.read_excel(file, engine='xlrd')
        except xlrd.biffh.XLRDError:
            st.error("The file is not a valid Excel file or is corrupted.")
        except Exception as e:
            st.error(f"An error occurred while reading the file: {str(e)}")
    except Exception as e:
        st.error(f"An unexpected error occurred: {str(e)}")
    return None
st.title("Excel File Browser")

uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx", "xls"])

if uploaded_file is not None:
    st.write("File selected:", uploaded_file.name)
    
    # Display file details
    file_details = {
        "Filename": uploaded_file.name,
        "File type": uploaded_file.type,
        "File size": f"{uploaded_file.size / 1024:.2f} KB"
    }
    st.write(file_details)
    
    # Preview the Excel file
    df = read_excel_file(uploaded_file)
    if df is not None:
        st.write("Preview of the Excel file:")
        st.dataframe(df.head())

        # **Additions start here**
        if 'Call to' in df.columns:
            unique_values = df['Call to'].unique()
            st.write("Unique values in the 'Call to' column:")
            st.write(unique_values)
        else:
            st.warning("The 'Call to' column is not found in the Excel file.")
        # **Additions end here**

    else:
        st.warning("Unable to read the file. Please make sure it's a valid Excel file.")