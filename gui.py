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

# Upload the Excel file
uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx", "xls"])

if uploaded_file is not None:
    # Read the file
    df = read_excel_file(uploaded_file)
    
    if df is not None:
        # Add title and dropdown after file is uploaded
        st.subheader("Time period for call data analysis")
        time_period = st.selectbox("Select time period", ["Month", "Week", "Day"])

        # You can now use the selected time period for further data analysis
        st.write(f"Selected time period: {time_period}")
    else:
        st.warning("Unable to read the file. Please make sure it's a valid Excel file.")
