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

def process_data(df, time_period):
    # Convert "Call Date" to datetime format
    df['Call Date'] = pd.to_datetime(df['Call Date'], errors='coerce')

    # Filter rows where "Call To" is "Dispatch Counter <5150>"
    call_to_values = ["Dispatch Counter <5150>"]  # You can extend this list if needed
    df_filtered = df[df['Call To'].isin(call_to_values)]
    
    # Group data based on selected time period
    if time_period == "Month":
        df_grouped = df_filtered.groupby(df_filtered['Call Date'].dt.to_period('M'))
    elif time_period == "Week":
        df_grouped = df_filtered.groupby(df_filtered['Call Date'].dt.to_period('W'))
    elif time_period == "Day":
        df_grouped = df_filtered.groupby(df_filtered['Call Date'].dt.to_period('D'))
    
    # Calculate total call time and total number of calls
    df_summary = df_grouped.agg(
        total_call_time=('Call Time', 'sum'),
        total_calls=('Call To', 'count')
    ).reset_index()

    return df_summary

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

        # Ensure the required columns are present in the dataset
        if 'Call Date' in df.columns and 'Call To' in df.columns and 'Call Time' in df.columns:
            # Process the data based on the selected time period
            df_summary = process_data(df, time_period)

            # Display the grouped results
            st.write(f"Grouped data based on {time_period}:")
            st.dataframe(df_summary)

        else:
            st.warning("The dataset does not contain the necessary columns: 'Call Date', 'Call To', and 'Call Time'.")
    else:
        st.warning("Unable to read the file. Please make sure it's a valid Excel file.")
