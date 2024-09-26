import streamlit as st
import pandas as pd
import xlrd
from zipfile import BadZipFile
from openpyxl.utils.exceptions import InvalidFileException
from datetime import timedelta
from io import StringIO

# Manually created dictionary from the uploaded Excel file
value_group_dict = {
    'Dispatch Counter <5150>': 'Dispatch',
    'Transfer Answering Serice <6200>': 'Call Center',
    'MTM Dispatch II <6103>': 'MTM',
    'Dispatch  Counter III <5152>': 'Dispatch',
    'Brittany Struble <5452>': 'CSC',
    'MTM Dispatch I <6101>': 'MTM',
    'CATARIDE QUEUE BUSY VM  <5588>': 'CSC',
    'Dispatch CounterII <5151>': 'Dispatch',
    'Kelly Saylor <5450>': 'CSC',
    'Erin LaPean <5177>': 'CSC',
    'Luke Keller <5120>': 'CSC',
    'Customer Service <5100>': 'CSC',
    'Customer Service <5900>': 'CSC',
    'MTM  <6100>': 'MTM',
    'Belinda Ilgen <6105>': 'MTM',
    'CATA GO Downtown <5700>': 'CSC',
    'Jason Barlick <6104>': 'MTM',
    'Jordan Robinson <6102>': 'MTM',
    'CATA B Line <5780>': 'CSC',
    'CATA GO Answering Sv <5760>': 'Call Center',
    'MTM Dispatch <6090>': 'MTM',
    'CATA CATARIDE <6080>': 'CSC',
    'Transfer Answering Serice <6200>': 'CATA GO',
    'CATARIDE QUEUE BUSY VM  <5588>': 'CATARIDE',
    'CATA GO Downtown <5700>': 'CATAGO',
    'CATA B Line <5780>': 'B Line',
    'CATA GO Answering Sv <5760>': 'CATAGO',
    'CATA CATARIDE <6080>': 'CATARIDE'
}
# Function to read the uploaded Excel file
def read_excel_file(file):
    try:
        return pd.read_excel(file, engine='openpyxl')
    except (BadZipFile, InvalidFileException):
        try:
            return pd.read_excel(file, engine='xlrd')
        except xlrd.biffh.XLRDError:
            st.error("The file is not a valid Excel file or is corrupted.")
        except Exception as e:
            st.error(f"An error occurred while reading the file: {str(e)}")
    except Exception as e:
        st.error(f"An unexpected error occurred: {str(e)}")
    return None

# Function to convert seconds to HH:MM:SS format
def seconds_to_hms(seconds):
    return str(timedelta(seconds=int(seconds)))

# Function to process the data based on the selected time period and the dictionary
def process_data(df, time_period, value_group_dict):
    # Convert "Call Date" to datetime format
    df['Call Date'] = pd.to_datetime(df['Call Date'], errors='coerce')

    # Filter and group data based on the dictionary (value-group pairs)
    df_filtered = df[df['Call To'].isin(value_group_dict.keys())]
    df_filtered['Group'] = df_filtered['Call To'].map(value_group_dict)

    # Group data based on selected time period and the group values
    if time_period == "Month":
        df_grouped = df_filtered.groupby([df_filtered['Call Date'].dt.to_period('M'), 'Group'])
    elif time_period == "Week":
        df_grouped = df_filtered.groupby([df_filtered['Call Date'].dt.to_period('W'), 'Group'])
    elif time_period == "Day":
        df_grouped = df_filtered.groupby([df_filtered['Call Date'].dt.to_period('D'), 'Group'])

    # Calculate total call time and total number of calls
    df_summary = df_grouped.agg(
        total_call_time=('Call Time (seconds)', 'sum'),
        total_calls=('Call To', 'count')
    ).reset_index()

    # Sort by total call time in seconds before converting to HH:MM:SS format
    df_summary = df_summary.sort_values(by='total_call_time', ascending=False)

    # Convert total call time from seconds to HH:MM:SS format
    df_summary['Total Call Time HH:MM:SS'] = df_summary['total_call_time'].apply(seconds_to_hms)
    df_summary = df_summary.drop(columns=['total_call_time'])

    return df_summary

# Function to display data by separate time periods
def display_data(df_summary, time_period):
    if time_period == "Week":
        week_count = 1
        for week, week_df in df_summary.groupby('Call Date'):
            start_date = week.start_time.strftime("%m/%d")
            end_date = week.end_time.strftime("%m/%d")
            st.subheader(f"Week {week_count} ({start_date} to {end_date})")
            st.dataframe(week_df)
            week_count += 1
    elif time_period == "Day":
        for day, day_df in df_summary.groupby('Call Date'):
            st.subheader(f"Day: {day.strftime('%m/%d')}")
            st.dataframe(day_df)
    else:
        st.dataframe(df_summary)

# Function to save the data as a CSV file
def save_csv(df):
    output = StringIO()
    df.to_csv(output, index=False)
    return output.getvalue()

def main():
    st.title("Excel File Browser with Grouping and Save Option")

    # Upload the Excel file for data analysis
    uploaded_file = st.file_uploader("Choose an Excel file for data analysis", type=["xlsx", "xls"])

    if uploaded_file is not None:
        # Read the data analysis file
        df = read_excel_file(uploaded_file)

        if df is not None:
            # Ensure the required columns are present in the dataset
            required_columns = ['Call Date', 'Call To', 'Call Time (seconds)']
            if all(col in df.columns for col in required_columns):
                # Add title and dropdown after file is uploaded
                st.subheader("Time period for call data analysis")
                time_period = st.selectbox("Select time period", ["Month", "Week", "Day"])

                # Process the data based on the selected time period and the value-group dictionary
                df_summary = process_data(df, time_period, value_group_dict)

                # Display the grouped results by period (if Week or Day is selected)
                display_data(df_summary, time_period)

                # Add a "Save as CSV" button
                if st.button('Save as CSV'):
                    csv_data = save_csv(df_summary)
                    st.download_button(
                        label="Download CSV file",
                        data=csv_data,
                        file_name='grouped_data.csv',
                        mime='text/csv'
                    )

            else:
                st.warning(f"The dataset must contain the following columns: {required_columns}")
        else:
            st.warning("Unable to read the file. Please make sure it's a valid Excel file.")

if __name__ == "__main__":
    main()