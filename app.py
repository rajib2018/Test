import streamlit as st
import pandas as pd
import numpy as np

# Function to process vibration data
def process_vibration_data(df):
    processed_data = []
    equipment_name = None
    sub_eqpt = None

    # Identify the row with column headers
    try:
        header_row_index = df[df.iloc[:, 0] == 'S.NO'].index[0]
    except IndexError:
        st.warning("Could not find 'S.NO' in the uploaded file. Please check the file format.")
        return pd.DataFrame() # Return empty DataFrame if header is not found

    # Extract dates from the row below the main header
    dates_row = df.iloc[header_row_index + 1, 3:]
    dates = [d for d in dates_row if pd.notna(d)]

    # Identify the row with vibration directions
    try:
        directions_row_index = df[df.iloc[:, 2] == 'DIRECTION'].index[0]
    except IndexError:
         st.warning("Could not find 'DIRECTION' in the uploaded file. Please check the file format.")
         return pd.DataFrame() # Return empty DataFrame if directions row is not found


    for i in range(directions_row_index + 1, len(df)):
        row = df.iloc[i]
        # Check if the row contains equipment name
        if pd.notna(row.iloc[1]) and row.iloc[1] != 'EQUIPMENT NAME':
            equipment_name = row.iloc[1]
            sub_eqpt = None # Reset sub_eqpt for a new equipment

        # Check if the row contains sub equipment and direction
        if pd.notna(row.iloc[2]) and row.iloc[2] != 'DIRECTION' and equipment_name is not None:
             sub_eqpt = row.iloc[1] if pd.notna(row.iloc[1]) and row.iloc[1] != equipment_name else sub_eqpt # Update sub_eqpt if available and different from equipment_name in the current row
             direction = row.iloc[2]

             # Extract vibration readings for the identified dates
             vibration_readings = row.iloc[3:3+len(dates)].tolist()

             # Extract health status and other information - Added checks for index bounds
             health_status = row.iloc[-5] if len(row) > 5 and pd.notna(row.iloc[-5]) else None
             observations = row.iloc[-4] if len(row) > 4 and pd.notna(row.iloc[-4]) else None
             recommendations = row.iloc[-3] if len(row) > 3 and pd.notna(row.iloc[-3]) else None
             spectrums = row.iloc[-2] if len(row) > 2 and pd.notna(row.iloc[-2]) else None
             magnetic_center = row.iloc[-1] if len(row) > 1 and pd.notna(row.iloc[-1]) else None


             row_data = {
                 'Eqpt Name & Tag': equipment_name,
                 'Sub Eqpt': sub_eqpt,
                 'Eqpt Health': health_status,
                 'Defect if any': observations,
                 'Nature of Defect': recommendations,
                 'Status': spectrums if pd.notna(spectrums) else magnetic_center
             }

             # Add vibration readings for each date
             for j, date in enumerate(dates):
                 if j < len(vibration_readings):
                     reading = vibration_readings[j]
                     if pd.notna(reading):
                          if direction == 'H':
                              row_data[f'H{j+1}'] = reading
                          elif direction == 'G-s':
                              row_data[f'G-s{j+1}'] = reading
                          elif direction == 'V':
                              row_data[f'V{j+1}'] = reading
                          elif direction == 'A':
                              row_data[f'A{j+1}'] = reading

             # Add health status columns
             row_data['Normal'] = 1 if health_status == 'NORMAL' else 0
             row_data['Satisfactory'] = 1 if health_status == 'SATISFACTORY' else 0
             row_data['Unsatisfactory'] = 1 if health_status == 'UNSATISFACTORY' else 0
             row_data['Date'] = dates[0] if dates else None # Assuming the first date is the primary date for the row

             processed_data.append(row_data)


    # Create a DataFrame and reorder columns
    processed_df = pd.DataFrame(processed_data)

    # Define the expected columns from df_blank_format (assuming it's loaded elsewhere or defined)
    # In a real Streamlit app, you might load this from a file upload as well
    expected_columns = ['Date', 'Eqpt Name & Tag', 'Sub Eqpt', 'H1', 'G-s1', 'V1', 'A1', 'H2', 'G-s2', 'V2', 'A2', 'Normal', 'Satisfactory', 'Unsatisfactory', 'Eqpt Health', 'Defect if any', 'Nature of Defect', 'Status']


    # Add missing columns with NaN if they don't exist in processed_df
    for col in expected_columns:
        if col not in processed_df.columns:
            processed_df[col] = np.nan

    # Reorder columns to match the blank format
    processed_df = processed_df[expected_columns]

    return processed_df

# Streamlit app
st.set_page_config(layout="wide") # Set page layout to wide

st.title("Vibration Analysis Report Converter")

st.markdown("""
    <style>
    .main {
        background-color: #f0f2f6;
        padding: 20px;
        border-radius: 10px;
    }
    .stButton>button {
        background-color: #4CAF50;
        color: white;
        padding: 10px 24px;
        text-align: center;
        text-decoration: none;
        display: inline-block;
        font-size: 16px;
        margin: 4px 2px;
        transition-duration: 0.4s;
        cursor: pointer;
        border: none;
        border-radius: 8px;
    }
    .stButton>button:hover {
        background-color: #45a049;
        color: white;
    }
    </style>
""", unsafe_allow_html=True)


st.header("Upload Files")

# File upload widgets
source_file = st.file_uploader("Upload Vibration Analysis Report (Excel)", type=["xlsx"])
blank_format_file = st.file_uploader("Upload Blank Format Template (Excel)", type=["xlsx"])

combined_df = pd.DataFrame() # Initialize an empty DataFrame

if source_file is not None and blank_format_file is not None:
    try:
        # Load data from the source Excel file sheets
        df_dust_collection = pd.read_excel(source_file, sheet_name='DUST COLLECTION FANS')
        df_process_fans = pd.read_excel(source_file, sheet_name='PROCESS FANS')

        # Load the blank format file to get expected columns
        df_blank_format = pd.read_excel(blank_format_file)
        expected_columns = df_blank_format.columns.tolist()


        # Apply the processing function to both loaded DataFrames
        processed_df_dust = process_vibration_data(df_dust_collection.copy())
        processed_df_process = process_vibration_data(df_process_fans.copy())

        # Concatenate the two processed DataFrames
        combined_df = pd.concat([processed_df_dust, processed_df_process], ignore_index=True)

        st.header("Combined Vibration Analysis Data")
        st.dataframe(combined_df)

        # Add download button for the combined data
        @st.cache_data
        def convert_df_to_excel(df):
            return df.to_excel(index=False).encode('utf-8')

        excel_data = convert_df_to_excel(combined_df)

        st.download_button(
            label="Export Combined Data to Excel",
            data=excel_data,
            file_name='combined_vibration_analysis_report.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except Exception as e:
        st.error(f"An error occurred: {e}")

elif source_file is None and blank_format_file is None:
    st.info("Please upload both the Vibration Analysis Report and the Blank Format Template.")
elif source_file is None:
    st.info("Please upload the Vibration Analysis Report.")
else:
    st.info("Please upload the Blank Format Template.")
