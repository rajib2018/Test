import streamlit as st
import pandas as pd
import openpyxl

st.set_page_config(page_title="Vibration Analysis Formatter", layout="wide")

st.title("📊 Vibration Analysis Report Formatter")

st.markdown(
    """
    Upload the **DRP Vibration Analysis Report** file and the **Blank Format** file.  
    The app will extract the **DUST COLLECTION FANS** and **PROCESS FANS** sheets  
    and show previews before restructuring.
    """
)

# File uploaders
report_file = st.file_uploader("Upload Vibration Analysis Report (.xlsx)", type=["xlsx"])
blank_file = st.file_uploader("Upload Blank Format File (.xlsx)", type=["xlsx"])

if report_file and blank_file:
    try:
        # Load report excel
        report_excel = pd.ExcelFile(report_file)

        # Extract required sheets
        dust_collection_fans = pd.read_excel(report_file, sheet_name="DUST COLLECTION FANS")
        process_fans = pd.read_excel(report_file, sheet_name="PROCESS FANS")

        # Load blank format structure (first sheet only)
        blank_excel = pd.ExcelFile(blank_file)
        blank_template = pd.read_excel(blank_file, sheet_name=blank_excel.sheet_names[0])

        st.subheader("✅ Successfully Loaded Files")

        # Show sheet previews
        st.write("### 📂 Dust Collection Fans (Preview)")
        st.dataframe(dust_collection_fans.head(10))

        st.write("### 📂 Process Fans (Preview)")
        st.dataframe(process_fans.head(10))

        st.write("### 📂 Blank Format (Preview)")
        st.dataframe(blank_template.head(10))

        # (Next step) — transformation logic can go here
        st.info("Currently showing previews only. Transformation into blank format can be added here.")

    except Exception as e:
        st.error(f"❌ Error while processing files: {e}")

else:
    st.warning("Please upload both files to continue.")
