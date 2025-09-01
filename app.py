import streamlit as st
import pandas as pd

def load_excel_sheets(file, sheetnames):
    """Load specific sheets from Excel file into a dictionary of DataFrames."""
    data = {}
    for sheet in sheetnames:
        try:
            data[sheet] = pd.read_excel(file, sheet_name=sheet)
        except Exception as e:
            st.error(f"⚠️ Could not load sheet '{sheet}': {e}")
    return data

st.set_page_config(page_title="Vibration Analysis Formatter", layout="wide")

st.title("📊 Vibration Analysis Report Formatter")

st.markdown(
    """
    Upload the **Vibration Analysis Report** and the **Blank Format Template**.  
    This app will extract:
    - **DUST COLLECTION FANS**
    - **PROCESS FANS**

    and display previews side by side.
    """
)

# File uploaders
report_file = st.file_uploader("Upload Vibration Analysis Report (.xlsx)", type=["xlsx"])
blank_file = st.file_uploader("Upload Blank Format File (.xlsx)", type=["xlsx"])

if report_file and blank_file:
    with st.spinner("Processing files..."):
        # Load required sheets from report
        report_sheets = load_excel_sheets(report_file, ["DUST COLLECTION FANS", "PROCESS FANS"])

        # Load blank format (first sheet only)
        blank_excel = pd.ExcelFile(blank_file)
        blank_template = pd.read_excel(blank_file, sheet_name=blank_excel.sheet_names[0])

    st.success("✅ Files successfully loaded!")

    # Show Dust Collection Fans
    if "DUST COLLECTION FANS" in report_sheets:
        st.subheader("📂 Dust Collection Fans")
        st.dataframe(report_sheets["DUST COLLECTION FANS"].head(15))

    # Show Process Fans
    if "PROCESS FANS" in report_sheets:
        st.subheader("📂 Process Fans")
        st.dataframe(report_sheets["PROCESS FANS"].head(15))

    # Show Blank Template
    st.subheader("📂 Blank Format Template (Preview)")
    st.dataframe(blank_template.head(15))

    st.info("✨ Transformation into blank format can be implemented in the next step.")

else:
    st.warning("⬆️ Please upload both Excel files to continue.")
