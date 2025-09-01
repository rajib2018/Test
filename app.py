import streamlit as st
import pandas as pd
import io

def load_excel_sheets(file, sheetnames):
    """Load specific sheets from Excel file into a dictionary of DataFrames."""
    data = {}
    for sheet in sheetnames:
        try:
            df = pd.read_excel(file, sheet_name=sheet)
            df = df.dropna(how="all")  # remove blank rows
            data[sheet] = df
        except Exception as e:
            st.error(f"⚠️ Could not load sheet '{sheet}': {e}")
    return data

st.set_page_config(page_title="Vibration Analysis Formatter", layout="wide")

st.title("📊 Vibration Analysis Report Formatter")

st.markdown(
    """
    Upload the **Vibration Analysis Report** and the **Blank Format Template**.  
    This app will:
    - Extract **DUST COLLECTION FANS** & **PROCESS FANS**
    - Clean + Align with the **Blank Format**
    - Provide a downloadable Excel file
    """
)

# File uploaders
report_file = st.file_uploader("Upload Vibration Analysis Report (.xlsx)", type=["xlsx"])
blank_file = st.file_uploader("Upload Blank Format File (.xlsx)", type=["xlsx"])

if report_file and blank_file:
    with st.spinner("Processing files..."):
        # Load required sheets from report
        report_sheets = load_excel_sheets(report_file, ["DUST COLLECTION FANS", "PROCESS FANS"])

        # Load blank format structure (first sheet only)
        blank_excel = pd.ExcelFile(blank_file)
        blank_template = pd.read_excel(blank_file, sheet_name=blank_excel.sheet_names[0])
        blank_template = blank_template.dropna(how="all")

    st.success("✅ Files successfully loaded!")

    # Show previews
    if "DUST COLLECTION FANS" in report_sheets:
        st.subheader("📂 Dust Collection Fans (Preview)")
        st.dataframe(report_sheets["DUST COLLECTION FANS"].head(10))

    if "PROCESS FANS" in report_sheets:
        st.subheader("📂 Process Fans (Preview)")
        st.dataframe(report_sheets["PROCESS FANS"].head(10))

    st.subheader("📂 Blank Format Template (Preview)")
    st.dataframe(blank_template.head(10))

    # ---- TRANSFORMATION ----
    st.header("🔄 Transforming Data into Blank Format")

    # Example column mapping between raw sheet and blank template
    # (adjust keys/values to match your Blank Format headers exactly)
    column_map = {
        "S.NO": "S.NO",
        "EQUIPMENT NAME": "EQUIPMENT NAME",
        "DIRECTION": "DIRECTION",
        "Attribute": "DATE",        # will come from unpivot step if applied
        "Value": "VALUE",           # vibration measurement
        "SOURCE_SHEET": "SOURCE",   # custom field we add
    }

    transformed = []
    for sheet_name, df in report_sheets.items():
        df_clean = df.copy()
        df_clean["SOURCE_SHEET"] = sheet_name  # keep track of origin
        df_clean = df_clean.dropna(how="all")

        # Convert datetime columns to date only
        for col in df_clean.select_dtypes(include=["datetime64[ns]"]).columns:
            df_clean[col] = df_clean[col].dt.date

        # Build aligned dataframe
        aligned = pd.DataFrame(columns=blank_template.columns)
        for raw_col, target_col in column_map.items():
            if raw_col in df_clean.columns and target_col in aligned.columns:
                aligned[target_col] = df_clean[raw_col]

        # Fill other blank format columns with empty strings
        for col in aligned.columns:
            if col not in aligned.columns:
                aligned[col] = ""

        transformed.append(aligned)

    # Combine Dust + Process into one DataFrame
    final_df = pd.concat(transformed, ignore_index=True)
    final_df = final_df.dropna(how="all")  # ensure no blank rows
    st.success("✅ Data transformed successfully!")

    # Ensure DATE column is clean (no timestamp)
    if "DATE" in final_df.columns:
        final_df["DATE"] = pd.to_datetime(final_df["DATE"], errors="coerce").dt.date

    st.dataframe(final_df.head(20))

    # ---- EXPORT TO EXCEL ----
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        final_df.to_excel(writer, index=False, sheet_name="Formatted Report")

    st.download_button(
        label="💾 Download Full Formatted Report",
        data=output.getvalue(),
        file_name="DRP_Vibration_Analysis_Formatted.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

else:
    st.warning("⬆️ Please upload both Excel files to continue.")
