import streamlit as st
import pandas as pd
import io

# --- Streamlit Page Configuration ---
st.set_page_config(
    page_title="Data Pairing and Export Tool",
    page_icon="🔄",
    layout="centered",
    initial_sidebar_state="auto"
)

# --- Main Application ---
st.title("Automated Data Pairing and Export Tool")
st.markdown("""
This tool helps you pair data from two spreadsheets and export the results into a multi-tab Excel file.

**Instructions:**
1.  Upload the 'Items' spreadsheet containing `Item #`, `Model`, `Region`, and `VBU`.
2.  Upload the 'Regions' spreadsheet containing the region patch codes.
3.  The tool will generate all possible pairings.
4.  Click the 'Download Paired Data' button to get your formatted Excel file.
""")

# --- File Uploaders ---
st.header("1. Upload Your Spreadsheets")

# Uploader for the first spreadsheet (Items)
file1 = st.file_uploader("Upload Items Spreadsheet (.xlsx)", type=["xlsx"])

# Uploader for the second spreadsheet (Regions)
file2 = st.file_uploader("Upload Region Patches Spreadsheet (.xlsx)", type=["xlsx"])


# --- Main Logic: Process files only when both are uploaded ---
if file1 and file2:
    try:
        st.header("2. Processing and Generating Output")

        # --- Data Loading ---
        df_items = pd.read_excel(file1)
        df_regions = pd.read_excel(file2)

        # --- Data Validation ---
        # Validate columns in the first spreadsheet
        required_item_cols = {'Item #', 'Model', 'Region', 'VBU'}
        if not required_item_cols.issubset(df_items.columns):
            st.error(f"Error: The first spreadsheet must contain the following columns: {', '.join(required_item_cols)}")
            st.stop()

        # Validate the second spreadsheet is not empty and has one column
        if df_regions.shape[1] != 1:
            st.error("Error: The second spreadsheet should contain exactly one column with the region patch codes.")
            st.stop()
        
        # Standardize the column name for merging
        region_patch_col_name = "Region patch code"
        df_regions.columns = [region_patch_col_name]

        # --- Data Pairing (Cross Join) ---
        # Create a temporary key for a cartesian product merge
        df_items['_key'] = 1
        df_regions['_key'] = 1

        # Perform the merge and drop the temporary key
        df_merged = pd.merge(df_items, df_regions, on='_key').drop('_key', axis=1)

        # --- Structure the Output ---
        # Reorder and select the desired columns for the final output
        output_df = df_merged[['VBU', 'Item #', 'Model', region_patch_col_name]]
        
        st.success("Data has been successfully paired and structured!")
        st.write("Preview of the first 20 paired rows:")
        st.dataframe(output_df.head(20))

        # --- Excel Export Preparation ---
        st.header("3. Download Your File")

        # Create an in-memory buffer for the Excel file
        output_buffer = io.BytesIO()

        # Use ExcelWriter to write to multiple tabs in the buffer
        with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
            # Get unique region patches to create a tab for each
            unique_regions = output_df[region_patch_col_name].unique()

            for region in unique_regions:
                # Filter the dataframe for the current region
                df_region_specific = output_df[output_df[region_patch_col_name] == region]
                
                # Write the filtered dataframe to a new sheet named after the region
                # Use a truncated and safe name for the Excel sheet
                sheet_name = str(region)[:31] # Excel sheet names have a 31-char limit
                df_region_specific.to_excel(writer, sheet_name=sheet_name, index=False)
        
        # After the 'with' block, the writer is saved. Now, prepare buffer for download.
        output_buffer.seek(0)

        # --- Download Button ---
        st.download_button(
            label="⬇️ Download Paired Data as Excel File",
            data=output_buffer,
            file_name="paired_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"An error occurred: {e}")
        st.warning("Please ensure your files are in the correct format and contain the expected data.")

else:
    st.info("Please upload both spreadsheets to begin processing.")

# --- Footer ---
st.markdown("---")
st.write("Developed by your friendly Web Developer")
