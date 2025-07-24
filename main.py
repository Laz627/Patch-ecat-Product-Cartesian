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
This tool intelligently pairs items with their regional patch codes and exports the results.

**Instructions:**
1.  Upload the 'Items' spreadsheet containing `Item #`, `Model`, `VBU`, and `Region`.
2.  Upload the 'Region Patches' spreadsheet where columns are the region names (e.g., "East", "West").
3.  The tool will match items to the patch codes based on their assigned region.
4.  Click the 'Download Paired Data' button to get your formatted Excel file.
""")

# --- File Uploaders ---
st.header("1. Upload Your Spreadsheets")

file1 = st.file_uploader("Upload Items Spreadsheet (.xlsx)", type=["xlsx"])
file2 = st.file_uploader("Upload Region Patches Spreadsheet (.xlsx)", type=["xlsx"])

# --- Main Logic: Process files only when both are uploaded ---
if file1 and file2:
    try:
        st.header("2. Processing and Generating Output")

        # --- Data Loading ---
        df_items = pd.read_excel(file1, dtype={'VBU': str}) # Read VBU as string to preserve formatting
        df_regions_wide = pd.read_excel(file2)

        # --- Data Validation ---
        required_item_cols = {'Item #', 'Model', 'VBU', 'Region'}
        if not required_item_cols.issubset(df_items.columns):
            st.error(f"Error: The Items spreadsheet must contain the following columns: {', '.join(required_item_cols)}")
            st.stop()

        if df_regions_wide.empty:
            st.error("Error: The Region Patches spreadsheet cannot be empty.")
            st.stop()

        # --- Data Transformation for Regions Spreadsheet ---
        st.write("Transforming Region Patches data...")
        
        # Melt the dataframe to unpivot it from a wide to a long format.
        # This creates a mapping from each Region to its list of patch codes.
        df_regions_long = pd.melt(df_regions_wide, var_name='Region', value_name='Region patch code')
        
        # Drop rows where the patch code is NaN/blank, which occurs with uneven columns.
        df_regions_long.dropna(subset=['Region patch code'], inplace=True)
        
        st.write("Preview of transformed region data (first 5 rows):")
        st.dataframe(df_regions_long.head())

        # --- Region-Based Pairing (Merge) ---
        # This is the key step. We merge the two tables on the 'Region' column.
        # This pairs each item only with the patch codes that match its region.
        st.write("Pairing items with patch codes based on matching regions...")
        df_merged = pd.merge(df_items, df_regions_long, on='Region', how='left')

        # Drop rows that didn't find a matching region in the patch file
        df_merged.dropna(subset=['Region patch code'], inplace=True)

        # --- Structure the Output ---
        # Reorder and select the desired columns for the final output
        output_df = df_merged[['VBU', 'Item #', 'Model', 'Region patch code', 'Region']]
        
        st.success("Data has been successfully paired and structured!")
        st.write("Preview of the first 20 paired rows:")
        st.dataframe(output_df.head(20))

        # --- Excel Export Preparation ---
        st.header("3. Download Your File")

        output_buffer = io.BytesIO()
        with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
            # Group by the 'Region' column (East, West, etc.) to create a tab for each
            for region_name, group_df in output_df.groupby('Region'):
                
                # For the output within each tab, we only need the specified four columns
                final_tab_df = group_df[['VBU', 'Item #', 'Model', 'Region patch code']]

                # Write the dataframe to a new sheet named after the region
                sheet_name = str(region_name)[:31] # Excel sheet names have a 31-char limit
                final_tab_df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        output_buffer.seek(0)

        # --- Download Button ---
        st.download_button(
            label="⬇️ Download Paired Data as Excel File",
            data=output_buffer,
            file_name="region_paired_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"An error occurred: {e}")
        st.warning("Please ensure your files are in the correct format and that region names match between files.")

else:
    st.info("Please upload both spreadsheets to begin processing.")

# --- Footer ---
st.markdown("---")
st.write("Developed by your friendly Web Developer")
