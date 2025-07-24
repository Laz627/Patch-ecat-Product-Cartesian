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

# --- Function to create Excel templates in memory ---
def create_template_files():
    """Creates two in-memory Excel files for user download."""
    
    # --- Template 1: Items Spreadsheet ---
    items_data = {
        'Item #': [1541187, 5452877, 5263949, 5452878, 9876543],
        'Model': ['1000008671', '1000011510', '1000010242', '1000011511', '1000099999'],
        'Region': ['East', 'West', 'West', 'West', 'National'],
        'VBU': ['1257', '1257', '1257', '1257', '1300']
    }
    df_items_template = pd.DataFrame(items_data)
    
    # Convert items template to an in-memory Excel file
    items_buffer = io.BytesIO()
    with pd.ExcelWriter(items_buffer, engine='openpyxl') as writer:
        df_items_template.to_excel(writer, sheet_name='Items', index=False)
    items_buffer.seek(0)

    # --- Template 2: Region Patches Spreadsheet ---
    patches_data = {
        'East': ['AB', 'AC', 'AD', 'AE', 'AF'],
        'West': ['HI', 'HU', 'IJ', 'IK', None], # Using None for uneven columns
        'Alaska': ['AK', 'AL', None, None, None],
        'National': ['AB', 'AC', 'AK', 'AN', None]
    }
    df_patches_template = pd.DataFrame(patches_data)

    # Convert patches template to an in-memory Excel file
    patches_buffer = io.BytesIO()
    with pd.ExcelWriter(patches_buffer, engine='openpyxl') as writer:
        df_patches_template.to_excel(writer, sheet_name='Region_Patches', index=False)
    patches_buffer.seek(0)
    
    return items_buffer, patches_buffer

# --- Main Application ---
st.title("Automated Data Pairing and Export Tool")
st.markdown("""
This tool intelligently pairs items with their regional patch codes and exports the results.

**Instructions:**
1.  (Optional) Download the template files to see the required format.
2.  Upload your completed 'Items' and 'Region Patches' spreadsheets.
3.  The tool will match items to patch codes based on their assigned region.
4.  Click the 'Download Paired Data' button to get your formatted Excel file.
""")

# --- Template Download Section ---
items_template_buffer, patches_template_buffer = create_template_files()

with st.expander("⬇️ Click here to download template files"):
    st.markdown("Use these templates to ensure your data is in the correct format.")
    
    st.download_button(
        label="Download Items Template (.xlsx)",
        data=items_template_buffer,
        file_name="template_items.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    st.download_button(
        label="Download Region Patches Template (.xlsx)",
        data=patches_template_buffer,
        file_name="template_region_patches.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# --- File Uploaders ---
st.header("1. Upload Your Spreadsheets")

file1 = st.file_uploader("Upload Items Spreadsheet (.xlsx)", type=["xlsx"])
file2 = st.file_uploader("Upload Region Patches Spreadsheet (.xlsx)", type=["xlsx"])

# --- Main Logic: Process files only when both are uploaded ---
if file1 and file2:
    try:
        st.header("2. Processing and Generating Output")

        # --- Data Loading ---
        df_items = pd.read_excel(file1, dtype={'VBU': str})
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
        df_regions_long = pd.melt(df_regions_wide, var_name='Region', value_name='Region patch code')
        df_regions_long.dropna(subset=['Region patch code'], inplace=True)
        
        # --- Region-Based Pairing (Merge) ---
        st.write("Pairing items with patch codes based on matching regions...")
        df_merged = pd.merge(df_items, df_regions_long, on='Region', how='left')
        df_merged.dropna(subset=['Region patch code'], inplace=True)

        # --- Structure the Output ---
        output_df = df_merged[['VBU', 'Item #', 'Model', 'Region patch code', 'Region']]
        
        st.success("Data has been successfully paired and structured!")
        st.write("Preview of the first 20 paired rows:")
        st.dataframe(output_df.head(20))

        # --- Excel Export Preparation ---
        st.header("3. Download Your File")

        output_buffer = io.BytesIO()
        with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
            for region_name, group_df in output_df.groupby('Region'):
                final_tab_df = group_df[['VBU', 'Item #', 'Model', 'Region patch code']]
                sheet_name = str(region_name)[:31]
                final_tab_df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        output_buffer.seek(0)

        # --- Download Button for Processed File ---
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
