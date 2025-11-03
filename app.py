import streamlit as st
import pandas as pd
import io
import zipfile

st.set_page_config(layout="centered")
st.title("Excel Tier Splitter ✂️")
st.markdown("Upload your Excel workbook to split it into individual workbooks for each price tier (Tier 0 to Tier 4).")

# Step 1: Upload file
uploaded_file = st.file_uploader("Upload your Master Price List Excel file (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        # Step 2: Read all sheets from the uploaded Excel file
        with st.spinner('Reading file and preparing data...'):
            # Explicitly load SNOMED CODE as string to prevent scientific notation
            sheets = pd.read_excel(uploaded_file, sheet_name=None, dtype={'SNOMED CODE': str})
            
        st.info(f"Loaded **{len(sheets)}** sheets successfully. Starting split...")
        
        # Define tiers and columns to keep
        tiers = ["Tier 0", "Tier 1", "Tier 2", "Tier 3", "Tier 4"]
        # The column names will be standardized to this list in the output
        cols_to_keep = ["S/N", "LINE ITEMS", "SNOMED CODE", "DESCRIPTION EN"] 
        
        # Use an in-memory buffer for the final zip file
        zip_buffer = io.BytesIO()
        zip_name = "RH_Tiers_Workbooks.zip"

        # Step 3: Create a workbook per tier in memory
        with st.spinner('Processing tiers and creating workbooks...'):
            with zipfile.ZipFile(zip_buffer, 'a', zipfile.ZIP_DEFLATED, False) as zip_file:
                for tier in tiers:
                    # Use an in-memory buffer for each Excel file
                    excel_buffer = io.BytesIO()
                    sheets_processed = 0

                    # Create the Excel file in memory
                    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                        for sheet_name, df in sheets.items():
                            # Clean column names for reliable matching
                            df.columns = df.columns.str.strip()
                            
                            # Skip if the tier column doesn't exist on the sheet
                            if tier not in df.columns:
                                continue

                            # 1. Identify available columns to select
                            available_cols = [col for col in cols_to_keep + [tier] if col in df.columns]

                            # 2. Select subset
                            if available_cols:
                                subset = df[available_cols].copy()
                                
                                # 3. CRITICAL FIX: Rename the tier column to 'PRICE'
                                subset.rename(columns={tier: 'PRICE'}, inplace=True)
                                
                                # 4. Standardize all column names for the output
                                # The current set of columns: ["S/N", "LINE ITEMS", "SNOMED CODE", "DESCRIPTION EN", "PRICE"]
                                
                                # 5. Write to the in-memory writer
                                safe_sheet_name = sheet_name[:31] 
                                subset.to_excel(writer, sheet_name=safe_sheet_name, index=False)
                                sheets_processed += 1
                        
                        # Only proceed if data was actually written for this tier
                        if sheets_processed > 0:
                            # Save the in-memory Excel file to the in-memory zip
                            writer.close() # Ensure data is written to the buffer
                            excel_buffer.seek(0) # Rewind the buffer to the beginning
                            tier_filename = f"{tier}_Price_List.xlsx"
                            zip_file.writestr(tier_filename, excel_buffer.getvalue())

        # Step 4: Finalize and provide download button
        zip_buffer.seek(0)
        
        st.success("✅ All workbooks processed and zipped successfully!")
        
        st.download_button(
            label="⬇️ Download Split Workbooks (.zip)",
            data=zip_buffer.getvalue(),
            file_name=zip_name,
            mime="application/zip"
        )
        
    except Exception as e:
        st.error(f"An error occurred during processing: {e}")