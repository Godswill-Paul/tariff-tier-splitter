import streamlit as st
import pandas as pd
import io
import zipfile

st.set_page_config(layout="centered")
st.title("Excel Tier Splitter ✂️")
st.markdown("Upload your Excel workbook. The code now includes robust column cleaning to handle extra spaces and casing inconsistencies.")

# Step 1: Upload file
uploaded_file = st.file_uploader("Upload your Master Price List Excel file (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        # Step 2: Read all sheets from the uploaded Excel file
        with st.spinner('Reading file and preparing data...'):
            # Load SNOMED CODE as string to prevent scientific notation
            sheets = pd.read_excel(uploaded_file, sheet_name=None, dtype={'SNOMED CODE': str})
            
        st.info(f"Loaded **{len(sheets)}** sheets successfully. Starting split...")
        
        # Define tiers and the target columns (what they *should* look like after cleaning)
        tiers = ["Tier 0", "Tier 1", "Tier 2", "Tier 3", "Tier 4"]
        # Standardized column names in uppercase without spaces for matching
        standard_cols = ["S/N", "LINE ITEMS", "SNOMED CODE", "DESCRIPTION EN"] 
        
        zip_buffer = io.BytesIO()
        zip_name = "RH_Tiers_Workbooks.zip"

        # Step 3: Create a workbook per tier in memory
        with st.spinner('Processing tiers and creating workbooks...'):
            with zipfile.ZipFile(zip_buffer, 'a', zipfile.ZIP_DEFLATED, False) as zip_file:
                for tier in tiers:
                    excel_buffer = io.BytesIO()
                    sheets_processed = 0

                    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                        for sheet_name, df in sheets.items():
                            
                            # --- ROBUST COLUMN CLEANING ---
                            # 1. Create a mapping of original column names to cleaned column names (UPPERCASE, stripped)
                            column_map = {col: col.strip().upper() for col in df.columns}
                            
                            # 2. Apply the mapping to the DataFrame
                            df.rename(columns=column_map, inplace=True)
                            
                            # 3. Define the cleaned tier column name
                            cleaned_tier_col = tier.upper()
                            
                            # ----------------------------

                            # Skip if the tier column doesn't exist on the sheet after cleaning
                            if cleaned_tier_col not in df.columns:
                                continue

                            # 4. Identify available columns (using the standardized list)
                            # We combine the standard columns and the single tier column we want
                            all_required_cols = standard_cols + [cleaned_tier_col]
                            available_cols = [col for col in all_required_cols if col in df.columns]

                            if available_cols:
                                subset = df[available_cols].copy()
                                
                                # 5. Rename the tier column to 'PRICE' for the final output
                                subset.rename(columns={cleaned_tier_col: 'PRICE'}, inplace=True)
                                
                                # 6. Write to the in-memory writer
                                safe_sheet_name = sheet_name[:31] 
                                subset.to_excel(writer, sheet_name=safe_sheet_name, index=False)
                                sheets_processed += 1
                        
                        if sheets_processed > 0:
                            writer.close() 
                            excel_buffer.seek(0) 
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