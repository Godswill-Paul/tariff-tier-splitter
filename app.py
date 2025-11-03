import streamlit as st
import pandas as pd
import io
import zipfile

st.set_page_config(layout="centered")
st.title("Excel Tier Splitter ✂️")
st.markdown("This tool splits your master workbook by price tier, ensuring robust column matching and correct final headers.")

# Step 1: Upload file
uploaded_file = st.file_uploader("Upload your Master Price List Excel file (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        # Step 2: Read all sheets from the uploaded Excel file
        with st.spinner('Reading file and preparing data...'):
            # Load SNOMED CODE as string to prevent scientific notation
            sheets = pd.read_excel(uploaded_file, sheet_name=None, dtype={'SNOMED CODE': str})
            
        st.info(f"Loaded **{len(sheets)}** sheets successfully. Starting split...")
        
        # Define tiers and the target columns for the final output
        tiers = ["Tier 0", "Tier 1", "Tier 2", "Tier 3", "Tier 4"]
        
        # --- CRITICAL UPDATE: Standardized column names for the FINAL output structure ---
        # The list includes the name we want to match from the input (e.g., LINE ITEMS) and 
        # the name we want in the output (e.g., TARIFF NAME).
        # We must use the standardized/cleaned names for matching: UPPERCASE and stripped.
        STANDARD_COLUMNS_MAPPING = {
            "LINE ITEMS": "TARIFF NAME", 
            "DESCRIPTION EN": "SNOMED DESCRIPTION EN"
        }
        
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
                            
                            # --- ROBUST COLUMN CLEANING AND MAPPING ---
                            # 1. Map original column names to their CLEANED (UPPERCASE, stripped) versions
                            column_map = {col: col.strip().upper() for col in df.columns}
                            
                            # 2. Apply the mapping to the DataFrame
                            df.rename(columns=column_map, inplace=True)
                            
                            # 3. Define the cleaned tier column name
                            cleaned_tier_col = tier.upper()
                            
                            # ----------------------------

                            # Skip if the tier column doesn't exist on the sheet after cleaning
                            if cleaned_tier_col not in df.columns:
                                continue

                            # 4. Define the list of columns to select from the cleaned DataFrame
                            # This list uses the cleaned (UPPERCASE) names for selection.
                            cols_to_select = ["S/N", "SNOMED CODE", cleaned_tier_col]
                            
                            # Add the "LINE ITEMS" column if it exists in the sheet (it will be renamed later)
                            if "LINE ITEMS" in df.columns:
                                cols_to_select.append("LINE ITEMS")
                            
                            # Add the "DESCRIPTION EN" column if it exists in the sheet (it will be renamed later)
                            if "DESCRIPTION EN" in df.columns:
                                cols_to_select.append("DESCRIPTION EN")

                            # 5. Select the subset
                            subset = df[cols_to_select].copy()
                            
                            # 6. CRITICAL FIX: Rename columns to the final desired structure (e.g., TARIFF NAME, PRICE)
                            # Rename the tier column to 'PRICE'
                            subset.rename(columns={cleaned_tier_col: 'PRICE'}, inplace=True)
                            
                            # Rename item and description columns to the requested final names
                            if "LINE ITEMS" in subset.columns:
                                subset.rename(columns={"LINE ITEMS": "TARIFF NAME"}, inplace=True)
                            if "DESCRIPTION EN" in subset.columns:
                                subset.rename(columns={"DESCRIPTION EN": "SNOMED DESCRIPTION EN"}, inplace=True)

                            # 7. Reorder columns to match the desired final sequence:
                            final_order = ["S/N", "TARIFF NAME", "PRICE", "SNOMED CODE", "SNOMED DESCRIPTION EN"]
                            
                            # Only keep columns that are actually present in the subset
                            subset = subset[[col for col in final_order if col in subset.columns]]


                            # 8. Write to the in-memory writer
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