import streamlit as st
import pandas as pd
import io
import zipfile

st.set_page_config(layout="centered")
st.title("Excel Tier Splitter ✂️")
st.markdown("This tool splits your master workbook by price tier. It uses robust logic to handle column name inconsistencies (like extra spaces or casing) and ensures the final output has the requested headers: **TARIFF NAME** and **PRICE**.")

# Step 1: Upload file
uploaded_file = st.file_uploader("Upload your Master Price List Excel file (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        # Step 2: Read all sheets from the uploaded Excel file
        with st.spinner('Reading file and preparing data...'):
            # Load SNOMED CODE as string to prevent scientific notation
            sheets = pd.read_excel(uploaded_file, sheet_name=None, dtype={'SNOMED CODE': str})
            
        st.info(f"Loaded **{len(sheets)}** sheets successfully. Starting split...")
        
        # Define tiers (base names) and the target columns for the final output
        tiers = ["Tier 0", "Tier 1", "Tier 2", "Tier 3", "Tier 4"]
        
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
                            
                            # 3. Define the base tier name for flexible matching
                            tier_clean = tier.upper().strip()
                            
                            # 4. Find the actual tier column (e.g., handles "TIER 4 ") using startswith
                            tier_col_matches = [c for c in df.columns if c.startswith(tier_clean)]
                            
                            if not tier_col_matches:
                                continue # Skip if no tier column is found
                            
                            # Use the first matched column as the actual price column
                            actual_tier_col = tier_col_matches[0] 
                            # ----------------------------

                            # 5. Define the list of core columns to select from the cleaned DataFrame
                            # This list uses the expected CLEANED (UPPERCASE) names for selection.
                            cols_to_select = ["S/N", "SNOMED CODE", actual_tier_col]
                            
                            # Columns we will select if they exist and then rename:
                            # LINE ITEMS -> TARIFF NAME
                            # SNOMED DESCRIPTION EN -> SNOMED DESCRIPTION EN (If it was "DESCRIPTION EN", it's covered by the cleaning and rename step)
                            
                            if "LINE ITEMS" in df.columns:
                                cols_to_select.append("LINE ITEMS")
                            if "SNOMED DESCRIPTION EN" in df.columns:
                                cols_to_select.append("SNOMED DESCRIPTION EN")


                            # 6. Select the subset and prepare for renaming
                            subset = df[[col for col in cols_to_select if col in df.columns]].copy()
                            
                            # 7. CRITICAL FIX: Rename columns to the final desired structure (e.g., TARIFF NAME, PRICE)
                            rename_dict = {}
                            
                            # Tier column to PRICE
                            rename_dict[actual_tier_col] = 'PRICE'
                            
                            # LINE ITEMS to TARIFF NAME
                            if "LINE ITEMS" in subset.columns:
                                rename_dict["LINE ITEMS"] = "TARIFF NAME"
                            
                            # SNOMED DESCRIPTION EN is used as the final name, but we ensure it's there
                            
                            subset.rename(columns=rename_dict, inplace=True)

                            # 8. Reorder columns to match the desired final sequence:
                            # S/N, TARIFF NAME, PRICE, SNOMED CODE, SNOMED DESCRIPTION EN
                            final_order = ["S/N", "TARIFF NAME", "PRICE", "SNOMED CODE", "SNOMED DESCRIPTION EN"]
                            
                            # Only keep columns that are actually present in the subset and in the correct order
                            subset = subset[[col for col in final_order if col in subset.columns]]


                            # 9. Write to the in-memory writer
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