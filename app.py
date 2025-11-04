import streamlit as st
import pandas as pd
import io
import zipfile

# Page configuration
st.set_page_config(page_title="Excel Tier Splitter", layout="centered")
st.title("Excel Tier Splitter ✂️")
st.markdown("""
This tool splits your master workbook by price tier.  
It handles column name inconsistencies (extra spaces, casing) and ensures the final output has headers:  
**TARIFF NAME**, **PRICE**, **SNOMED CODE**, **SNOMED DESCRIPTION EN**.
""")

# Step 1: Upload file
uploaded_file = st.file_uploader("Upload your Master Price List Excel file (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        # Step 2: Read all sheets
        with st.spinner('Reading file...'):
            sheets = pd.read_excel(uploaded_file, sheet_name=None, dtype={'SNOMED CODE': str})

        st.info(f"Loaded **{len(sheets)}** sheets. Starting split...")

        # Define tiers
        tiers = ["Tier 0", "Tier 1", "Tier 2", "Tier 3", "Tier 4"]
        zip_buffer = io.BytesIO()
        zip_name = "RH_Tiers_Workbooks.zip"

        # Step 3: Process tiers and create workbooks
        with st.spinner('Processing tiers and creating workbooks...'):
            with zipfile.ZipFile(zip_buffer, 'a', zipfile.ZIP_DEFLATED, False) as zip_file:
                for tier in tiers:
                    excel_buffer = io.BytesIO()
                    sheets_processed = 0

                    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                        for sheet_name, df in sheets.items():
                            # --- Clean column names ---
                            df.columns = [col.strip().upper().replace(" ", "") for col in df.columns]

                            # Match tier column
                            tier_col_matches = [c for c in df.columns if c.startswith(tier.upper().replace(" ", ""))]
                            if not tier_col_matches:
                                continue
                            actual_tier_col = tier_col_matches[0]

                            # Prepare columns to select
                            cols_to_select = ["S/N", "SNOMEDCODE", actual_tier_col]

                            # Handle TARIFF NAME (LINE ITEMS)
                            if "LINEITEMS" in df.columns:
                                cols_to_select.append("LINEITEMS")
                            elif "LINEITEM" in df.columns:  # in case it's singular
                                cols_to_select.append("LINEITEM")

                            # Handle SNOMED DESCRIPTION EN
                            if "SNOMEDDESCRIPTIONEN" in df.columns:
                                cols_to_select.append("SNOMEDDESCRIPTIONEN")

                            # Subset dataframe
                            subset = df[[col for col in cols_to_select if col in df.columns]].copy()

                            # Rename columns to final structure
                            rename_dict = {}
                            rename_dict[actual_tier_col] = "PRICE"
                            if "LINEITEMS" in subset.columns:
                                rename_dict["LINEITEMS"] = "TARIFF NAME"
                            elif "LINEITEM" in subset.columns:
                                rename_dict["LINEITEM"] = "TARIFF NAME"
                            if "SNOMEDDESCRIPTIONEN" in subset.columns:
                                rename_dict["SNOMEDDESCRIPTIONEN"] = "SNOMED DESCRIPTION EN"
                            if "SNOMEDCODE" in subset.columns:
                                rename_dict["SNOMEDCODE"] = "SNOMED CODE"

                            subset.rename(columns=rename_dict, inplace=True)

                            # Reorder columns
                            final_order = ["S/N", "TARIFF NAME", "PRICE", "SNOMED CODE", "SNOMED DESCRIPTION EN"]
                            subset = subset[[col for col in final_order if col in subset.columns]]

                            # Write to Excel
                            safe_sheet_name = sheet_name[:31]
                            subset.to_excel(writer, sheet_name=safe_sheet_name, index=False)
                            sheets_processed += 1

                        if sheets_processed > 0:
                            writer.close()
                            excel_buffer.seek(0)
                            tier_filename = f"{tier}_Price_List.xlsx"
                            zip_file.writestr(tier_filename, excel_buffer.getvalue())

        # Step 4: Provide download
        zip_buffer.seek(0)
        st.success("✅ All workbooks processed and zipped successfully!")
        st.download_button(
            label="⬇️ Download Split Workbooks (.zip)",
            data=zip_buffer.getvalue(),
            file_name=zip_name,
            mime="application/zip"
        )

    except Exception as e:
        st.error(f"An error occurred: {e}")
