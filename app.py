import streamlit as st
import pandas as pd
import shutil
import os
import tempfile

st.title("Excel Tier Splitter")

# Step 1: Upload file
uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])
if uploaded_file is not None:
    sheets = pd.read_excel(uploaded_file, sheet_name=None)
    
    # Step 2: Define tiers and columns to keep
    tiers = ["Tier 0", "Tier 1", "Tier 2", "Tier 3", "Tier 4"]
    cols_to_keep = ["S/N", "LINE ITEMS", "SNOMED CODE", "SNOMED DESCRIPTION EN"]
    
    output_files = []

    # Use a temporary folder to save files
    with tempfile.TemporaryDirectory() as tmpdirname:
        # Step 3: Create workbooks per tier
        for tier in tiers:
            out_path = os.path.join(tmpdirname, f"{tier}.xlsx")
            with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
                for sheet_name, df in sheets.items():
                    df.columns = df.columns.str.strip()
                    if tier not in df.columns:
                        continue
                    available_cols = [col for col in cols_to_keep + [tier] if col in df.columns]
                    subset = df[available_cols].copy()
                    subset.to_excel(writer, sheet_name=sheet_name[:31], index=False)
            output_files.append(out_path)
        
        # Step 4: Zip only the Excel files
        zip_path = os.path.join(tmpdirname, "RH_Tiers_Workbooks.zip")
        with shutil.make_archive(zip_path.replace('.zip',''), 'zip', tmpdirname):
            pass  # Already created by make_archive
        
        # Step 5: Provide download link
        with open(zip_path, "rb") as f:
            st.download_button(
                label="Download Split Workbooks",
                data=f,
                file_name="RH_Tiers_Workbooks.zip",
                mime="application/zip"
            )
    
    st.success("✅ All workbooks processed and zipped successfully!")
