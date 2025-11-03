import streamlit as st
import pandas as pd
import shutil
import os

st.title("Excel Tier Splitter")

# Step 1: Upload file
uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])
if uploaded_file is not None:
    sheets = pd.read_excel(uploaded_file, sheet_name=None)
    
    # Step 2: Define tiers and columns to keep
    tiers = ["Tier 0", "Tier 1", "Tier 2", "Tier 3", "Tier 4"]
    cols_to_keep = ["S/N", "LINE ITEMS", "SNOMED CODE", "SNOMED DESCRIPTION EN"]
    
    output_files = []
    
    # Step 3: Create workbooks per tier
    for tier in tiers:
        out_path = f"{tier}.xlsx"
        with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
            for sheet_name, df in sheets.items():
                df.columns = df.columns.str.strip()
                if tier not in df.columns:
                    continue
                available_cols = [col for col in cols_to_keep + [tier] if col in df.columns]
                subset = df[available_cols].copy()
                subset.to_excel(writer, sheet_name=sheet_name[:31], index=False)
        output_files.append(out_path)
    
    # Step 4: Zip all tier workbooks
    zip_name = "RH_Tiers_Workbooks"
    shutil.make_archive(zip_name, "zip", ".")
    
    # Step 5: Show success message first
    st.success("✅ All workbooks processed and zipped successfully!")
    
    # Step 6: Provide download button
    with open(f"{zip_name}.zip", "rb") as f:
        st.download_button(
            label="Download Split Workbooks",
            data=f,
            file_name=f"{zip_name}.zip",
            mime="application/zip"
        )
