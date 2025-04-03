import os
import re
import pandas as pd
from datetime import datetime, timedelta
import streamlit as st

# Function to normalize names
def normalize_name(name):
    """Normalize a name by converting to uppercase, removing extra spaces, and splitting into components"""
    if pd.isna(name) or name == "NaN":
        return {"first": "", "middle": "", "surname": ""}
    name = " ".join(name.split()).upper()
    parts = name.split()
    if len(parts) == 1:
        return {"first": parts[0], "middle": "", "surname": ""}
    elif len(parts) == 2:
        return {"first": parts[0], "middle": "", "surname": parts[1]}
    else:
        return {"first": parts[0], "middle": " ".join(parts[1:-1]), "surname": parts[-1]}

# Streamlit app
st.title("Name Matching and Processing App")

# --- Process Extracted Names ---
st.header("Step 1: Upload Text Files for Name Extraction")
multiple_txt_files = st.radio("Do you have multiple text files for name extraction?", ("Yes", "No"))

all_extracted_names = []
line_number = 0

if multiple_txt_files == "Yes":
    txt_files = st.file_uploader("Upload your text files", type=["txt"], accept_multiple_files=True, key="txt_multiple")
    if txt_files:
        for txt_file in txt_files:
            for line in txt_file.read().decode('utf-8').splitlines():
                line_number += 1
                matches = re.findall(r'NPR\s{0,4}([A-Z][A-Z.\s]+)', line)
                if matches:
                    for match in matches:
                        all_extracted_names.append(match.strip())
                else:
                    all_extracted_names.append("NaN")
else:
    txt_file = st.file_uploader("Upload your text file", type=["txt"], accept_multiple_files=False, key="txt_single")
    if txt_file:
        for line in txt_file.read().decode('utf-8').splitlines():
            line_number += 1
            matches = re.findall(r'NPR\s{0,4}([A-Z][A-Z.\s]+)', line)
            if matches:
                for match in matches:
                    all_extracted_names.append(match.strip())
            else:
                all_extracted_names.append("NaN")

if all_extracted_names:
    extracted_df = pd.DataFrame({'EXTRACTED_NAME': all_extracted_names})
    st.success(f"Extracted {len(extracted_df)} names from text file(s)")

# --- Process HCS Files ---
st.header("Step 2: Upload HCS Files (.xls)")
multiple_hcs_files = st.radio("Do you have multiple HCS files to process?", ("Yes", "No"))

all_hcs_data = []

if multiple_hcs_files == "Yes":
    hcs_files = st.file_uploader("Upload your HCS files (.xls)", type=["xls"], accept_multiple_files=True, key="hcs_multiple")
    if hcs_files:
        for hcs_file in hcs_files:
            try:
                hcs_content = hcs_file.read().decode('utf-8')
                hcs_list = pd.read_html(hcs_content)
                hcs = hcs_list[0]
                hcs.columns = hcs.iloc[0]
                hcs = hcs[1:].reset_index(drop=True)
                all_hcs_data.append(hcs)
                st.write(f"Successfully processed {hcs_file.name}")
            except Exception as e:
                st.error(f"Error processing {hcs_file.name}: {e}")
        if all_hcs_data:
            hcs_df = pd.concat(all_hcs_data, ignore_index=True)
else:
    hcs_file = st.file_uploader("Upload your HCS file (.xls)", type=["xls"], accept_multiple_files=False, key="hcs_single")
    if hcs_file:
        try:
            hcs_content = hcs_file.read().decode('utf-8')
            hcs_list = pd.read_html(hcs_content)
            hcs_df = hcs_list[0]
            hcs_df.columns = hcs_df.iloc[0]
            hcs_df = hcs_df[1:].reset_index(drop=True)
            st.write(f"Successfully processed {hcs_file.name}")
        except Exception as e:
            st.error(f"Error processing {hcs_file.name}: {e}")

if 'hcs_df' in locals():
    st.success(f"Loaded {len(hcs_df)} rows from HCS file(s)")

# --- Name Matching ---
if 'extracted_df' in locals() and 'hcs_df' in locals():
    st.header("Step 3: Processing and Matching")
    hcs_name_col = 'E_NAME'
    account_col = 'ACCOUNT'
    pan_col = 'PAN'
    card_code_col = 'CAR_CODE'
    expiry_date_col = 'EXPIRYDATE'

    # Normalize names
    hcs_df['norm'] = hcs_df[hcs_name_col].apply(normalize_name)
    extracted_df['norm'] = extracted_df['EXTRACTED_NAME'].apply(normalize_name)

    # Create matching keys
    hcs_df['match_key'] = hcs_df['norm'].apply(lambda x: f"{x['first']}|{x['middle']}|{x['surname']}")
    extracted_df['match_key'] = extracted_df['norm'].apply(lambda x: f"{x['first']}|{x['middle']}|{x['surname']}")

    # Merge with extracted_df as base
    matched_df = extracted_df.merge(
        hcs_df,
        how='left',
        on='match_key',
        suffixes=('_extracted', '_hcs')
    )
    matched_df['MATCHED_EXTRACTED_NAME'] = matched_df['EXTRACTED_NAME']
    matched_df = matched_df.dropna(subset=[hcs_name_col])

    # --- Deduplication and Splitting ---
    duplicate_mask = matched_df.duplicated(subset=['EXTRACTED_NAME', account_col], keep=False)
    duplicate_df = matched_df[duplicate_mask].drop(columns=['norm_extracted', 'norm_hcs', 'match_key'])
    unique_df = matched_df[~duplicate_mask].drop(columns=['norm_extracted', 'norm_hcs', 'match_key'])

    # --- Compare E_Name and EXTRACTED_NAME in unique_df ---
    def check_length_match(row):
        hcs_name = str(row[hcs_name_col]) if pd.notna(row[hcs_name_col]) else ""
        extracted_name = str(row['EXTRACTED_NAME']) if pd.notna(row['EXTRACTED_NAME']) else ""
        return len(hcs_name) == len(extracted_name)

    unique_df['length_match'] = unique_df.apply(check_length_match, axis=1)
    mismatched_df = unique_df[~unique_df['length_match']].drop(columns=['length_match'])
    matched_unique_df = unique_df[unique_df['length_match']].drop(columns=['length_match'])

    # --- Generate Excel Output for mismatched_df ---
    output_columns = ['ISSTYPE', 'CARD_NUMBER', 'CRDH_NAME', 'ATM_ACCT', 'ISS_DATE', 'EXPIR_DATE', 'CARD_ID']
    output_df = pd.DataFrame(index=mismatched_df.index, columns=output_columns)

    output_df['ISSTYPE'] = 'NEW'
    output_df['CARD_NUMBER'] = mismatched_df[pan_col]
    output_df['CRDH_NAME'] = mismatched_df['EXTRACTED_NAME']
    output_df['ATM_ACCT'] = mismatched_df[account_col]
    output_df['EXPIR_DATE'] = pd.to_datetime(mismatched_df[expiry_date_col]).dt.strftime('%Y-%m-%d')
    output_df['ISS_DATE'] = (pd.to_datetime(mismatched_df[expiry_date_col]) - timedelta(days=4*365)).dt.strftime('%Y-%m-%d')
    output_df['CARD_ID'] = mismatched_df.get(card_code_col, '')

    # --- Display Results ---
    st.subheader("Results")
    st.write(f"Original HCS rows: {len(hcs_df)}")
    st.write(f"Extracted names: {len(extracted_df)}")
    st.write(f"Matched rows before splitting: {len(matched_df)}")
    st.write(f"Rows with duplicate ACCOUNT and EXTRACTED_NAME: {len(duplicate_df)}")
    st.write(f"Unique or non-duplicate rows: {len(unique_df)}")
    st.write(f"Rows with mismatched lengths in unique_df: {len(mismatched_df)}")
    st.write(f"Rows with matched lengths in unique_df: {len(matched_unique_df)}")
    st.write("First few rows of Excel output DataFrame:")
    st.dataframe(output_df.head())

    # --- Download Buttons ---
    st.subheader("Download Results")
    output_dir = os.getcwd()

    duplicate_csv = duplicate_df.to_csv(index=False).encode('utf-8')
    mismatched_csv = mismatched_df.to_csv(index=False).encode('utf-8')
    matched_unique_csv = matched_unique_df.to_csv(index=False).encode('utf-8')
    excel_buffer = pd.ExcelWriter('mismatched_output.xlsx', engine='xlsxwriter')
    output_df.to_excel(excel_buffer, index=False)
    excel_buffer.close()
    with open('mismatched_output.xlsx', 'rb') as f:
        excel_data = f.read()

    st.download_button("Download Duplicate Matches (CSV)", duplicate_csv, "duplicate_matches.csv", "text/csv")
    st.download_button("Download Mismatched Lengths (CSV)", mismatched_csv, "mismatched_lengths.csv", "text/csv")
    st.download_button("Download Matched Unique (CSV)", matched_unique_csv, "matched_unique.csv", "text/csv")
    st.download_button("Download Mismatched Output (Excel)", excel_data, "mismatched_output.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    st.success("Processing complete! Download the files above.")
else:
    st.warning("Please upload both text and HCS (.xls) files to proceed.")