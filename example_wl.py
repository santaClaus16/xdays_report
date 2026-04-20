import streamlit as st
import pandas as pd
import msoffcrypto
import re
from io import BytesIO
from datetime import datetime
import pyodbc
import os
from dotenv import load_dotenv

load_dotenv()

from utils import (
    load_mapping_file,
    apply_header_mapping,
    save_to_folder,
    clean_account_number_column,
    format_date_columns,
    force_columns_to_str,
)

MAPPING_PATH_WL = r"C:\Users\SPM\Desktop\eod_report\mapping\mapping_file_wl.xlsx"


def decrypt_file(file_bytes, password):
    """Helper function to decrypt msoffcrypto file"""
    try:
        decrypted = BytesIO()
        office_file = msoffcrypto.OfficeFile(BytesIO(file_bytes))
        office_file.load_key(password=password)
        office_file.decrypt(decrypted)
        decrypted.seek(0)
        return decrypted, None
    except msoffcrypto.exceptions.DecryptionError:
        return None, "Wrong password"
    except Exception as e:
        return None, str(e)


# ====================== DEFAULT SHEETS ======================
sheet_default = ["ActiveQry", "Active_updt", "Pulled Out", "!!DNC!!"]


# ====================== CLEAN COLUMN NAMES ======================
def clean_column_names(df: pd.DataFrame) -> pd.DataFrame:
    """Fix mixed-type column names warning"""
    new_cols = []
    for col in df.columns:
        col_str = str(col).strip()
        col_str = re.sub(r'[\r\n\t]', ' ', col_str)
        col_str = re.sub(r'\s+', ' ', col_str)
        new_cols.append(col_str)
    
    df = df.copy()
    df.columns = new_cols
    return df


# ====================== CACHED TEMPLATE LOADER ======================
@st.cache_data(show_spinner="Loading template sheets...", ttl=600)
def load_template_sheets(template_file):
    """Cached function to load only default sheets from template"""
    if template_file is None:
        return {sheet: None for sheet in sheet_default}
    
    SHEET_DICT = {sheet: None for sheet in sheet_default}
    
    try:
        with st.spinner("Reading template sheets..."):
            excel_file = pd.ExcelFile(template_file)
            available_sheets = set(excel_file.sheet_names)

            for sheet_name in sheet_default:
                if sheet_name in available_sheets:
                    header_row = 1 if sheet_name in ["Active_updt", "Pulled Out"] else 0
                    
                    df = pd.read_excel(
                        template_file,
                        sheet_name=sheet_name,
                        header=header_row,
                        dtype=str,
                        engine="openpyxl"
                    )
                    
                    df = clean_column_names(df)
                    SHEET_DICT[sheet_name] = df
                    
                    # st.success(f"✅ Loaded sheet: **{sheet_name}** ({len(df):,} rows)")
                else:
                    st.warning(f"⚠️ Sheet '**{sheet_name}**' not found in template.")
        
        return SHEET_DICT
    
    except Exception as e:
        st.error(f"❌ Failed to load template: {e}")
        return {sheet: None for sheet in sheet_default}


# ====================== MAIN APP ======================
st.title("🔄 Worklist Processor")

# 1. Template Uploader
upload_file_template = st.file_uploader(
    "Upload your worklist template here",
    type=["xlsx"],
    key="template_uploader"
)

# Load template with caching (this prevents reloading when uploading worklists)
SHEET_DICT = load_template_sheets(upload_file_template)

if upload_file_template is None:
    st.info("👆 Please upload the **worklist template** first.")
    st.stop()

# Summary
loaded_count = sum(1 for v in SHEET_DICT.values() if v is not None)
st.success(f"✅ Template loaded successfully — {loaded_count}/{len(sheet_default)} default sheets ready")

# Safe DataFrame extraction
df_active       = SHEET_DICT.get("ActiveQry")
df_active_updt  = SHEET_DICT.get("Active_updt")
df_pulled_out   = SHEET_DICT.get("Pulled Out")
df_dnc          = SHEET_DICT.get("!!DNC!!")

if df_active is None:
    st.error("❌ Critical: 'ActiveQry' sheet is missing from template!")

# ====================== LOAD MAPPING RULES (once) ======================
# Assuming these are needed for apply_header_mapping
try:
    rules = load_mapping_file(MAPPING_PATH_WL)             # Load once
    standard_order = rules.get("standard_order", []) if isinstance(rules, dict) else []
except Exception as e:
    st.error(f"Failed to load mapping rules: {e}")
    rules = {"mapping": {}}
    standard_order = []

# ====================== WORKLIST FILES UPLOADER ======================
uploaded_files = st.file_uploader(
    "Upload Worklist Files (c9, c14, c15, etc.)",
    type=["xlsx"],
    accept_multiple_files=True,
    key="worklist_uploader"
)

if uploaded_files:
    results = []
    all_dfs = []
    
    progress_bar = st.progress(0)
    status_text = st.empty()

    for i, uploaded_file in enumerate(uploaded_files):
        filename = uploaded_file.name
        status_text.text(f"Processing: {filename} ({i+1}/{len(uploaded_files)})")

        match = re.search(r'[cC](\d+)', filename)
        if not match:
            st.error(f"❌ Could not find cycle number in: **{filename}**")
            continue

        c_number = match.group(1)
        auto_password = f"CYCLE_{c_number}*"

        with st.container(border=True):
            with st.expander(f"📁 Processing: {filename}", expanded=True):
                st.subheader(f"📄 {filename}")
                status = st.container()

                file_bytes = uploaded_file.read()
                decrypted_io = None

                # Auto password attempt
                with status:
                    st.info(f"🔄 Trying automatic password: `{auto_password}`")

                decrypted_io, error = decrypt_file(file_bytes, auto_password)

                if decrypted_io is None:
                    if "wrong password" in str(error).lower():
                        st.toast("❌ Wrong Password!", icon="🚫")
                        st.error(f"❌ Wrong password for **{filename}**")
                    else:
                        st.error(f"Decryption failed: {error}")

                    # Manual password input
                    st.markdown("### 🔓 Enter Correct Password Manually")
                    manual_password = st.text_input(
                        "Enter password:",
                        value="",
                        type="password",
                        key=f"manual_pass_{i}_{filename}"
                    )

                    col1, col2, col3 = st.columns([2, 1, 1])
                    retry_btn = col1.button("🔓 Decrypt with this Password", key=f"retry_{i}")
                    skip_btn = col2.button("⏭️ Skip this file", key=f"skip_{i}")

                    if skip_btn:
                        st.info(f"⏭️ Skipped: {filename}")
                        continue

                    if retry_btn:
                        if not manual_password.strip():
                            st.error("Please enter a password")
                            continue
                        
                        decrypted_io, error = decrypt_file(file_bytes, manual_password)
                        if decrypted_io:
                            st.success(f"✅ Decrypted **{filename}** successfully with manual password!")
                        else:
                            st.error("❌ Password still incorrect. Please try again.")
                            continue
                else:
                    st.success(f"✅ Auto password worked for **{filename}**!")

                # Process decrypted file
                if decrypted_io is None:
                    continue

                try:
                    df = pd.read_excel(decrypted_io, dtype=str, engine="openpyxl")
                    df = clean_column_names(df)   # Clean column names here too

                    with status:
                        st.success(f"✅ Loaded — {df.shape[0]:,} rows × {df.shape[1]} columns")

                    # Processing pipeline
                    WL_STR_COLS = ["CUST_ID", "OFFICE_PH", "HOME_PH", "MOBILE_NO", "OB", "BOS", "AOD", "MAD", "PDA", "LPA", "PTP_AMT"]
                    df = force_columns_to_str(df, WL_STR_COLS)

                    WL_DATE_COLS = ["LAST_PAYMENT_DATE", "PTP_DATE", "BIRTHDATE", "LAST_DUE_DATE", "D_CUST_OPN", "Birthdate"]
                    df = format_date_columns(df, date_cols=WL_DATE_COLS)

                    df = clean_account_number_column(df, show_success=(i == 0), st=st)

                    # Apply header mapping
                    df, map_logs = apply_header_mapping(df, rules.get("mapping", {}), standard_order)

                    # Collection Cycle
                    cycle_value = c_number
                    if "COLLECTION_CYCLE" in df.columns:
                        df["COLLECTION_CYCLE"] = (
                            df["COLLECTION_CYCLE"]
                            .fillna(cycle_value)
                            .astype(str)
                            .str.upper()
                            .str.replace(r"\s+", "_", regex=True)
                        )
                    else:
                        df["COLLECTION_CYCLE"] = cycle_value
                        st.warning(f"Created missing column COLLECTION_CYCLE = {cycle_value}")

                    with status:
                        st.success(f"✅ Mapping & processing completed for {filename}")

                    with st.expander("🔍 Preview first 10 rows", expanded=False):
                        st.dataframe(df.head(10), use_container_width=True, hide_index=True)

                    results.append({
                        "filename": filename,
                        "dataframe": df,
                        "c_number": c_number,
                    })
                    all_dfs.append(df)

                except Exception as e:
                    st.error(f"❌ Error processing {filename}: {str(e)}")

        # Update progress
        progress_bar.progress((i + 1) / len(uploaded_files))

    st.success(f"🎉 Finished processing all {len(uploaded_files)} worklist files!")
    
    # Optional: Show final summary or download button here later