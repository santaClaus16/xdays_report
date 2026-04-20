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

st.set_page_config(page_title="Worklist Decrypter + Mapper", layout="wide")


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


def worklist_app():
    # ================== Database Connection ==================
    try:
        conn_str = (
            f"DRIVER={{{os.getenv('DB_DRIVER')}}};"
            f"SERVER={os.getenv('DB_HOST')};"
            f"DATABASE={os.getenv('DB_NAME')};"
            f"UID={os.getenv('DB_USER')};"
            f"PWD={os.getenv('DB_PASS')};"
        )
        conn = pyodbc.connect(conn_str)
        st.success("✅ Successfully connected to the database!")
    except Exception as e:
        st.error(f"❌ Failed to connect to the database: {e}")
        st.stop()

    # ================== Constants & Mapping ==================
    MAPPING_PATH_WL = r"C:\Users\SPM\Desktop\eod_report\mapping\mapping_file_wl.xlsx"
    OUTPUT_FOLDER = r"C:\Users\SPM\Desktop\eod_report\clean_wl"

    rules, load_logs, standard_order = load_mapping_file(MAPPING_PATH_WL)

    if any("❌" in log for log in load_logs):
        st.error("❌ Mapping file issue")
        st.dataframe(pd.DataFrame(load_logs, columns=["Mapping Load Logs"]), use_container_width=True)
        st.stop()
    else:
        st.sidebar.success("✅ Mapping loaded successfully")
    
    sheet_default = ["ActiveQry", "Active_updt", "Pulled Out","!!DNC!!"]
        
    with st.sidebar:
        # Dropdown for sheet selection
        selected_sheet = st.selectbox(
            "Select a sheet",
            sheet_default
        )
        
        selected_header = st.selectbox(
            "Select a sheet",
            [1,2,3]
        )
        
    upload_file_template = st.file_uploader(
        "Upload your worklist template here",
        type=["xlsx"],
        help="You can upload multiple files at once. The system will attempt to auto-decrypt using the cycle-based password format."
    )
    
    if upload_file_template is not None:
        # Load Excel file
        excel_file = pd.ExcelFile(upload_file_template)
  

        # Read selected sheet (header = row 2)
        df = pd.read_excel(
            excel_file,
            sheet_name=selected_sheet,
            header=selected_header
        )
        with st.expander(f"🔍 Preview of '{selected_sheet}' with header row {selected_header}", expanded=True):
            st.dataframe(df)

        # ================== File Upload ==================
        uploaded_files = st.file_uploader(
            "Upload Worklist Files (c9, c14, c15, etc.)",
            type=["xlsx"],
            accept_multiple_files=True
        )

        if not uploaded_files:
            st.info("👆 Upload your password-protected worklist files to begin.")
            return

        results = []
        all_dfs = []

        for i, uploaded_file in enumerate(uploaded_files):
            filename = uploaded_file.name

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

                    # ====================== Try Auto Password ======================
                    with status:
                        st.info(f"🔄 Trying automatic password: `{auto_password}`")

                    decrypted_io, error = decrypt_file(file_bytes, auto_password)

                    if decrypted_io is None:
                        if "Wrong password" in error.lower():
                            st.toast("❌ **Wrong Password!**", icon="🚫")
                            st.error(f"❌ Wrong password for **{filename}**")
                            st.warning(f"Automatic password `{auto_password}` did not work.")
                        else:
                            st.toast("⚠️ Decryption Error", icon="⚠️")
                            st.error(f"Decryption failed: {error}")

                        # ====================== Manual Password Input ======================
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
                            st.toast("⏭️ File skipped", icon="⏭️")
                            st.info(f"Skipped: {filename}")
                            continue

                        if retry_btn:
                            if not manual_password.strip():
                                st.toast("❌ Please enter a password", icon="❗")
                                st.error("Please enter a password")
                            else:
                                st.toast("🔄 Trying manual password...", icon="🔄")
                                
                                decrypted_io, error = decrypt_file(file_bytes, manual_password)

                                if decrypted_io is None:
                                    if "Wrong password" in error.lower():
                                        st.toast("❌ Wrong Password - Try again", icon="🚫")
                                        st.error("❌ The password you entered is still incorrect.")
                                    else:
                                        st.toast("⚠️ Decryption Failed", icon="⚠️")
                                        st.error(f"Decryption error: {error}")
                                else:
                                    st.toast("✅ Decryption Successful!", icon="✅")
                                    st.success(f"✅ Successfully decrypted **{filename}** with manual password!")
                                    used_password = manual_password
                    else:
                        st.toast("✅ Auto password worked!", icon="✅")
                        st.success(f"✅ Decrypted successfully with auto password: `{auto_password}`")

                    # ====================== Process Decrypted File ======================
                    if decrypted_io is None:
                        continue

                    try:
                        df = pd.read_excel(decrypted_io, dtype=str, engine="openpyxl")

                        with status:
                            st.success(f"✅ File loaded — {df.shape[0]:,} rows × {df.shape[1]} columns")

                        # Processing steps (unchanged)
                        WL_STR_COLS = ["CUST_ID", "OFFICE_PH", "HOME_PH", "MOBILE_NO", "OB", "BOS", "AOD", "MAD", "PDA", "LPA", "PTP_AMT"]
                        df = force_columns_to_str(df, WL_STR_COLS)

                        WL_DATE_COLS = ["LAST_PAYMENT_DATE", "PTP_DATE", "BIRTHDATE", "LAST_DUE_DATE", "D_CUST_OPN", "Birthdate"]
                        df = format_date_columns(df, date_cols=WL_DATE_COLS)

                        df = clean_account_number_column(df, show_success=(i == 0), st=st)

                        df, map_logs = apply_header_mapping(df, rules["mapping"], standard_order)

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
                            st.warning(f"➕ Created missing column: COLLECTION_CYCLE = {cycle_value}")

                        with status:
                            st.success(f"✅ Mapping applied → Final columns: {len(df.columns)}")

                        with st.expander("🔍 Preview first 10 rows", expanded=False):
                            st.dataframe(df.head(10), use_container_width=True, hide_index=True)

                        results.append({
                            "filename": filename,
                            "dataframe": df,
                            "c_number": c_number,
                        })
                        all_dfs.append(df)

                    except Exception as e:
                        st.error(f"❌ Error processing file: {e}")

        # ====================== Save Section ======================
        if results:
            st.subheader("💾 Save Processed Files")
            timestamp = datetime.now().strftime("%m%d_%H%M")

            if st.button("💾 Save All Individual Files", type="primary"):
                files_to_save = {
                    f"CYCLE{result['c_number']}_{timestamp}--{result['filename']}": result["dataframe"]
                    for result in results
                }

                folder_path = save_to_folder(
                    output_dir=OUTPUT_FOLDER,
                    folder_name="WORKLIST",
                    files=files_to_save,
                    timestamp=timestamp,
                )
                st.success(f"✅ All files saved to: `{folder_path}`")
                st.toast("💾 Files saved successfully!", icon="✅")

if __name__ == "__main__":
    worklist_app()