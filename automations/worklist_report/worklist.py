import streamlit as st
import pandas as pd
import numpy as np
import msoffcrypto
import re
from io import BytesIO
from pathlib import Path
import os
from datetime import datetime

from utils import (
    load_mapping_file,
    apply_header_mapping,
    save_to_folder,
    clean_account_number_column,
    format_date_columns,
)

# ─────────────────────────────────────────────────────────────────────────────
# Constants
# ─────────────────────────────────────────────────────────────────────────────
MAPPING_PATH  = r"C:\Users\SOLIZA\Documents\secret\mapping\mapping_file_wl.xlsx"
OUTPUT_FOLDER = r"C:\Users\SOLIZA\Documents\secret\clean_wl"

# ─────────────────────────────────────────────────────────────────────────────
# Main App
# ─────────────────────────────────────────────────────────────────────────────
def worklist_app():
    st.set_page_config(page_title="Worklist Decrypter + Mapper", layout="wide")
    st.title("🔓 Worklist Decrypter + Advanced Column Mapper")
    st.markdown("Decrypts `c9`, `c14`, `c15`... files + Applies mapping from **mapping_file_wl.xlsx**")

    # ── Load mapping ──────────────────────────────────────────────────────────
    rules, load_logs, standard_order = load_mapping_file(MAPPING_PATH)

    if any("❌" in log for log in load_logs):
        st.error("Mapping file issue")
        for log in load_logs:
            st.write(log)
        st.stop()
    else:
        st.sidebar.success("✅ Mapping loaded successfully")
        with st.sidebar.expander("Mapping Preview"):
            for log in load_logs[:20]:
                st.write(log)

    # ── File upload ───────────────────────────────────────────────────────────
    uploaded_files = st.file_uploader(
        "Upload Worklist Files (c9 Worklist..., c14 Worklist..., etc.)",
        type=["xlsx"],
        accept_multiple_files=True
    )

    if not uploaded_files:
        st.info("👆 Upload your password-protected worklist files to begin.")
        st.caption("Decryption: CYCLE_9* • Mapping from mapping_file_wl.xlsx (Row 1 = Standard, lower rows = Aliases)")
        return

    # ── Process each file ─────────────────────────────────────────────────────
    results = []

    for i, uploaded_file in enumerate(uploaded_files):
        filename = uploaded_file.name

        # Extract cycle number from filename for password
        match = re.search(r'[cC](\d+)', filename)
        if not match:
            st.error(f"❌ Could not find 'c' number in: **{filename}**")
            continue

        c_number = match.group(1)
        password = f"CYCLE_{c_number}*"

        st.info(f"🔄 Decrypting **{filename}** → Password: `{password}`")

        try:
            # ── Decrypt ───────────────────────────────────────────────────────
            file_bytes  = uploaded_file.read()
            decrypted   = BytesIO()
            office_file = msoffcrypto.OfficeFile(BytesIO(file_bytes))
            office_file.load_key(password=password)
            office_file.decrypt(decrypted)
            decrypted.seek(0)

            df = pd.read_excel(decrypted, engine="openpyxl")
            st.success(f"✅ Decrypted: {filename} ({df.shape[0]:,} rows × {df.shape[1]} columns)")

            WL_DATE_COLS = [
                "LAST_PAYMENT_DATE",
                "PTP DATE",
                "BIRTHDATE",
                "LAST DUE DATE",
                "D_CUST_OPN",
            ]
            # ── Format dates ──────────────────────────────────────────────────
            df = format_date_columns(df, date_cols=WL_DATE_COLS)

            # ── Clean account numbers (show message only for first file) ──────
            df = clean_account_number_column(df, show_success=(i == 0), st=st)

            # ── Apply mapping ─────────────────────────────────────────────────
            df, map_logs = apply_header_mapping(df, rules["mapping"], standard_order)
            st.success(f"✅ Column mapping applied → Final columns: {len(df.columns)} (strict order)")

            with st.expander(f"Preview after mapping — {filename}"):
                st.dataframe(df.head(10), use_container_width=True)
                for log in map_logs:
                    if   "✅" in log: st.success(log)
                    elif "➕" in log: st.warning(log)
                    elif "⚠️" in log: st.warning(log)
                    else:             st.info(log)

            results.append({
                "filename": filename,
                "dataframe": df,
                "c_number": c_number,
            })

        except Exception as e:
            st.error(f"❌ Failed to process {filename}: {e}")

    # ── Download section ──────────────────────────────────────────────────────
    if not results:
        return

    st.subheader("📥 Download Processed Files")
    timestamp = datetime.now().strftime("%m%d_%H%M")

    for result in results:
        download_name = f"{timestamp}_MAPPED_c{result['c_number']}_{result['filename']}"

        output_bytes = BytesIO()
        with pd.ExcelWriter(output_bytes, engine="openpyxl") as writer:
            result["dataframe"].to_excel(writer, index=False, sheet_name="Data")
        output_bytes.seek(0)

        st.download_button(
            label=f"⬇️ Download {download_name}",
            data=output_bytes.getvalue(),
            file_name=download_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # ── Save all locally ──────────────────────────────────────────────────────
    if st.button("💾 Save All to Local Folder"):
        files_to_save = {
            f"CYCLE{result['c_number']}_{timestamp}--{result['filename']}": result["dataframe"]
            for result in results
        }

        folder_path = save_to_folder(
            output_dir  = OUTPUT_FOLDER,
            folder_name = "WORKLIST",
            files       = files_to_save,
            timestamp   = timestamp,
        )
        st.success(f"✅ All files saved to folder: `{folder_path}`")


# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    worklist_app()