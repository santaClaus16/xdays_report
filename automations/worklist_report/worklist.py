import streamlit as st
import pandas as pd
import msoffcrypto
import re
from io import BytesIO
from datetime import datetime

from utils import (
    load_mapping_file,
    apply_header_mapping,
    save_to_folder,
    clean_account_number_column,
    format_date_columns,
    force_columns_to_str,
)

# ─────────────────────────────────────────────────────────────────────────────
# Constants
# ─────────────────────────────────────────────────────────────────────────────
MAPPING_PATH_WL = r"C:\Users\SPM\Desktop\eod_report\mapping\mapping_file_wl.xlsx"
OUTPUT_FOLDER = r"C:\Users\SPM\Desktop\eod_report\clean_wl"

# ─────────────────────────────────────────────────────────────────────────────
# Main App
# ─────────────────────────────────────────────────────────────────────────────
def worklist_app():
    st.set_page_config(page_title="Worklist Decrypter + Mapper", layout="wide")
    st.title("🔓 Worklist Decrypter + Advanced Column Mapper")
    st.markdown("Decrypts `c9`, `c14`, `c15`... files + Applies mapping from **mapping_file_wl.xlsx**")

    # ── Load mapping ──────────────────────────────────────────────────────────
    rules, load_logs, standard_order = load_mapping_file(MAPPING_PATH_WL)

    if any("❌" in log for log in load_logs):
        st.error("Mapping file issue")
        log_df = pd.DataFrame(load_logs, columns=["Mapping Load Logs"])
        st.dataframe(log_df, height=300, use_container_width=True, hide_index=True)
        st.stop()
    else:
        st.sidebar.success("✅ Mapping loaded successfully")
        with st.sidebar.expander(f"Mapping Preview (**{len(load_logs)}**)", expanded=False):
            preview_df = pd.DataFrame(load_logs[:20], columns=["Mapping Load Logs"])
            st.dataframe(preview_df, height=250, use_container_width=True, hide_index=True)

    # ── File upload ───────────────────────────────────────────────────────────
    uploaded_files = st.file_uploader(
        "Upload Worklist Files (c9 Worklist..., c14 Worklist..., etc.)",
        type=["xlsx"],
        accept_multiple_files=True
    )

    if not uploaded_files:
        st.info("👆 Upload your password-protected worklist files to begin.")
        st.caption("Decryption: CYCLE_9* • Mapping from mapping_file_wl.xlsx")
        return

    # ── Process each file ─────────────────────────────────────────────────────
    results = []
    all_dfs = []   # For combining later

    for i, uploaded_file in enumerate(uploaded_files):
        filename = uploaded_file.name

        match = re.search(r'[cC](\d+)', filename)
        if not match:
            st.error(f"❌ Could not find cycle number in: **{filename}**")
            continue

        c_number = match.group(1)
        password = f"CYCLE_{c_number}*"

        with st.container(border=True):
            with st.expander(f"📁 Processing: {filename}", expanded=True):
                st.subheader(f"📄 {filename}")
                status_container = st.container()

                try:
                    with status_container:
                        st.info(f"🔄 Decrypting → Password: `{password}`")

                    # Decrypt
                    file_bytes = uploaded_file.read()
                    decrypted = BytesIO()
                    office_file = msoffcrypto.OfficeFile(BytesIO(file_bytes))
                    office_file.load_key(password=password)
                    office_file.decrypt(decrypted)
                    decrypted.seek(0)

                    df = pd.read_excel(decrypted, dtype=str, engine="openpyxl")

                    with status_container:
                        st.success(f"✅ Decrypted — {df.shape[0]:,} rows × {df.shape[1]} columns")

                    # Force string columns
                    WL_STR_COLS = [
                        "CUST_ID", "OFFICE_PH", "HOME_PH", "MOBILE_NO",
                        "OB", "BOS", "AOD", "MAD", "PDA", "LPA", "PTP_AMT"
                    ]
                    df = force_columns_to_str(df, WL_STR_COLS)

                    # Date formatting
                    WL_DATE_COLS = [
                        "LAST_PAYMENT_DATE", "PTP_DATE", "BIRTHDATE",
                        "LAST_DUE_DATE", "D_CUST_OPN", "Birthdate"
                    ]
                    df = format_date_columns(df, date_cols=WL_DATE_COLS)

                    # Clean account numbers
                    df = clean_account_number_column(df, show_success=(i == 0), st=st)

                    # Apply mapping
                    df, map_logs = apply_header_mapping(df, rules["mapping"], standard_order)

                    # Handle COLLECTION_CYCLE
                    cycle_value = f"C{c_number}"   # Changed to C9, C14 format (recommended)

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
                        with status_container:
                            st.warning(f"➕ Created missing column: COLLECTION_CYCLE = {cycle_value}")

                    with status_container:
                        st.success(f"✅ Mapping applied → Final columns: {len(df.columns)}")

                    # Preview
                    with st.expander("🔍 Preview first 10 rows", expanded=False):
                        st.dataframe(df.head(10), use_container_width=True, hide_index=True)

                    with st.expander("📋 Mapping Logs", expanded=False):
                        for log in map_logs:
                            if "✅" in log:
                                st.success(log)
                            elif "➕" in log or "⚠️" in log:
                                st.warning(log)
                            else:
                                st.info(log)

                    results.append({
                        "filename": filename,
                        "dataframe": df,
                        "c_number": c_number,
                    })

                    all_dfs.append(df)   # Collect for overall

                except Exception as e:
                    with status_container:
                        st.error(f"❌ Failed to process: {e}")

    # ── Overall Combined Data Section ─────────────────────────────────────────
    # if results and all_dfs:
    #     st.subheader("📊 Overall Combined Worklist")
        
    #     overall_df = pd.concat(all_dfs, ignore_index=True)
        
    #     col1, col2, col3 = st.columns(3)
    #     with col1:
    #         st.metric("Total Files Processed", len(results))
    #     with col2:
    #         st.metric("Total Rows", f"{len(overall_df):,}")
    #     with col3:
    #         st.metric("Total Columns", len(overall_df.columns))

    #     # Preview of combined data
    #     with st.expander("🔍 Preview of Combined Data (first 10 rows)", expanded=False):
    #         st.dataframe(overall_df.head(10), use_container_width=True, hide_index=True)

    #     # Download combined file
    #     output_bytes = BytesIO()
    #     with pd.ExcelWriter(output_bytes, engine="openpyxl") as writer:
    #         overall_df.to_excel(writer, index=False, sheet_name="Overall_Worklist")
    #     output_bytes.seek(0)

    #     st.download_button(
    #         label="⬇️ Download Combined Overall Worklist",
    #         data=output_bytes.getvalue(),
    #         file_name=f"OVERALL_WORKLIST_{datetime.now().strftime('%m%d_%H%M')}.xlsx",
    #         mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    #         type="primary"
    #     )

    # ── Individual Save Section ───────────────────────────────────────────────
    if not results:
        return

    st.subheader("💾 Save to Local files")
    timestamp = datetime.now().strftime("%m%d_%H%M")

    if st.button("💾 Save All Individual Files to Local Folder", type="primary"):
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
        st.success(f"✅ Individual files saved to: `{folder_path}`")

        # Also save the overall file to the same folder
        if 'overall_df' in locals():
            overall_filename = f"OVERALL_WORKLIST_{timestamp}.xlsx"
            overall_path = save_to_folder(
                output_dir=OUTPUT_FOLDER,
                folder_name="WORKLIST",
                files={overall_filename: overall_df},
                timestamp=timestamp,
            )
            st.success(f"✅ Overall combined file also saved as: `{overall_filename}`")

# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    worklist_app()