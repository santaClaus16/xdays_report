import streamlit as st
import pandas as pd
import numpy as np
import io
import os
import time
from datetime import datetime

from utils import (
    clean_account_number_column,
    load_mapping_file,
    apply_header_mapping,
    apply_cleaning,
    optimize_dtypes,
    format_date_columns,
    save_to_folder,
)

# ─────────────────────────────────────────────────────────────────────────────
# Constants
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="DataFlow Cleaner", page_icon="⚡", layout="wide")

MAPPING_PATH = r"C:\Users\SOLIZA\Documents\secret\mapping\mapping_file.xlsx"
OUTPUT_DIR   = r"C:\Users\SOLIZA\Documents\secret\clean_drr"

# ─────────────────────────────────────────────────────────────────────────────
# Main App
# ─────────────────────────────────────────────────────────────────────────────
def drr_cleaning_app():
    st.title("⚡ DataFlow Cleaner - Ultra Optimized")
    st.caption("Chunked processing + ETA + Memory optimization | Best for 100k–500k rows")

    # ── Load mapping ──────────────────────────────────────────────────────────
    rules, load_logs, standard_order = load_mapping_file(MAPPING_PATH)

    if any("❌" in log for log in load_logs):
        st.error("Mapping file error")
        for log in load_logs:
            st.write(log)
        st.stop()
    else:
        st.success("✅ Mapping file loaded successfully")

    # ── File upload ───────────────────────────────────────────────────────────
    data_upload = st.file_uploader(
        "Upload your data file (CSV or Excel)",
        type=["csv", "xlsx", "xls"]
    )

    if not (data_upload and st.button("🚀 Start Processing", type="primary")):
        return

    # ── Processing setup ──────────────────────────────────────────────────────
    start_time    = time.time()
    status_text   = st.empty()
    file_size_mb  = len(data_upload.getvalue()) / (1024 * 1024)
    is_csv        = data_upload.name.lower().endswith('.csv')
    chunk_size    = 80_000 if file_size_mb < 50 else 40_000 if file_size_mb < 150 else 25_000

    all_chunks     = []
    deleted_chunks = []
    total_rows     = 0
    all_logs       = list(load_logs)

    try:
        # ── Build reader ──────────────────────────────────────────────────────
        if is_csv:
            reader = pd.read_csv(data_upload, chunksize=chunk_size, low_memory=False)
        else:
            status_text.info("Loading Excel file (this may take a moment for large files)...")
            reader = [pd.read_excel(data_upload, engine='calamine')]

        # ── Chunk loop ────────────────────────────────────────────────────────
        for i, chunk in enumerate(reader):
            chunk       = chunk.copy()
            total_rows += len(chunk)

            status_text.info(
                f"Processing chunk {i+1} • {len(chunk):,} rows • Total so far: {total_rows:,}"
            )

            # Format dates
            chunk = format_date_columns(chunk)

            # Clean account numbers (show message only on first chunk)
            chunk = clean_account_number_column(chunk, show_success=(i == 0), st=st)

            # Memory optimization
            chunk = optimize_dtypes(chunk)

            # Apply mapping & cleaning rules
            chunk, map_logs   = apply_header_mapping(chunk, rules["mapping"], standard_order)
            chunk, deleted_chunk, clean_logs, _ = apply_cleaning(chunk, rules["cleaning"])

            all_logs.extend(map_logs)
            all_logs.extend(clean_logs)
            all_chunks.append(chunk)

            if not deleted_chunk.empty:
                deleted_chunks.append(deleted_chunk)

        # ── Combine chunks ────────────────────────────────────────────────────
        df         = pd.concat(all_chunks,     ignore_index=True) if all_chunks     else pd.DataFrame()
        deleted_df = pd.concat(deleted_chunks, ignore_index=True) if deleted_chunks else pd.DataFrame()

        # ── Split by SYSTEM remarks ───────────────────────────────────────────
        if "Remark By" in df.columns:
            system_mask    = df["Remark By"].astype(str).str.contains("SYSTEM", case=False, na=False)
            no_system_df   = df[~system_mask].copy()
            with_system_df = df[ system_mask].copy()
        else:
            no_system_df   = df.copy()
            with_system_df = pd.DataFrame(columns=df.columns)

        # ── Metrics ───────────────────────────────────────────────────────────
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Original Rows", f"{total_rows:,}")
        col2.metric("Rows Kept",     f"{len(df):,}",            delta=f"-{len(deleted_df):,}")
        col3.metric("No SYSTEM",     f"{len(no_system_df):,}")
        col4.metric("With SYSTEM",   f"{len(with_system_df):,}")

        # ── Tabs ──────────────────────────────────────────────────────────────
        tab1, tab2, tab3, tab4, tab5 = st.tabs([
            "✅ Cleaned Data", "📌 No SYSTEM", "📌 With SYSTEM", "🗑️ Deleted", "🔍 Logs"
        ])

        with tab1:
            st.success(f"✅ Cleaned Data — {len(df):,} rows")
            st.dataframe(df.head(10), use_container_width=True)

        with tab2:
            st.success(f"No SYSTEM — {len(no_system_df):,} rows")
            st.dataframe(no_system_df.head(10), use_container_width=True)

        with tab3:
            st.success(f"With SYSTEM — {len(with_system_df):,} rows")
            st.dataframe(with_system_df.head(10), use_container_width=True)

        with tab4:
            if deleted_df.empty:
                st.info("No rows were deleted.")
            else:
                st.error(f"🗑️ Deleted {len(deleted_df):,} rows")
                st.dataframe(deleted_df.head(10_000), use_container_width=True)

        with tab5:
            st.subheader("Processing Log")
            for line in all_logs[-60:]:
                if   "✅" in line: st.success(line)
                elif "🗑️" in line: st.error(line)
                elif "⚠️" in line: st.warning(line)
                else:              st.info(line)

        
            # ── Export ────────────────────────────────────────────────────────────────
        timestamp = datetime.now().strftime("%m-%d-%Y_%H%M")

        output_sheets = {
            "DAILY REMARKS REPORT | AGENT":  no_system_df,
            "DAILY REMARKS REPORT | SYSTEM": with_system_df,
        }

        # Save to local folder
        files_to_save = {
            f"{timestamp} - CLEAN DRR (AGENT).xlsx":  no_system_df,
            f"{timestamp} - CLEAN DRR (SYSTEM).xlsx": with_system_df,
        }

        folder_path = save_to_folder(
            output_dir  = OUTPUT_DIR,
            folder_name = "CLEAN DRR",
            files       = files_to_save,
            timestamp   = timestamp,
        )
        st.success(f"✅ Files saved to folder: `{folder_path}`")

        # In-memory download (single combined Excel)
        excel_bytes = io.BytesIO()
        with pd.ExcelWriter(excel_bytes, engine="openpyxl") as writer:
            for sheet_name, frame in output_sheets.items():
                frame.to_excel(writer, sheet_name=sheet_name[:31], index=False)

        st.download_button(
            "⬇️ Download Full Output",
            data=excel_bytes.getvalue(),
            file_name=f"dataflow_output_{timestamp}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    except Exception as e:
        st.error(f"❌ Processing failed: {str(e)}")
        st.exception(e)


# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    drr_cleaning_app()