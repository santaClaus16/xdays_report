import streamlit as st
import pandas as pd
import numpy as np
import io
import os
import time
from datetime import timedelta
from datetime import datetime

# ─────────────────────────────────────────────────────────────────────────────
# Page Config
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="DataFlow Cleaner", page_icon="⚡", layout="wide")

# ─────────────────────────────────────────────────────────────────────────────
# Constants
# ─────────────────────────────────────────────────────────────────────────────
MAPPING_PATH = r"C:\Users\SOLIZA\Documents\secret\mapping\mapping_file.xlsx"
output_clean_drr = r"C:\Users\SOLIZA\Documents\secret\clean_drr"

# ─────────────────────────────────────────────────────────────────────────────
# Core Functions
# ─────────────────────────────────────────────────────────────────────────────

def load_mapping_file(path: str):
    rules = {"mapping": {}, "cleaning": {}}
    logs = []
    standard_order = []

    if not os.path.exists(path):
        return rules, ["❌ Mapping file not found"], []

    try:
        xls = pd.ExcelFile(path)
    except Exception as e:
        return rules, [f"❌ Cannot open mapping: {e}"], []

    # ====================== MAPPING SHEET ======================
    if "mapping" in xls.sheet_names:
        dm = pd.read_excel(xls, sheet_name="mapping", header=None)

        for ci in range(dm.shape[1]):
            col = dm.iloc[:, ci].dropna().tolist()
            if not col:
                continue

            std = str(col[0]).strip()
            if not std:
                continue

            aliases = [str(v).strip() for v in col[1:] if str(v).strip()]

            # <<< KEY FIX: Always add to standard_order, even if no aliases >>>
            standard_order.append(std)

            if aliases:
                rules["mapping"][std] = aliases
                logs.append(f"✅ Mapping '{std}' ← {aliases}")
            else:
                # Still useful — this column should appear in the final order
                logs.append(f"✅ Standard column added: '{std}' (header only)")

    # ====================== CLEANING SHEET ======================
    if "cleaning" in xls.sheet_names:
        dc = pd.read_excel(xls, sheet_name="cleaning", header=None)
        for ci in range(dc.shape[1]):
            col = dc.iloc[:, ci].dropna().tolist()
            if col:
                header = str(col[0]).strip()
                del_vals = [str(v).strip() for v in col[1:] if str(v).strip()]
                if del_vals:
                    rules["cleaning"][header] = del_vals
                    logs.append(f"🗑️ Cleaning '{header}' → delete if: {del_vals}")

    return rules, logs, standard_order

def apply_header_mapping(df: pd.DataFrame, mapping: dict, standard_order: list):
    logs = []
    reverse = {}

    for std, aliases in mapping.items():
        for alias in aliases:
            reverse[alias.lower()] = std
        reverse[std.lower()] = std   # exact standard name also maps to itself

    rename_map = {}
    for col in df.columns:
        cl = str(col).strip().lower()
        if cl in reverse:
            std = reverse[cl]
            if str(col).strip() != std:
                rename_map[col] = std
                logs.append(f"✅ Renamed '{col}' → '{std}'")
            # else: already correct name, no rename needed

    df = df.rename(columns=rename_map)

    # Reorder columns: standards first, then any extra columns
    # present = [std for std in standard_order if std in df.columns]
    # extra = [c for c in df.columns if c not in present]
    
    # return df[present + extra], logs
    
    # ✅ Ensure ALL mapping headers exist (even if missing in upload)
    for col in standard_order:
        if col not in df.columns:
            df[col] = np.nan   # or np.nan if you prefer
            logs.append(f"➕ Added missing column '{col}' (empty)")

    # ✅ STRICT ORDER — ONLY mapping headers
    df = df[standard_order]

    return df, logs

def apply_cleaning(df: pd.DataFrame, cleaning: dict):
    logs = []
    deleted_mask = pd.Series([False] * len(df), index=df.index)

    for header, del_vals in cleaning.items():
        match_col = next((c for c in df.columns if str(c).strip().lower() == header.strip().lower()), None)
        if match_col is None:
            logs.append(f"⚠️ Cleaning column '{header}' not found — skipped")
            continue

        mask = pd.Series([False] * len(df), index=df.index)

        # Special handling for blanks
        has_blank_rule = any(pd.isna(v) or str(v).strip() == "" for v in del_vals)
        
        if has_blank_rule:
            blank_mask = df[match_col].isna() | (df[match_col].astype(str).str.strip() == "")
            mask |= blank_mask
            logs.append(f"🗑️ Will delete rows where '{match_col}' is blank/empty")

        # Normal value matching
        non_blank_vals = [str(v).strip() for v in del_vals if pd.notna(v) and str(v).strip() != ""]
        if non_blank_vals:
            lower_vals = [v.lower() for v in non_blank_vals]
            value_mask = df[match_col].astype(str).str.strip().str.lower().isin(lower_vals)
            mask |= value_mask
            logs.append(f"🗑️ Will delete rows where '{match_col}' matches: {non_blank_vals}")

        count = int(mask.sum())
        deleted_mask |= mask
        
        if count:
            logs.append(f"🗑️ Deleted {count} row(s) from '{match_col}' cleaning rule")

    deleted_df = df[deleted_mask].copy()
    cleaned_df = df[~deleted_mask].copy()
    
    return cleaned_df, deleted_df, logs, int(deleted_mask.sum())


def optimize_dtypes(df: pd.DataFrame) -> pd.DataFrame:
    for col in df.select_dtypes(include=['float64']).columns:
        df[col] = df[col].astype('float32')
    for col in df.select_dtypes(include=['int64']).columns:
        df[col] = df[col].astype('int32')
    for col in ['Remark By', 'STATUS', 'Client Type']:
        if col in df.columns:
            df[col] = df[col].astype('category')
    return df


# ─────────────────────────────────────────────────────────────────────────────
# Main App
# ─────────────────────────────────────────────────────────────────────────────
def drr_cleaning_app():
    st.title("⚡ DataFlow Cleaner - Ultra Optimized")
    st.caption("Chunked processing + ETA + Memory optimization | Best for 100k–500k rows")

    rules, load_logs, standard_order = load_mapping_file(MAPPING_PATH)

    if any("❌" in log for log in load_logs):
        st.error("Mapping file error")
        for log in load_logs:
            st.write(log)
        st.stop()
    else:
        st.success("✅ Mapping file loaded successfully")

    data_upload = st.file_uploader("Upload your data file (CSV or Excel)", 
                                  type=["csv", "xlsx", "xls"])

    if data_upload and st.button("🚀 Start Processing", type="primary"):
        start_time = time.time()
        # progress_bar = st.progress(0)
        status_text = st.empty()
        eta_text = st.empty()

        file_size_mb = len(data_upload.getvalue()) / (1024 * 1024)
        is_csv = data_upload.name.lower().endswith('.csv')

        # Auto adjust chunk size
        chunk_size = 80_000 if file_size_mb < 50 else 40_000 if file_size_mb < 150 else 25_000

        all_chunks = []
        deleted_chunks = []
        total_rows = 0
        all_logs = list(load_logs)

        try:
            if is_csv:
                reader = pd.read_csv(data_upload, chunksize=chunk_size, low_memory=False)
            else:
                status_text.info("Loading Excel file (this may take a moment for large files)...")
                df_full = pd.read_excel(data_upload, engine='calamine')
                reader = [df_full]

            for i, chunk in enumerate(reader):
                chunk_start = time.time()
                chunk = chunk.copy()
                total_rows += len(chunk)

                status_text.info(f"Processing chunk {i+1} • {len(chunk):,} rows • Total so far: {total_rows:,}")

                # Fast Formatting
                for col in ["Date", "PTP Date", "Claim Paid Date"]:
                    if col in chunk.columns:
                        chunk[col] = pd.to_datetime(chunk[col], errors='coerce').dt.strftime('%m/%d/%Y').fillna('')

                # for col in ["Balance", "PTP Amount", "Claim Paid Amount"]:
                #     if col in chunk.columns:
                #         chunk[col] = pd.to_numeric(chunk[col], errors='coerce').round(2)

                # Account No. with leading zero (only show message once)
                if i == 0:
                    for col in chunk.columns:
                        if str(col).strip().lower() in ["account no.", "account no", "account_no", "account number"]:
                            s = chunk[col].astype(str).str.strip()
                            s = s.str.replace(r'\.0$', '', regex=True)           # Remove .0 from float
                            # s = s.str.replace(r'[^0-9]', '', regex=True)         # Keep only digits
                            s = np.where((s != '') & (~s.str.startswith('0')) & (s != 'nan'), '0' + s, s)
                            s = np.where(s == 'nan', '', s)
                            chunk[col] = s
                            st.success(f"✅ Account numbers cleaned in column: **{col}**")
                            break

                # Memory optimization
                chunk = optimize_dtypes(chunk)

                # Mapping & Cleaning ← FIXED HERE
                chunk, map_logs = apply_header_mapping(chunk, rules["mapping"], standard_order)
                chunk, deleted_chunk, clean_logs, deleted_count = apply_cleaning(chunk, rules["cleaning"])

                all_logs.extend(map_logs)
                all_logs.extend(clean_logs)

                all_chunks.append(chunk)
                if not deleted_chunk.empty:
                    deleted_chunks.append(deleted_chunk)

                # # ETA Calculation
                # elapsed = time.time() - start_time
                # avg_time_per_chunk = elapsed / (i + 1)
                # estimated_total_chunks = total_rows / len(chunk) if len(chunk) > 0 else 1
                # remaining_chunks = estimated_total_chunks - (i + 1)
                # eta_seconds = avg_time_per_chunk * max(remaining_chunks, 0)
                # eta_str = str(timedelta(seconds=int(eta_seconds))) if eta_seconds > 0 else "Finishing..."

                # eta_text.info(f"⏱️ Elapsed: {str(timedelta(seconds=int(elapsed)))} | ETA: {eta_str}")

                # progress = min(0.95 * (total_rows / max(total_rows, 1)), 0.95)
                # progress_bar.progress(progress)

            # Combine chunks
            df = pd.concat(all_chunks, ignore_index=True) if all_chunks else pd.DataFrame()
            deleted_df = pd.concat(deleted_chunks, ignore_index=True) if deleted_chunks else pd.DataFrame()

            # Split by Remark By
            if "Remark By" in df.columns:
                no_system_df = df[~df["Remark By"].astype(str).str.contains("SYSTEM", case=False, na=False)].copy()
                with_system_df = df[df["Remark By"].astype(str).str.contains("SYSTEM", case=False, na=False)].copy()
            else:
                no_system_df = df.copy()
                with_system_df = pd.DataFrame(columns=df.columns)

            cleaned_rows = len(df)

            # ====================== RESULTS ======================
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Original Rows", f"{total_rows:,}")
            col2.metric("Rows Kept", f"{cleaned_rows:,}", delta=f"-{len(deleted_df):,}")
            col3.metric("No SYSTEM", f"{len(no_system_df):,}")
            col4.metric("With SYSTEM", f"{len(with_system_df):,}")

            tab1, tab2, tab3, tab4, tab5 = st.tabs([
                "✅ Cleaned Data", "📌 No SYSTEM", "📌 With SYSTEM", "🗑️ Deleted", "🔍 Logs"
            ])

            with tab1:
                st.success(f"✅ Cleaned Data — {cleaned_rows:,} rows")
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
                    if "✅" in line:
                        st.success(line)
                    elif "🗑️" in line:
                        st.error(line)
                    elif "⚠️" in line:
                        st.warning(line)
                    else:
                        st.info(line)

            # ====================== DOWNLOAD ======================
            
            
            
            output_requires = {
                "DAILY REMARKS REPORT | AGENT": no_system_df,
                "DAILY REMARKS REPORT | SYSTEM": with_system_df
            }

            
            timestamp = datetime.now().strftime("%m-%d-%Y_%H%M")
            os.makedirs(output_clean_drr, exist_ok=True)

            export_path = os.path.join(output_clean_drr, f"{timestamp} - CLEAN DRR.xlsx")

            with pd.ExcelWriter(export_path, engine="openpyxl") as writer:
                for sheet_name, frame in output_requires.items():
                    frame.to_excel(writer, sheet_name=sheet_name, index=False)

            st.success(f"✅ File saved locally: `{export_path}`")

            # In-memory download
            excel_bytes = io.BytesIO()
            with pd.ExcelWriter(excel_bytes, engine="openpyxl") as writer:
                for sheet_name, frame in output_requires.items():
                    frame.to_excel(writer, index=False, sheet_name=sheet_name[:31])

            st.download_button(
                "⬇️ Download Full Output",
                data=excel_bytes.getvalue(),
                file_name=f"dataflow_output_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        except Exception as e:
            st.error(f"❌ Processing failed: {str(e)}")
            st.exception(e)   # This helps debug during development


# Run the app
if __name__ == "__main__":
    drr_cleaning_app()