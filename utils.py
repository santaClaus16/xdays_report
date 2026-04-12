# utils.py
import numpy as np
import pandas as pd
import os
from datetime import datetime


# ─────────────────────────────────────────────────────────────────────────────
# 1. Clean account number column
# ─────────────────────────────────────────────────────────────────────────────
def clean_account_number_column(chunk: pd.DataFrame, show_success: bool = True, st=None) -> pd.DataFrame:
    """
    Detects and cleans account number columns by adding leading zeros.

    Args:
        chunk       : DataFrame chunk to process.
        show_success: Whether to show a Streamlit success message.
        st          : Streamlit module (pass `st` from your app).

    Returns:
        Modified DataFrame with cleaned account numbers.
    """
    account_col_names = {"account no.", "account no", "account_no", "account number", "CUST_ID", "CUST ID","cust_id"}

    for col in chunk.columns:
        if str(col).strip().lower() in account_col_names:
            s = chunk[col].astype(str).str.strip()
            s = s.str.replace(r'\.0$', '', regex=True)
            s = np.where((s != '') & (~s.str.startswith('0')) & (s != 'nan'), '0' + s, s)
            s = np.where(s == 'nan', '', s)
            chunk[col] = s
            if show_success and st:
                st.success(f"✅ Account numbers cleaned in column: **{col}**")
            break

    return chunk


# ─────────────────────────────────────────────────────────────────────────────
# 2. Load mapping file
# ─────────────────────────────────────────────────────────────────────────────
def load_mapping_file(path: str):
    """
    Reads mapping and cleaning rules from an Excel file.

    Args:
        path: Full path to the mapping Excel file.

    Returns:
        Tuple of (rules dict, logs list, standard_order list)
    """
    rules = {"mapping": {}, "cleaning": {}}
    logs = []
    standard_order = []

    if not os.path.exists(path):
        return rules, ["❌ Mapping file not found"], []

    try:
        xls = pd.ExcelFile(path)
    except Exception as e:
        return rules, [f"❌ Cannot open mapping: {e}"], []

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
            standard_order.append(std)
            if aliases:
                rules["mapping"][std] = aliases
                logs.append(f"✅ Mapping '{std}' ← {aliases}")
            else:
                logs.append(f"✅ Standard column added: '{std}' (header only)")

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


# ─────────────────────────────────────────────────────────────────────────────
# 3. Apply header mapping
# ─────────────────────────────────────────────────────────────────────────────
def apply_header_mapping(df: pd.DataFrame, mapping: dict, standard_order: list):
    """
    Renames columns based on mapping rules and reorders to standard order.

    Args:
        df            : Input DataFrame.
        mapping       : Dict of {standard_name: [aliases]}.
        standard_order: List of column names in desired order.

    Returns:
        Tuple of (modified DataFrame, logs list)
    """
    logs = []
    reverse = {}

    for std, aliases in mapping.items():
        for alias in aliases:
            reverse[alias.lower()] = std
        reverse[std.lower()] = std

    rename_map = {}
    for col in df.columns:
        cl = str(col).strip().lower()
        if cl in reverse:
            std = reverse[cl]
            if str(col).strip() != std:
                rename_map[col] = std
                logs.append(f"✅ Renamed '{col}' → '{std}'")

    df = df.rename(columns=rename_map)

    for col in standard_order:
        if col not in df.columns:
            df[col] = np.nan
            logs.append(f"➕ Added missing column '{col}' (empty)")

    df = df[standard_order]
    return df, logs


# ─────────────────────────────────────────────────────────────────────────────
# 4. Apply cleaning rules
# ─────────────────────────────────────────────────────────────────────────────
def apply_cleaning(df: pd.DataFrame, cleaning: dict):
    """
    Deletes rows that match cleaning rules.

    Args:
        df      : Input DataFrame.
        cleaning: Dict of {column_header: [values_to_delete]}.

    Returns:
        Tuple of (cleaned_df, deleted_df, logs, deleted_count)
    """
    logs = []
    deleted_mask = pd.Series([False] * len(df), index=df.index)

    for header, del_vals in cleaning.items():
        match_col = next((c for c in df.columns if str(c).strip().lower() == header.strip().lower()), None)
        if match_col is None:
            logs.append(f"⚠️ Cleaning column '{header}' not found — skipped")
            continue

        mask = pd.Series([False] * len(df), index=df.index)

        has_blank_rule = any(pd.isna(v) or str(v).strip() == "" for v in del_vals)
        if has_blank_rule:
            blank_mask = df[match_col].isna() | (df[match_col].astype(str).str.strip() == "")
            mask |= blank_mask
            logs.append(f"🗑️ Will delete rows where '{match_col}' is blank/empty")

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


# ─────────────────────────────────────────────────────────────────────────────
# 5. Optimize data types
# ─────────────────────────────────────────────────────────────────────────────
def optimize_dtypes(df: pd.DataFrame, category_cols: list = None) -> pd.DataFrame:
    """
    Downcasts numeric columns and converts specified columns to category dtype.

    Args:
        df             : Input DataFrame.
        category_cols  : List of column names to cast as 'category'.
                         Defaults to ['Remark By', 'STATUS', 'Client Type'].

    Returns:
        Optimized DataFrame.
    """
    if category_cols is None:
        category_cols = ['Remark By', 'STATUS', 'Client Type']

    for col in df.select_dtypes(include=['float64']).columns:
        df[col] = df[col].astype('float32')
    for col in df.select_dtypes(include=['int64']).columns:
        df[col] = df[col].astype('int32')
    for col in category_cols:
        if col in df.columns:
            df[col] = df[col].astype('category')

    return df


# ─────────────────────────────────────────────────────────────────────────────
# 6. Format date columns
# ─────────────────────────────────────────────────────────────────────────────
def format_date_columns(chunk: pd.DataFrame, date_cols: list = None) -> pd.DataFrame:
    """
    Converts date columns to MM/DD/YYYY string format.

    Args:
        chunk    : DataFrame chunk to process.
        date_cols: List of date column names.
                   Defaults to ['Date', 'PTP Date', 'Claim Paid Date'].

    Returns:
        DataFrame with formatted date columns.
    """
    if date_cols is None:
        date_cols = ["Date", "PTP Date", "Claim Paid Date"] #default

    for col in date_cols:
        if col in chunk.columns:
            chunk[col] = pd.to_datetime(chunk[col], errors='coerce').dt.strftime('%m/%d/%Y').fillna('')

    return chunk


# ─────────────────────────────────────────────────────────────────────────────
# 7. Save multiple DataFrames into a timestamped folder
# ─────────────────────────────────────────────────────────────────────────────
def save_to_folder(output_dir: str, folder_name: str, files: dict, timestamp: str = None) -> str:
    """
    Creates a timestamped subfolder and saves multiple DataFrames as Excel files.

    Args:
        output_dir  : Base directory
        folder_name : Subfolder label (e.g. "CLEAN DRR" or "WORKLIST")
        files       : Dict of {filename: DataFrame} to save
        timestamp   : Optional timestamp string. Auto-generated if None.

    Returns:
        Full path of the created folder.
    """
    if timestamp is None:
        timestamp = datetime.now().strftime("%m-%d-%Y_%H%M")

    folder_path = os.path.join(output_dir, f"{timestamp} - {folder_name}")
    os.makedirs(folder_path, exist_ok=True)

    for filename, df in files.items():
        file_path = os.path.join(folder_path, filename)
        df.to_excel(file_path, index=False, engine="openpyxl")

    return folder_path