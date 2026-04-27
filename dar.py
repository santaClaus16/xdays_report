import io
from io import BytesIO

import msoffcrypto
import pandas as pd
import streamlit as st

# ─────────────────────────────────────────────
# Constants
# ─────────────────────────────────────────────

PAGE_TITLE = "Excel Report Reader"
APP_TITLE = "CURING REPORT OMTOOL"
DEFAULT_PASSWORD = "spmadrid0348"

HEADER_DEFAULTS: dict = {
    "Deb Collection Agency": "SP Madrid",
    "Email Address": "bpalegal@spmadridlaw.com",
    "I agree that": "I agree",
    "Account Type": "",
    "CTL4 / Unit": "",
    "Account Number/LAN/PAN": "",
    "Loan Account Name": "",
    "Endo Date": "",
    "Transaction": "Report",
    "Report Submission": "",
    "Activity": "CALB - Collector Called Borrower",
    "Client Contact Number": "",
    "Client Email Address": "",
    "With New Contact Information?": "No",
    "New Contact Type": "",
    "New Contact Number": "",
    "New Email Address": "",
    "Client Type": "Primary Borrower",
    "Negotiation Remarks": "",
    "Reason for Default (DAR)": "",
    "Negotiation Status": "",
    "PTP / PTVS Date (DAR)": "",
    "PTP Amount": "",
    "STATUS": "FOR INPUT",
    "DCA Coordinator": "",
    "Clean Status": "",
    "Hierarchy": 0,
    "OB": "",
    "DPR Type": "Cash Settlement",
    "Cash Settlement": "Full Payment",
    "Payment Amount": "",
    "Payment / Repo Date": "",
}

# Maps remark keyword → (negotiation status label, hierarchy)
NEGOTIATION_STATUS_MAP: dict[str, tuple[str, int]] = {
    "PTPA":           ("PTVS - Promised-to-Voluntary Surrender", 4),
    "PTPB":           ("PTPB - Promised-to-Pay", 4),
    "FPTP":           ("INS - With Insurance Dispute", 3),
    "NEGO":           ("NEGO - No Commitment To Pay", 2),
    "BUSY":           ("BUSY - Phone Uncontacted - Busy Signal", 1),
    "COFF":           ("COFF - Cellphone Out of Coverage/Turned-Off", 1),
    "DISC":           ("DISP - With Payment Claims / Dispute", 1),
    "KRNG":           ("KRNG - Keeps on Ringing", 1),
    "NYIS":           ("NYIS - Number Dialled Not Yet In Service", 1),
    "WRBR":           ("REPA - With Pending Restructuring", 1),
    "MSSG":           ("MSSG - Left Message/Payment Reminder", 1),
    "PLAYED MESSAGE": ("MSSG - Left Message/Payment Reminder", 1),
    "NO ANSWER":      ("KRNG - Keeps on Ringing", 1),
}

DCA_COORDINATOR_MAP: dict[str, str] = {
    "1-30 dpd":  "Recca Marie",
    "30-60 dpd": "Charmaine",
}

DAILY_ACTIVITY_COLS: list[str] = [
    "Deb Collection Agency", "Email Address", "I agree that", "Account Type",
    "CTL4 / Unit", "Account Number/LAN/PAN", "Loan Account Name", "Endo Date",
    "Transaction", "Report Submission", "Activity", "Client Contact Number",
    "Client Email Address", "With New Contact Information?", "New Contact Type",
    "New Contact Number", "New Email Address", "Client Type", "Negotiation Remarks",
    "Reason for Default (DAR)", "Negotiation Status", "PTP / PTVS Date (DAR)",
    "PTP Amount", "STATUS", "DCA Coordinator",
]

DAILY_PRODUCTIVITY_COLS: list[str] = [
    "Deb Collection Agency", "Email Address", "I agree that", "Account Type",
    "CTL4 / Unit", "Account Number/LAN/PAN", "Loan Account Name", "Endo Date",
    "Transaction", "Report Submission", "OB", "DPR Type", "Cash Settlement",
    "Payment Amount", "Payment / Repo Date", "STATUS", "DCA Coordinator",
]

# ─────────────────────────────────────────────
# Excel Helpers
# ─────────────────────────────────────────────

def decrypt_excel(file_upload: io.BytesIO, password: str) -> pd.ExcelFile:
    """Decrypt (if needed) and return a pd.ExcelFile."""
    decrypted = io.BytesIO()
    try:
        file_upload.seek(0)
        office_file = msoffcrypto.OfficeFile(file_upload)
        if office_file.is_encrypted():
            office_file.load_key(password=password)
            office_file.decrypt(decrypted)
        else:
            file_upload.seek(0)
            decrypted.write(file_upload.read())
    except Exception:
        file_upload.seek(0)
        decrypted.write(file_upload.read())
    decrypted.seek(0)
    return pd.ExcelFile(decrypted)


def resolve_column(df: pd.DataFrame, *candidates: str) -> str | None:
    """Return the first candidate column name present in df, else None."""
    return next((c for c in candidates if c in df.columns), None)

# ─────────────────────────────────────────────
# Data Processing
# ─────────────────────────────────────────────

def process_amount_column(df: pd.DataFrame, *candidates: str, target: str) -> None:
    """Coerce an amount column to numeric; replace non-positive values with ''."""
    col = resolve_column(df, *candidates)
    if col is None:
        return
    series = pd.to_numeric(
        df[col].astype(str).str.replace(",", "", regex=False),
        errors="coerce",
    ).fillna(0)
    df[target] = series.apply(lambda x: "" if x <= 0 else x)


def process_date_column(
    df: pd.DataFrame,
    *candidates: str,
    target: str,
    fmt: str = "%m/%d/%Y",
) -> None:
    """Parse and reformat a date column; invalid dates become ''."""
    col = resolve_column(df, *candidates)
    if col is None:
        return
    df[target] = pd.to_datetime(df[col], errors="coerce").dt.strftime(fmt).fillna("")


def clean_remarks(df: pd.DataFrame, *candidates: str, target: str) -> None:
    """Keep only alphanumeric and whitespace characters in a remarks column."""
    col = resolve_column(df, *candidates)
    if col is None:
        return
    df[target] = (
        df[col]
        .fillna("")
        .astype(str)
        .apply(lambda t: "".join(c for c in t if c.isalnum() or c.isspace()).strip())
    )


def extract_negotiation_status(df: pd.DataFrame, remark_col: str = "Remark") -> None:
    """
    Derive 'Clean Status', 'Negotiation Status', and 'Hierarchy'
    from the remark column using NEGOTIATION_STATUS_MAP.
    """
    if remark_col not in df.columns:
        return

    def _find_code(remark) -> str:
        if pd.isna(remark):
            return ""
        upper = str(remark).upper()
        return next((code for code in NEGOTIATION_STATUS_MAP if code in upper), "")

    df["Clean Status"] = df[remark_col].apply(_find_code)
    df["Negotiation Status"] = df["Clean Status"].map(
        lambda c: NEGOTIATION_STATUS_MAP.get(c, ("", 0))[0]
    )
    df["Hierarchy"] = df["Clean Status"].map(
        lambda c: NEGOTIATION_STATUS_MAP.get(c, ("", 0))[1]
    )


def deduplicate_by_hierarchy(df: pd.DataFrame) -> pd.DataFrame:
    """Keep only the highest-hierarchy row per LAN."""
    if "LAN" not in df.columns or "Hierarchy" not in df.columns:
        st.warning("LAN or Hierarchy column missing — using all rows.")
        return df
    return (
        df.sort_values("Hierarchy", ascending=False)
        .drop_duplicates(subset=["LAN"], keep="first")
    )


def transform_columns(df: pd.DataFrame) -> None:
    """Run all column transformations in-place on the matched dataframe."""
    process_amount_column(df, "PTP Amount", target="PTP Amount")
    process_amount_column(
        df,
        "Claim Paid Amount", "Claim Paid Amout", "Paid Amount",
        target="Claim Paid Amount",
    )
    process_amount_column(
        df,
        "Outstanding Balance", "Outstanding Bal", "OS Balance",
        target="OB",
    )
    process_date_column(df, "ENDORSEMENT DATE", target="ENDORSEMENT DATE")
    process_date_column(
        df,
        "PTP DATE", "PTP Date", "PTP / PTVS Date (DAR)", "PTP Date (DAR)",
        target="PTP Date",
    )
    process_date_column(df, "PTP / PTVS Date (DAR)", target="PTP / PTVS Date (DAR)")
    clean_remarks(df, "Remark", "Remarks", "REMARK", target="Negotiation Remarks")
    extract_negotiation_status(df)

# ─────────────────────────────────────────────
# Report Builder
# ─────────────────────────────────────────────

def build_mapped_df(
    df: pd.DataFrame,
    account_type: str,
    dca_coordinator: str,
) -> pd.DataFrame:
    """Map the cleaned dataframe onto HEADER_DEFAULTS."""
    records = []
    for _, row in df.iterrows():
        record = HEADER_DEFAULTS.copy()

        # Auto-map exact column name matches (skip NaN)
        for key in record:
            if key in df.columns and pd.notna(row[key]):
                record[key] = row[key]

        # Explicit overrides
        record.update({
            "Account Number/LAN/PAN": row.get("LAN", ""),
            "CTL4 / Unit":            row.get("CTL4", ""),
            "Loan Account Name":      row.get("ACCOUNT NAME", ""),
            "Endo Date":              row.get("ENDORSEMENT DATE", ""),
            "Account Type":           account_type,
            "Client Contact Number":  row.get("Dialed Number", ""),
            "Report Submission":      "",  # set per-report slice below
            "DCA Coordinator":        dca_coordinator,
            "PTP / PTVS Date (DAR)":  row.get("PTP Date", ""),
            "Clean Status":           row.get("Clean Status", ""),
            "Hierarchy":              row.get("Hierarchy", 0),
            "Negotiation Status":     row.get("Negotiation Status", ""),
            "Negotiation Remarks":    row.get("Negotiation Remarks", row.get("Remark", "")),
            "Payment Amount":         row.get("PTP Amount", ""),
            "Payment / Repo Date":    row.get("PTP Date", ""),
            "OB":                     row.get("OB", ""),
        })
        records.append(record)

    return pd.DataFrame(records)


def build_report_slices(
    mapped_df: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Return (daily_activity, daily_productivity).
    Daily productivity is filtered to PTPA/PTPB rows only.
    """
    daily_activity = mapped_df[DAILY_ACTIVITY_COLS].copy()
    daily_activity["Report Submission"] = "Daily Activity Report"

    daily_productivity = (
        mapped_df[mapped_df["Clean Status"].isin(["PTPA", "PTPB"])]
        [DAILY_PRODUCTIVITY_COLS]
        .copy()
    )
    daily_productivity["Report Submission"] = "Daily Productivity Report"

    return daily_activity, daily_productivity


def to_excel_bytes(*sheet_pairs: tuple[pd.DataFrame, str]) -> bytes:
    """Serialize one or more (DataFrame, sheet_name) pairs to Excel bytes."""
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        for df, sheet_name in sheet_pairs:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    return buf.getvalue()


def to_zip_bytes(files: dict[str, bytes]) -> bytes:
    """
    Pack multiple files into a zip archive.
    files: {filename: raw_bytes}
    """
    import zipfile
    buf = BytesIO()
    with zipfile.ZipFile(buf, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for filename, data in files.items():
            zf.writestr(filename, data)
    return buf.getvalue()

# ─────────────────────────────────────────────
# UI Components
# ─────────────────────────────────────────────

def render_sidebar() -> tuple[str, str, str]:
    """Render sidebar controls; return (password, account_type, dca_coordinator)."""
    st.sidebar.markdown("### Password")
    pwd_option = st.sidebar.radio(
        "Password Option:",
        ["Automatic Password (Default: spmadrid0348)", "Manual Password"],
    )
    password = DEFAULT_PASSWORD
    if pwd_option == "Manual Password":
        password = st.sidebar.text_input("Enter password:", type="password")

    st.sidebar.markdown("---")
    st.sidebar.markdown("### Account Type")
    account_type = st.sidebar.radio("Select Account Type:", list(DCA_COORDINATOR_MAP))
    dca_coordinator = DCA_COORDINATOR_MAP[account_type]

    return password, account_type, dca_coordinator


def render_sheet_selectors(
    xls_tad: pd.ExcelFile, xls_drr: pd.ExcelFile
) -> tuple[str, str]:
    st.sidebar.markdown("---")
    st.sidebar.markdown("### Sheet Selection")
    tad_sheet = st.sidebar.selectbox("Select TAD Sheet", xls_tad.sheet_names)
    drr_sheet = st.sidebar.selectbox("Select DRR Sheet", xls_drr.sheet_names)
    return tad_sheet, drr_sheet


def render_status_summary(matched_df: pd.DataFrame, final_df: pd.DataFrame) -> None:
    status_counts = matched_df["Clean Status"].value_counts()
    with st.expander("Status Summary"):
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("PTPA",   status_counts.get("PTPA", 0))
        c2.metric("PTPB",   status_counts.get("PTPB", 0))
        c3.metric("Others", status_counts.drop(["PTPA", "PTPB"], errors="ignore").sum())
        c4.metric("Total",  len(matched_df))

        c1, c2 = st.columns(2)
        c1.metric("Original rows", len(matched_df))
        c2.metric("After dedup",   len(final_df))

        preview_cols = ["LAN", "Status", "Negotiation Status"]
        if "ACCOUNT NAME" in final_df.columns:
            preview_cols.insert(1, "ACCOUNT NAME")

        with st.expander("Preview Clean Data"):
            st.dataframe(
                final_df[[c for c in preview_cols if c in final_df.columns]].reset_index(drop=True),
                use_container_width=True,
            )


def render_report_previews(
    daily_activity: pd.DataFrame, daily_productivity: pd.DataFrame
) -> None:
    with st.expander("Daily Activity Report"):
        st.dataframe(daily_activity, use_container_width=True)
    with st.expander("Daily Productivity Report"):
        st.dataframe(daily_productivity, use_container_width=True)

    c1, c2 = st.columns(2)
    c1.metric("Daily Activity Records",     len(daily_activity))
    c2.metric("Daily Productivity Records", len(daily_productivity))


def render_downloads(
    daily_activity: pd.DataFrame,
    daily_productivity: pd.DataFrame,
    account_type: str,
) -> None:
    zip_data = to_zip_bytes({
        f"Combined_Daily_Reports_{account_type}.xlsx": to_excel_bytes(
            (daily_activity,     "Daily Activity"),
            (daily_productivity, "Daily Productivity"),
        ),
        f"Daily_Activity_Report_{account_type}.xlsx": to_excel_bytes(
            (daily_activity, "Daily Activity"),
        ),
        f"PTP_Productivity_Report_{account_type}.xlsx": to_excel_bytes(
            (daily_productivity, "PTP Productivity"),
        ),
    })
    st.download_button(
        label="📦 Download All Reports (.zip)",
        data=zip_data,
        file_name=f"Curing_Reports_{account_type}.zip",
        mime="application/zip",
    )

# ─────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────

def main() -> None:
    st.set_page_config(page_title=PAGE_TITLE, layout="wide")
    st.title(APP_TITLE)

    password, account_type, dca_coordinator = render_sidebar()

    # ── File uploaders ────────────────────────
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("Upload a **TAD** file")
        tad_file = st.file_uploader("Select TAD File", type=["xlsx", "xls"])
    with col2:
        st.markdown("Upload a **DRR** file")
        drr_file = st.file_uploader("Select DRR File", type=["xlsx", "xls"])

    if tad_file is None or drr_file is None:
        if tad_file or drr_file:
            st.info("ℹ️ Please upload both the TAD and DRR files to proceed.")
        return

    if not password:
        st.warning("Please enter a manual password to proceed.")
        return

    # ── Decrypt ───────────────────────────────
    try:
        xls_tad = decrypt_excel(tad_file, password)
        xls_drr = decrypt_excel(drr_file, password)
    except msoffcrypto.exceptions.InvalidKeyError:
        st.error("❌ Incorrect password. Failed to decrypt one or both files.")
        return

    tad_sheet, drr_sheet = render_sheet_selectors(xls_tad, xls_drr)

    if not st.button("Match and Read Files", type="primary"):
        return

    # ── Load sheets ───────────────────────────
    try:
        # dtype=str prevents pandas from coercing numbers (e.g. LAN) prematurely
        df_tad = pd.read_excel(xls_tad, sheet_name=tad_sheet, dtype=str)
        df_drr = pd.read_excel(xls_drr, sheet_name=drr_sheet, dtype=str)
    except ValueError as e:
        st.error(f"❌ Could not read sheets: {e}")
        return

    if "LAN" not in df_tad.columns:
        st.error(f"Column 'LAN' not found in TAD sheet '{tad_sheet}'.")
        return
    if "Account No." not in df_drr.columns:
        st.error(f"Column 'Account No.' not found in DRR sheet '{drr_sheet}'.")
        return

    # ── Merge ─────────────────────────────────
    df_tad["LAN"]         = df_tad["LAN"].str.strip()
    df_drr["Account No."] = df_drr["Account No."].str.strip()
    matched_df = pd.merge(df_tad, df_drr, left_on="LAN", right_on="Account No.", how="inner")

    c1, c2 = st.columns(2)
    c1.metric("TAD Active Accounts", len(df_tad))
    c2.metric("DRR Records",         len(df_drr))

    c1, c2 = st.columns(2)
    c1.metric("Matched Rows", len(matched_df))
    ptp_count = (
        matched_df["Status"].str.contains("PTP", na=False).sum()
        if "Status" in matched_df.columns else 0
    )
    c2.metric("PTP Records", ptp_count)

    # ── Transform → Deduplicate → Map → Slice ─
    transform_columns(matched_df)
    final_df = deduplicate_by_hierarchy(matched_df)
    render_status_summary(matched_df, final_df)

    mapped_df = build_mapped_df(final_df, account_type, dca_coordinator)
    daily_activity, daily_productivity = build_report_slices(mapped_df)

    render_report_previews(daily_activity, daily_productivity)
    render_downloads(daily_activity, daily_productivity, account_type)


if __name__ == "__main__":
    main()