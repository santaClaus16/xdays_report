import streamlit as st
import pandas as pd
import msoffcrypto
import re
from io import BytesIO
from datetime import datetime
import pyodbc
import os
from dotenv import load_dotenv
import openpyxl

MAPPING_PATH_WL = r"C:\Users\SPM\Desktop\eod_report\mapping\mapping_file_wl.xlsx"
FILE_PATH = r"C:\Users\SPM\Desktop\eod_report\wl\data_ref.xlsx"

st.set_page_config(page_title="WORKLIST", layout="wide")

load_dotenv()
notif = st.empty()

# Database connection
try:
    conn_str = (
        f"DRIVER={{{os.getenv('DB_DRIVER')}}};"
        f"SERVER={os.getenv('DB_HOST')};"
        f"DATABASE={os.getenv('DB_NAME')};"
        f"UID={os.getenv('DB_USER')};"
        f"PWD={os.getenv('DB_PASS')};"
    )
    conn = pyodbc.connect(conn_str)
    notif.success("✅ Successfully connected to the database!")
except Exception as e:
    notif.error(f"❌ Failed to connect to the database: {e}")
    st.stop()

# Refresh Button
refresh_col, _ = st.columns([1, 3])
with refresh_col:
    if st.button("🔄 Refresh Data", type="primary"):
        st.cache_data.clear()

# Load master sheets function
@st.cache_data(show_spinner=False)
def load_master_sheets():
    df_master_local = {}
    if os.path.exists(FILE_PATH):
        with st.spinner("Fetching master file..."):
            try:
                xl = pd.ExcelFile(FILE_PATH, engine="openpyxl")
                for sheet_name in xl.sheet_names:
                    df = xl.parse(sheet_name)
                    df_master_local[sheet_name] = df
            except Exception as e:
                notif.error(f"Error fetching file: {e}")
                st.stop()
        notif.success("✅ Master data fetched successfully!")
    return df_master_local

# Query
query_bcrm = """
SELECT
    `leads`.`leads_chcode` AS 'CHCODE',
    `leads`.`leads_chname` AS 'CH NAME',
    `leads`.`leads_acctno` AS 'ACCOUNT NUMBER',
    leads.`leads_endo_date` AS 'ENDO DATE',
    `leads_status`.`leads_status_name` AS 'STATUS',
    `leads_substatus`.`leads_substatus_name` AS 'SUBSTATUS',
    `leads_result`.`leads_result_sdate` AS 'START DATE',
    `leads_result`.`leads_result_edate` AS 'END DATE',
    `leads_result`.`leads_result_comment` AS 'NOTES',
    `leads_result`.`leads_result_ts` AS 'RESULT DATE',
    `leads_result`.`leads_result_id` AS 'RESULT ID',
    `users`.`users_username` AS 'AGENT',
    users.users_name AS 'AGENT NAME'
FROM `bcrm`.`leads_result`
LEFT JOIN `bcrm`.`leads` ON `leads_result`.`leads_result_lead` = `leads`.`leads_id`
LEFT JOIN `bcrm`.`client` ON `leads`.`leads_client_id` = `client`.`client_id`
LEFT JOIN `bcrm`.`users` ON `leads_result`.`leads_result_users` = `users`.`users_id`
LEFT JOIN `bcrm`.`users` AS leads_users ON leads_users.`users_id` = leads.`leads_users_id`
LEFT JOIN `bcrm`.`leads_status` ON `leads_result`.`leads_result_status_id` = `leads_status`.`leads_status_id`
LEFT JOIN `bcrm`.`leads_substatus` ON `leads_result`.`leads_result_substatus_id` = `leads_substatus`.`leads_substatus_id`
WHERE `client`.`client_id` = '174' 
AND `leads_users`.`users_username` != 'POUT' 
AND `leads_result`.`leads_result_hidden` <> 1 
AND `leads_status`.`leads_status_name` <> "LETTER SENT"
AND DATE_FORMAT(`leads_result`.`leads_result_barcode_date`, '%Y-%m-%d') BETWEEN (CURDATE() - INTERVAL 16 DAY) AND CURDATE()
ORDER BY leads_result.`leads_result_barcode_date` DESC
"""

@st.cache_data(show_spinner=False)
def fetch_active_queries(_conn, query):
    with st.spinner("Fetching active queries..."):
        try:
            res = pd.read_sql(query, _conn)
        except Exception as e:
            notif.error(f"Error fetching metrics: {e}")
            st.stop()
        notif.success("✅ Active queries fetched successfully!")
        return res

# Fetch active queries
active_qry = fetch_active_queries(conn, query_bcrm)

# Load master data
df_master = load_master_sheets()

# Extract sheets safely
def get_df_by_name(master_dict, name, default=None):
    if master_dict is None:
        return default
    return master_dict.get(name, default)

df_active = get_df_by_name(df_master, "ACTIVE_UPDATE", pd.DataFrame())
df_pullout = get_df_by_name(df_master, "PULLOUT", pd.DataFrame())
df_dnc = get_df_by_name(df_master, "DNC", pd.DataFrame())

# Upload section (unchanged)
with st.form("Upload Files", border=False):
    upload_overall_wl = st.file_uploader(
        "Upload your overall files here",
        type=["xlsx"],
        key="overall_wl"
    )
    submit_btn = st.form_submit_button("Process", type="primary", use_container_width=True)


with st.expander("Overall data", expanded=True):
    tab1, tab2,tab3, tab4 = st.tabs(["Active", "Pull out", "DNC", "Active Query"])
    # Display sheets
    with tab1:
        st.subheader("📄 ACTIVE UPDATE")
        st.dataframe(df_active, use_container_width=True)
    with tab2:
        st.subheader("📄 PULLOUT")
        st.dataframe(df_pullout, use_container_width=True)
    with tab3:
        st.subheader("📄 DNC")
        st.dataframe(df_dnc, use_container_width=True)
    with tab4:
        # Show active queries
        st.subheader("📊 Active Queries")
        st.dataframe(active_qry, use_container_width=True)

