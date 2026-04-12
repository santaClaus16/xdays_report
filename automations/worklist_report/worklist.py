import streamlit as st
import io
import pandas as pd
from openpyxl import load_workbook
import msoffcrypto

def worklist():
        
    # Password mapping: base_filename -> password
    PASSWORD_MAP = {
        "c9 Worklist 0410 - Xdays (result)": "CYCLE_9",
        "c14 Worklist 0410 - Xdays (result)": "CYCLE_14",
        # Add more: "filename": "password"
    }

    @st.cache_data
    def get_password(filename):
        base_name = filename.rsplit('.', 1)[0]  # Remove extension
        for key, pwd in PASSWORD_MAP.items():
            if key in base_name or base_name.startswith(key):
                return pwd
        return None

    def unlock_and_read(file_bytes, filename):
        password = get_password(filename)
        if not password:
            st.error(f"No password for {filename}")
            return None
        office_file = msoffcrypto.OfficeFile(io.BytesIO(file_bytes))
        office_file.load_key(password=password)
        decrypted = io.BytesIO()
        office_file.decrypt(decrypted)
        decrypted.seek(0)
        df = pd.read_excel(decrypted)
        return df

    def append_to_template(dfs, template_bytes):
        template_io = io.BytesIO(template_bytes)
        wb = load_workbook(template_io)
        ws = wb.active
        for df in dfs:
            for _, row in df.iterrows():
                ws.append(row.tolist())
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        return output.getvalue()

    st.title("Excel Append Automation")
    st.write("Upload password-protected Excel files. Appends data to template.")

    # Upload template (with formulas)
    template_file = st.file_uploader("Upload Template (with formulas)", type=["xlsx"], key="template")
    if template_file:
        template_bytes = template_file.read()
        st.success("Template loaded!")

    # Multiple uploads
    uploaded_files = st.file_uploader("Upload Data Files", type=["xlsx"], accept_multiple_files=True)

    if st.button("Process Files") and template_file and uploaded_files:
        dfs = []
        for f in uploaded_files:
            df = unlock_and_read(f.read(), f.name)
            if df is not None:
                dfs.append(df)
                st.success(f"Loaded {f.name}")
        
        if dfs:
            updated_bytes = append_to_template(dfs, template_bytes)
            st.download_button(
                "Download Updated Template",
                updated_bytes,
                "updated_template.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("No valid files processed.")