import streamlit as st
import pandas as pd
import msoffcrypto
import re
from io import BytesIO
from datetime import datetime
import os
import shutil

from utils import (
    load_mapping_file,
    apply_header_mapping,
    clean_account_number_column,
    format_date_columns,
    force_columns_to_str,
)

# -------------------------------
# SESSION STATE INIT
# -------------------------------
if "decrypted_files" not in st.session_state:
    st.session_state.decrypted_files = {}

def decrypt_file(file_bytes, password):
    """Helper function to decrypt msoffcrypto file"""
    try:
        decrypted = BytesIO()
        office_file = msoffcrypto.OfficeFile(BytesIO(file_bytes))
        office_file.load_key(password=password)
        office_file.decrypt(decrypted)
        decrypted.seek(0)
        return decrypted.getvalue(), None
    except msoffcrypto.exceptions.DecryptionError:
        return None, "Wrong password"
    except Exception as e:
        return None, str(e)

def process_dataframe(decrypted_bytes, c_number, rules, standard_order):
    """Process the decrypted dataframe according to mapping and cleaning rules."""
    df = pd.read_excel(BytesIO(decrypted_bytes), dtype=str, engine="openpyxl")

    WL_STR_COLS = ["CUST_ID", "OFFICE_PH", "HOME_PH", "MOBILE_NO", "OB", "BOS", "AOD", "MAD", "PDA", "LPA", "PTP_AMT"]
    df = force_columns_to_str(df, WL_STR_COLS)

    WL_DATE_COLS = ["LAST_PAYMENT_DATE", "PTP_DATE", "BIRTHDATE", "LAST_DUE_DATE", "D_CUST_OPN", "Birthdate"]
    df = format_date_columns(df, date_cols=WL_DATE_COLS)

    df = clean_account_number_column(df, show_success=False, st=st)

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

    return df

def wl_v2():
    st.title("📁 Worklist Processor v2")
    st.markdown("Automated worklist decryption, mapping, and consolidation.")

    # ================== Constants & Mapping ==================
    MAPPING_PATH_WL = r"C:\Users\SPM\Desktop\eod_report\mapping\mapping_file_wl.xlsx"
    OUTPUT_FOLDER = r"C:\Users\SPM\Desktop\eod_report\clean_wl"

    rules, load_logs, standard_order = load_mapping_file(MAPPING_PATH_WL)

    uploaded_files = st.file_uploader("UPLOAD WL FILES", type=["xlsx", "xls"], accept_multiple_files=True)
    template_file = st.file_uploader("UPLOAD TEMPLATE (Optional)", type=["xlsx", "xls"])

    if not uploaded_files:
        st.info("👆 Upload your password-protected worklist files to begin.")
        st.session_state.decrypted_files.clear() # Reset on new batch
        return

    # Track files
    files_ready = []
    files_needing_password = []

    # 1. First Pass - Attempt Auto Decryption
    for uploaded_file in uploaded_files:
        filename = uploaded_file.name
        
        match = re.search(r'[cC](\d+)', filename)
        if not match:
            st.error(f"❌ Could not find cycle number in: **{filename}**")
            continue
            
        c_number = match.group(1)
        file_bytes = uploaded_file.read()
        
        # Check if already decrypted
        if filename in st.session_state.decrypted_files:
            files_ready.append((filename, c_number))
            continue

        auto_password = f"CYCLE_{c_number}*"
        decrypted_bytes, error = decrypt_file(file_bytes, auto_password)
        
        if decrypted_bytes:
            st.session_state.decrypted_files[filename] = decrypted_bytes
            files_ready.append((filename, c_number))
        else:
            files_needing_password.append((filename, c_number, file_bytes))

    # 2. Handle Manual Passwords if any failed
    if files_needing_password:
        st.warning(f"⚠️ {len(files_needing_password)} file(s) require manual passwords.")
        
        with st.form("manual_passwords_form", border=True):
            st.subheader("🔓 Enter Passwords")
            manual_inputs = {}
            for filename, c_number, _ in files_needing_password:
                manual_inputs[filename] = st.text_input(
                    f"Password for {filename} (Cycle {c_number}):", 
                    type="password"
                )
                
            if st.form_submit_button("Decrypt & Continue", type="primary"):
                for filename, c_number, file_bytes in files_needing_password:
                    pw = manual_inputs[filename]
                    if pw:
                        decrypted_bytes, error = decrypt_file(file_bytes, pw)
                        if decrypted_bytes:
                            st.session_state.decrypted_files[filename] = decrypted_bytes
                            st.toast(f"✅ Decrypted {filename}")
                        else:
                            st.error(f"❌ Incorrect password for {filename}")
                st.rerun()
        return # Block further processing until all files are decrypted

    # 3. All files decrypted - Proceed to processing
    st.success(f"✅ All {len(files_ready)} file(s) successfully decrypted! Ready to process.")
    
    if st.button("🚀 Process & Generate 1-File Output", type="primary", use_container_width=True):
        with st.spinner("Processing data..."):
            all_dfs = []
            results = []
            
            progress_bar = st.progress(0)
            for idx, (filename, c_number) in enumerate(files_ready):
                decrypted_bytes = st.session_state.decrypted_files[filename]
                
                try:
                    df = process_dataframe(decrypted_bytes, c_number, rules, standard_order)
                    results.append({"filename": filename, "c_number": c_number, "dataframe": df})
                    all_dfs.append(df)
                except Exception as e:
                    st.error(f"❌ Error processing {filename}: {str(e)}")
                    
                progress_bar.progress((idx + 1) / len(files_ready))
            
            if not all_dfs:
                st.error("No data could be processed.")
                return

            # Combine data
            combined_df = pd.concat(all_dfs, ignore_index=True)
            timestamp = datetime.now().strftime("%m%d_%H%M")
            
            # Generate single output file containing both COMBINED and INDIVIDUAL sheets
            output_filename = f"WORKLIST_CONSOLIDATED_{timestamp}.xlsx"
            
            # Paths
            worklist_dir = os.path.join(OUTPUT_FOLDER, "WORKLIST")
            overall_cycle_dir = os.path.join(OUTPUT_FOLDER, "Overall_Cycle")
            os.makedirs(worklist_dir, exist_ok=True)
            os.makedirs(overall_cycle_dir, exist_ok=True)
            
            output_path = os.path.join(worklist_dir, output_filename)
            overall_cycle_path = os.path.join(overall_cycle_dir, output_filename)

            # Write to Excel with multiple sheets
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                combined_df.to_excel(writer, sheet_name='COMBINED_DATA', index=False)
                for res in results:
                    sheet_name = f"CYCLE_{res['c_number']}"[:31] # Max 31 chars
                    res['dataframe'].to_excel(writer, sheet_name=sheet_name, index=False)
            
            # Copy to Overall Cycle folder
            shutil.copy2(output_path, overall_cycle_path)

            if template_file:
                try:
                    # Read template headers
                    template_df = pd.read_excel(template_file, sheet_name="OVERALL_CYCLE")
                    template_headers = template_df.columns.tolist()
                    
                    # Create mapped dataframe
                    mapped_data = {}
                    for header in template_headers:
                        if header in combined_df.columns:
                            mapped_data[header] = combined_df[header]
                        else:
                            mapped_data[header] = [None] * len(combined_df)
                            
                    mapped_df = pd.DataFrame(mapped_data)
                    
                    template_output_filename = f"OVERALL_CYCLE_TEMPLATE_{timestamp}.xlsx"
                    template_output_path = os.path.join(overall_cycle_dir, template_output_filename)
                    
                    import openpyxl
                    from openpyxl.utils.dataframe import dataframe_to_rows
                    
                    template_file.seek(0)
                    wb = openpyxl.load_workbook(template_file)
                    
                    if "OVERALL_CYCLE" in wb.sheetnames:
                        ws = wb["OVERALL_CYCLE"]
                        
                        # Append mapped rows
                        for r in dataframe_to_rows(mapped_df, index=False, header=False):
                            ws.append(r)
                            
                        # Save the updated workbook
                        wb.save(template_output_path)
                        st.success(f"✅ Template mapped and saved: {template_output_path}")
                    else:
                        st.error("❌ 'OVERALL_CYCLE' sheet not found in the template.")
                except Exception as e:
                    st.error(f"❌ Error processing template: {str(e)}")
            
            st.success("🎉 Processing Complete!")
            st.info(f"📄 Generated 1 file containing both Combined & Individual Cycle data.")
            st.code(f"Output Path 1: {output_path}\nOutput Path 2: {overall_cycle_path}")
            
            # Provide download button for convenience
            with open(output_path, "rb") as file:
                st.download_button(
                    label="⬇️ Download Excel Report",
                    data=file,
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

if __name__ == "__main__":
    wl_v2()
