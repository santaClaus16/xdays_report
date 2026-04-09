import streamlit as st
import os
import pandas as pd

st.set_page_config(
    page_title="DRR Automation", 
    page_icon="🚀",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("🚀 Daily Remarks Report Automation")
st.markdown("Professional DRR processing tool with highlighting and export features.")

# Sidebar for inputs
st.sidebar.header("📁 Input Selection")
clean_drr_path = r"C:\Users\SOLIZA\Documents\secret\clean_drr"

try:
    files = [f for f in os.listdir(clean_drr_path) if f.lower().endswith(('.xlsx', '.xls'))]
    if not files:
        st.sidebar.error("No Excel files in clean_drr/")
        st.stop()
    
    selected_file = st.sidebar.selectbox("Select file", files, key="file_sel")
    file_path = os.path.join(clean_drr_path, selected_file)
    
    xl = pd.ExcelFile(file_path)
    sheet_names = xl.sheet_names
    selected_sheet = st.sidebar.selectbox("Select sheet", sheet_names, key="sheet_sel")
    
    if st.sidebar.button("⚡ Load & Process", type="primary", use_container_width=True):
        with st.spinner("Processing..."):
            df = pd.read_excel(file_path, sheet_name=selected_sheet)
            
            original_cols = set(df.columns)
            processed = False
            
            if "DAILY REMARKS REPORT | SYSTEM" in selected_sheet:
                if "Balance" in df.columns:
                    df["AMOUNT FIX"] = df["Balance"]
                if "Card No." in df.columns:
                    df["CYCLE"] = "Cycle " + df["Card No."].astype(str).str[:2]
                processed = True
                st.sidebar.success("✅ SYSTEM processed")
            
            elif "DAILY REMARKS REPORT | AGENT" in selected_sheet:
                first_col = df.columns[0]
                if "Card No." in df.columns:
                    df["CYCLE"] = "Cycle " + df["Card No."].astype(str).str[:2]
                st.sidebar.success(f"✅ AGENT: First col '{first_col}' + CYCLE")
                processed = True
            
            # Main preview
            st.subheader("📊 Processed Preview")
            new_cols = set(df.columns) - original_cols
            
            def highlight_processed(x):
                styles = []
                for col in x.index:
                    if col in new_cols:
                        styles.append('background-color: #4CAF50; color: white; font-weight: bold')
                    elif "AGENT" in selected_sheet and col == df.columns[0]:
                        styles.append('background-color: #FFEB3B; color: black; font-weight: bold')
                    else:
                        styles.append('')
                return styles
            
            styled_df = df.style.apply(highlight_processed, axis=1)
            st.dataframe(styled_df, width='stretch')
            
            st.metric("Rows", len(df), delta=None)
            st.metric("Columns", len(df.columns), delta=len(new_cols))
            
            # Download & Save
            col1, col2, col3 = st.columns(3)
            with col1:
                csv = df.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="📥 Download CSV",
                    data=csv,
                    file_name=f"{selected_file}_{selected_sheet}.csv",
                    mime="text/csv"
                )
            with col2:
                if st.button("💾 Save Excel", type="secondary"):
                    output_dir = "automations/daily_remarks_report/output"
                    os.makedirs(output_dir, exist_ok=True)
                    output_path = os.path.join(output_dir, f"processed_{selected_sheet.replace('|', '_')}_{os.path.splitext(selected_file)[0]}.xlsx")
                    df.to_excel(output_path, index=False)
                    st.success(f"✅ Saved: {output_path}")
                    st.balloons()
            with col3:
                st.markdown("**[View Output Folder](automations/daily_remarks_report/output/)**")
            
except Exception as e:
    st.error(f"❌ {str(e)}")
    st.stop()

st.markdown("---")
st.markdown("*Professional DRR Automation - BLACKBOXAI*")
