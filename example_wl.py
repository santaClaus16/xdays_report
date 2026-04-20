import streamlit as st
import pandas as pd
import win32com.client as win32
import tempfile
import os
import time
import pythoncom  # ADD THIS
# Function to refresh Excel and extract ActiveQry
def refresh_and_extract(file_path):
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        wb = excel.Workbooks.Open(file_path)

        # Refresh all queries
        wb.RefreshAll()

        # Wait until refresh is done
        excel.CalculateUntilAsyncQueriesDone()

        # Small buffer (important for stability)
        time.sleep(5)

        # Access sheet
        sheet = wb.Sheets("ActiveQry")

        # Get used range
        data = sheet.UsedRange.Value

        # Convert to DataFrame
        df = pd.DataFrame(data)

        # Set first row as header
        df.columns = df.iloc[0]
        df = df[1:].reset_index(drop=True)

        # Clean empty columns
        df = df.loc[:, df.columns.notna()]

        wb.Close(SaveChanges=True)

    except Exception as e:
        wb.Close(SaveChanges=False)
        excel.Quit()
        raise e

    excel.Quit()
    return df



def refresh_and_extract(file_path):
    pythoncom.CoInitialize()  # ✅ Initialize COM

    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        wb = excel.Workbooks.Open(file_path)

        wb.RefreshAll()
        excel.CalculateUntilAsyncQueriesDone()
        time.sleep(5)

        sheet = wb.Sheets("ActiveQry")
        data = sheet.UsedRange.Value

        df = pd.DataFrame(data)
        df.columns = df.iloc[0]
        df = df[1:].reset_index(drop=True)
        df = df.loc[:, df.columns.notna()]

        wb.Close(SaveChanges=True)

    except Exception as e:
        wb.Close(SaveChanges=False)
        excel.Quit()
        pythoncom.CoUninitialize()  # ✅ Clean up
        raise e

    excel.Quit()
    pythoncom.CoUninitialize()  # ✅ Clean up
    return df
# Streamlit UI
st.title("📊 ActiveQry Automation")

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file:
    with st.spinner("Processing and refreshing queries..."):

        # Save uploaded file to temp location
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            tmp.write(uploaded_file.read())
            temp_path = tmp.name

        try:
            df = refresh_and_extract(temp_path)

            st.success("✅ File processed successfully!")

            # Show preview
            st.dataframe(df)

            # Download button
            output_file = "processed_activeqry.xlsx"
            df.to_excel(output_file, index=False)

            with open(output_file, "rb") as f:
                st.download_button(
                    label="📥 Download Result",
                    data=f,
                    file_name=output_file,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
            st.error(f"❌ Error: {str(e)}")

        finally:
            # Clean up temp file
            if os.path.exists(temp_path):
                os.remove(temp_path)