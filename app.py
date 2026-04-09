import streamlit as st
import webbrowser
from automations.drr_cleaning.drr import drr_cleaning_app;
from automations.daily_remarks_report.drr_automate import drr_automate;



automates = {
    "DRR CLEANER":drr_cleaning_app,
    "DRR AUTOMATE": drr_automate

}

default_app = "DRR CLEANER"

st.set_page_config(
    page_title="Auto Status",
    page_icon="🛠️",
    layout="wide",
)

# Initialize toggle state
if "show_links" not in st.session_state:
    st.session_state.show_links = True

with st.sidebar:
    st.title("🛠️ Automation Hub")
    st.caption("Select what to display")

    selected = st.selectbox(
        "Automation to show",
        options=list(automates.keys()),
        index=list(automates.keys()).index(default_app),
        key="selected_automation_selectbox"
    )


automates[selected]()