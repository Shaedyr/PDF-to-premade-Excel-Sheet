import streamlit as st
import pandas as pd

TEMPLATE_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQZgo_lI3n1uTuOz6DzJnKUU--_Cs991MzQ_NNtkqxUmEq5k8W6Qki_O0hwngLVxHoD9GcAxRG-mq7w/pub?output=xlsx"

def load_template():
    try:
        df = pd.read_excel(TEMPLATE_URL)
        st.session_state["template_df"] = df
        st.success("Excel template loaded from Google Sheets")
        return df
    except Exception as e:
        st.error(f"Could not load Excel template: {e}")
        st.stop()

def run():
    st.title("📁 Template Loader")
    st.write("This module loads the Excel template directly from Google Sheets.")
    st.info("Used by the main page to fetch the template.")
