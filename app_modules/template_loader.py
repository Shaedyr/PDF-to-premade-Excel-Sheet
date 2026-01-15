import streamlit as st
import pandas as pd

# Direct OneDrive download link (your working link)
TEMPLATE_URL = "https://onedrive.live.com/download?resid=F5E2800FEEB07258!sd128a89fc2ac495d9e1ea17ad5de4f00&authkey=!K0HaUp"


def load_template():
    """
    Downloads the Excel template directly from OneDrive
    and stores it in session_state.
    """

    try:
        df = pd.read_excel(TEMPLATE_URL)
        st.session_state["template_df"] = df
        st.success("Excel-mal lastet ned fra OneDrive")
        return df

    except Exception as e:
        st.error(f"Kunne ikke laste Excel-malen: {e}")
        st.stop()


def run():
    st.title("📁 Template Loader")
    st.write("Denne modulen laster Excel-malen direkte fra OneDrive.")
    st.info("Brukes av hovedsiden for å hente Excel-malen.")
