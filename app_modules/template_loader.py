import streamlit as st
import requests

TEMPLATE_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQZgo_lI3n1uTuOz6DzJnKUU--_Cs991MzQ_NNtkqxUmEq5k8W6Qki_O0hwngLVxHoD9GcAxRG-mq7w/pub?output=xlsx"

def load_template():
    try:
        response = requests.get(TEMPLATE_URL)
        response.raise_for_status()

        template_bytes = response.content  # <-- THIS is what fill_excel needs

        st.session_state["template_bytes"] = template_bytes
        st.success("Excel template loaded from Google Sheets")
        return template_bytes

    except Exception as e:
        st.error(f"Could not load Excel template: {e}")
        st.stop()
