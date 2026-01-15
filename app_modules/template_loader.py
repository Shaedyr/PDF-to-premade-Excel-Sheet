import streamlit as st
from app_modules.graph_client import download_onedrive_file


CLOUD_TEMPLATE_PATH = "/ExcelTemplates/PremadeExcelTemplate.xlsx"


def load_template_secure():
    """
    Loads the Excel template from the user's OneDrive using Microsoft Graph.
    Requires the user to be logged in.
    """

    st.info("🔄 Henter Excel-mal fra OneDrive (sikker tilgang)...")

    try:
        content = download_onedrive_file(CLOUD_TEMPLATE_PATH)

        # Validate Excel file (Excel files start with PK)
        if len(content) < 50 or content[:2] != b"PK":
            raise Exception("Filen som ble hentet er ikke en gyldig Excel-fil.")

        st.success("✅ Excel-mal lastet ned via Microsoft Graph!")
        return content

    except Exception as e:
        st.error(f"❌ Kunne ikke hente Excel-malen: {e}")
        st.stop()


def run():
    st.title("📁 Template Loader (Secure)")
    st.write("Denne modulen laster Excel-malen fra OneDrive via Microsoft Graph.")
    st.info("Brukes av hovedsiden for å hente Excel-malen.")
