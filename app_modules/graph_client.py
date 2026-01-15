import streamlit as st
import requests


GRAPH_BASE = "https://graph.microsoft.com/v1.0"


def _get_access_token():
    """
    Returns the access token stored after login.
    Raises a clear error if the user is not logged in.
    """
    token_data = st.session_state.get("token")

    if not token_data or "access_token" not in token_data:
        st.error("Du må logge inn med Microsoft før du kan hente filer.")
        st.stop()

    return token_data["access_token"]


def download_onedrive_file(cloud_path: str) -> bytes:
    """
    Downloads a file from the user's personal OneDrive using Microsoft Graph.

    Example cloud_path:
        '/ExcelTemplates/PremadeExcelTemplate.xlsx'
    """

    access_token = _get_access_token()

    # Build Graph API URL
    url = f"{GRAPH_BASE}/me/drive/root:{cloud_path}:/content"

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Accept": "*/*"
    }

    with st.spinner("🔄 Laster ned fil fra OneDrive..."):
        response = requests.get(url, headers=headers)

    if response.status_code != 200:
        st.error(f"Kunne ikke hente filen fra OneDrive. HTTP {response.status_code}")
        st.stop()

    return response.content


def run():
    """
    Optional page view so this module can appear in the sidebar.
    """
    st.title("☁️ Microsoft Graph Client")
    st.write("Dette modulen håndterer nedlasting av filer fra OneDrive via Microsoft Graph.")
    st.info("Brukes av Template Loader for å hente Excel-malen.")
