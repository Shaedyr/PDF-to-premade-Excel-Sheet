import requests
import streamlit as st
from io import BytesIO

# Your OneDrive share link (replace with your own)
ONEDRIVE_URL = "https://onedrive.live.com/:x:/g/personal/F5E2800FEEB07258/IQBBPI2scMXjQ6bi18LIvXFGASbYGUC3UFrRl-H3_zYplRI?resid=F5E2800FEEB07258!sac8d3c41c57043e3a6e2d7c2c8bd7146&ithint=file%2Cxlsx&e=DY54v7&migratedtospo=true&redeem=aHR0cHM6Ly8xZHJ2Lm1zL3gvYy9mNWUyODAwZmVlYjA3MjU4L0lRQkJQSTJzY01YalE2YmkxOExJdlhGR0FTYllHVUMzVUZyUmwtSDNfellwbFJJP2U9RFk1NHY3&download=1"


def _convert_onedrive_url(url: str) -> str:
    """
    Converts a public OneDrive share link into a direct download link.
    Handles both OneDrive and SharePoint redirects.
    """

    try:
        session = requests.Session()
        response = session.get(url, allow_redirects=True, timeout=30)
        final_url = response.url

        # If redirected to SharePoint
        if "sharepoint.com" in final_url:
            # Convert /forms/ link to /download?download=1
            if "/forms/" in final_url:
                return final_url.replace("/forms/", "/download?") + "&download=1"
            return final_url + "?download=1"

        # If still a OneDrive short link
        if "1drv.ms" in url:
            # Extract share ID
            import re
            match = re.search(r'/([A-Za-z0-9_\-!]+)$', url)
            if match:
                share_id = match.group(1)
                encoded = requests.utils.quote(share_id, safe='')
                return f"https://api.onedrive.com/v1.0/shares/u!{encoded}/root/content"

        # Fallback
        return url + "?download=1"

    except Exception:
        return url + "?download=1"


def load_template_from_onedrive():
    """
    Downloads the Excel template from OneDrive.
    Ensures the file is valid and returns raw bytes.
    """

    st.info("🔄 Henter Excel-mal fra OneDrive...")

    download_url = _convert_onedrive_url(ONEDRIVE_URL)

    headers = {
        "User-Agent": "Mozilla/5.0",
        "Accept": "*/*"
    }

    try:
        response = requests.get(download_url, headers=headers, timeout=60)

        if response.status_code != 200:
            raise Exception(f"HTTP {response.status_code}")

        content = response.content

        # Excel files are ZIP files starting with PK
        if len(content) < 50 or content[:2] != b"PK":
            raise Exception("Filen som ble lastet ned er ikke en gyldig Excel-fil.")

        st.success("✅ Excel-mal lastet fra OneDrive!")
        return content

    except Exception as e:
        st.error(f"❌ Kunne ikke laste Excel-malen fra OneDrive: {e}")
        st.stop()
