import streamlit as st
import requests

BRREG_SEARCH_URL = "https://data.brreg.no/enhetsregisteret/api/enheter"
BRREG_ENTITY_URL = "https://data.brreg.no/enhetsregisteret/api/enheter/{}"


# ---------------------------------------------------------
# LIVE SEARCH
# ---------------------------------------------------------
def search_brreg_live(name: str):
    """
    Live search for companies in Brønnøysund.
    Returns a list of raw API objects.
    """

    name = (name or "").strip()
    if len(name) < 2:
        return []

    try:
        r = requests.get(
            BRREG_SEARCH_URL,
            params={"navn": name, "size": 10},
            timeout=10
        )
        r.raise_for_status()

        data = r.json()
        return data.get("_embedded", {}).get("enheter", []) or []

    except Exception:
        return []


# ---------------------------------------------------------
# FETCH FULL COMPANY DATA
# ---------------------------------------------------------
def fetch_company_by_org(org_number: str):
    """
    Fetch full company details using org number.
    Returns raw API JSON or None.
    """

    org_number = (org_number or "").strip()
    if not org_number.isdigit():
        return None

    try:
        r = requests.get(
            BRREG_ENTITY_URL.format(org_number),
            timeout=10
        )
        r.raise_for_status()
        return r.json()

    except Exception:
        return None


# ---------------------------------------------------------
# FORMAT RAW API DATA INTO CLEAN DICT
# ---------------------------------------------------------
def format_company_data(api_data):
    """
    Converts raw Brønnøysund API data into a clean dictionary
    that the rest of the app can use.
    """

    if not api_data:
        return {}

    out = {
        "company_name": api_data.get("navn", ""),
        "org_number": api_data.get("organisasjonsnummer", ""),
        "homepage": api_data.get("hjemmeside", ""),
        "employees": api_data.get("antallAnsatte", ""),
        "registration_date": api_data.get("stiftelsesdato", ""),
        "nace_code": "",
        "nace_description": "",
        "address": "",
        "post_nr": "",
        "city": "",
    }

    # Address
    addr = api_data.get("forretningsadresse") or {}
    if addr:
        a = addr.get("adresse", [])
        out["address"] = ", ".join(a) if isinstance(a, list) else a
        out["post_nr"] = addr.get("postnummer", "")
        out["city"] = addr.get("poststed", "")

    # NACE
    nace = api_data.get("naeringskode1") or {}
    if nace:
        out["nace_code"] = nace.get("kode", "")
        out["nace_description"] = nace.get("beskrivelse", "")

    return out


# ---------------------------------------------------------
# OPTIONAL DEBUG PAGE
# ---------------------------------------------------------
def run():
    st.title("🔍 Company Data Module")
    st.write("Dette er et backend-modul og brukes av andre sider.")
    st.info("Ingen interaktiv funksjon her. Brukes av Input- og Summary-moduler.")
