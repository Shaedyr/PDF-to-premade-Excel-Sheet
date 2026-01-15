import streamlit as st
import requests
import re


def _clean_text(t: str) -> str:
    """Remove weird whitespace and shorten long text."""
    if not t:
        return ""
    t = re.sub(r"\s+", " ", t).strip()
    return t[:600]  # keep it readable


# ---------------------------------------------------------
# 1) Brønnøysund-based summary (most reliable)
# ---------------------------------------------------------
def summary_from_brreg(data: dict) -> str:
    """
    Creates a simple summary using Brønnøysund data.
    Always safe and predictable.
    """

    if not data:
        return ""

    name = data.get("company_name", "")
    employees = data.get("employees", "")
    reg = data.get("registration_date", "")
    nace = data.get("nace_description", "")

    parts = []

    if name:
        parts.append(f"{name} er et registrert norsk selskap.")

    if reg:
        parts.append(f"Selskapet ble registrert i {reg}.")

    if employees:
        try:
            e = int(employees)
            if e > 200:
                parts.append(f"Det er en større arbeidsgiver med {e} ansatte.")
            elif e > 50:
                parts.append(f"Det er et mellomstort foretak med {e} ansatte.")
            else:
                parts.append(f"Selskapet har {e} ansatte.")
        except Exception:
            pass

    if nace:
        parts.append(f"Virksomheten opererer innen bransjen: {nace}.")

    return " ".join(parts)


# ---------------------------------------------------------
# 2) Wikipedia summary (if available)
# ---------------------------------------------------------
def summary_from_wikipedia(name: str) -> str:
    """
    Attempts to fetch a short summary from Wikipedia.
    Returns empty string if not found.
    """

    if not name:
        return ""

    try:
        url = f"https://no.wikipedia.org/api/rest_v1/page/summary/{name}"
        r = requests.get(url, timeout=10)

        if r.status_code == 200:
            data = r.json()
            extract = data.get("extract", "")
            return _clean_text(extract)

    except Exception:
        pass

    return ""


# ---------------------------------------------------------
# 3) DuckDuckGo fallback summary
# ---------------------------------------------------------
def summary_from_duckduckgo(query: str) -> str:
    """
    Fallback summary using DuckDuckGo Instant Answer API.
    """

    if not query:
        return ""

    try:
        url = "https://api.duckduckgo.com/"
        r = requests.get(url, params={"q": query, "format": "json"}, timeout=10)

        if r.status_code == 200:
            abstract = r.json().get("AbstractText", "")
            return _clean_text(abstract)

    except Exception:
        pass

    return ""


# ---------------------------------------------------------
# MASTER FUNCTION — used by the app
# ---------------------------------------------------------
def generate_company_summary(company_data: dict) -> str:
    """
    Attempts summary in this order:
    1) Brønnøysund-based summary
    2) Wikipedia summary
    3) DuckDuckGo summary
    4) Fallback to Brønnøysund summary again
    """

    # 1) Brønnøysund summary
    base = summary_from_brreg(company_data)
    if len(base) > 40:
        return base

    # 2) Wikipedia
    name = company_data.get("company_name", "")
    wiki = summary_from_wikipedia(name)
    if len(wiki) > 40:
        return wiki

    # 3) DuckDuckGo
    ddg = summary_from_duckduckgo(name)
    if len(ddg) > 40:
        return ddg

    # 4) Fallback
    return base or "Ingen tilgjengelig selskapsbeskrivelse."


# ---------------------------------------------------------
# PAGE VIEW (so it works as a selectable page)
# ---------------------------------------------------------
def run():
    st.title("📝 Summary Module")
    st.write("Dette modulen genererer selskapsbeskrivelser fra flere kilder.")
    st.info("Brukes av hovedsiden for å lage 'Om oss'‑tekst.")
