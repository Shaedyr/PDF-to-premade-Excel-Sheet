import streamlit as st
import pdfplumber
import re
from io import BytesIO

# ---------------------------------------------------------
# REGEX PATTERNS
# ---------------------------------------------------------

ORG_RE = re.compile(r"\b(\d{9})\b")
ORG_IN_TEXT_RE = re.compile(
    r"(organisasjonsnummer|org\.?nr|org nr|orgnummer)[:\s]*?(\d{9})",
    flags=re.I
)

COMPANY_WITH_SUFFIX_RE = re.compile(
    r"([A-ZÆØÅ][A-Za-zÆØÅæøå0-9.\-&\s]{1,120}?)\s+(AS|ASA|ANS|DA|ENK|KS|BA)\b",
    flags=re.I
)

POST_CITY_RE = re.compile(
    r"(\d{4})\s+([A-ZÆØÅa-zæøå\-\s]{2,50})"
)

ADDRESS_RE = re.compile(
    r"([A-ZÆØÅa-zæøå.\-\s]{3,60}\s+\d{1,4}[A-Za-z]?)"
)

REVENUE_RE = re.compile(
    r"omsetning\s*(?:2024)?[:\s]*([\d\s\.,]+(?:kr)?)",
    flags=re.I
)

DEADLINE_RE = re.compile(
    r"(?:anbudsfrist|frist)[:\s]*([0-3]?\d[./-][01]?\d[./-]\d{2,4})",
    flags=re.I
)

# ---------------------------------------------------------
# PDF TEXT EXTRACTION
# ---------------------------------------------------------

def extract_text_from_pdf(pdf_bytes: bytes) -> str:
    """
    Extracts text from the first 6 pages of a PDF.
    Returns a single string.
    """

    if not pdf_bytes:
        return ""

    try:
        text = ""
        with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages[:6]:
                extracted = page.extract_text()
                if extracted:
                    text += extracted + "\n"
        return text

    except Exception:
        return ""

# ---------------------------------------------------------
# FIELD EXTRACTION
# ---------------------------------------------------------

def extract_fields_from_pdf(pdf_bytes: bytes) -> dict:
    """
    Extracts useful fields from a PDF:
    - org number
    - company name
    - address
    - post nr + city
    - revenue
    - deadline
    """

    txt = extract_text_from_pdf(pdf_bytes)
    fields = {}

    if not txt:
        return fields

    # 1) Org number
    m = ORG_IN_TEXT_RE.search(txt)
    if m:
        fields["org_number"] = m.group(2)
    else:
        m2 = ORG_RE.search(txt)
        if m2:
            fields["org_number"] = m2.group(1)

    # 2) Company name
    m3 = COMPANY_WITH_SUFFIX_RE.search(txt)
    if m3:
        fields["company_name"] = m3.group(0).strip()
    else:
        # fallback: first title-cased line
        for line in txt.splitlines():
            line = line.strip()
            if len(line) > 3 and line == line.title():
                fields["company_name"] = line
                break

    # 3) Postnummer + city
    mpc = POST_CITY_RE.search(txt)
    if mpc:
        fields["post_nr"] = mpc.group(1)
        fields["city"] = mpc.group(2).strip()

    # 4) Address
    maddr = ADDRESS_RE.search(txt)
    if maddr:
        fields["address"] = maddr.group(1).strip()

    # 5) Revenue
    mrev = REVENUE_RE.search(txt)
    if mrev:
        fields["revenue_2024"] = mrev.group(1).strip()

    # 6) Deadline
    mdate = DEADLINE_RE.search(txt)
    if mdate:
        fields["tender_deadline"] = mdate.group(1).strip()

    return fields

# ---------------------------------------------------------
# PAGE VIEW (so it works as a selectable page)
# ---------------------------------------------------------
def run():
    st.title("📄 PDF Parser Module")
    st.write("Dette modulen ekstraherer tekst og felter fra PDF-dokumenter.")
    st.info("Brukes av hovedsiden for å hente data fra PDF.")
