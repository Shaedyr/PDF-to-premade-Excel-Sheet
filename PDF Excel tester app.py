import os
import re
import requests
import wikipedia
import streamlit as st
import pandas as pd
import pdfplumber
import difflib
from io import BytesIO
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Alignment


# =========================
# CONFIGURATION
# =========================
st.set_page_config(page_title="PDF → Excel (Brønnøysund)", layout="wide", page_icon="📊")
for k, v in {"extracted_data": {}, "api_response": None, "excel_ready": False, "company_summary": ""}.items():
    if k not in st.session_state:
        st.session_state[k] = v

# Keep existing session keys used by UI
for k in ('selected_company_data', 'companies_list', 'current_search', 'last_search', 'show_dropdown'):
    if k not in st.session_state:
        st.session_state[k] = None if k == 'selected_company_data' else [] if k == 'companies_list' else "" if k in (
            'current_search', 'last_search') else False

# =========================
# WIKIPEDIA / WEB SEARCH HELPERS
# =========================
def _strip_suffix(name: str):
    return re.sub(r'\b(AS|ASA|ANS|DA|ENK|KS|BA)\b\.?$', '', (name or ''), flags=re.I).strip()


def _wiki_summary(name: str):
    if not name:
        return None
    base = _strip_suffix(name)
    attempts = [base, name, base + " (bedrift)", base + " (company)"]
    for lang in ("no", "en"):
        try:
            wikipedia.set_lang(lang)
        except Exception:
            pass
        for a in attempts:
            try:
                results = wikipedia.search(a)
                if not results:
                    continue
                target = next((r for r in results[:5] if any(
                    w in r.lower() for w in ("as", "asa", "bedrift", "selskap", "company", "group"))), results[0])
                page = wikipedia.page(target, auto_suggest=False)
                s = page.summary or ""
                sent = [x.strip() for x in s.split('. ') if x.strip()]
                if len(sent) > 2:
                    return '. '.join(sent[:2]) + '.'
                short = s[:300] + '...' if len(s) > 300 else s
                return short if lang == "no" else short.replace(" is a ", " er et ").replace(" company",
                                                                                             " selskap").replace(
                    " based in ", " med hovedkontor i ")
            except (wikipedia.exceptions.DisambiguationError, wikipedia.exceptions.PageError):
                continue
            except Exception:
                continue
    return None


def _web_summary(name: str):
    if not name:
        return None
    q = _strip_suffix(name)
    for term in ("bedrift", "company"):
        try:
            url = f"https://api.duckduckgo.com/?q={requests.utils.quote(q + ' ' + term)}&format=json&no_html=1&skip_disambig=1"
            r = requests.get(url, timeout=10)
            if r.status_code == 200:
                txt = r.json().get("AbstractText", "") or ""
                if len(txt) > 50:
                    return txt if term == "bedrift" else txt.replace(" is a ", " er et ").replace(" company",
                                                                                                  " selskap").replace(
                        " based in ", " med hovedkontor i ")
        except Exception:
            continue
    return None


def create_summary_from_brreg_data(d: dict):
    name = d.get("company_name", "")
    if not name:
        return "Ingen informasjon tilgjengelig om dette selskapet."
    parts = []
    industry, city, emp, reg = d.get("nace_description", ""), d.get("city", ""), d.get("employees", ""), d.get(
        "registration_date", "")
    if industry and city:
        parts.append(f"{name} driver {industry.lower()} virksomhet fra {city}.")
    elif industry:
        parts.append(f"{name} opererer innen {industry.lower()}.")
    else:
        parts.append(f"{name} er et registrert norsk selskap.")
    if reg:
        try:
            year = int(reg.split('-')[0]) if '-' in reg else int(reg)
            years_old = datetime.now().year - year
            if years_old > 30:
                parts.append(f"Etablert i {year}, har selskapet over {years_old} års bransjeerfaring.")
            elif years_old > 10:
                parts.append(f"Selskapet har utviklet seg over {years_old} år siden etableringen i {year}.")
            else:
                parts.append(f"Etablert i {year}, er dette et yngre selskap i vekstfasen.")
        except Exception:
            parts.append(f"Selskapet ble registrert i {reg}.")
    if emp:
        try:
            emp_count = int(emp)
            if emp_count > 200:
                parts.append(
                    f"Som en større arbeidsgiver med {emp_count} ansatte, har det betydelig samfunnspåvirkning.")
            elif emp_count > 50:
                parts.append(f"Med {emp_count} ansatte representerer det et mellomstort foretak.")
            elif emp_count > 10:
                parts.append(f"Selskapet sysselsetter {emp_count} personer.")
            else:
                parts.append(f"Dette er et mindre selskap med {emp_count} ansatte.")
        except:
            pass
    if len(parts) < 3:
        parts.append("Virksomheten er registrert og i god stand i Brønnøysundregistrene.")
    summary = ' '.join(parts)
    return summary[:797] + "..." if len(summary) > 800 else summary


# =========================
# BRØNNØYSUND LIVE SEARCH (unchanged)
# =========================
@st.cache_data(ttl=3600)
def search_companies_live(name: str):
    if not name or len(name.strip()) < 2:
        return []
    try:
        r = requests.get("https://data.brreg.no/enhetsregisteret/api/enheter",
                         params={"navn": name.strip(), "size": 10}, timeout=30)
        if r.status_code == 200:
            return r.json().get("_embedded", {}).get("enheter", []) or []
    except Exception:
        pass
    return []


def get_company_details(company: dict):
    if not company:
        return None
    formatted = {
        "company_name": company.get("navn", ""),
        "org_number": company.get("organisasjonsnummer", ""),
        "nace_code": "",
        "nace_description": "",
        "homepage": company.get("hjemmeside", ""),
        "employees": company.get("antallAnsatte", ""),
        "address": "",
        "post_nr": "",
        "city": "",
        "registration_date": company.get("stiftelsesdato", "")
    }
    addr = company.get("forretningsadresse", {})
    if addr:
        address_lines = addr.get("adresse", [])
        if isinstance(address_lines, list):
            formatted["address"] = ", ".join(filter(None, address_lines))
        formatted["post_nr"] = addr.get("postnummer", "")
        formatted["city"] = addr.get("poststed", "")
    nace = company.get("naeringskode1", {})
    if nace:
        formatted["nace_code"] = nace.get("kode", "")
        formatted["nace_description"] = nace.get("beskrivelse", "")
    return formatted


# =========================
# EXCEL TEMPLATE HANDLING (updated)
# - Do NOT rename sheets (user requested)
# - Detect fillable cells by exact RGB match to F2F2F2 (only those will be filled)
# - Fill Brønnøysund-derived data on the FIRST sheet only
# - Provide debug listing of detected cells so you can confirm detection
# =========================

TARGET_FILL_HEX = "F2F2F2"  # cells to fill must have this exact color

def load_template_from_github():
    try:
        if os.path.exists("Grundmall.xlsx"):
            with open("Grundmall.xlsx", "rb") as f:
                return f.read()
        github_url = "https://raw.githubusercontent.com/Shaedyr/PDF-to-premade-Excel-Sheet/main/PremadeExcelTemplate.xlsx"
        r = requests.get(github_url, timeout=30)
        if r.status_code == 200:
            return r.content
        else:
            st.error("Kunne ikke laste Excel-malen fra GitHub")
            return None
    except Exception as e:
        st.error(f"Feil ved lasting av mal: {e}")
        return None


def _rgb_hex_from_color(col):
    """
    Given an openpyxl Color object, try to return a 6-char hex string like 'F2F2F2', or None.
    Handles 'FF' alpha prefix.
    """
    try:
        if not col:
            return None
        rgb = getattr(col, "rgb", None)
        if rgb:
            rgb = rgb.upper()
            if len(rgb) == 8:  # possibly 'FFRRGGBB'
                rgb = rgb[2:]
            if len(rgb) == 6:
                return rgb
        # other color types (theme/indexed) we do not resolve -> None
        return None
    except Exception:
        return None


def _is_fillable(cell):
    """
    Only treat a cell as fillable when its fill color exactly matches TARGET_FILL_HEX.
    Conservative: if color cannot be resolved, we do NOT fill.
    """
    try:
        f = cell.fill
        fg = getattr(f, "fgColor", None) or getattr(f, "start_color", None)
        hexcol = _rgb_hex_from_color(fg)
        if hexcol and hexcol.upper() == TARGET_FILL_HEX:
            return True
        return False
    except Exception:
        return False


# mapping keywords for detecting which field a fillable cell corresponds to
FIELD_KEYWORDS = {
    "company_name": ["selskapsnavn", "selskap", "navn", "firmanavn", "firma", "firma navn", "company", "orgnavn"],
    "org_number": ["organisasjonsnummer", "org.nr", "org nr", "orgnummer", "organisasjons nr", "org", "orgnummer:"],
    "address": ["adresse", "gate", "gatenavn", "street", "postadresse", "adr"],
    "post_nr": ["postnummer", "post nr", "postkode", "post", "postnr"],
    "city": ["poststed", "sted", "city", "post", "by"],
    "employees": ["ansatte", "antall ansatte", "antal ansatte", "ansatt", "employees"],
    "homepage": ["hjemmeside", "web", "website", "url", "nettside"],
    "nace_description": ["nace", "bransje", "næring", "naeringskode", "næringstype", "bransjetekst"],
    "registration_date": ["stiftelsesdato", "registrert", "registration", "registrering", "etablert"]
}

def _normalize_label(text):
    if not text:
        return ""
    return re.sub(r'[^a-zA-Z0-9æøåÆØÅ]+', ' ', str(text).lower()).strip()


def _match_field_by_label(label_text):
    """
    Try to match a nearby label (string) to a field name.
    Steps:
      1) Normalize and do substring checks (fast, exact-ish).
      2) Token-level partial matches (e.g. label contains token 'ansatte').
      3) Fuzzy match against keywords using SequenceMatcher; require conservative threshold.
    Returns field_name or None.
    """
    if not label_text:
        return None

    lab = _normalize_label(label_text)  # normalized lower-case tokens

    # 1) Exact substring match of any keyword
    for field, keywords in FIELD_KEYWORDS.items():
        for kw in keywords:
            if kw in lab:
                return field

    # 2) Token-level presence (label tokens vs keyword fragment)
    lab_tokens = lab.split()
    for field, keywords in FIELD_KEYWORDS.items():
        for kw in keywords:
            kw_tokens = kw.split()
            # if any keyword token is in label tokens -> match
            if any(tok in lab_tokens for tok in kw_tokens):
                return field

    # 3) Fuzzy match: compute best ratio across all keywords → require threshold
    best_field = None
    best_score = 0.0
    for field, keywords in FIELD_KEYWORDS.items():
        for kw in keywords:
            score = difflib.SequenceMatcher(None, lab, kw).ratio()
            if score > best_score:
                best_score = score
                best_field = field

    # Conservative threshold: 0.60 — adjust upward if you get false positives
    if best_score >= 0.60:
        return best_field

    return None

def fill_workbook_bytes(template_bytes: bytes, field_values: dict):
    """
    Fill mapped cells, but only on the FIRST sheet for Brønnøysund-derived data.
    Returns (filled_bytes, report).
    """
    report = {"filled": [], "skipped": [], "errors": [], "unmapped_cells": [], "debug_cells": []}
    wb_scan = load_workbook(BytesIO(template_bytes), data_only=False)
    sheet_names = wb_scan.sheetnames
    first_sheet_name = sheet_names[0] if sheet_names else None

    # scan mapping/unmatched/debug for whole workbook (for debug), but we will fill only first sheet
    mapping, unmatched, debug_cells = scan_and_map_fill_cells(template_bytes)
    report["debug_cells"] = debug_cells

    # Only consider mapping on the first sheet for actual filling
    first_map = mapping.get(first_sheet_name, {}) if first_sheet_name else {}

    # Prepare workbook to write
    wb = load_workbook(BytesIO(template_bytes))
    if not first_sheet_name:
        report["errors"].append(("NO_SHEET", None, "No sheets found in template"))
        return template_bytes, report
    ws = wb[first_sheet_name]

    # Fill fields that have a mapped cell on the first sheet
    for field_name, value in field_values.items():
        if field_name in first_map:
            coord = first_map[field_name][1]
            try:
                if value not in (None, ""):
                    ws[coord].value = str(value)
                    report["filled"].append((first_sheet_name, coord, field_name))
                else:
                    report["skipped"].append((first_sheet_name, coord, field_name, "No value provided"))
            except Exception as e:
                report["errors"].append((first_sheet_name, coord, field_name, str(e)))
        else:
            # Not mapped on first sheet; skip for now (per user: fill on first sheet)
            report["skipped"].append((first_sheet_name, None, field_name, "No mapped cell on first sheet"))

    # For unmatched fillable cells on the first sheet, try to auto-fill remaining values (if any)
    remaining = {k: v for k, v in field_values.items() if v not in (None, "")}
    for _, _, f, *rest in report["filled"]:
        # remove filled fields
        remaining.pop(f, None)

    if unmatched:
        # filter unmatched to first sheet only
        unmatched_first = [t for t in unmatched if t[0] == first_sheet_name]
        if unmatched_first and remaining:
            # fill unmatched first-sheet cells in insertion order with remaining fields
            for (sheetname, coord, label) in unmatched_first:
                if not remaining:
                    break
                field_name, val = remaining.popitem()
                try:
                    wb[sheetname][coord].value = str(val)
                    report["filled"].append((sheetname, coord, field_name, "auto-mapped"))
                except Exception as e:
                    report["errors"].append((sheetname, coord, field_name, str(e)))
        else:
            # nothing to auto-map or no remaining values
            pass

    # Save workbook
    out = BytesIO()
    wb.save(out)
    out.seek(0)
    report["unmapped_cells"] = [u for u in unmatched if u[0] == first_sheet_name]
    return out.getvalue(), report


# =========================
# PDF SCANNING HELPERS
# - naive extraction heuristics (org number, company name, address, postcode/city)
# - if org number found we fetch Brreg details (more reliable)
# =========================
ORG_RE = re.compile(r'\b(\d{9})\b')
ORG_IN_TEXT_RE = re.compile(r'(organisasjonsnummer|org.nr|org nr|orgnummer)[:\s]*?(\d{9})', flags=re.I)
COMPANY_WITH_SUFFIX_RE = re.compile(r'([A-ZÆØÅ][A-Za-zÆØÅæøå\.\-&\s]{1,120}?)\s+(AS|ASA|ANS|DA|ENK|KS|BA)\b',
                                     flags=re.I)


def extract_text_from_pdf(file_bytes):
    try:
        text = ""
        with pdfplumber.open(BytesIO(file_bytes)) as pdf:
            # combine first few pages to be safe
            for i, page in enumerate(pdf.pages[:4]):
                text += page.extract_text() or ""
        return text
    except Exception:
        return ""


def extract_fields_from_pdf_bytes(pdf_bytes):
    txt = extract_text_from_pdf(pdf_bytes)
    fields = {}
    # org number
    m = ORG_IN_TEXT_RE.search(txt)
    if m:
        fields["org_number"] = m.group(2)
    else:
        m2 = ORG_RE.search(txt)
        if m2:
            fields["org_number"] = m2.group(1)
    # company name
    m3 = COMPANY_WITH_SUFFIX_RE.search(txt)
    if m3:
        name = m3.group(0).strip()
        fields["company_name"] = name
    else:
        # fallback: first non-empty line that looks like a name (capitalized)
        lines = [l.strip() for l in txt.splitlines() if l.strip()]
        for ln in lines[:20]:
            if len(ln) > 3 and any(ch.isalpha() for ch in ln) and ln == ln.title():
                fields["company_name"] = ln
                break
    # postcode + city
    mpc = re.search(r'(\d{4})\s+([A-ZÆØÅa-zæøå\-\s]{2,50})', txt)
    if mpc:
        fields["post_nr"] = mpc.group(1)
        fields["city"] = mpc.group(2).strip()
    # address - heuristic: look for lines with street + number
    maddr = re.search(r'([A-ZÆØÅa-zæøå\.\-\s]{3,60}\s+\d{1,4}[A-Za-z]?)', txt)
    if maddr:
        fields["address"] = maddr.group(1).strip()
    return fields


def fetch_brreg_by_org(org_number: str):
    try:
        url = f"https://data.brreg.no/enhetsregisteret/api/enheter/{org_number}"
        r = requests.get(url, timeout=20)
        if r.status_code == 200:
            return r.json()
    except Exception:
        pass
    return None


def format_brreg_data(api_data):
    if not api_data:
        return {}
    formatted = {
        "company_name": api_data.get("navn", ""),
        "org_number": api_data.get("organisasjonsnummer", ""),
        "nace_code": "",
        "nace_description": "",
        "homepage": api_data.get("hjemmeside", ""),
        "employees": api_data.get("antallAnsatte", ""),
        "address": "",
        "post_nr": "",
        "city": "",
        "registration_date": api_data.get("stiftelsesdato", "")
    }
    addr = api_data.get("forretningsadresse", {})
    if addr:
        address_lines = addr.get("adresse", [])
        if isinstance(address_lines, list):
            formatted["address"] = ", ".join(filter(None, address_lines))
        formatted["post_nr"] = addr.get("postnummer", "")
        formatted["city"] = addr.get("poststed", "")
    nace = api_data.get("naeringskode1", {})
    if nace:
        formatted["nace_code"] = nace.get("kode", "")
        formatted["nace_description"] = nace.get("beskrivelse", "")
    return formatted


# =========================
# STREAMLIT UI (unchanged layout)
# - UI kept same per user request
# - Now fills FIRST sheet only and only cells with color F2F2F2
# =========================
def main():
    st.title("📄 PDF → Excel (Brønnøysund)")
    st.markdown("Hent selskapsinformasjon og oppdater Excel automatisk")
    st.markdown("---")

    col1, col2 = st.columns(2)

    with col1:
        pdf_file = st.file_uploader("PDF dokument (valgfritt)", type="pdf", help="Last opp PDF for referanse")

    with col2:
        company_name_input = st.text_input(
            "Selskapsnavn *",
            placeholder="F.eks. Equinor ASA",
            help="Skriv inn navn og velg fra listen",
            key="company_search_input"
        )

        if company_name_input and len(company_name_input.strip()) >= 2:
            companies = search_companies_live(company_name_input)
            st.session_state.companies_list = companies

            if companies:
                options = ["-- Velg selskap --"]
                company_dict = {}

                for company in companies:
                    name = company.get('navn', 'Ukjent navn')
                    org_num = company.get('organisasjonsnummer', '')
                    city = company.get('forretningsadresse', {}).get('poststed', '')

                    display_text = f"{name}"
                    if org_num:
                        display_text += f" (Org.nr: {org_num})"
                    if city:
                        display_text += f" - {city}"

                    options.append(display_text)
                    company_dict[display_text] = company

                selected = st.selectbox(
                    "🔍 Velg fra søkeresultater:",
                    options,
                    key="company_dropdown"
                )

                if selected and selected != "-- Velg selskap --":
                    selected_company = company_dict[selected]
                    st.session_state.selected_company_data = format_brreg_data(selected_company)
                    st.success(f"✅ Valgt: {selected_company.get('navn')}")
                else:
                    st.session_state.selected_company_data = None
                    if company_name_input and len(company_name_input.strip()) >= 3:
                        st.warning("Vennligst velg et selskap fra listen")
            else:
                if company_name_input and len(company_name_input.strip()) >= 3:
                    st.warning("Ingen selskaper funnet. Prøv et annet navn.")
                st.session_state.selected_company_data = None
        else:
            st.session_state.selected_company_data = None

    st.markdown("---")

    # Load Excel template (only once)
    if 'template_loaded' not in st.session_state:
        with st.spinner("Laster Excel-mal..."):
            template_bytes = load_template_from_github()
            if template_bytes:
                st.session_state.template_bytes = template_bytes
                st.session_state.template_loaded = True
                st.success("✅ Excel-mal lastet")
            else:
                st.session_state.template_loaded = False
                st.error("❌ Kunne ikke laste Excel-mal")

    # Provide inspection UI (optional) but does not change main UI layout
    st.markdown("---")
    st.markdown("### 🔎 Inspeksjon (valgfritt)")
    ins_col1, ins_col2 = st.columns(2)
    with ins_col1:
        uploaded_xlsx = st.file_uploader("Last opp Excel for inspeksjon (valgfritt)", type=["xlsx"])
        if uploaded_xlsx:
            try:
                info = {}
                wb = load_workbook(BytesIO(uploaded_xlsx.read()), data_only=True)
                info["sheets"] = wb.sheetnames
                ws = wb.worksheets[0]
                info["sheet_title"] = ws.title
                info["merged_ranges"] = [str(r) for r in ws.merged_cells.ranges]
                info["A2"] = (ws["A2"].value or "")[:1000]
                # show detected fillable cells (by inspecting fills)
                dbg_map = []
                wb_full = load_workbook(BytesIO(uploaded_xlsx.read()), data_only=False)
                for w in wb_full.worksheets:
                    for row in w.iter_rows():
                        for c in row:
                            try:
                                fg = getattr(c.fill, "fgColor", None) or getattr(c.fill, "start_color", None)
                                hexcol = _rgb_hex_from_color(fg)
                                if hexcol:
                                    dbg_map.append((w.title, c.coordinate, hexcol, True if hexcol.upper() == TARGET_FILL_HEX else False))
                            except Exception:
                                continue
                info["detected_colors_sample"] = dbg_map[:200]
                st.json(info)
            except Exception as e:
                st.error(f"Kunne ikke lese filen: {e}")

    with ins_col2:
        if st.button("Vis lastet mal (om tilgjengelig)"):
            tb = st.session_state.get("template_bytes")
            if not tb:
                st.warning("Ingen mal lastet i session_state.")
            else:
                try:
                    wb = load_workbook(BytesIO(tb), data_only=True)
                    info = {"sheets": wb.sheetnames}
                    ws = wb.worksheets[0]
                    info["sheet_title"] = ws.title
                    info["merged_ranges"] = [str(r) for r in ws.merged_cells.ranges]
                    info["A2"] = (ws["A2"].value or "")[:1000]
                    # show colors sample for first sheet
                    wb_full = load_workbook(BytesIO(tb), data_only=False)
                    dbg_map = []
                    w = wb_full.worksheets[0]
                    for row in w.iter_rows():
                        for c in row:
                            try:
                                fg = getattr(c.fill, "fgColor", None) or getattr(c.fill, "start_color", None)
                                hexcol = _rgb_hex_from_color(fg)
                                if hexcol:
                                    dbg_map.append((w.title, c.coordinate, hexcol, True if hexcol.upper() == TARGET_FILL_HEX else False))
                            except Exception:
                                continue
                    info["first_sheet_color_sample"] = dbg_map[:200]
                    st.json(info)
                except Exception as e:
                    st.error(f"Feil ved inspeksjon av mal: {e}")

    st.markdown("---")

    # Processing button: now uses PDF scanning and template scanning + mapping (fills first sheet only)
    if st.button("🚀 Prosesser & Oppdater Excel", use_container_width=True):
        # Ensure template loaded
        if not st.session_state.get('template_loaded'):
            st.error("❌ Excel-mal ikke tilgjengelig")
            st.stop()

        # Start with empty field values
        field_values = {}

        # 1) If user selected a company from dropdown (Brreg), use that data as base
        if st.session_state.selected_company_data:
            field_values.update(st.session_state.selected_company_data)

        # 2) If PDF uploaded, extract fields from PDF and prefer them (or use them to find org and then Brreg)
        if pdf_file:
            try:
                pdf_bytes = pdf_file.read()
                extracted = extract_fields_from_pdf_bytes(pdf_bytes)
                # If org_number found, fetch Brreg to obtain richer data
                if "org_number" in extracted:
                    br = fetch_brreg_by_org(extracted["org_number"])
                    if br:
                        br_data = format_brreg_data(br)
                        # Brreg data takes precedence
                        for k, v in br_data.items():
                            if v:
                                field_values[k] = v
                        # overlay PDF provided fields where useful
                        for k, v in extracted.items():
                            if v:
                                field_values[k] = v
                    else:
                        field_values.update(extracted)
                else:
                    # No org number, just merge the extracted PDF fields
                    field_values.update(extracted)
            except Exception as e:
                st.error(f"❌ Feil ved PDF-parsing: {e}")

        # If still no data and no selection -> error
        if not field_values:
            st.error("❌ Ingen selskapsdata funnet. Velg et selskap fra listen eller last opp en PDF som inneholder selskapets informasjon.")
            st.stop()

        st.session_state.extracted_data = field_values

        # 3) Fill the workbook by scanning fillable cells and mapping labels (first sheet only)
        try:
            updated_bytes, report = fill_workbook_bytes(st.session_state.template_bytes, field_values)
            st.session_state.excel_bytes = updated_bytes
            # Evaluate report for errors/warnings
            if report["errors"]:
                st.error("Noen celler kunne ikke fylles. Se detaljer under.")
                for err in report["errors"]:
                    st.write(f"Feil: {err}")
            if report["skipped"]:
                st.warning("Noen felter ble hoppet over (ingen mapping eller verdi):")
                for s in report["skipped"]:
                    st.write(f"{s}")
            if report["unmapped_cells"]:
                st.info("Fylleceller på første ark uten entydig label (kan være auto-mappet):")
                for um in report["unmapped_cells"]:
                    st.write(f"{um}")
            if report["filled"]:
                st.success(f"✅ Fylte {len(report['filled'])} celler i første arket.")
            else:
                st.warning("Kunne ikke fylle noen celler på første arket — sjekk malen og feltene.")
            # Show detection debug so you can verify which cells were considered fillable vs headers
            if report.get("debug_cells"):
                st.markdown("**Oppdagede celler (debug)**")
                df_dbg = pd.DataFrame(report["debug_cells"], columns=["sheet", "cell", "rgb_hex", "is_fillable", "near_label"])
                st.dataframe(df_dbg)
            st.session_state.excel_ready = True
        except Exception as e:
            st.error(f"❌ Feil ved utfylling av Excel: {e}")
            st.session_state.excel_ready = False

    # =========================
    # DISPLAY EXTRACTED DATA (UI unchanged)
    # =========================
    if st.session_state.extracted_data:
        st.markdown("---")
        st.subheader("📋 Ekstraherte data")
        col_data1, col_data2 = st.columns(2)
        with col_data1:
            st.write("**Selskapsinformasjon:**")
            data = st.session_state.extracted_data
            st.write(f"**Selskapsnavn:** {data.get('company_name', '')}")
            st.write(f"**Organisasjonsnummer:** {data.get('org_number', '')}")
            st.write(f"**Adresse:** {data.get('address', '')}")
            st.write(f"**Postnummer:** {data.get('post_nr', '')}")
            st.write(f"**Poststed:** {data.get('city', '')}")
            st.write(f"**Antall ansatte:** {data.get('employees', '')}")
            st.write(f"**Hjemmeside:** {data.get('homepage', '')}")
            nace_code = data.get('nace_code', '')
            nace_description = data.get('nace_description', '')
            if nace_code and nace_description:
                st.write(f"**NACE-bransje/nummer:** {nace_description} ({nace_code})")
            elif nace_code:
                st.write(f"**NACE-nummer:** {nace_code}")
            elif nace_description:
                st.write(f"**NACE-bransje:** {nace_description}")
        with col_data2:
            if st.session_state.company_summary:
                st.write("**Sammendrag (går i celle A2:D13):**")
                st.info(st.session_state.company_summary)

    # =========================
    # DOWNLOAD (UI unchanged)
    # =========================
    if st.session_state.get('excel_ready') and st.session_state.get('excel_bytes'):
        st.markdown("---")
        st.subheader("📥 Last ned")
        company_name_dl = st.session_state.extracted_data.get('company_name', 'selskap')
        safe_name = re.sub(r'[^\w\s-]', '', company_name_dl, flags=re.UNICODE)
        safe_name = re.sub(r'[-\s]+', '_', safe_name)
        st.download_button(
            label="⬇️ Last ned oppdatert Excel",
            data=st.session_state.excel_bytes,
            file_name=f"{safe_name}_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    st.markdown("---")
    st.caption("Drevet av Brønnøysund Enhetsregisteret API | Data mellomlagret i 1 time")


if __name__ == "__main__":
    main()
