import os
import re
import difflib
import requests
import wikipedia
import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Alignment
import pdfplumber

# =========================
# CONFIGURATION
# =========================
st.set_page_config(page_title="PDF → Excel (Brønnøysund)", layout="wide", page_icon="📊")
for k, v in {"extracted_data": {}, "api_response": None, "excel_ready": False, "company_summary": ""}.items():
    if k not in st.session_state:
        st.session_state[k] = v

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
# BRØNNØYSUND LIVE SEARCH
# =========================
@st.cache_data(ttl=3600)
def search_companies_live(name: str):
    if not name or len(name.strip()) < 2:
        return []
    try:
        r = requests.get("https://data.brreg.no/enhetsregisteret/api/enheter",
                         params={"navn": name.strip(), "size": 10}, timeout=30)
        if r.status_code == 200:
            data = r.json()
            companies = data.get("_embedded", {}).get("enheter", [])
            return companies if companies else []
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
# EXCEL TEMPLATE HANDLING
# - exact target color: F2F2F2
# - fill only FIRST sheet with Brreg data
# - robust Norwegian keyword/fuzzy matching
# =========================
TARGET_FILL_HEX = "F2F2F2"  # exact hex to detect fillable cells


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
    try:
        if not col:
            return None
        rgb = getattr(col, "rgb", None)
        if rgb:
            rgb = rgb.upper()
            if len(rgb) == 8:
                rgb = rgb[2:]
            if len(rgb) == 6:
                return rgb
        return None
    except Exception:
        return None


# Expanded keywords (Norwegian variants + common labels)
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
    if not label_text:
        return None
    lab = _normalize_label(label_text)

    # 1) Exact substring match
    for field, keywords in FIELD_KEYWORDS.items():
        for kw in keywords:
            if kw in lab:
                return field

    # 2) Token-level presence
    lab_tokens = lab.split()
    for field, keywords in FIELD_KEYWORDS.items():
        for kw in keywords:
            kw_tokens = kw.split()
            if any(tok in lab_tokens for tok in kw_tokens):
                return field

    # 3) Conservative fuzzy match
    best_field = None
    best_score = 0.0
    for field, keywords in FIELD_KEYWORDS.items():
        for kw in keywords:
            score = difflib.SequenceMatcher(None, lab, kw).ratio()
            if score > best_score:
                best_score = score
                best_field = field

    if best_score >= 0.60:
        return best_field

    return None


# --------- REPLACE scan_and_map_fill_cells WITH THIS ----------
def scan_and_map_fill_cells(wb_bytes):
    """
    Two-pass mapping:
    1) Find all fillable cells (exact F2F2F2).
    2) For each fillable cell try local label matching (left/above/right/below, diags).
    3) For remaining unmapped fields, search whole sheet for label keywords and assign
       nearest unassigned fillable cell.
    Returns mapping, unmatched, debug_cells
      - mapping: { sheetname: { field_name: (cell_coordinate, assigned_by, label_text) } }
      - unmatched: list of (sheetname, coord, nearby_label)
      - debug_cells: list of (sheet, coord, hex, is_fillable, nearby_label)
    """
    mapping = {}
    unmatched = []
    debug_cells = []

    wb = load_workbook(BytesIO(wb_bytes), data_only=False)
    for ws in wb.worksheets:
        sheet_fillables = []  # list of (row, col, coord)
        assigned = {}         # field -> (coord, assigned_by, label)
        cell_by_coord = {}    # coord -> cell object for nearest calc
        # collect fillable cells and debug info
        for row in ws.iter_rows():
            for cell in row:
                try:
                    fg = getattr(cell.fill, "fgColor", None) or getattr(cell.fill, "start_color", None)
                    hexcol = _rgb_hex_from_color(fg)
                    is_fill = True if (hexcol and hexcol.upper() == TARGET_FILL_HEX) else False
                    # nearby label attempt (simple left/above)
                    label = ""
                    if cell.column > 1:
                        left = ws.cell(row=cell.row, column=cell.column - 1).value
                        if left:
                            label = str(left)
                    if not label and cell.row > 1:
                        above = ws.cell(row=cell.row - 1, column=cell.column).value
                        if above:
                            label = str(above)
                    combined = " ".join(filter(None, [label, str(ws.cell(row=cell.row - 1, column=cell.column).value or "")]))
                    debug_cells.append((ws.title, cell.coordinate, hexcol, is_fill, label))
                    if is_fill:
                        sheet_fillables.append((cell.row, cell.column, cell.coordinate))
                        cell_by_coord[cell.coordinate] = cell
                except Exception:
                    continue

        # local matching: for each fillable look around in a neighborhood for labels
        for r, c, coord in sheet_fillables:
            # build list of neighbor coords to probe in order of priority
            neighbors = [
                (r, c - 1), (r, c - 2),     # left, left2
                (r - 1, c), (r - 2, c),     # above, above2
                (r, c + 1), (r, c + 2),     # right, right2
                (r + 1, c), (r + 2, c),     # below, below2
                (r - 1, c - 1), (r - 1, c + 1), (r + 1, c - 1), (r + 1, c + 1)
            ]
            found_label = ""
            found_field = None
            for rr, cc in neighbors:
                if rr < 1 or cc < 1:
                    continue
                try:
                    val = ws.cell(row=rr, column=cc).value
                    if val and str(val).strip():
                        found_label = str(val)
                        found_field = _match_field_by_label(found_label)
                        if found_field:
                            assigned[found_field] = (coord, "local", found_label)
                            break
                except Exception:
                    continue

        # global mapping for remaining fields: search sheet for label cells (keywords)
        # precompute list of all candidate label cells with their normalized text
        label_cells = []
        for row in ws.iter_rows():
            for cell in row:
                try:
                    txt = cell.value
                    if txt and str(txt).strip():
                        label_cells.append((cell.row, cell.column, cell.coordinate, str(txt)))
                except Exception:
                    continue

        # for fields not yet assigned, find the best label cell and nearest unassigned fillable
        all_fields = list(FIELD_KEYWORDS.keys())
        unassigned_fields = [f for f in all_fields if f not in assigned]
        # compute set of unassigned fillable coords
        unassigned_fill_coords = [coord for (_r, _c, coord) in sheet_fillables]

        def manhattan(r1, c1, r2, c2):
            return abs(r1 - r2) + abs(c1 - c2)

        for field in unassigned_fields:
            # search label_cells for any that match the field keywords (normalized/fuzzy)
            best_label_cell = None
            best_label_score = 0.0
            for rr, cc, coord_label, raw in label_cells:
                lab = _normalize_label(raw)
                # quick substring check
                for kw in FIELD_KEYWORDS[field]:
                    if kw in lab:
                        # immediate strong match
                        best_label_cell = (rr, cc, coord_label, raw)
                        best_label_score = 1.0
                        break
                # otherwise fuzzy check
                if best_label_cell is None:
                    for kw in FIELD_KEYWORDS[field]:
                        score = difflib.SequenceMatcher(None, lab, kw).ratio()
                        if score > best_label_score:
                            best_label_score = score
                            best_label_cell = (rr, cc, coord_label, raw)
            # accept if strong enough
            if best_label_cell and best_label_score >= 0.55:
                rr, cc, coord_label, raw = best_label_cell
                # find nearest unassigned fillable cell
                best_fill = None
                best_dist = None
                for (_r, _c, coord_fill) in sheet_fillables:
                    if coord_fill not in unassigned_fill_coords:
                        continue
                    d = manhattan(rr, cc, _r, _c)
                    if best_dist is None or d < best_dist:
                        best_dist = d
                        best_fill = coord_fill
                if best_fill:
                    assigned[field] = (best_fill, "global", raw)
                    # mark that fill coord as assigned
                    if best_fill in unassigned_fill_coords:
                        unassigned_fill_coords.remove(best_fill)

        # build mapping for sheet
        smap = {}
        for f, v in assigned.items():
            smap[f] = v  # (coord, assigned_by, label)
        # unmatched are fillables left without assignment
        remaining_unmapped = [coord for coord in [t[2] for t in sheet_fillables] if coord not in [v[0] for v in assigned.values()]]
        for coord_un in remaining_unmapped:
            # find nearby label for info
            lab = ""
            r_cell = ws[coord_un].row
            c_cell = ws[coord_un].column
            if c_cell > 1:
                left = ws.cell(row=r_cell, column=c_cell - 1).value
                if left:
                    lab = str(left)
            if not lab and r_cell > 1:
                above = ws.cell(row=r_cell - 1, column=c_cell).value
                if above:
                    lab = str(above)
            unmatched.append((ws.title, coord_un, lab or None))

        if smap:
            mapping[ws.title] = smap

    return mapping, unmatched, debug_cells


def fill_workbook_bytes(template_bytes: bytes, field_values: dict):
    """
    Fill using mapping produced by scan_and_map_fill_cells.
    Writes only to the FIRST sheet (per your request).
    After filling mapped fields it will also replace the "Skriv her" cell
    (if present) with the company summary (field_values['company_summary']).
    Returns (filled_bytes, report) where report includes mapping, debug, errors, etc.
    """
    report = {"filled": [], "skipped": [], "errors": [], "unmapped_cells": [], "debug_cells": [], "mapping": {}}
    wb_scan = load_workbook(BytesIO(template_bytes), data_only=False)
    sheet_names = wb_scan.sheetnames
    first_sheet_name = sheet_names[0] if sheet_names else None

    mapping, unmatched, debug_cells = scan_and_map_fill_cells(template_bytes)
    report["debug_cells"] = debug_cells
    report["mapping"] = mapping

    first_map = mapping.get(first_sheet_name, {}) if first_sheet_name else {}

    wb = load_workbook(BytesIO(template_bytes))
    if not first_sheet_name:
        report["errors"].append(("NO_SHEET", None, "No sheets found in template"))
        return template_bytes, report
    ws = wb[first_sheet_name]

    # 1) Fill mapped fields on first sheet
    for field_name, value in field_values.items():
        if field_name in first_map:
            coord = first_map[field_name][0]  # coord stored at index 0
            try:
                if value not in (None, ""):
                    ws[coord].value = str(value)
                    report["filled"].append((first_sheet_name, coord, field_name))
                else:
                    report["skipped"].append((first_sheet_name, coord, field_name, "No value provided"))
            except Exception as e:
                report["errors"].append((first_sheet_name, coord, field_name, str(e)))
        else:
            report["skipped"].append((first_sheet_name, None, field_name, "No mapped cell on first sheet"))

    # 2) Auto-map remaining values to unmatched fillable cells on first sheet
    remaining = {k: v for k, v in field_values.items() if v not in (None, "")}
    for (_s, _coord, f) in report["filled"]:
        remaining.pop(f, None)

    unmatched_first = [t for t in unmatched if t[0] == first_sheet_name]
    if unmatched_first and remaining:
        for (sheetname, coord, label) in unmatched_first:
            if not remaining:
                break
            field_name, val = remaining.popitem()
            try:
                wb[sheetname][coord].value = str(val)
                report["filled"].append((sheetname, coord, field_name, "auto-mapped"))
            except Exception as e:
                report["errors"].append((sheetname, coord, field_name, str(e)))

    # 3) SPECIAL HANDLING: place company_summary into the "Skriv her" cell (or A46 fallback)
    summary = field_values.get("company_summary") or ""
    if summary:
        wrote_summary = False
        # search the first sheet for a cell that contains "skriv her" (case-insensitive)
        try:
            for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
                for c in row:
                    try:
                        if c.value and isinstance(c.value, str) and "skriv her" in c.value.strip().lower():
                            c.value = str(summary)
                            report["filled"].append((first_sheet_name, c.coordinate, "company_summary", "replaced 'Skriv her'"))
                            wrote_summary = True
                            break
                    except Exception:
                        continue
                if wrote_summary:
                    break
            # fallback: write to A46 if not found and sheet has that cell
            if not wrote_summary:
                try:
                    ws["A46"] = str(summary)
                    report["filled"].append((first_sheet_name, "A46", "company_summary", "fallback to A46"))
                    wrote_summary = True
                except Exception as e:
                    report["errors"].append((first_sheet_name, "A46", "company_summary", f"Fallback write failed: {e}"))
        except Exception as e:
            report["errors"].append((first_sheet_name, None, "company_summary", f"Search/replace failed: {e}"))

    # Save workbook
    out = BytesIO()
    wb.save(out)
    out.seek(0)
    report["unmapped_cells"] = [u for u in unmatched if u[0] == first_sheet_name]
    return out.getvalue(), report


# =========================
# PDF EXTRACTION HELPERS
# =========================
ORG_RE = re.compile(r'\b(\d{9})\b')
ORG_IN_TEXT_RE = re.compile(r'(organisasjonsnummer|org\.?nr|org nr|orgnummer)[:\s]*?(\d{9})', flags=re.I)
COMPANY_WITH_SUFFIX_RE = re.compile(r'([A-ZÆØÅ][A-Za-zÆØÅæøå\.\-&\s]{1,120}?)\s+(AS|ASA|ANS|DA|ENK|KS|BA)\b',
                                     flags=re.I)


def extract_text_from_pdf(file_bytes):
    try:
        text = ""
        with pdfplumber.open(BytesIO(file_bytes)) as pdf:
            for i, page in enumerate(pdf.pages[:6]):
                text += page.extract_text() or ""
        return text
    except Exception:
        return ""


def extract_fields_from_pdf_bytes(pdf_bytes):
    txt = extract_text_from_pdf(pdf_bytes)
    fields = {}
    m = ORG_IN_TEXT_RE.search(txt)
    if m:
        fields["org_number"] = m.group(2)
    else:
        m2 = ORG_RE.search(txt)
        if m2:
            fields["org_number"] = m2.group(1)
    m3 = COMPANY_WITH_SUFFIX_RE.search(txt)
    if m3:
        fields["company_name"] = m3.group(0).strip()
    else:
        lines = [l.strip() for l in txt.splitlines() if l.strip()]
        for ln in lines[:30]:
            if len(ln) > 3 and any(ch.isalpha() for ch in ln) and ln == ln.title():
                fields["company_name"] = ln
                break
    mpc = re.search(r'(\d{4})\s+([A-ZÆØÅa-zæøå\-\s]{2,50})', txt)
    if mpc:
        fields["post_nr"] = mpc.group(1)
        fields["city"] = mpc.group(2).strip()
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
# STREAMLIT UI (kept layout)
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

    # Inspector UI
    st.markdown("---")
    st.markdown("### 🔎 Inspeksjon (valgfritt)")
    ins_col1, ins_col2 = st.columns(2)
    with ins_col1:
        uploaded_xlsx = st.file_uploader("Last opp Excel for inspeksjon (valgfritt)", type=["xlsx"])
        if uploaded_xlsx:
            try:
                info = {}
                data_bytes = uploaded_xlsx.read()
                wb = load_workbook(BytesIO(data_bytes), data_only=True)
                info["sheets"] = wb.sheetnames
                ws = wb.worksheets[0]
                info["sheet_title"] = ws.title
                info["merged_ranges"] = [str(r) for r in ws.merged_cells.ranges]
                info["A2"] = (ws["A2"].value or "")[:1000]
                dbg_map = []
                wb_full = load_workbook(BytesIO(data_bytes), data_only=False)
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
                info["detected_colors_sample"] = dbg_map[:400]
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
                    info["first_sheet_color_sample"] = dbg_map[:400]
                    st.json(info)
                except Exception as e:
                    st.error(f"Feil ved inspeksjon av mal: {e}")

    st.markdown("---")

    # Process
    if st.button("🚀 Prosesser & Oppdater Excel", use_container_width=True):
        if not st.session_state.get('template_loaded'):
            st.error("❌ Excel-mal ikke tilgjengelig")
            st.stop()

        field_values = {}
        if st.session_state.selected_company_data:
            field_values.update(st.session_state.selected_company_data)

        if pdf_file:
            try:
                pdf_bytes = pdf_file.read()
                extracted = extract_fields_from_pdf_bytes(pdf_bytes)
                if "org_number" in extracted:
                    br = fetch_brreg_by_org(extracted["org_number"])
                    if br:
                        br_data = format_brreg_data(br)
                        for k, v in br_data.items():
                            if v:
                                field_values[k] = v
                        for k, v in extracted.items():
                            if v:
                                field_values[k] = v
                    else:
                        field_values.update(extracted)
                else:
                    field_values.update(extracted)
            except Exception as e:
                st.error(f"❌ Feil ved PDF-parsing: {e}")

        if not field_values:
            st.error("❌ Ingen selskapsdata funnet. Velg et selskap fra listen eller last opp en PDF som inneholder selskapets informasjon.")
            st.stop()

        st.session_state.extracted_data = field_values

        try:
            updated_bytes, report = fill_workbook_bytes(st.session_state.template_bytes, field_values)
            st.session_state.excel_bytes = updated_bytes
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
            if report.get("debug_cells"):
                st.markdown("**Oppdagede celler (debug)**")
                df_dbg = pd.DataFrame(report["debug_cells"], columns=["sheet", "cell", "rgb_hex", "is_fillable", "near_label"])
                st.dataframe(df_dbg)
            st.session_state.excel_ready = True
        except Exception as e:
            st.error(f"❌ Feil ved utfylling av Excel: {e}")
            st.session_state.excel_ready = False

    # Display extracted data
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

    # Download
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
