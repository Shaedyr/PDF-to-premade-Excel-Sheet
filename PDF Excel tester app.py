import streamlit as st
import pdfplumber
import re
import requests
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from io import BytesIO
import pandas as pd
from datetime import datetime
import wikipedia
import os

# =========================
# CONFIGURATION
# =========================
st.set_page_config(page_title="PDF → Excel (Brønnøysund)", layout="wide", page_icon="📊")
if 'extracted_data' not in st.session_state: st.session_state.extracted_data = {}
if 'api_response' not in st.session_state: st.session_state.api_response = None
if 'excel_ready' not in st.session_state: st.session_state.excel_ready = False
if 'company_summary' not in st.session_state: st.session_state.company_summary = ""

# =========================
# WIKIPEDIA SEARCH
# =========================
def strip_company_suffix(name):
    """Remove common Norwegian company suffixes (case-insensitive)."""
    if not name:
        return name
    # remove trailing suffix like " AS", "ASA", etc.
    stripped = re.sub(r'\b(AS|ASA|ANS|DA|ENK|KS|BA)\b\.?$', '', name, flags=re.I).strip()
    return stripped

def get_company_summary_from_wikipedia(company_name):
    """Get company summary from Wikipedia"""
    try:
        if not company_name:
            return None

        search_name = strip_company_suffix(company_name)

        search_attempts = [search_name, company_name, search_name + " (bedrift)", search_name + " (company)"]

        # Try Norwegian first
        try:
            wikipedia.set_lang("no")
        except Exception:
            pass

        for attempt in search_attempts:
            try:
                search_results = wikipedia.search(attempt)
                if search_results:
                    company_results = []
                    for result in search_results[:5]:
                        result_lower = result.lower()
                        if any(kw in result_lower for kw in ["as", "asa", "bedrift", "selskap", "company", "group"]):
                            company_results.append(result)

                    target = company_results[0] if company_results else search_results[0]
                    page = wikipedia.page(target, auto_suggest=False)
                    summary = page.summary
                    sentences = [s.strip() for s in summary.split('. ') if s.strip()]
                    if len(sentences) > 2:
                        return '. '.join(sentences[:2]) + '.'
                    else:
                        return summary[:300] + '...' if len(summary) > 300 else summary
            except (wikipedia.exceptions.DisambiguationError, wikipedia.exceptions.PageError):
                continue
            except Exception:
                continue

        # Try English and translate a few key phrases to Norwegian
        try:
            wikipedia.set_lang("en")
        except Exception:
            pass

        for attempt in search_attempts:
            try:
                search_results = wikipedia.search(attempt)
                if search_results:
                    page = wikipedia.page(search_results[0], auto_suggest=False)
                    summary = page.summary
                    sentences = [s.strip() for s in summary.split('. ') if s.strip()]
                    short_summary = '. '.join(sentences[:2]) + '.' if len(sentences) > 2 else summary[:300] + '...' if len(summary) > 300 else summary
                    return short_summary.replace(" is a ", " er et ").replace(" company", " selskap").replace(" based in ", " med hovedkontor i ")
            except (wikipedia.exceptions.DisambiguationError, wikipedia.exceptions.PageError):
                continue
            except Exception:
                continue
    except Exception:
        pass
    return None

def get_company_summary_from_web(company_name):
    """Search web for company information via DuckDuckGo instant answer API"""
    try:
        if not company_name:
            return None

        search_name = strip_company_suffix(company_name)

        ddg_url = f"https://api.duckduckgo.com/?q={requests.utils.quote(search_name + ' bedrift')}&format=json&no_html=1&skip_disambig=1"
        response = requests.get(ddg_url, timeout=10)

        if response.status_code == 200:
            data = response.json()
            abstract = data.get('AbstractText', '')
            if abstract and len(abstract) > 50:
                return abstract

        ddg_url_en = f"https://api.duckduckgo.com/?q={requests.utils.quote(search_name + ' company')}&format=json&no_html=1&skip_disambig=1"
        response_en = requests.get(ddg_url_en, timeout=10)
        if response_en.status_code == 200:
            data_en = response_en.json()
            abstract_en = data_en.get('AbstractText', '')
            if abstract_en and len(abstract_en) > 50:
                return abstract_en.replace(" is a ", " er et ").replace(" company", " selskap").replace(" based in ", " med hovedkontor i ")
    except Exception:
        pass
    return None

def create_summary_from_brreg_data(company_data):
    """Create analytical business summary"""
    company_name = company_data.get('company_name', '')
    industry = company_data.get('nace_description', '')
    city = company_data.get('city', '')
    employees = company_data.get('employees', '')
    founded = company_data.get('registration_date', '')

    if not company_name:
        return "Ingen informasjon tilgjengelig om dette selskapet."

    parts = []

    if industry and city:
        parts.append(f"{company_name} driver {industry.lower()} virksomhet fra {city}.")
    elif industry:
        parts.append(f"{company_name} opererer innen {industry.lower()}.")
    else:
        parts.append(f"{company_name} er et registrert norsk selskap.")

    if founded:
        try:
            year = founded.split('-')[0] if '-' in founded else founded
            years_old = datetime.now().year - int(year)
            if years_old > 30:
                parts.append(f"Etablert i {year}, har selskapet over {years_old} års bransjeerfaring.")
            elif years_old > 10:
                parts.append(f"Selskapet har utviklet seg over {years_old} år siden etableringen i {year}.")
            else:
                parts.append(f"Etablert i {year}, er dette et yngre selskap i vekstfasen.")
        except:
            parts.append(f"Selskapet ble registrert i {founded}.")

    if employees:
        try:
            emp_count = int(employees)
            if emp_count > 200:
                parts.append(f"Som en større arbeidsgiver med {emp_count} ansatte, har det betydelig samfunnspåvirkning.")
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
# BRØNNØYSUND API - LIVE SEARCH WITH DROPDOWN
# =========================
@st.cache_data(ttl=3600)
def search_companies_live(name):
    """Search companies and return list for dropdown"""
    if not name or len(name.strip()) < 2:
        return []

    search_term = name.strip()

    try:
        url = "https://data.brreg.no/enhetsregisteret/api/enheter"
        params = {"navn": search_term, "size": 10}

        response = requests.get(url, params=params, timeout=30)

        if response.status_code == 200:
            data = response.json()
            companies = data.get("_embedded", {}).get("enheter", [])
            return companies if companies else []
    except Exception:
        pass

    return []

def get_company_details(company):
    """Extract details from company API response"""
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
# FORMAT COMPANY DATA FUNCTION
# =========================
def format_company_data(api_data):
    """Format API response"""
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
# EXCEL PROCESSING
# =========================
def load_template_from_github():
    """Load Excel template from local file or GitHub; return bytes"""
    try:
        if os.path.exists("Grundmall.xlsx"):
            with open("Grundmall.xlsx", "rb") as f:
                return f.read()

        github_url = "https://raw.githubusercontent.com/Shaedyr/PDF-to-premade-Excel-Sheet/main/PremadeExcelTemplate.xlsx"
        response = requests.get(github_url, timeout=30)

        if response.status_code == 200:
            return response.content
        else:
            st.error("Kunne ikke laste Excel-malen fra GitHub")
            return None
    except Exception as e:
        st.error(f"Feil ved lasting av mal: {e}")
        return None

def update_excel_template(template_bytes, company_data, company_summary):
    """Update Excel template; template_bytes is raw bytes (not BytesIO)"""
    try:
        if not template_bytes:
            raise ValueError("Ingen mal tilgjengelig")

        # Create a fresh BytesIO for each update so load_workbook always reads from start
        template_stream = BytesIO(template_bytes)
        template_stream.seek(0)
        wb = load_workbook(template_stream)
        ws = wb.worksheets[0]

        company_name = company_data.get('company_name', 'Selskap')
        ws.title = re.sub(r'[\\/*?:\[\]]', '', f"{company_name} Info")[:31]

        if len(wb.worksheets) > 1:
            wb.worksheets[1].title = re.sub(r'[\\/*?:\[\]]', '', f"{company_name} Anbud")[:31]

        # Update info window (A2:D13)
        for merged_range in list(ws.merged_cells.ranges):
            if str(merged_range) == 'A2:D13':
                try:
                    ws.unmerge_cells(str(merged_range))
                except Exception:
                    pass

        ws.merge_cells('A2:D13')
        ws['A2'] = company_summary if company_summary else f"Informasjon om {company_name}"
        ws['A2'].alignment = Alignment(wrap_text=True, vertical='top', horizontal='left')

        for row in range(2, 14):
            try:
                ws.row_dimensions[row].height = 18
            except Exception:
                pass

        # Update company data with merging
        data_mapping = {
            'company_name': {'cell': 'B14', 'merge_to': 'D14', 'max_len': 50},
            'org_number': {'cell': 'B15', 'merge_to': None, 'max_len': 20},
            'address': {'cell': 'B16', 'merge_to': 'D16', 'max_len': 100},
            'post_nr': {'cell': 'B17', 'merge_to': 'C17', 'max_len': 15},
            'homepage': {'cell': 'B20', 'merge_to': 'D20', 'max_len': 100},
            'employees': {'cell': 'B21', 'merge_to': None, 'max_len': 10}
        }

        for field, config in data_mapping.items():
            value = company_data.get(field, '')
            if value:
                cell = config['cell']
                merge_to = config['merge_to']
                max_len = config['max_len']

                if len(str(value)) > max_len:
                    value = str(value)[:max_len-3] + "..."

                if field == 'org_number' and len(str(value)) == 9 and str(value).isdigit():
                    ws[cell] = f"'{value}"
                else:
                    ws[cell] = str(value)

                if merge_to and len(str(value)) > 20:
                    try:
                        ws.merge_cells(f"{cell}:{merge_to}")
                        ws[cell].alignment = Alignment(wrap_text=True, vertical='center', horizontal='left')
                    except Exception:
                        pass

        # Handle NACE separately - combine code and description in B18
        nace_code = company_data.get('nace_code', '')
        nace_description = company_data.get('nace_description', '')

        if nace_code and nace_description:
            nace_combined = f"{nace_description} ({nace_code})"
            ws['B18'] = nace_combined
            try:
                ws.merge_cells('B18:D18')
                ws['B18'].alignment = Alignment(wrap_text=True, vertical='center', horizontal='left')
            except Exception:
                pass
        elif nace_code:
            ws['B18'] = nace_code
        elif nace_description:
            ws['B18'] = nace_description
        else:
            ws['B18'] = "Data ikke tilgjengelig"

        ws.column_dimensions['A'].width = 15
        ws.column_dimensions['B'].width = 25
        ws.column_dimensions['C'].width = 15
        ws.column_dimensions['D'].width = 15

        output = BytesIO()
        wb.save(output)
        output.seek(0)
        return output.getvalue()  # return bytes
    except Exception as e:
        st.error(f"Excel oppdatering feilet: {str(e)}")
        # fallback: return a simple excel from dataframe
        try:
            df = pd.DataFrame([company_data])
            output = BytesIO()
            df.to_excel(output, index=False)
            output.seek(0)
            return output.getvalue()
        except Exception:
            return None

# =========================
# STREAMLIT UI
# =========================
def main():
    # REMOVED THE ENTIRE SIDEBAR SECTION
    
    st.title("📄 PDF → Excel (Brønnøysund)")
    st.markdown("Hent selskapsinformasjon og oppdater Excel automatisk")
    st.markdown("---")
    
    # Initialize ALL session state variables needed
    if 'selected_company_data' not in st.session_state:
        st.session_state.selected_company_data = None
    if 'companies_list' not in st.session_state:
        st.session_state.companies_list = []
    if 'current_search' not in st.session_state:
        st.session_state.current_search = ""
    if 'last_search' not in st.session_state:
        st.session_state.last_search = ""
    if 'show_dropdown' not in st.session_state:
        st.session_state.show_dropdown = False
    
    col1, col2 = st.columns(2)
    
    with col1:
        pdf_file = st.file_uploader("PDF dokument (valgfritt)", type="pdf", help="Last opp PDF for referanse")
    
    with col2:
        # Create a container for the search results
        search_results_container = st.container()
        
        # Search input - we'll track changes manually
        company_name_input = st.text_input(
            "Selskapsnavn *",
            placeholder="Skriv her... (minst 2 bokstaver)",
            help="Søk starter automatisk når du skriver",
            key="live_search_input"
        )
        
        # Clear previous selection if search changes
        if 'current_search' in st.session_state and st.session_state.current_search != company_name_input:
            st.session_state.selected_company_data = None
        
        # Store current search
        st.session_state.current_search = company_name_input
        
        # Perform search if we have at least 2 characters
        if company_name_input and len(company_name_input.strip()) >= 2:
            with search_results_container:
                # Add a small delay indicator for better UX
                with st.spinner("Søker..."):
                    companies = search_companies_live(company_name_input)
                
                if companies:
                    # Create dropdown options
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
                    
                    # Show dropdown
                    selected = st.selectbox(
                        "🔍 Søkeresultater:",
                        options,
                        key="dynamic_company_dropdown"
                    )
                    
                    # If user selects a company
                    if selected and selected != "-- Velg selskap --":
                        selected_company = company_dict[selected]
                        st.session_state.selected_company_data = get_company_details(selected_company)
                        st.success(f"✅ Valgt: {selected_company.get('navn')}")
                else:
                    if len(company_name_input.strip()) >= 3:
                        st.warning("Ingen selskaper funnet. Prøv et annet navn.")
                    st.session_state.selected_company_data = None

    st.markdown("---")

    # Load Excel template (only once)
    if 'template_loaded' not in st.session_state:
        with st.spinner("Laster Excel-mal..."):
            template_bytes = load_template_from_github()
            if template_bytes:
                st.session_state.template_bytes = template_bytes  # store raw bytes
                st.session_state.template_loaded = True
                st.success("✅ Excel-mal lastet")
            else:
                st.session_state.template_loaded = False
                st.error("❌ Kunne ikke laste Excel-mal")

    # Process button
    if st.button("🚀 Prosesser & Oppdater Excel", use_container_width=True):
        if not st.session_state.selected_company_data:
            st.error("❌ Vennligst velg et selskap fra listen først")
            st.stop()

        if not st.session_state.get('template_loaded'):
            st.error("❌ Excel-mal ikke tilgjengelig")
            st.stop()

        # Get the selected company data
        formatted_data = st.session_state.selected_company_data
        st.session_state.extracted_data = formatted_data

        # Get company summary (3-step approach)
        st.write("**Trinn 1:** Søker etter selskapsopplysninger...")

        company_summary = None
        company_name = formatted_data.get('company_name', '')

        # Step 1: Try Wikipedia
        with st.spinner("Søker på Wikipedia..."):
            company_summary = get_company_summary_from_wikipedia(company_name)

        # Step 2: If Wikipedia fails, try web search
        if not company_summary:
            st.info("Fant ikke på Wikipedia. Prøver websøk...")
            with st.spinner("Søker på nettet..."):
                company_summary = get_company_summary_from_web(company_name)

        # Step 3: If web search also fails, use Brønnøysund data analysis
        if not company_summary:
            st.info("Fant ikke på nettet. Lager analyse fra Brønnøysund-data...")
            company_summary = create_summary_from_brreg_data(formatted_data)

        st.session_state.company_summary = company_summary

        # Update Excel
        st.write("**Trinn 2:** Oppdaterer Excel...")

        try:
            updated_excel_bytes = update_excel_template(
                st.session_state.template_bytes,
                formatted_data,
                company_summary
            )

            if updated_excel_bytes:
                st.session_state.excel_ready = True
                st.session_state.excel_bytes = updated_excel_bytes
                st.success("✅ Excel-fil oppdatert!")

                st.info(f"""
                **Informasjon plassert i:**
                - **Ark 1:** {re.sub(r'[\\/*?:\\[\\]]', '', f"{formatted_data.get('company_name', 'Selskap')} Info")[:31]}
                - **Stort informasjonsvindu:** Celle A2:D13
                - **Selskapsdata:** Celler B14-B21
                - **Ark 2:** {re.sub(r'[\\/*?:\\[\\]]', '', f"{formatted_data.get('company_name', 'Selskap')} Anbud")[:31]}
                """)
            else:
                st.error("❌ Kunne ikke generere Excel-fil")
        except Exception as e:
            st.error(f"❌ Feil ved Excel-oppdatering: {str(e)}")

    # Display extracted data
    if st.session_state.extracted_data:
        st.markdown("---")
        st.subheader("📋 Ekstraherte data")

        col_data1, col_data2 = st.columns(2)
        with col_data1:
            st.write("**Selskapsinformasjon:**")
            data = st.session_state.extracted_data

            # Display company info
            st.write(f"**Selskapsnavn:** {data.get('company_name', '')}")
            st.write(f"**Organisasjonsnummer:** {data.get('org_number', '')}")
            st.write(f"**Adresse:** {data.get('address', '')}")
            st.write(f"**Postnummer:** {data.get('post_nr', '')}")
            st.write(f"**Poststed:** {data.get('city', '')}")
            st.write(f"**Antall ansatte:** {data.get('employees', '')}")
            st.write(f"**Hjemmeside:** {data.get('homepage', '')}")

            # Display combined NACE
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

    # Download button
    if st.session_state.get('excel_ready', False) and st.session_state.get('excel_bytes'):
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
