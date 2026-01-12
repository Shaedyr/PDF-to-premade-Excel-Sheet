import streamlit as st
import pdfplumber
import re
import requests
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from io import BytesIO
import pandas as pd
import json
from datetime import datetime
import wikipedia
import os


# =========================
# CONFIGURATION
# =========================
st.set_page_config(
    page_title="PDF → Excel (Brønnøysund)",
    layout="wide",
    page_icon="📊",
    initial_sidebar_state="expanded"
)

# Session state initialization
if 'extracted_data' not in st.session_state:
    st.session_state.extracted_data = {}
if 'api_response' not in st.session_state:
    st.session_state.api_response = None
if 'excel_ready' not in st.session_state:
    st.session_state.excel_ready = False
if 'company_summary' not in st.session_state:
    st.session_state.company_summary = ""


# =========================
# WIKIPEDIA SEARCH
# =========================
def get_company_summary_from_wikipedia(company_name):
    """Get a short company summary from Wikipedia"""
    try:
        # Try Norwegian Wikipedia first
        wikipedia.set_lang("no")
        
        try:
            search_results = wikipedia.search(company_name)
            
            if search_results:
                # Get the page for the first result
                page = wikipedia.page(search_results[0], auto_suggest=False)
                summary = page.summary
                
                # Make it shorter (3-4 sentences max)
                sentences = [s.strip() for s in summary.split('. ') if s.strip()]
                if len(sentences) > 3:
                    short_summary = '. '.join(sentences[:3]) + '.'
                else:
                    short_summary = summary
                
                return short_summary
                
        except (wikipedia.exceptions.DisambiguationError, wikipedia.exceptions.PageError):
            # Try English Wikipedia
            wikipedia.set_lang("en")
            search_results = wikipedia.search(company_name)
            
            if search_results:
                try:
                    page = wikipedia.page(search_results[0], auto_suggest=False)
                    summary = page.summary
                    
                    sentences = [s.strip() for s in summary.split('. ') if s.strip()]
                    if len(sentences) > 3:
                        short_summary = '. '.join(sentences[:3]) + '.'
                    else:
                        short_summary = summary
                    
                    return short_summary
                except:
                    pass
                    
    except Exception as e:
        st.warning(f"Wikipedia search note: {str(e)[:100]}")
    
    return None


def create_summary_from_brreg_data(company_data):
    """Create summary from Brønnøysund data if Wikipedia fails"""
    company_name = company_data.get('company_name', '')
    industry = company_data.get('nace_description', '')
    city = company_data.get('city', '')
    employees = company_data.get('employees', '')
    founded = company_data.get('registration_date', '')
    
    parts = []
    
    if company_name:
        if industry:
            parts.append(f"{company_name} er et {industry.lower()} selskap")
        else:
            parts.append(f"{company_name} er et norsk selskap")
        
        if city:
            parts.append(f"med hovedkontor i {city}.")
        else:
            parts.append(".")
    
    if founded:
        year = founded.split('-')[0] if '-' in founded else founded
        parts.append(f"Selskapet ble etablert i {year}.")
    
    if employees:
        parts.append(f"Det har omtrent {employees} ansatte.")
    
    if not parts:
        return f"{company_name} er et norsk selskap."
    
    return ' '.join(parts)


# =========================
# BRØNNØYSUND API - COMPANY NAME SEARCH ONLY
# =========================
@st.cache_data(ttl=3600)
def search_company_by_name(name):
    """Search company by name only (no org number search)"""
    if not name or len(name.strip()) < 2:
        st.warning("Vennligst skriv inn minst 2 tegn for selskapsnavn")
        return None
    
    search_term = name.strip()
    
    try:
        url = "https://data.brreg.no/enhetsregisteret/api/enheter"
        headers = {
            "User-Agent": "CompanyDataExtractor/1.0",
            "Accept": "application/json"
        }
        params = {
            "navn": search_term,
            "size": 5,
            "organisasjonsform": "AS,ASA,ENK,ANS,DA"
        }
        
        with st.spinner(f"Søker etter '{search_term}'..."):
            response = requests.get(url, headers=headers, params=params, timeout=30)
        
        if response.status_code == 200:
            data = response.json()
            companies = data.get("_embedded", {}).get("enheter", [])
            
            if not companies:
                st.warning(f"Ingen selskaper funnet med navn: '{search_term}'")
                return None
            
            # Find best match
            best_match = None
            best_score = 0
            
            for company in companies:
                score = 0
                company_name = company.get('navn', '').lower()
                search_lower = search_term.lower()
                
                # Exact match
                if company_name == search_lower:
                    score += 100
                
                # Starts with search term
                if company_name.startswith(search_lower):
                    score += 50
                
                # Contains search term
                if search_lower in company_name:
                    score += 30
                
                # Active companies
                if company.get('konkurs') == False:
                    score += 20
                
                if score > best_score:
                    best_score = score
                    best_match = company
            
            if best_match:
                st.success(f"✅ Funnet: {best_match.get('navn')}")
                return best_match
            
        elif response.status_code == 404:
            st.warning(f"Ingen selskaper funnet for: '{search_term}'")
        else:
            st.error(f"API feil: {response.status_code}")
            
    except requests.RequestException as e:
        st.error(f"Søk feilet: {str(e)}")
    
    return None


# =========================
# PDF TEXT EXTRACTION (OPTIONAL)
# =========================
def extract_pdf_text_improved(pdf_file):
    """Extract text from PDF (optional - for reference only)"""
    all_text = ""
    
    try:
        with pdfplumber.open(pdf_file) as pdf:
            progress_bar = st.progress(0)
            
            for i, page in enumerate(pdf.pages):
                text = page.extract_text()
                if text:
                    all_text += text + "\n\n"
                
                # Update progress
                progress_bar.progress((i + 1) / len(pdf.pages))
            
            progress_bar.empty()
            
    except Exception as e:
        st.warning(f"Kunne ikke lese PDF: {str(e)}")
    
    return all_text


# =========================
# EXCEL PROCESSING
# =========================
def load_template_from_github():
    """Load the Excel template from GitHub or local file"""
    try:
        # Try local file first (for development)
        if os.path.exists("Grundmall.xlsx"):
            with open("Grundmall.xlsx", "rb") as f:
                return BytesIO(f.read())
        
        # Try to load from GitHub
        github_url = "https://raw.githubusercontent.com/YOUR_USERNAME/YOUR_REPO/main/Grundmall.xlsx"
        response = requests.get(github_url, timeout=30)
        
        if response.status_code == 200:
            return BytesIO(response.content)
        else:
            st.error("Kunne ikke laste Excel-malen fra GitHub")
            return None
            
    except Exception as e:
        st.error(f"Feil ved lasting av mal: {str(e)}")
        return None


def update_excel_template(template_stream, company_data, company_summary):
    """Update Excel template with company data"""
    try:
        # Load workbook
        wb = load_workbook(template_stream)
        
        # Get first sheet (always use first sheet)
        ws = wb.worksheets[0]
        
        # Rename sheet 1 to company name
        company_name = company_data.get('company_name', 'Selskap')
        safe_sheet_name = clean_sheet_name(f"{company_name} Info")
        ws.title = safe_sheet_name[:31]  # Excel limit: 31 chars
        
        # Rename sheet 2 if it exists
        if len(wb.worksheets) > 1:
            ws2 = wb.worksheets[1]
            ws2.title = clean_sheet_name(f"{company_name} Anbud")[:31]
        
        # Update BIG INFORMATION WINDOW (A2:D13 merged cell)
        # First, check if cells are already merged
        cell_range = 'A2:D13'
        
        # Unmerge if already merged
        for merged_range in list(ws.merged_cells.ranges):
            if str(merged_range) == cell_range:
                ws.unmerge_cells(str(merged_range))
        
        # Merge the cells
        ws.merge_cells(cell_range)
        
        # Set the company summary text
        if company_summary:
            ws['A2'] = company_summary
        else:
            ws['A2'] = f"Informasjon om {company_name}"
        
        # Style the merged cell
        ws['A2'].alignment = Alignment(
            wrap_text=True,
            vertical='top',
            horizontal='left'
        )
        
        # Adjust row height for better visibility
        for row in range(2, 14):
            ws.row_dimensions[row].height = 18
        
        # Update company data in column B
        data_mapping = {
            'company_name': 'B14',    # Kunde
            'org_number': 'B15',      # Org-nr
            'address': 'B16',         # Adresse
            'post_nr': 'B17',         # Post-nr
            'nace_code': 'B18',       # NACE-kode
            'homepage': 'B20',        # Hjemmeside
            'employees': 'B21'        # Number of Employees
        }
        
        # Update each field
        for field, cell in data_mapping.items():
            value = company_data.get(field, '')
            if value:
                if field == 'org_number' and len(str(value)) == 9:
                    ws[cell] = f"'{value}"  # Keep leading zeros
                else:
                    ws[cell] = str(value)
        
        # Revenue 2024 (B19) - placeholder for now
        ws['B19'] = "Data ikke tilgjengelig"
        
        # Save to BytesIO
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        
        return output
        
    except Exception as e:
        st.error(f"Excel oppdatering feilet: {str(e)}")
        
        # Fallback: create simple Excel
        try:
            df = pd.DataFrame([company_data])
            output = BytesIO()
            df.to_excel(output, index=False, engine='openpyxl')
            output.seek(0)
            return output
        except:
            # Last resort
            wb = load_workbook()
            ws = wb.active
            ws.title = company_data.get('company_name', 'Selskap')[:31]
            ws['A1'] = "Feil ved oppdatering. Data:"
            for i, (key, value) in enumerate(company_data.items(), 2):
                ws.cell(row=i, column=1, value=key)
                ws.cell(row=i, column=2, value=value)
            
            output = BytesIO()
            wb.save(output)
            output.seek(0)
            return output


def clean_sheet_name(name):
    """Clean sheet name for Excel (remove invalid characters)"""
    invalid_chars = ['\\', '/', '*', '?', ':', '[', ']']
    for char in invalid_chars:
        name = name.replace(char, '')
    return name.strip()


def format_company_data(api_data):
    """Format API response into structured data"""
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
    
    # Address information
    addr = api_data.get("forretningsadresse", {})
    if addr:
        address_lines = addr.get("adresse", [])
        if isinstance(address_lines, list):
            formatted["address"] = ", ".join(filter(None, address_lines))
        
        formatted["post_nr"] = addr.get("postnummer", "")
        formatted["city"] = addr.get("poststed", "")
    
    # NACE code
    nace = api_data.get("naeringskode1", {})
    if nace:
        formatted["nace_code"] = nace.get("kode", "")
        formatted["nace_description"] = nace.get("beskrivelse", "")
    
    return formatted


# =========================
# STREAMLIT UI
# =========================
def main():
    # Sidebar
    with st.sidebar:
        st.title("⚙️ Innstillinger")
        
        st.markdown("---")
        st.subheader("Instruksjoner")
        st.markdown("""
        1. **Last opp PDF** (valgfritt) - for referanse
        2. **Skriv inn selskapsnavn**
        3. Klikk **Prosesser** for å hente data
        4. **Last ned** oppdatert Excel-fil
        
        **Funksjoner:**
        - Automatisk søk i Brønnøysund
        - Wikipedia-sammendrag om selskapet
        - Excel-mal fra GitHub
        - Arknavn endres til selskapsnavn
        """)
        
        st.markdown("---")
        st.caption(f"Sist oppdatert: {datetime.now().strftime('%d.%m.%Y %H:%M')}")
    
    # Main content
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.title("📄 PDF → Excel (Brønnøysund)")
        st.markdown("Hent selskapsinformasjon og oppdater Excel automatisk")
    
    with col2:
        st.image("https://img.icons8.com/color/96/000000/parse-from-clipboard.png", width=80)
    
    st.markdown("---")
    
    # File upload section
    col_upload1, col_upload2 = st.columns(2)
    
    with col_upload1:
        st.subheader("📂 Last opp filer")
        
        pdf_file = st.file_uploader(
            "PDF dokument (valgfritt)",
            type="pdf",
            help="Last opp PDF for referanse",
            key="pdf_upload"
        )
    
    with col_upload2:
        st.subheader("🏢 Selskapsinformasjon")
        
        company_name = st.text_input(
            "Selskapsnavn *",
            placeholder="F.eks. Equinor ASA",
            help="Skriv inn fullt navn på selskapet",
            key="company_name_input"
        )
    
    st.markdown("---")
    
    # Load template at startup
    if 'template_loaded' not in st.session_state:
        with st.spinner("Laster Excel-mal..."):
            template_stream = load_template_from_github()
            if template_stream:
                st.session_state.template_stream = template_stream
                st.session_state.template_loaded = True
                st.success("✅ Excel-mal lastet")
            else:
                st.error("❌ Kunne ikke laste Excel-mal")
    
    # Process button
    if st.button("🚀 Prosesser & Oppdater Excel", type="primary", use_container_width=True):
        if not company_name or not company_name.strip():
            st.error("❌ Vennligst skriv inn et selskapsnavn")
            st.stop()
        
        if not st.session_state.get('template_loaded', False):
            st.error("❌ Excel-mal ikke tilgjengelig")
            st.stop()
        
        # Initialize
        extracted_data = {}
        api_data = None
        
        # Step 1: Extract from PDF (optional)
        if pdf_file:
            with st.expander("📊 PDF-ekstraksjon", expanded=False):
                st.write("**Trinn 1:** Leser PDF...")
                pdf_text = extract_pdf_text_improved(pdf_file)
                
                if pdf_text:
                    st.info(f"Lest {len(pdf_text)} tegn fra PDF")
                    # Display first 500 characters
                    st.text_area("Forhåndsvisning:", pdf_text[:500] + "..." if len(pdf_text) > 500 else pdf_text, height=150)
        
        # Step 2: Search Brønnøysund
        st.write("**Trinn 2:** Søker i Brønnøysund...")
        api_data = search_company_by_name(company_name.strip())
        
        if not api_data:
            st.error("❌ Fant ikke selskapet i Brønnøysund. Sjekk navnet og prøv igjen.")
            st.stop()
        
        # Step 3: Format data
        formatted_data = format_company_data(api_data)
        st.session_state.extracted_data = formatted_data
        st.session_state.api_response = api_data
        
        # Step 4: Get Wikipedia summary
        st.write("**Trinn 3:** Søker etter selskapsopplysninger...")
        
        company_summary = None
        with st.spinner("Søker på Wikipedia..."):
            company_summary = get_company_summary_from_wikipedia(company_name)
            
            if not company_summary:
                st.info("Fant ikke Wikipedia-artikkel. Lager sammendrag fra Brønnøysund-data.")
                company_summary = create_summary_from_brreg_data(formatted_data)
        
        st.session_state.company_summary = company_summary
        
        # Step 5: Update Excel
        st.write("**Trinn 4:** Oppdaterer Excel...")
        
        try:
            updated_excel = update_excel_template(
                st.session_state.template_stream,
                formatted_data,
                company_summary
            )
            
            st.session_state.excel_ready = True
            st.session_state.excel_file = updated_excel
            
            st.success("✅ Excel-fil oppdatert!")
            
            # Show where information was placed
            st.info(f"""
            **Informasjon plassert i:**
            - **Ark 1:** {clean_sheet_name(f"{formatted_data.get('company_name', 'Selskap')} Info")[:31]}
            - **Stort informasjonsvindu:** Celle A2:D13 (sammendrag)
            - **Selskapsdata:** Celler B14-B21
            - **Ark 2:** {clean_sheet_name(f"{formatted_data.get('company_name', 'Selskap')} Anbud")[:31]}
            - **Ark 3:** Skader (uendret)
            """)
            
        except Exception as e:
            st.error(f"❌ Feil ved Excel-oppdatering: {str(e)}")
    
    # Display results if available
    if st.session_state.extracted_data:
        st.markdown("---")
        st.subheader("📋 Ekstraherte data")
        
        col_data1, col_data2 = st.columns(2)
        
        with col_data1:
            st.write("**Selskapsinformasjon:**")
            data = st.session_state.extracted_data
            
            fields_to_show = [
                ("company_name", "Selskapsnavn"),
                ("org_number", "Organisasjonsnummer"),
                ("address", "Adresse"),
                ("post_nr", "Postnummer"),
                ("city", "Poststed"),
                ("nace_description", "NACE-bransje"),
                ("employees", "Antall ansatte"),
                ("homepage", "Hjemmeside")
            ]
            
            for field_key, field_label in fields_to_show:
                value = data.get(field_key)
                if value:
                    st.write(f"**{field_label}:** {value}")
        
        with col_data2:
            if st.session_state.company_summary:
                st.write("**Sammendrag (går i celle A2:D13):**")
                st.info(st.session_state.company_summary)
    
    # Download section
    if st.session_state.get('excel_ready', False):
        st.markdown("---")
        st.subheader("📥 Last ned")
        
        company_name = st.session_state.extracted_data.get('company_name', 'selskap')
        safe_name = re.sub(r'[^\w\s-æøåÆØÅ]', '', company_name)
        safe_name = re.sub(r'[-\s]+', '_', safe_name)
        
        download_filename = f"{safe_name}_{datetime.now().strftime('%Y%m%d')}.xlsx"
        
        st.download_button(
            label="⬇️ Last ned oppdatert Excel",
            data=st.session_state.excel_file,
            file_name=download_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary"
        )
    
    # Footer
    st.markdown("---")
    st.caption("""
    Drevet av Brønnøysund Enhetsregisteret API og Wikipedia | 
    Data er mellomlagret i 1 time
    """)


if __name__ == "__main__":
    main()
