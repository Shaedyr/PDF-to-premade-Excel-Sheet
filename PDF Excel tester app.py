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
def get_company_summary_from_wikipedia(company_name):
    """Get company summary from Wikipedia"""
    try:
        search_name = company_name
        for suffix in [" AS", " ASA", " ANS", " DA", " ENK", " KS"]:
            if search_name.endswith(suffix):
                search_name = search_name[:-len(suffix)].strip()
                break
        
        search_attempts = [search_name, company_name, search_name + " (bedrift)", search_name + " (company)"]
        
        wikipedia.set_lang("no")
        for attempt in search_attempts:
            try:
                search_results = wikipedia.search(attempt)
                if search_results:
                    company_results = []
                    for result in search_results[:3]:
                        result_lower = result.lower()
                        if any(kw in result_lower for kw in ["as", "asa", "bedrift", "selskap", "company", "group"]):
                            company_results.append(result)
                    
                    if company_results:
                        page = wikipedia.page(company_results[0], auto_suggest=False)
                        summary = page.summary
                        sentences = [s.strip() for s in summary.split('. ') if s.strip()]
                        if len(sentences) > 2:
                            return '. '.join(sentences[:2]) + '.'
                        else:
                            return summary[:300] + '...' if len(summary) > 300 else summary
            except (wikipedia.exceptions.DisambiguationError, wikipedia.exceptions.PageError):
                continue
        
        wikipedia.set_lang("en")
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
    except:
        pass
    return None

def get_company_summary_from_web(company_name):
    """Search web for company information"""
    try:
        search_name = company_name
        for suffix in [" AS", " ASA", " ANS", " DA", " ENK", " KS", " BA"]:
            if search_name.endswith(suffix):
                search_name = search_name[:-len(suffix)].strip()
                break
        
        ddg_url = f"https://api.duckduckgo.com/?q={search_name}+bedrift&format=json&no_html=1&skip_disambig=1"
        response = requests.get(ddg_url, timeout=10)
        
        if response.status_code == 200:
            data = response.json()
            abstract = data.get('AbstractText', '')
            if abstract and len(abstract) > 50:
                return abstract
            
            ddg_url_en = f"https://api.duckduckgo.com/?q={search_name}+company&format=json&no_html=1&skip_disambig=1"
            response_en = requests.get(ddg_url_en, timeout=10)
            
            if response_en.status_code == 200:
                data_en = response_en.json()
                abstract_en = data_en.get('AbstractText', '')
                if abstract_en and len(abstract_en) > 50:
                    return abstract_en.replace(" is a ", " er et ").replace(" company", " selskap").replace(" based in ", " med hovedkontor i ")
    except:
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
# BRØNNØYSUND API
# =========================
@st.cache_data(ttl=3600)
def search_company_by_name(name):
    """Search company by name"""
    if not name or len(name.strip()) < 2:
        st.warning("Skriv inn minst 2 tegn")
        return None
    
    search_term = name.strip().upper()  # Convert to uppercase for consistent comparison
    try:
        url = "https://data.brreg.no/enhetsregisteret/api/enheter"
        params = {"navn": search_term, "size": 20, "organisasjonsform": "AS,ASA,ENK,ANS,DA"}
        
        with st.spinner(f"Søker etter '{search_term}'..."):
            response = requests.get(url, params=params, timeout=30)
        
        if response.status_code == 200:
            data = response.json()
            companies = data.get("_embedded", {}).get("enheter", [])
            
            if companies:
                # First, try exact case-insensitive match
                exact_match = None
                for company in companies:
                    company_name_upper = company.get('navn', '').upper()
                    if company_name_upper == search_term:
                        exact_match = company
                        break
                
                if exact_match:
                    st.success(f"✅ Funnet eksakt treff: {exact_match.get('navn')}")
                    return exact_match
                
                # If no exact match, try case-insensitive startswith
                startswith_match = None
                for company in companies:
                    company_name_upper = company.get('navn', '').upper()
                    if company_name_upper.startswith(search_term):
                        startswith_match = company
                        break
                
                if startswith_match:
                    st.success(f"✅ Funnet delvis treff: {startswith_match.get('navn')}")
                    return startswith_match
                
                # If still no match, use the original scoring system but with exact word matching
                best_match = None
                best_score = 0
                
                # Split search term into words for better matching
                search_words = search_term.split()
                
                for company in companies:
                    score = 0
                    company_name_upper = company.get('navn', '').upper()
                    
                    # Check if all search words are in company name
                    all_words_match = all(word in company_name_upper for word in search_words)
                    if all_words_match:
                        score += 80
                    
                    # Calculate percentage match
                    if search_term in company_name_upper:
                        match_percentage = (len(search_term) / len(company_name_upper)) * 100
                        score += match_percentage
                    
                    if company.get('konkurs') == False:
                        score += 10
                    
                    if score > best_score:
                        best_score = score
                        best_match = company
                
                if best_match:
                    st.success(f"✅ Funnet: {best_match.get('navn')}")
                    return best_match
                
                # If no good match found, return the first result with warning
                if companies:
                    st.warning(f"Ingen eksakt treff. Fant: {companies[0].get('navn')}")
                    return companies[0]
                    
            else:
                st.warning(f"Ingen selskaper funnet: '{search_term}'")
                return None
    except Exception as e:
        st.error(f"Søk feilet: {str(e)}")
    
    return None

# =========================
# Dropdown Function
# =========================
def search_company_by_name_with_selection(name):
    """Search company by name with manual selection option"""
    if not name or len(name.strip()) < 2:
        st.warning("Skriv inn minst 2 tegn")
        return None
    
    search_term = name.strip().upper()
    try:
        url = "https://data.brreg.no/enhetsregisteret/api/enheter"
        params = {"navn": search_term, "size": 10, "organisasjonsform": "AS,ASA,ENK,ANS,DA"}
        
        with st.spinner(f"Søker etter '{search_term}'..."):
            response = requests.get(url, params=params, timeout=30)
        
        if response.status_code == 200:
            data = response.json()
            companies = data.get("_embedded", {}).get("enheter", [])
            
            if companies:
                if len(companies) == 1:
                    st.success(f"✅ Funnet: {companies[0].get('navn')}")
                    return companies[0]
                else:
                    # Show selection dropdown if multiple companies found
                    company_names = [f"{c.get('navn')} ({c.get('organisasjonsnummer')})" for c in companies]
                    selected = st.selectbox(
                        "Flere selskaper funnet. Velg ett:",
                        company_names,
                        key="company_selection"
                    )
                    
                    if selected:
                        org_num = selected.split('(')[-1].strip(')')
                        for company in companies:
                            if company.get('organisasjonsnummer') == org_num:
                                return company
                    
                    return None
            else:
                st.warning(f"Ingen selskaper funnet: '{search_term}'")
    except Exception as e:
        st.error(f"Søk feilet: {str(e)}")
    
    return None

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
    """Load Excel template from GitHub"""
    try:
        if os.path.exists("Grundmall.xlsx"):
            with open("Grundmall.xlsx", "rb") as f:
                return BytesIO(f.read())
        
        github_url = "https://raw.githubusercontent.com/Shaedyr/PDF-to-premade-Excel-Sheet/main/PremadeExcelTemplate.xlsx"
        response = requests.get(github_url, timeout=30)
        
        if response.status_code == 200:
            return BytesIO(response.content)
        else:
            st.error("Kunne ikke laste Excel-malen")
            return None
    except:
        st.error("Feil ved lasting av mal")
        return None

def update_excel_template(template_stream, company_data, company_summary):
    """Update Excel template"""
    try:
        wb = load_workbook(template_stream)
        ws = wb.worksheets[0]
        
        company_name = company_data.get('company_name', 'Selskap')
        ws.title = re.sub(r'[\\/*?:[\]]', '', f"{company_name} Info")[:31]
        
        if len(wb.worksheets) > 1:
            wb.worksheets[1].title = re.sub(r'[\\/*?:[\]]', '', f"{company_name} Anbud")[:31]
        
        # Update info window (A2:D13)
        for merged_range in list(ws.merged_cells.ranges):
            if str(merged_range) == 'A2:D13':
                ws.unmerge_cells(str(merged_range))
        
        ws.merge_cells('A2:D13')
        ws['A2'] = company_summary if company_summary else f"Informasjon om {company_name}"
        ws['A2'].alignment = Alignment(wrap_text=True, vertical='top', horizontal='left')
        
        for row in range(2, 14):
            ws.row_dimensions[row].height = 18
        
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
                
                if field == 'org_number' and len(str(value)) == 9:
                    ws[cell] = f"'{value}"
                else:
                    ws[cell] = str(value)
                
                if merge_to and len(str(value)) > 20:
                    try:
                        ws.merge_cells(f"{cell}:{merge_to}")
                        ws[cell].alignment = Alignment(wrap_text=True, vertical='center', horizontal='left')
                    except:
                        pass
        
        # Handle NACE separately - combine code and description in B18
        nace_code = company_data.get('nace_code', '')
        nace_description = company_data.get('nace_description', '')
        
        if nace_code and nace_description:
            # Combine both: "Description (Code)"
            nace_combined = f"{nace_description} ({nace_code})"
            ws['B18'] = nace_combined
            try:
                ws.merge_cells('B18:D18')  # Merge B18 to D18
                ws['B18'].alignment = Alignment(wrap_text=True, vertical='center', horizontal='left')
            except:
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
        return output
        
    except Exception as e:
        st.error(f"Excel oppdatering feilet: {str(e)}")
        df = pd.DataFrame([company_data])
        output = BytesIO()
        df.to_excel(output, index=False)
        output.seek(0)
        return output

# =========================
# STREAMLIT UI
# =========================
def main():
    # REMOVED THE ENTIRE SIDEBAR SECTION
    
    st.title("📄 PDF → Excel (Brønnøysund)")
    st.markdown("Hent selskapsinformasjon og oppdater Excel automatisk")
    st.markdown("---")
    
    col1, col2 = st.columns(2)
    with col1:
        pdf_file = st.file_uploader("PDF dokument (valgfritt)", type="pdf", help="Last opp PDF for referanse")
    with col2:
        company_name = st.text_input("Selskapsnavn *", placeholder="F.eks. Equinor ASA", help="Skriv inn fullt navn")
    
    st.markdown("---")
    
    if 'template_loaded' not in st.session_state:
        with st.spinner("Laster Excel-mal..."):
            template_stream = load_template_from_github()
            if template_stream:
                st.session_state.template_stream = template_stream
                st.session_state.template_loaded = True
                st.success("✅ Excel-mal lastet")
            else:
                st.error("❌ Kunne ikke laste Excel-mal")
    
    if st.button("🚀 Prosesser & Oppdater Excel", type="primary", use_container_width=True):
        if not company_name:
            st.error("❌ Vennligst skriv inn selskapsnavn")
            st.stop()
        
        if not st.session_state.get('template_loaded'):
            st.error("❌ Excel-mal ikke tilgjengelig")
            st.stop()
        
        # Step 1: Search Brønnøysund
        st.write("**Trinn 1:** Søker i Brønnøysund...")
        api_data = search_company_by_name(company_name.strip())
        
        if not api_data:
            st.error("❌ Fant ikke selskapet. Sjekk navnet.")
            st.stop()
        
        # Step 2: Format data
        formatted_data = format_company_data(api_data)
        st.session_state.extracted_data = formatted_data
        st.session_state.api_response = api_data
        
        # Step 3: Get company summary (3-step approach)
        st.write("**Trinn 2:** Søker etter selskapsopplysninger...")
        
        company_summary = None
        
        # Step 1: Try Wikipedia
        with st.spinner("Trinn 1: Søker på Wikipedia..."):
            company_summary = get_company_summary_from_wikipedia(company_name)
        
        # Step 2: If Wikipedia fails, try web search
        if not company_summary:
            st.info("Fant ikke på Wikipedia. Prøver websøk...")
            with st.spinner("Trinn 2: Søker på nettet..."):
                company_summary = get_company_summary_from_web(company_name)
        
        # Step 3: If web search also fails, use Brønnøysund data analysis
        if not company_summary:
            st.info("Fant ikke på nettet. Lager analyse fra Brønnøysund-data...")
            company_summary = create_summary_from_brreg_data(formatted_data)
        
        st.session_state.company_summary = company_summary
        
        # Step 4: Update Excel
        st.write("**Trinn 3:** Oppdaterer Excel...")
        
        try:
            updated_excel = update_excel_template(
                st.session_state.template_stream,
                formatted_data,
                company_summary
            )
            
            st.session_state.excel_ready = True
            st.session_state.excel_file = updated_excel
            st.success("✅ Excel-fil oppdatert!")
            
            st.info(f"""
            **Informasjon plassert i:**
            - **Ark 1:** {re.sub(r'[\\/*?:[\]]', '', f"{formatted_data.get('company_name', 'Selskap')} Info")[:31]}
            - **Stort informasjonsvindu:** Celle A2:D13
            - **Selskapsdata:** Celler B14-B21
            - **Ark 2:** {re.sub(r'[\\/*?:[\]]', '', f"{formatted_data.get('company_name', 'Selskap')} Anbud")[:31]}
            """)
            
        except Exception as e:
            st.error(f"❌ Feil ved Excel-oppdatering: {str(e)}")
    
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
    
    if st.session_state.get('excel_ready', False):
        st.markdown("---")
        st.subheader("📥 Last ned")
        
        company_name_dl = st.session_state.extracted_data.get('company_name', 'selskap')
        safe_name = re.sub(r'[^\w\s-]', '', company_name_dl, flags=re.UNICODE)
        safe_name = re.sub(r'[-\s]+', '_', safe_name)
        
        st.download_button(
            label="⬇️ Last ned oppdatert Excel",
            data=st.session_state.excel_file,
            file_name=f"{safe_name}_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary"
        )
    
    st.markdown("---")
    st.caption("Drevet av Brønnøysund Enhetsregisteret API | Data mellomlagret i 1 time")

if __name__ == "__main__":
    main()
