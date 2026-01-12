import streamlit as st
import pdfplumber
import re
import requests
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from io import BytesIO
import time


# =========================
# BRØNNØYSUND API
# =========================
def lookup_by_org_number(org_number):
    """Lookup company by organization number with better error handling"""
    if not org_number or not str(org_number).strip().isdigit():
        return None
    
    org_number = str(org_number).strip()
    if len(org_number) != 9:
        return None
    
    try:
        url = f"https://data.brreg.no/enhetsregisteret/api/enheter/{org_number}"
        headers = {
            "User-Agent": "CompanyDataExtractor/1.0",
            "Accept": "application/json"
        }
        r = requests.get(url, headers=headers, timeout=15)
        
        if r.status_code == 200:
            return r.json()
        elif r.status_code == 404:
            st.warning(f"Organization number {org_number} not found in Brønnøysund")
        else:
            st.warning(f"API returned status code: {r.status_code}")
            
    except requests.Timeout:
        st.error("Request to Brønnøysund API timed out. Please try again.")
    except requests.RequestException as e:
        st.error(f"Network error: {str(e)}")
    except Exception as e:
        st.error(f"Unexpected error during API call: {str(e)}")
    
    return None


def search_company_by_name(name):
    """Search company by name with improved matching"""
    if not name or len(name.strip()) < 2:
        return None
    
    try:
        url = "https://data.brreg.no/enhetsregisteret/api/enheter"
        headers = {
            "User-Agent": "CompanyDataExtractor/1.0",
            "Accept": "application/json"
        }
        params = {
            "navn": name.strip(),
            "size": 5  # Get top 5 results
        }
        r = requests.get(url, headers=headers, params=params, timeout=15)
        
        if r.status_code == 200:
            data = r.json()
            companies = data.get("_embedded", {}).get("enheter", [])
            
            if companies:
                # Try to find exact match first
                exact_matches = [c for c in companies if c.get("navn", "").lower() == name.lower()]
                if exact_matches:
                    return exact_matches[0]
                # Otherwise return the first result
                return companies[0]
        elif r.status_code == 404:
            st.warning(f"No companies found with name: '{name}'")
            
    except requests.RequestException as e:
        st.error(f"Search failed: {str(e)}")
    
    return None


# =========================
# PDF TEXT EXTRACTION
# =========================
def extract_pdf_text(pdf_file):
    """Extract text from PDF with improved error handling"""
    text = ""
    try:
        with pdfplumber.open(pdf_file) as pdf:
            total_pages = len(pdf.pages)
            progress_bar = st.progress(0)
            
            for i, page in enumerate(pdf.pages):
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"
                
                # Update progress
                progress_bar.progress((i + 1) / total_pages)
                time.sleep(0.01)  # Small delay for smoother progress bar
                
            progress_bar.empty()
            
            if not text.strip():
                st.warning("No readable text found in PDF. It might be scanned or image-based.")
                
    except pdfplumber.PDFSyntaxError:
        st.error("Invalid or corrupted PDF file.")
        return ""
    except Exception as e:
        st.error(f"Failed to read PDF: {str(e)}")
        return ""
    
    return text


# =========================
# IMPROVED PDF FIELD EXTRACTION
# =========================
def extract_fields_from_text(text):
    """Extract company information from text with multiple patterns"""
    fields = {}
    
    if not text.strip():
        return fields
    
    # Multiple patterns for organization number
    org_patterns = [
        r"organisasjonsnummer[:\s]*([0-9]{9})",
        r"org\.?nr\.?[:\s]*([0-9]{9})",
        r"org[:\s]*([0-9]{9})",
        r"([0-9]{9})\s*\(orgnr\)",
        r"\b([0-9]{9})\b(?=[^0-9]*\b(?:org|organisasjonsnummer)\b)"
    ]
    
    org_number = None
    for pattern in org_patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            org_number = match.group(1)
            break
    
    # Try to find any 9-digit number that could be an org number
    if not org_number:
        potential_numbers = re.findall(r'\b\d{9}\b', text)
        for num in potential_numbers:
            # Basic validation: Norwegian org numbers have specific patterns
            if num.startswith(('8', '9')) or (800000000 <= int(num) <= 999999999):
                org_number = num
                break
    
    fields["org_number"] = org_number if org_number else ""
    
    # Try to extract company name
    company_patterns = [
        r"(?:Firma|Selskap|Company)[:\s]*(.+)",
        r"(.+(?:AS|ASA|DA|ANS|ENK|KS)\b)"
    ]
    
    company_name = None
    for pattern in company_patterns:
        match = re.search(pattern, text, re.IGNORECASE | re.MULTILINE)
        if match:
            company_name = match.group(1).strip()
            # Clean up the name
            company_name = re.sub(r'[\n\r]+', ' ', company_name)
            break
    
    fields["company_name"] = company_name if company_name else ""
    
    return fields


# =========================
# EXCEL UPDATE
# =========================
def update_excel(template_file, data, summary):
    """Update Excel template with company data"""
    try:
        wb = load_workbook(template_file)
        ws = wb.active
        
        # Verify template has required cells
        required_cells = ["B14", "B15", "B16", "B17", "B18", "B21", "B22", "B10"]
        for cell in required_cells:
            if ws[cell] is None:
                st.warning(f"Cell {cell} not found in template. Using available cells.")
        
    except Exception as e:
        st.error(f"Invalid Excel template: {str(e)}")
        raise

    # Cell mapping with validation
    cell_mapping = {
        "company_name": "B14",
        "org_number": "B15",
        "address": "B16",
        "post_nr": "B17",
        "nace_code": "B18",
        "homepage": "B21",
        "employees": "B22",
    }

    # Update cells with data
    for field, cell in cell_mapping.items():
        value = data.get(field, "")
        if value:
            ws[cell] = str(value)
    
    # Update summary with wrapping
    if summary:
        ws["B10"] = f"Kort info om företaget:\n{summary}"
        ws["B10"].alignment = Alignment(wrap_text=True, vertical='top')
    
    # Adjust column width for better display
    for column in ['A', 'B', 'C']:
        ws.column_dimensions[column].width = 25
    
    # Save to BytesIO
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# =========================
# STREAMLIT UI
# =========================
st.set_page_config(
    page_title="PDF → Excel (Brreg)",
    layout="centered",
    page_icon="📊"
)

st.title("📄➡️📊 PDF → Excel (Brønnøysund)")
st.markdown("---")

# Sidebar for instructions
with st.sidebar:
    st.header("Instructions")
    st.markdown("""
    1. **Upload Excel Template** (required)
    2. **Upload PDF** (optional - for automatic data extraction)
    3. **Enter company name** if not found in PDF
    4. Click **"Extract & Update Excel"**
    
    The app will:
    - Extract data from PDF (if provided)
    - Fetch company details from Brønnøysund
    - Update your Excel template
    - Provide download link
    """)

# Main content area
col1, col2 = st.columns(2)

with col1:
    st.subheader("Upload Files")
    excel_file = st.file_uploader(
        "Upload Excel template *",
        type="xlsx",
        help="Required: Your Excel template file"
    )
    
    pdf_file = st.file_uploader(
        "Upload PDF (optional)",
        type="pdf",
        help="Optional: PDF containing company information"
    )

with col2:
    st.subheader("Company Information")
    
    manual_company_name = st.text_input(
        "Company name *",
        placeholder="e.g. Eksempel AS",
        help="Required if not found in PDF"
    )
    
    manual_org_number = st.text_input(
        "Organization number (optional)",
        placeholder="e.g. 999999999",
        help="Optional: 9-digit Norwegian organization number"
    )
    
    st.caption("(*) Required fields")

st.markdown("---")

# Processing section
if st.button("🚀 Extract & Update Excel", type="primary", use_container_width=True):
    if not excel_file:
        st.error("❌ Please upload an Excel template file.")
        st.stop()
    
    if not manual_company_name and not pdf_file:
        st.error("❌ Please provide either a PDF file or a company name.")
        st.stop()
    
    with st.spinner("Processing your request..."):
        st.info("Step 1: Extracting data from sources")
        extracted = {}
        
        # STEP 1 — Extract from PDF if provided
        if pdf_file:
            with st.expander("PDF Extraction Details", expanded=False):
                st.write("Extracting text from PDF...")
                pdf_text = extract_pdf_text(pdf_file)
                if pdf_text:
                    extracted = extract_fields_from_text(pdf_text)
                    st.code(pdf_text[:500] + "..." if len(pdf_text) > 500 else pdf_text)
        
        # STEP 2 — Determine lookup keys (manual input overrides PDF)
        company_name = manual_company_name.strip() if manual_company_name else extracted.get("company_name", "")
        org_number = manual_org_number.strip() if manual_org_number else extracted.get("org_number", "")
        
        st.info("Step 2: Fetching from Brønnøysund")
        
        # STEP 3 — Brønnøysund lookup (try org number first, then name)
        company_data = None
        lookup_source = ""
        
        if org_number:
            with st.status("Looking up by organization number..."):
                company_data = lookup_by_org_number(org_number)
                if company_data:
                    lookup_source = f"Organization number: {org_number}"
        
        if not company_data and company_name:
            with st.status(f"Searching by company name: '{company_name}'..."):
                company_data = search_company_by_name(company_name)
                if company_data:
                    lookup_source = f"Company name: {company_name}"
        
        # STEP 4 — Normalize and prepare data
        if company_data:
            st.success(f"✅ Company found via {lookup_source}")
            
            # Extract and format data
            extracted["company_name"] = company_data.get("navn", "")
            extracted["org_number"] = company_data.get("organisasjonsnummer", "")
            
            # Address handling
            addr = company_data.get("forretningsadresse") or {}
            address_parts = addr.get("adresse", [])
            if isinstance(address_parts, list):
                extracted["address"] = ", ".join(filter(None, address_parts))
            else:
                extracted["address"] = str(address_parts)
            
            extracted["post_nr"] = addr.get("postnummer", "")
            extracted["poststed"] = addr.get("poststed", "")
            
            # NACE code
            nace = company_data.get("naeringskode1", {})
            extracted["nace_code"] = nace.get("kode", "")
            nace_description = nace.get("beskrivelse", "")
            
            # Other info
            extracted["homepage"] = company_data.get("hjemmeside", "")
            extracted["employees"] = str(company_data.get("antallAnsatte", ""))
            
            # Create summary
            summary_parts = [extracted["company_name"]]
            if nace_description:
                summary_parts.append(f"NACE: {nace_description}")
            if extracted["employees"]:
                summary_parts.append(f"Employees: {extracted['employees']}")
            
            summary = " | ".join(summary_parts)
            
        else:
            st.warning("⚠️ Company not found in Brønnøysund. Using extracted data only.")
            summary = extracted.get("company_name", "Company information not found")
        
        # Display extracted data
        st.info("Step 3: Data to be inserted")
        with st.expander("View extracted data", expanded=True):
            st.json({k: v for k, v in extracted.items() if v})
        
        st.info("Step 4: Updating Excel template")
        
        # Update Excel
        try:
            updated_excel = update_excel(excel_file, extracted, summary)
            
            st.success("🎉 Excel updated successfully!")
            
            # Download button
            col1, col2, col3 = st.columns([1, 2, 1])
            with col2:
                st.download_button(
                    label="📥 Download Updated Excel",
                    data=updated_excel,
                    file_name=f"updated_{extracted.get('company_name', 'template').replace(' ', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            
            
        except Exception as e:
            st.error(f"❌ Failed to update Excel: {str(e)}")

# Footer
st.markdown("---")
st.caption("Powered by Brønnøysund Enhetsregisteret API | Data extraction tool")
