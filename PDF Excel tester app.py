import os, re, requests, wikipedia
import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Alignment

st.set_page_config(page_title="PDF → Excel (Brønnøysund)", layout="wide", page_icon="📊")
for k, v in {"extracted_data": {}, "api_response": None, "excel_ready": False, "company_summary": ""}.items():
    if k not in st.session_state: st.session_state[k] = v

def _strip_suffix(name: str):
    return re.sub(r'\b(AS|ASA|ANS|DA|ENK|KS|BA)\b\.?$', '', (name or ''), flags=re.I).strip()

def _wiki_summary(name: str):
    if not name: return None
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
                if not results: continue
                target = next((r for r in results[:5] if any(w in r.lower() for w in ("as","asa","bedrift","selskap","company","group"))), results[0])
                page = wikipedia.page(target, auto_suggest=False)
                s = page.summary or ""
                sent = [x.strip() for x in s.split('. ') if x.strip()]
                if len(sent) > 2: return '. '.join(sent[:2]) + '.'
                short = s[:300] + '...' if len(s) > 300 else s
                return short if lang == "no" else short.replace(" is a ", " er et ").replace(" company", " selskap").replace(" based in ", " med hovedkontor i ")
            except (wikipedia.exceptions.DisambiguationError, wikipedia.exceptions.PageError):
                continue
            except Exception:
                continue
    return None

def _web_summary(name: str):
    if not name: return None
    q = _strip_suffix(name)
    for term in ("bedrift", "company"):
        try:
            url = f"https://api.duckduckgo.com/?q={requests.utils.quote(q + ' ' + term)}&format=json&no_html=1&skip_disambig=1"
            r = requests.get(url, timeout=10)
            if r.status_code == 200:
                txt = r.json().get("AbstractText", "") or ""
                if len(txt) > 50:
                    return txt if term == "bedrift" else txt.replace(" is a ", " er et ").replace(" company", " selskap").replace(" based in ", " med hovedkontor i ")
        except Exception:
            continue
    return None

def create_summary_from_brreg_data(d: dict):
    name = d.get("company_name","")
    if not name: return "Ingen informasjon tilgjengelig om dette selskapet."
    parts = []
    industry, city, emp, reg = d.get("nace_description",""), d.get("city",""), d.get("employees",""), d.get("registration_date","")
    if industry and city: parts.append(f"{name} driver {industry.lower()} virksomhet fra {city}.")
    elif industry: parts.append(f"{name} opererer innen {industry.lower()}.")
    else: parts.append(f"{name} er et registrert norsk selskap.")
    if reg:
        try:
            year = int(reg.split('-')[0]) if '-' in reg else int(reg)
            age = datetime.now().year - year
            if age > 30: parts.append(f"Etablert i {year}, har selskapet over {age} års bransjeerfaring.")
            elif age > 10: parts.append(f"Selskapet har utviklet seg over {age} år siden etableringen i {year}.")
            else: parts.append(f"Etablert i {year}, er dette et yngre selskap i vekstfasen.")
        except Exception:
            parts.append(f"Selskapet ble registrert i {reg}.")
    if emp:
        try:
            e = int(emp)
            if e > 200: parts.append(f"Som en større arbeidsgiver med {e} ansatte, har det betydelig samfunnspåvirkning.")
            elif e > 50: parts.append(f"Med {e} ansatte representerer det et mellomstort foretak.")
            elif e > 10: parts.append(f"Selskapet sysselsetter {e} personer.")
            else: parts.append(f"Dette er et mindre selskap med {e} ansatte.")
        except Exception:
            pass
    if len(parts) < 3: parts.append("Virksomheten er registrert og i god stand i Brønnøysundregistrene.")
    s = ' '.join(parts)
    return s[:797] + "..." if len(s) > 800 else s

@st.cache_data(ttl=3600)
def search_companies_live(name: str):
    if not name or len(name.strip()) < 2: return []
    try:
        r = requests.get("https://data.brreg.no/enhetsregisteret/api/enheter", params={"navn": name.strip(), "size": 10}, timeout=30)
        if r.status_code == 200:
            return r.json().get("_embedded", {}).get("enheter", []) or []
    except Exception:
        pass
    return []

def get_company_details(company: dict):
    if not company: return None
    out = {"company_name": company.get("navn",""), "org_number": company.get("organisasjonsnummer",""),
           "nace_code": "", "nace_description": "", "homepage": company.get("hjemmeside",""),
           "employees": company.get("antallAnsatte",""), "address": "", "post_nr": "", "city": "",
           "registration_date": company.get("stiftelsesdato","")}
    addr = company.get("forretningsadresse",{}) or {}
    if addr:
        a = addr.get("adresse",[])
        out["address"] = ", ".join(filter(None, a)) if isinstance(a, list) else a or ""
        out["post_nr"] = addr.get("postnummer","")
        out["city"] = addr.get("poststed","")
    nace = company.get("naeringskode1",{}) or {}
    if nace:
        out["nace_code"] = nace.get("kode","")
        out["nace_description"] = nace.get("beskrivelse","")
    return out

def load_template_from_github():
    try:
        if os.path.exists("Grundmall.xlsx"):
            return open("Grundmall.xlsx","rb").read()
        r = requests.get("https://raw.githubusercontent.com/Shaedyr/PDF-to-premade-Excel-Sheet/main/PremadeExcelTemplate.xlsx", timeout=30)
        if r.status_code == 200: return r.content
        st.error("Kunne ikke laste Excel-malen fra GitHub")
    except Exception as e:
        st.error(f"Feil ved lasting av mal: {e}")
    return None

def update_excel_template(template_bytes: bytes, data: dict, summary: str):
    if not template_bytes: return None
    try:
        s = BytesIO(template_bytes); s.seek(0)
        wb = load_workbook(s); ws = wb.worksheets[0]
        cname = data.get("company_name","Selskap")
        safe = lambda x: re.sub(r'[\\/*?:\[\]]', '', x)[:31]
        try: ws.title = safe(f"{cname} Info")
        except Exception: pass
        if len(wb.worksheets) > 1:
            try: wb.worksheets[1].title = safe(f"{cname} Anbud")
            except Exception: pass
        # info box
        for rng in list(ws.merged_cells.ranges):
            if str(rng) == 'A2:D13':
                try: ws.unmerge_cells(str(rng))
                except Exception: pass
        try: ws.merge_cells('A2:D13')
        except Exception: pass
        ws['A2'] = summary or f"Informasjon om {cname}"
        ws['A2'].alignment = Alignment(wrap_text=True, vertical='top', horizontal='left')
        for r in range(2,14):
            try: ws.row_dimensions[r].height = 18
            except Exception: pass
        mapping = {
            'company_name': ('B14','D14',50),
            'org_number': ('B15',None,20),
            'address': ('B16','D16',100),
            'post_nr': ('B17','C17',15),
            'homepage': ('B20','D20',100),
            'employees': ('B21',None,10)
        }
        for k,(cell,merge_to,maxlen) in mapping.items():
            val = data.get(k,"")
            if val:
                v = str(val)
                if len(v) > maxlen: v = v[:maxlen-3] + "..."
                if k=='org_number' and len(v)==9 and v.isdigit(): ws[cell] = f"'{v}"
                else: ws[cell] = v
                if merge_to and len(v) > 20:
                    try: ws.merge_cells(f"{cell}:{merge_to}"); ws[cell].alignment = Alignment(wrap_text=True, vertical='center', horizontal='left')
                    except Exception: pass
        nace_code = data.get("nace_code",""); nace_desc = data.get("nace_description","")
        if nace_code and nace_desc:
            ws['B18'] = f"{nace_desc} ({nace_code})"
            try: ws.merge_cells('B18:D18'); ws['B18'].alignment = Alignment(wrap_text=True, vertical='center', horizontal='left')
            except Exception: pass
        elif nace_code: ws['B18'] = nace_code
        elif nace_desc: ws['B18'] = nace_desc
        else: ws['B18'] = "Data ikke tilgjengelig"
        ws.column_dimensions['A'].width = 15; ws.column_dimensions['B'].width = 25
        ws.column_dimensions['C'].width = 15; ws.column_dimensions['D'].width = 15
        out = BytesIO(); wb.save(out); out.seek(0); return out.getvalue()
    except Exception as e:
        st.error(f"Excel oppdatering feilet: {e}")
        try:
            df = pd.DataFrame([data]); o = BytesIO(); df.to_excel(o, index=False); o.seek(0); return o.getvalue()
        except Exception:
            return None

def main():
    # session keys
    for k in ('selected_company_data','companies_list','current_search','last_search','show_dropdown'):
        if k not in st.session_state: st.session_state[k] = None if k=='selected_company_data' else [] if k=='companies_list' else "" if k in ('current_search','last_search') else False

    st.title("📄 PDF → Excel (Brønnøysund)")
    st.markdown("Hent selskapsinformasjon og oppdater Excel automatisk")
    st.markdown("---")
    c1, c2 = st.columns(2)
    with c1:
        _ = st.file_uploader("PDF dokument (valgfritt)", type="pdf", help="Last opp PDF for referanse")
    with c2:
        container = st.container()
        q = st.text_input("Selskapsnavn *", placeholder="Skriv her... (minst 2 bokstaver)", help="Søk starter automatisk når du skriver", key="live_search_input")
        if st.session_state.get('current_search','') != q:
            st.session_state.selected_company_data = None
        st.session_state.current_search = q
        if q and len(q.strip()) >= 2:
            with container:
                with st.spinner("Søker..."):
                    comps = search_companies_live(q)
                if comps:
                    opts = ["-- Velg selskap --"]
                    cd = {}
                    for c in comps:
                        name = c.get('navn','Ukjent navn'); org = c.get('organisasjonsnummer',''); city = c.get('forretningsadresse',{}).get('poststed','')
                        disp = f"{name}" + (f" (Org.nr: {org})" if org else "") + (f" - {city}" if city else "")
                        opts.append(disp); cd[disp]=c
                    sel = st.selectbox("🔍 Søkeresultater:", opts, key="dynamic_company_dropdown")
                    if sel and sel != "-- Velg selskap --":
                        st.session_state.selected_company_data = get_company_details(cd[sel]); st.success(f"✅ Valgt: {cd[sel].get('navn')}")
                else:
                    if len(q.strip()) >= 3: st.warning("Ingen selskaper funnet. Prøv et annet navn.")
                    st.session_state.selected_company_data = None

    st.markdown("---")
    if 'template_loaded' not in st.session_state:
        with st.spinner("Laster Excel-mal..."):
            tb = load_template_from_github()
            if tb: st.session_state.template_bytes = tb; st.session_state.template_loaded = True; st.success("✅ Excel-mal lastet")
            else: st.session_state.template_loaded = False; st.error("❌ Kunne ikke laste Excel-mal")
    if st.button("🚀 Prosesser & Oppdater Excel", use_container_width=True):
        if not st.session_state.selected_company_data:
            st.error("❌ Vennligst velg et selskap fra listen først"); st.stop()
        if not st.session_state.get('template_loaded'):
            st.error("❌ Excel-mal ikke tilgjengelig"); st.stop()
        data = st.session_state.selected_company_data; st.session_state.extracted_data = data
        st.write("**Trinn 1:** Søker etter selskapsopplysninger...")
        summary = None; name = data.get('company_name','')
        with st.spinner("Søker på Wikipedia..."): summary = _wiki_summary(name)
        if not summary:
            st.info("Fant ikke på Wikipedia. Prøver websøk...")
            with st.spinner("Søker på nettet..."): summary = _web_summary(name)
        if not summary:
            st.info("Fant ikke på nettet. Lager analyse fra Brønnøysund-data...")
            summary = create_summary_from_brreg_data(data)
        st.session_state.company_summary = summary
        st.write("**Trinn 2:** Oppdaterer Excel...")
        try:
            updated = update_excel_template(st.session_state.template_bytes, data, summary)
            if updated:
                st.session_state.excel_ready = True; st.session_state.excel_bytes = updated; st.success("✅ Excel-fil oppdatert!")
                st.info(f"**Informasjon plassert i:**\n- **Ark 1:** {re.sub(r'[\\/*?:\\[\\]]','', f\"{data.get('company_name','Selskap')} Info\")[:31]}\n- **Stort informasjonsvindu:** Celle A2:D13\n- **Selskapsdata:** Celler B14-B21\n- **Ark 2:** {re.sub(r'[\\/*?:\\[\\]]','', f\"{data.get('company_name','Selskap')} Anbud\")[:31]}")
            else:
                st.error("❌ Kunne ikke generere Excel-fil")
        except Exception as e:
            st.error(f"❌ Feil ved Excel-oppdatering: {e}")

    if st.session_state.extracted_data:
        st.markdown("---"); st.subheader("📋 Ekstraherte data")
        d1, d2 = st.columns(2)
        with d1:
            d = st.session_state.extracted_data
            st.write(f"**Selskapsnavn:** {d.get('company_name','')}")
            st.write(f"**Organisasjonsnummer:** {d.get('org_number','')}")
            st.write(f"**Adresse:** {d.get('address','')}")
            st.write(f"**Postnummer:** {d.get('post_nr','')}")
            st.write(f"**Poststed:** {d.get('city','')}")
            st.write(f"**Antall ansatte:** {d.get('employees','')}")
            st.write(f"**Hjemmeside:** {d.get('homepage','')}")
            nc, nd = d.get('nace_code',''), d.get('nace_description','')
            if nc and nd: st.write(f"**NACE-bransje/nummer:** {nd} ({nc})")
            elif nc: st.write(f"**NACE-nummer:** {nc}")
            elif nd: st.write(f"**NACE-bransje:** {nd}")
        with d2:
            if st.session_state.company_summary:
                st.write("**Sammendrag (går i celle A2:D13):**"); st.info(st.session_state.company_summary)

    if st.session_state.get('excel_ready') and st.session_state.get('excel_bytes'):
        st.markdown("---"); st.subheader("📥 Last ned")
        cname = st.session_state.extracted_data.get('company_name','selskap')
        safe = re.sub(r'[^\w\s-]','', cname, flags=re.UNICODE); safe = re.sub(r'[-\s]+','_', safe)
        st.download_button(label="⬇️ Last ned oppdatert Excel", data=st.session_state.excel_bytes,
                           file_name=f"{safe}_{datetime.now().strftime('%Y%m%d')}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    st.markdown("---"); st.caption("Drevet av Brønnøysund Enhetsregisteret API | Data mellomlagret i 1 time")

if __name__ == "__main__":
    main()
