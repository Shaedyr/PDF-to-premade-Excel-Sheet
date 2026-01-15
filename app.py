import streamlit as st

from step_1_input import get_user_inputs
from step_2_template_loader import load_template_from_onedrive
from step_3_company_data import fetch_company_by_org, format_company_data
from step_4_summary import generate_company_summary
from step_5_pdf_parser import extract_fields_from_pdf
from step_6_excel_filler import fill_excel
from step_7_download import download_excel_file


def main():
    st.set_page_config(page_title="PDF → Excel (Brønnøysund)", layout="wide")

    st.title("📄 PDF → Excel (Brønnøysund)")
    st.caption("Hent selskapsinformasjon og oppdater Excel automatisk")
    st.divider()

    # -----------------------------
    # STEP 1: INPUTS
    # -----------------------------
    pdf_bytes, selected_company_raw = get_user_inputs()

    if not selected_company_raw:
        st.info("Velg et selskap for å fortsette.")
        return

    # -----------------------------
    # STEP 2: LOAD TEMPLATE
    # -----------------------------
    if "template_bytes" not in st.session_state:
        st.session_state.template_bytes = load_template_from_onedrive()

    template_bytes = st.session_state.template_bytes

    # -----------------------------
    # STEP 3: COMPANY DATA
    # -----------------------------
    org_number = selected_company_raw.get("organisasjonsnummer")
    raw_company_data = fetch_company_by_org(org_number) if org_number else selected_company_raw
    company_data = format_company_data(raw_company_data)

    # -----------------------------
    # STEP 4: SUMMARY
    # -----------------------------
    summary_text = generate_company_summary(company_data)

    # -----------------------------
    # STEP 5: PDF FIELDS
    # -----------------------------
    pdf_fields = extract_fields_from_pdf(pdf_bytes) if pdf_bytes else {}

    # -----------------------------
    # MERGE FIELDS
    # -----------------------------
    merged_fields = {}
    merged_fields.update(company_data)
    merged_fields.update(pdf_fields)
    merged_fields["company_summary"] = summary_text

    st.divider()
    st.subheader("📋 Ekstraherte data")

    col_left, col_right = st.columns(2)

    with col_left:
        st.write("**Selskapsnavn:**", merged_fields.get("company_name", ""))
        st.write("**Organisasjonsnummer:**", merged_fields.get("org_number", ""))
        st.write("**Adresse:**", merged_fields.get("address", ""))
        st.write("**Postnummer:**", merged_fields.get("post_nr", ""))
        st.write("**Poststed:**", merged_fields.get("city", ""))
        st.write("**Antall ansatte:**", merged_fields.get("employees", ""))
        st.write("**Hjemmeside:**", merged_fields.get("homepage", ""))
        st.write("**NACE-kode:**", merged_fields.get("nace_code", ""))
        st.write("**NACE-beskrivelse:**", merged_fields.get("nace_description", ""))

    with col_right:
        st.markdown("**Sammendrag (går i 'Om oss' / 'Skriv her' celle):**")
        st.info(summary_text or "Ingen tilgjengelig selskapsbeskrivelse.")

    st.divider()

    # -----------------------------
    # STEP 6 + 7: PROCESS & DOWNLOAD
    # -----------------------------
    if st.button("🚀 Prosesser & Oppdater Excel", use_container_width=True):
        with st.spinner("Behandler og fyller inn Excel..."):
            excel_bytes = fill_excel(
                template_bytes=template_bytes,
                field_values=merged_fields,
                summary_text=summary_text,
            )

        download_excel_file(
            excel_bytes=excel_bytes,
            company_name=merged_fields.get("company_name", "Selskap")
        )


if __name__ == "__main__":
    main()
