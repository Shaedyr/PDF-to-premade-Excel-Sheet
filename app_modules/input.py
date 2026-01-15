import streamlit as st
from step_3_company_data import search_brreg_live

def get_user_inputs():
    """
    Handles:
    - PDF upload
    - Company name input
    - Brønnøysund live search
    - Dropdown selection

    Returns:
        pdf_bytes (bytes or None)
        selected_company (dict or None)
    """

    st.header("1️⃣ Last opp PDF og velg selskap")

    col1, col2 = st.columns(2)

    # PDF upload
    with col1:
        pdf_file = st.file_uploader(
            "PDF dokument (valgfritt)",
            type="pdf",
            help="Last opp PDF for ekstra informasjon"
        )
        pdf_bytes = pdf_file.read() if pdf_file else None

    # Company search
    with col2:
        query = st.text_input(
            "Selskapsnavn *",
            placeholder="Skriv inn minst 2 bokstaver..."
        )

        selected_company = None

        if query and len(query.strip()) >= 2:
            with st.spinner("Søker i Brønnøysund..."):
                results = search_brreg_live(query)

            if results:
                display_list = ["-- Velg selskap --"]
                mapping = {}

                for c in results:
                    name = c.get("navn", "Ukjent")
                    org = c.get("organisasjonsnummer", "")
                    city = c.get("forretningsadresse", {}).get("poststed", "")

                    label = f"{name} (Org.nr: {org}) - {city}"
                    display_list.append(label)
                    mapping[label] = c

                choice = st.selectbox("🔍 Søkeresultater:", display_list)

                if choice != "-- Velg selskap --":
                    selected_company = mapping[choice]
                    st.success(f"Valgt: {selected_company.get('navn')}")

            else:
                st.warning("Ingen selskaper funnet.")

    return pdf_bytes, selected_company
