import streamlit as st

# Import all modules
from app_modules import (
    login_page,
    main_page,
    input,
    company_data,
    pdf_parser,
    summary,
    excel_filler,
    template_loader,
    graph_client,
    download,
)

# Sidebar page mapping
PAGES = {
    "🔐 Login": login_page,
    "🏠 Hovedside": main_page,
    "📄 Input-modul": input,
    "🏢 Company Data": company_data,
    "📄 PDF Parser": pdf_parser,
    "📝 Summary Generator": summary,
    "📊 Excel Filler": excel_filler,
    "📁 Template Loader": template_loader,
    "☁️ Graph Client": graph_client,
    "📥 Download": download,
}


def main():
    st.set_page_config(page_title="PDF → Excel Automator", layout="wide")

    # Sidebar navigation
    st.sidebar.title("Navigasjon")
    choice = st.sidebar.radio("Velg side:", list(PAGES.keys()))

    # If user is not logged in, force login page
    if "token" not in st.session_state and choice != "🔐 Login":
        st.warning("Du må logge inn først.")
        login_page.run()
        return

    # Run selected page
    page = PAGES[choice]
    page.run()


if __name__ == "__main__":
    main()
