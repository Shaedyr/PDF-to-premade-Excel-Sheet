import streamlit as st

# Import only the modules you actually use
from app_modules import (
    main_page,
    input,
    company_data,
    pdf_parser,
    summary,
    excel_filler,
    template_loader,
    download,
)

# Sidebar page mapping
PAGES = {
    "🏠 Hovedside": main_page,
    "📄 Input-modul": input,
    "🏢 Company Data": company_data,
    "📄 PDF Parser": pdf_parser,
    "📝 Summary Generator": summary,
    "📊 Excel Filler": excel_filler,
    "📁 Template Loader": template_loader,
    "📥 Download": download,
}


def main():
    st.set_page_config(page_title="PDF → Excel Automator", layout="wide")

    # Sidebar navigation
    st.sidebar.title("Navigasjon")
    choice = st.sidebar.radio("Velg side:", list(PAGES.keys()))

    # Run selected page
    page = PAGES[choice]
    page.run()


if __name__ == "__main__":
    main()
