import streamlit as st

# Import your modules from the app_modules folder
from app_modules import login_page
from app_modules import input
from app_modules import template_loader
from app_modules import company_data
from app_modules import summary
from app_modules import pdf_parser
from app_modules import excel_filler
from app_modules import download

st.set_page_config(page_title="PDF → Excel (Brønnøysund)", layout="wide")

st.sidebar.title("Navigasjon")

page = st.sidebar.radio(
    "Velg side",
    [
        "Login",
        "Input",
        "Template Loader",
        "PDF Parser",
        "Company Summary",
        "Excel Filler",
        "Download"
    ]
)

if page == "Login":
    login_page.run()

elif page == "Input":
    input.run()

elif page == "Template Loader":
    template_loader.run()

elif page == "PDF Parser":
    pdf_parser.run()

elif page == "Company Summary":
    summary.run()

elif page == "Excel Filler":
    excel_filler.run()

elif page == "Download":
    download.run()
