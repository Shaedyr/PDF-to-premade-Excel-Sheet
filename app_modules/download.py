import streamlit as st
from datetime import datetime


def download_excel_file(excel_bytes, company_name="Selskap"):
    """
    Displays a download button for the final Excel file.
    Generates a clean filename with timestamp.
    """

    if not excel_bytes:
        st.error("Ingen Excel-fil å laste ned.")
        return

    # Clean filename
    safe_name = "".join(c for c in company_name if c.isalnum() or c in " _-").strip()
    timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    filename = f"{safe_name}_{timestamp}.xlsx"

    st.subheader("📥 Last ned ferdig Excel-fil")

    st.download_button(
        label="⬇️ Last ned oppdatert Excel",
        data=excel_bytes,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.success("Excel-filen er klar for nedlasting!")


# ---------------------------------------------------------
# PAGE VIEW (so it works as a selectable page)
# ---------------------------------------------------------
def run():
    st.title("📥 Download Excel")
    st.write("Denne siden viser nedlastingsfunksjonen.")
    st.info("Excel-filen genereres på hovedsiden og lastes ned her.")
