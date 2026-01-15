import streamlit as st
import msal

# ---------------------------------------------------------
# CONFIGURATION
# ---------------------------------------------------------
TENANT_ID = "consumers"
CLIENT_ID = "8dcb3914-362f-4c33-9348-d8f9b41347c2"   # keep your real client ID here

AUTHORITY = "https://login.microsoftonline.com/consumers"
SCOPES = ["Files.Read"]



# ---------------------------------------------------------
# MSAL APP FACTORY
# ---------------------------------------------------------
def get_msal_app():
    return msal.PublicClientApplication(
        CLIENT_ID,
        authority=AUTHORITY
    )


# ---------------------------------------------------------
# LOGIN PAGE
# ---------------------------------------------------------
def run():
    st.title("🔐 Login to Microsoft")

    # If already logged in
    if "token" in st.session_state:
        st.success("Du er allerede logget inn.")
        return

    st.write("Klikk på knappen under for å logge inn med Microsoft.")

    # Start login flow
    if st.button("Login med Microsoft"):
        app = get_msal_app()
        flow = app.initiate_device_flow(scopes=SCOPES)

        if "user_code" not in flow:
            st.error("Kunne ikke starte innloggingsflyten.")
            return

        st.info("Følg disse stegene:")
        st.write("1. Gå til https://microsoft.com/devicelogin")
        st.write(f"2. Skriv inn denne koden: **{flow['user_code']}**")
        st.write("3. Fullfør innloggingen i nettleseren.")

        # Button to confirm login is completed
        if st.button("Jeg har fullført innloggingen"):
            result = app.acquire_token_by_device_flow(flow)

            if "access_token" in result:
                # Store the entire token result
                st.session_state["token"] = result
                st.success("Innlogging vellykket!")
            else:
                st.error("Innlogging feilet. Prøv igjen.")
