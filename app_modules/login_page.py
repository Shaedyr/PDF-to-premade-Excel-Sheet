import streamlit as st
import msal

TENANT_ID = "YOUR_TENANT_ID"
CLIENT_ID = "YOUR_CLIENT_ID"
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["Files.Read"]

def get_msal_app():
    return msal.PublicClientApplication(
        CLIENT_ID,
        authority=AUTHORITY
    )

def run():
    st.title("🔐 Login to Microsoft")

    if "access_token" in st.session_state:
        st.success("You are already logged in.")
        return

    st.write("Click the button below to sign in to Microsoft.")

    if st.button("Login with Microsoft"):
        app = get_msal_app()
        flow = app.initiate_device_flow(scopes=SCOPES)

        if "user_code" not in flow:
            st.error("Could not start login flow.")
            return

        st.info("Follow these steps:")
        st.write("1. Go to https://microsoft.com/devicelogin")
        st.write(f"2. Enter this code: **{flow['user_code']}**")
        st.write("3. Complete the login in your browser.")

        if st.button("I have completed login"):
            result = app.acquire_token_by_device_flow(flow)

            if "access_token" in result:
                st.session_state["access_token"] = result["access_token"]
                st.success("Login successful!")
            else:
                st.error("Login failed. Try again.")
