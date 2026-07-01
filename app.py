import os
import streamlit as st
from modules.database import init_db, get_stats
from modules.auth import authenticate, current_user, ensure_default_admin, logout

st.set_page_config(
    page_title="SLC Course Content Checker",
    page_icon="✅",
    layout="wide",
)

init_db()
ensure_default_admin()

st.title("SLC Course Content Checker")
st.caption("Python Streamlit version - simpler local/internal QA tool")

user = current_user()

if not user:
    st.subheader("Login")
    st.write("Use your username and password to access the checker.")

    with st.form("login_form"):
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")
        submitted = st.form_submit_button("Log in", type="primary")

    if submitted:
        ok, message = authenticate(username, password)
        if ok:
            st.success(message)
            st.rerun()
        else:
            st.error(message)

    with st.expander("First-time admin login"):
        st.write(
            "The app creates a default admin from `.env`. If you did not set it, the temporary default is:"
        )
        st.code("Username: admin\nPassword: admin123")
        st.warning("Change this by setting ADMIN_USERNAME and ADMIN_PASSWORD in `.env`, then create real user accounts from Admin.")

    st.stop()

with st.sidebar:
    st.write(f"Logged in as **{user['username']}**")
    st.caption(f"Role: {user['role']}")
    if st.button("Log out"):
        logout()
        st.rerun()

stats = get_stats()

col1, col2, col3, col4, col5 = st.columns(5)
col1.metric("Courses", stats["total_courses"])
col2.metric("Specifications", stats["total_specifications"])
col3.metric("Extracted Specs", stats["extracted_specifications"])
col4.metric("Ready Courses", stats["ready_courses"])
col5.metric("Reports", stats["total_reports"])

st.divider()

st.subheader("Recommended workflow")
st.write(
    """
    1. Admin logs in and opens **Admin**.  
    2. Upload the Course Information Tracker Sheet and import courses.  
    3. Extract specification text from PDF links.  
    4. Users open **Dashboard** and run checks for ready courses.  
    5. Open **Reports** to download Excel or Word documents.
    """
)

st.info(
    "This version avoids Supabase Edge Functions. Everything runs in one Python app, "
    "making it easier to debug and maintain."
)
