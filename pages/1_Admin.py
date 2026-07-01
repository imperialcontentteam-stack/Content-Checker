import pandas as pd
import streamlit as st
from modules.auth import require_login, create_user, list_users, set_user_active
from modules.database import init_db, get_stats, fetch_all
from modules.tracker_import import read_tracker, validate_tracker, import_tracker
from modules.pdf_extractor import extract_pending_specs, reset_failed_specs

st.set_page_config(page_title="Admin - SLC Checker", layout="wide")
init_db()
require_login(required_role="admin")

st.title("Admin Control Panel")
st.caption("Import tracker data, extract specification text, create users, and check data readiness.")

stats = get_stats()
col1, col2, col3, col4, col5, col6 = st.columns(6)
col1.metric("Courses", stats["total_courses"])
col2.metric("Specifications", stats["total_specifications"])
col3.metric("Extracted", stats["extracted_specifications"])
col4.metric("Pending", stats["pending_specifications"])
col5.metric("Failed", stats["failed_specifications"])
col6.metric("Ready Courses", stats["ready_courses"])

st.divider()

st.header("1. Upload and import tracker sheet")
uploaded_file = st.file_uploader("Upload Course Information Tracker Sheet", type=["xlsx", "xls"])

if uploaded_file:
    df = read_tracker(uploaded_file)
    missing = validate_tracker(df)

    st.write(f"Rows found: **{len(df)}**")
    st.dataframe(df.head(20), use_container_width=True)

    if missing:
        st.error("Missing required columns: " + ", ".join(missing))
    else:
        if st.button("Import / Update Courses", type="primary"):
            result = import_tracker(df)
            st.success("Import completed.")
            st.json(result)
            st.rerun()

st.divider()

st.header("2. Extract specification text")
st.write("This downloads pending specification PDF links and saves extracted text into the local database.")

limit = st.number_input("Batch size", min_value=1, max_value=300, value=300, step=1)

col_a, col_b = st.columns(2)
with col_a:
    if st.button("Extract Pending Specifications", type="primary"):
        with st.spinner("Extracting specifications. Please wait..."):
            results = extract_pending_specs(limit=int(limit))
        success = sum(1 for r in results if r["ok"])
        failed = sum(1 for r in results if not r["ok"])
        st.success(f"Extraction completed. Success: {success}, Failed: {failed}")
        st.dataframe(pd.DataFrame(results), use_container_width=True)
        st.rerun()

with col_b:
    if st.button("Reset Failed Specifications to Pending"):
        count = reset_failed_specs()
        st.warning(f"Reset {count} failed specifications to pending.")
        st.rerun()

st.divider()

st.header("3. User accounts")
st.write("Create usernames for team members. Admin users can access this page; normal users can run checks and view reports.")

with st.form("create_user_form"):
    c1, c2 = st.columns(2)
    with c1:
        new_username = st.text_input("New username")
        full_name = st.text_input("Full name / display name")
    with c2:
        new_password = st.text_input("Temporary password", type="password")
        role = st.selectbox("Role", ["user", "admin"])
    submitted = st.form_submit_button("Create user", type="primary")

if submitted:
    ok, message = create_user(new_username, new_password, role, full_name)
    if ok:
        st.success(message)
        st.rerun()
    else:
        st.error(message)

users = list_users()
if users:
    st.dataframe(pd.DataFrame(users), use_container_width=True)

    with st.expander("Activate / deactivate user"):
        user_options = {f"{u['username']} ({u['role']})": u for u in users}
        selected = st.selectbox("Select user", list(user_options.keys()))
        selected_user = user_options[selected]
        active = st.checkbox("Active", value=bool(selected_user["is_active"]))
        if st.button("Save user active status"):
            set_user_active(selected_user["id"], active)
            st.success("User status updated.")
            st.rerun()

st.divider()

st.header("4. Failed specification links")
failed_specs = fetch_all(
    """
    select id, source_url, extraction_error, extracted_at
    from specifications
    where extraction_status = 'failed'
    order by extracted_at desc
    limit 100
    """
)

if failed_specs:
    st.dataframe(pd.DataFrame(failed_specs), use_container_width=True)
else:
    st.info("No failed specification links.")
