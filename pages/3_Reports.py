import streamlit as st
from modules.auth import require_login
from modules.database import init_db, fetch_all
from modules.report_exporter import (
    reports_to_dataframe,
    dataframe_to_excel_bytes,
    all_reports_to_docx_bytes,
    build_single_report_docx_bytes,
)

st.set_page_config(page_title="Reports - SLC Checker", layout="wide")
init_db()
require_login()

st.title("Reports")
st.caption("View and export saved course check reports.")

df = reports_to_dataframe()

if df.empty:
    st.info("No reports yet. Run a check from the Dashboard first.")
    st.stop()

st.dataframe(df, use_container_width=True)

col1, col2 = st.columns(2)
with col1:
    excel_bytes = dataframe_to_excel_bytes(df)
    st.download_button(
        label="Download Excel Report",
        data=excel_bytes,
        file_name="slc_course_check_reports.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

with col2:
    docx_bytes = all_reports_to_docx_bytes(df)
    st.download_button(
        label="Download All Reports as Word Document",
        data=docx_bytes,
        file_name="slc_course_check_reports.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )

st.divider()
st.subheader("Download one report as Word document")
report_rows = fetch_all(
    """
    select id, course_name, summary_status, created_at
    from check_reports
    order by created_at desc
    """
)

if report_rows:
    options = {
        f"#{r['id']} - {r['course_name']} - {r['summary_status']} - {r['created_at']}": r["id"]
        for r in report_rows
    }
    selected = st.selectbox("Select report", list(options.keys()))
    selected_id = options[selected]
    single_docx = build_single_report_docx_bytes(selected_id)
    st.download_button(
        label="Download selected report as Word document",
        data=single_docx,
        file_name=f"slc_course_check_report_{selected_id}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
