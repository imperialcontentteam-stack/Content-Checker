import json
import re

import streamlit as st

from modules.auth import require_login
from modules.database import init_db, fetch_all, fetch_one, execute
from modules.course_scraper import fetch_course_page_text, find_relevant_snippet
from modules.ai_checker import (
    build_prompt,
    call_openrouter,
    course_overview_plagiarism_percentage,
)
from modules.report_exporter import report_json_to_docx_bytes

st.set_page_config(page_title="Dashboard - SLC Checker", layout="wide")
init_db()
require_login()

st.title("Course Checker Dashboard")
st.caption("Select a ready course and compare website content against tracker and specification text.")

courses = fetch_all(
    """
    select c.id, c.course_name, c.course_url,
           s.extraction_status,
           length(coalesce(s.specification_text, '')) as specification_text_length
    from courses c
    left join specifications s on s.id = c.specification_id
    order by c.course_name
    """
)

if not courses:
    st.warning("No courses found. Import the tracker sheet from Admin first.")
    st.stop()

ready_only = st.checkbox("Show ready courses only", value=True)
shown_courses = [c for c in courses if (c["specification_text_length"] or 0) > 0] if ready_only else courses

if not shown_courses:
    st.warning("No ready courses found yet. Extract specification text from Admin first.")
    st.stop()

course_options = {f"{c['course_name']} | Spec text: {c['specification_text_length']}": c["id"] for c in shown_courses}
selected_label = st.selectbox("Select course", list(course_options.keys()))
course_id = course_options[selected_label]

course = fetch_one(
    """
    select c.*, s.specification_text, s.extraction_status, s.extraction_error
    from courses c
    left join specifications s on s.id = c.specification_id
    where c.id = ?
    """,
    (course_id,),
)

col1, col2 = st.columns(2)
with col1:
    st.write("**Course URL:**", course.get("course_url") or "Missing")
    st.write("**Specification status:**", course.get("extraction_status") or "Missing")
with col2:
    spec_len = len(course.get("specification_text") or "")
    st.write("**Specification text length:**", spec_len)
    if spec_len <= 0:
        st.error("This course is not ready because specification text is missing.")

fields = st.multiselect(
    "Fields to check",
    ["Entry Requirements", "Method of Assessment", "Course Overview", "Awarding Body Details"],
    default=["Entry Requirements"],
)

st.info("Plagiarism percentage is calculated for **Course Overview only**. Other fields show decision and wording similarity risk, but not a plagiarism percentage.")

if st.button("Run Check", type="primary"):
    if not fields:
        st.error("Select at least one field.")
        st.stop()

    if not course.get("course_url"):
        st.error("Course URL is missing.")
        st.stop()

    if not course.get("specification_text"):
        st.error("Specification text is missing. Extract specification text from Admin first.")
        st.stop()

    try:
        with st.spinner("Fetching live course page..."):
            page_text = fetch_course_page_text(course["course_url"])

        results = []
        for field in fields:
            with st.spinner(f"Checking {field}..."):
                live_snippet = find_relevant_snippet(page_text, field)
                prompt = build_prompt(course, field, live_snippet, course["specification_text"])
                result = call_openrouter(prompt)
                result["field_checked"] = result.get("field_checked") or field

                # Only Course Overview receives a plagiarism percentage.
                if field == "Course Overview":
                    local_percent = course_overview_plagiarism_percentage(
                        course.get("course_overview") or live_snippet,
                        course.get("specification_text") or "",
                    )
                    ai_percent = result.get("plagiarism_percentage")
                    result["plagiarism_percentage"] = local_percent if local_percent is not None else ai_percent
                else:
                    result["plagiarism_percentage"] = None

                results.append(result)

        decisions = [r.get("decision", "Needs Review") for r in results]
        if any(d in ["Incorrect", "Missing"] for d in decisions):
            summary_status = "Action Required"
        elif any(d == "Needs Review" for d in decisions):
            summary_status = "Needs Review"
        else:
            summary_status = "Correct"

        report_id = execute(
            """
            insert into check_reports (course_id, course_name, checked_fields, summary_status, result_json)
            values (?, ?, ?, ?, ?)
            """,
            (
                course_id,
                course["course_name"],
                json.dumps(fields),
                summary_status,
                json.dumps(results, ensure_ascii=False),
            ),
        )

        st.success(f"Check completed. Report ID: {report_id}. Summary: {summary_status}")

        docx_bytes = report_json_to_docx_bytes(
            course_name=course["course_name"],
            summary_status=summary_status,
            results=results,
            report_id=report_id,
        )
        safe_name = re.sub(r"[^A-Za-z0-9_-]+", "_", course["course_name"])[:80]
        st.download_button(
            "Download this answer as Word document",
            data=docx_bytes,
            file_name=f"course_check_{report_id}_{safe_name}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

        for item in results:
            field_name = item.get("field_checked") or "Checked Field"
            st.subheader(field_name)

            if field_name == "Course Overview":
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Decision", item.get("decision", ""))
                c2.metric("Priority", item.get("priority", ""))
                c3.metric("Similarity Risk", item.get("wording_similarity_risk", ""))
                percentage = item.get("plagiarism_percentage")
                c4.metric("Plagiarism %", "N/A" if percentage is None else f"{percentage}%")
            else:
                c1, c2, c3 = st.columns(3)
                c1.metric("Decision", item.get("decision", ""))
                c2.metric("Priority", item.get("priority", ""))
                c3.metric("Similarity Risk", item.get("wording_similarity_risk", ""))

            st.write("**Explanation**")
            st.write(item.get("explanation", ""))

            st.write("**Suggested corrected wording**")
            st.info(item.get("suggested_corrected_wording", ""))

            st.write("**Low-risk rewritten wording**")
            st.success(item.get("low_risk_rewritten_wording", ""))

            with st.expander("Evidence"):
                st.write("**Website evidence**")
                st.write(item.get("current_website_evidence", ""))
                st.write("**Tracker evidence**")
                st.write(item.get("tracker_sheet_evidence", ""))
                st.write("**Specification evidence**")
                st.write(item.get("specification_evidence", ""))

    except Exception as exc:
        st.error(str(exc))
