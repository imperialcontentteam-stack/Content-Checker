import os
from urllib.parse import urlparse, unquote
import pandas as pd
from modules.database import get_conn

REQUIRED_COLUMNS = [
    "Course Name",
    "Course URL",
    "Specification Document",
    "Course Overview",
    "Entry Requirements",
    "Method of Assessment",
]


def clean(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def to_float(value):
    value = clean(value)
    if not value:
        return None
    try:
        return float(value)
    except ValueError:
        return None


def file_name_from_url(url):
    url = clean(url)
    if not url:
        return "Specification Document"
    try:
        parsed = urlparse(url)
        name = os.path.basename(parsed.path)
        return unquote(name) if name else "Specification Document"
    except Exception:
        return "Specification Document"


def read_tracker(uploaded_file):
    df = pd.read_excel(uploaded_file)
    df = df.fillna("")
    return df


def validate_tracker(df):
    missing = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    return missing


def import_tracker(df):
    inserted_courses = 0
    updated_courses = 0
    inserted_specs = 0
    reused_specs = 0
    skipped_rows = 0

    with get_conn() as conn:
        for _, row in df.iterrows():
            course_name = clean(row.get("Course Name"))
            course_url = clean(row.get("Course URL"))
            spec_url = clean(row.get("Specification Document"))

            if not course_name:
                skipped_rows += 1
                continue

            spec_id = None
            if spec_url:
                existing_spec = conn.execute(
                    "select id from specifications where source_url = ? limit 1",
                    (spec_url,),
                ).fetchone()

                if existing_spec:
                    spec_id = existing_spec["id"]
                    reused_specs += 1
                else:
                    cur = conn.execute(
                        """
                        insert into specifications (source_url, file_name, extraction_status)
                        values (?, ?, 'pending')
                        """,
                        (spec_url, file_name_from_url(spec_url)),
                    )
                    spec_id = cur.lastrowid
                    inserted_specs += 1

            payload = {
                "course_name": course_name,
                "course_url": course_url or None,
                "specification_document": spec_url or None,
                "specification_id": spec_id,
                "course_overview": clean(row.get("Course Overview")) or None,
                "entry_requirements": clean(row.get("Entry Requirements")) or None,
                "method_of_assessment": clean(row.get("Method of Assessment")) or None,
                "human_score": to_float(row.get("Human Score")),
                "readability": to_float(row.get("Readability")),
                "status": "active",
            }

            if course_url:
                existing_course = conn.execute(
                    "select id from courses where course_url = ? limit 1",
                    (course_url,),
                ).fetchone()
            else:
                existing_course = conn.execute(
                    "select id from courses where course_name = ? limit 1",
                    (course_name,),
                ).fetchone()

            if existing_course:
                conn.execute(
                    """
                    update courses
                    set course_name = ?, course_url = ?, specification_document = ?,
                        specification_id = ?, course_overview = ?, entry_requirements = ?,
                        method_of_assessment = ?, human_score = ?, readability = ?,
                        status = ?, updated_at = current_timestamp
                    where id = ?
                    """,
                    (
                        payload["course_name"],
                        payload["course_url"],
                        payload["specification_document"],
                        payload["specification_id"],
                        payload["course_overview"],
                        payload["entry_requirements"],
                        payload["method_of_assessment"],
                        payload["human_score"],
                        payload["readability"],
                        payload["status"],
                        existing_course["id"],
                    ),
                )
                updated_courses += 1
            else:
                conn.execute(
                    """
                    insert into courses (
                        course_name, course_url, specification_document, specification_id,
                        course_overview, entry_requirements, method_of_assessment,
                        human_score, readability, status
                    ) values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        payload["course_name"],
                        payload["course_url"],
                        payload["specification_document"],
                        payload["specification_id"],
                        payload["course_overview"],
                        payload["entry_requirements"],
                        payload["method_of_assessment"],
                        payload["human_score"],
                        payload["readability"],
                        payload["status"],
                    ),
                )
                inserted_courses += 1

    return {
        "inserted_courses": inserted_courses,
        "updated_courses": updated_courses,
        "inserted_specs": inserted_specs,
        "reused_specs": reused_specs,
        "skipped_rows": skipped_rows,
    }
