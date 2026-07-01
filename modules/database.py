import os
import sqlite3
from contextlib import contextmanager
from dotenv import load_dotenv

load_dotenv()

DB_PATH = os.getenv("DATABASE_PATH", "data/slc_checker.db")


def ensure_data_dir():
    folder = os.path.dirname(DB_PATH)
    if folder:
        os.makedirs(folder, exist_ok=True)


@contextmanager
def get_conn():
    ensure_data_dir()
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    try:
        yield conn
        conn.commit()
    finally:
        conn.close()


def init_db():
    with get_conn() as conn:
        conn.executescript(
            """
            create table if not exists specifications (
                id integer primary key autoincrement,
                source_url text unique,
                file_name text,
                awarding_body text,
                specification_text text,
                extraction_status text default 'pending',
                extraction_error text,
                extracted_at text,
                created_at text default current_timestamp
            );

            create table if not exists courses (
                id integer primary key autoincrement,
                course_name text not null,
                course_url text,
                specification_document text,
                specification_id integer,
                course_overview text,
                entry_requirements text,
                method_of_assessment text,
                human_score real,
                readability real,
                status text default 'active',
                created_at text default current_timestamp,
                updated_at text default current_timestamp,
                foreign key (specification_id) references specifications(id)
            );

            create table if not exists check_reports (
                id integer primary key autoincrement,
                course_id integer,
                course_name text,
                checked_fields text,
                summary_status text,
                result_json text,
                created_at text default current_timestamp,
                foreign key (course_id) references courses(id)
            );
            """
        )


def get_stats():
    init_db()
    with get_conn() as conn:
        total_courses = conn.execute("select count(*) from courses").fetchone()[0]
        total_specifications = conn.execute("select count(*) from specifications").fetchone()[0]
        extracted_specifications = conn.execute(
            "select count(*) from specifications where length(coalesce(specification_text, '')) > 0"
        ).fetchone()[0]
        pending_specifications = conn.execute(
            "select count(*) from specifications where extraction_status = 'pending'"
        ).fetchone()[0]
        failed_specifications = conn.execute(
            "select count(*) from specifications where extraction_status = 'failed'"
        ).fetchone()[0]
        total_reports = conn.execute("select count(*) from check_reports").fetchone()[0]
        ready_courses = conn.execute(
            """
            select count(*)
            from courses c
            join specifications s on s.id = c.specification_id
            where c.course_url is not null
              and length(coalesce(s.specification_text, '')) > 0
            """
        ).fetchone()[0]

    return {
        "total_courses": total_courses,
        "total_specifications": total_specifications,
        "extracted_specifications": extracted_specifications,
        "pending_specifications": pending_specifications,
        "failed_specifications": failed_specifications,
        "total_reports": total_reports,
        "ready_courses": ready_courses,
    }


def fetch_all(query, params=()):
    with get_conn() as conn:
        rows = conn.execute(query, params).fetchall()
        return [dict(row) for row in rows]


def fetch_one(query, params=()):
    with get_conn() as conn:
        row = conn.execute(query, params).fetchone()
        return dict(row) if row else None


def execute(query, params=()):
    with get_conn() as conn:
        cur = conn.execute(query, params)
        return cur.lastrowid
