"""
SLC Course Content Checker — v4 (Auth + Spec Upload + Filtered Run Check)
=========================================================================
Internal tool for the South London College course content team.

What's new in v4
----------------
1.  User Authentication
      • Admin and User accounts with login
      • Admin  → Import courses · upload/manage specification documents · manage course data
      • User   → Run Check page only · view & download generated reports
2.  Admin Features
      • 📥 Import Courses — upload the Excel tracker (now with Category ID / Level / Type /
        Qualification Specification columns) and store courses in the database
      • 📑 Specification documents — upload a PDF/DOCX spec per course (or extract from URL)
        and automatically extract & store: Entry Requirements · Qualification Specification ·
        Method of Assessment · other spec fields
3.  Run Check (User)
      • Filters before running the validation: Category ID · Level · Type (Award /
        Certificate / Diploma)
      • Left panel  = selected course details
      • Right panel = specification document + extracted requirements
      • Compares ONLY: Qualification Specification · Entry Requirements ·
        Method of Assessment (wording only)
4.  Report — rendered in the same format as the reference screenshot:
      Level | Type | Number header, one box per section with
      ❌ Errors identified 01..N (red) and ✅ Recommended Action 01..N (green),
      plus ✅ correct / ❌ incorrect wording / ❌ missing / ⚠ mismatch markers and a summary.
5.  Report Download — ⬇️ Download Report (Word) button on the same Run Check page,
    available immediately after the check completes.

Stack: Python · Streamlit · SQLite · OpenRouter API
Run:   streamlit run app.py

Default accounts (change the passwords after first login!):
    admin / admin123   (role: admin)
    user  / user123    (role: user)
"""

import hashlib
import html
import io
import json
import os
import re
import secrets as pysecrets
import sqlite3
import time
from datetime import datetime

import pandas as pd
import requests
import streamlit as st
from bs4 import BeautifulSoup
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor

# ═══════════════════════════════════════════════════════════════════
#  CONSTANTS
# ═══════════════════════════════════════════════════════════════════

DB_PATH = "slc_checker.db"
MODEL = "deepseek/deepseek-v4-pro"

# The ONLY sections compared during a Run Check (per the requirements)
CHECK_SECTIONS = [
    ("qualification_specification", "Qualification Specification"),
    ("entry_requirements", "Entry Requirements"),
    ("method_of_assessment", "Method of Assessment"),
]

TYPE_OPTIONS = ["Award", "Certificate", "Diploma"]

STATUS_META = {
    "correct":           {"icon": "✅", "label": "Correct",           "color": "#1D7A46"},
    "incorrect_wording": {"icon": "❌", "label": "Incorrect wording", "color": "#C62828"},
    "missing":           {"icon": "❌", "label": "Missing information","color": "#C62828"},
    "mismatch":          {"icon": "⚠️", "label": "Mismatched requirements", "color": "#E67E22"},
}

USER_AGENT = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                  "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
}

# ═══════════════════════════════════════════════════════════════════
#  DATABASE
# ═══════════════════════════════════════════════════════════════════

def get_conn():
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    return conn


def _ensure_columns(c, table: str, cols: dict):
    """Add any missing columns (simple forward migrations)."""
    existing = [r[1] for r in c.execute(f"PRAGMA table_info({table})")]
    for name, ddl in cols.items():
        if name not in existing:
            c.execute(f"ALTER TABLE {table} ADD COLUMN {name} {ddl}")


def init_db():
    with get_conn() as c:
        c.executescript(
            """
            CREATE TABLE IF NOT EXISTS users (
                id            INTEGER PRIMARY KEY AUTOINCREMENT,
                username      TEXT UNIQUE NOT NULL,
                password_hash TEXT NOT NULL,
                salt          TEXT NOT NULL,
                role          TEXT NOT NULL DEFAULT 'user',   -- 'admin' | 'user'
                created_at    TEXT
            );

            CREATE TABLE IF NOT EXISTS courses (
                id                   INTEGER PRIMARY KEY AUTOINCREMENT,
                course_name          TEXT UNIQUE NOT NULL,
                course_url           TEXT,
                spec_url             TEXT,
                category_id          TEXT,
                level                TEXT,
                course_type          TEXT,        -- Award | Certificate | Diploma
                course_number        TEXT,
                entry_requirements   TEXT,
                method_of_assessment TEXT,
                qualification_specification TEXT,
                course_overview      TEXT,
                spec_text            TEXT,        -- full extracted spec document text
                spec_filename        TEXT,        -- uploaded spec file name
                spec_entry_requirements   TEXT,   -- sections extracted FROM the spec
                spec_qualification        TEXT,
                spec_assessment           TEXT,
                spec_other_json           TEXT,   -- other extracted spec fields (json)
                updated_at           TEXT
            );

            CREATE TABLE IF NOT EXISTS check_reports (
                id           INTEGER PRIMARY KEY AUTOINCREMENT,
                course_id    INTEGER NOT NULL REFERENCES courses(id) ON DELETE CASCADE,
                checked_by   TEXT,
                checked_at   TEXT,
                status       TEXT,          -- 'pass' | 'errors' | 'failed'
                report_json  TEXT,          -- full structured report (sections/errors/actions)
                summary      TEXT
            );
            """
        )
        # migrations for databases created by earlier versions
        _ensure_columns(c, "courses", {
            "category_id": "TEXT", "level": "TEXT", "course_type": "TEXT",
            "course_number": "TEXT", "qualification_specification": "TEXT",
            "spec_filename": "TEXT", "spec_entry_requirements": "TEXT",
            "spec_qualification": "TEXT", "spec_assessment": "TEXT",
            "spec_other_json": "TEXT",
        })


# ── users / auth ───────────────────────────────────────────────────

def _hash_password(password: str, salt: str) -> str:
    return hashlib.sha256((salt + password).encode("utf-8")).hexdigest()


def create_user(username: str, password: str, role: str = "user") -> bool:
    salt = pysecrets.token_hex(16)
    try:
        with get_conn() as c:
            c.execute(
                "INSERT INTO users (username, password_hash, salt, role, created_at) VALUES (?,?,?,?,?)",
                (username.strip(), _hash_password(password, salt), salt, role,
                 datetime.now().isoformat(timespec="seconds")),
            )
        return True
    except sqlite3.IntegrityError:
        return False


def verify_user(username: str, password: str):
    with get_conn() as c:
        row = c.execute("SELECT * FROM users WHERE username = ?", (username.strip(),)).fetchone()
    if row and _hash_password(password, row["salt"]) == row["password_hash"]:
        return {"id": row["id"], "username": row["username"], "role": row["role"]}
    return None


def change_password(username: str, new_password: str):
    salt = pysecrets.token_hex(16)
    with get_conn() as c:
        c.execute("UPDATE users SET password_hash=?, salt=? WHERE username=?",
                  (_hash_password(new_password, salt), salt, username))


def all_users() -> list:
    with get_conn() as c:
        return [dict(r) for r in c.execute("SELECT id, username, role, created_at FROM users ORDER BY username")]


def delete_user(user_id: int):
    with get_conn() as c:
        c.execute("DELETE FROM users WHERE id=?", (user_id,))


def seed_default_users():
    """Create default admin/user accounts on first run."""
    with get_conn() as c:
        n = c.execute("SELECT COUNT(*) FROM users").fetchone()[0]
    if n == 0:
        create_user("admin", "admin123", "admin")
        create_user("user", "user123", "user")


# ── courses ────────────────────────────────────────────────────────

def upsert_course(row: dict) -> str:
    """Insert or update a course by name. Returns 'inserted' or 'updated'."""
    now = datetime.now().isoformat(timespec="seconds")
    fields = ["course_url", "spec_url", "category_id", "level", "course_type",
              "course_number", "entry_requirements", "method_of_assessment",
              "qualification_specification", "course_overview"]
    with get_conn() as c:
        cur = c.execute("SELECT id FROM courses WHERE course_name = ?", (row["course_name"],))
        existing = cur.fetchone()
        if existing:
            sets = ", ".join(f"{f}=?" for f in fields)
            c.execute(f"UPDATE courses SET {sets}, updated_at=? WHERE id=?",
                      [row.get(f) for f in fields] + [now, existing["id"]])
            return "updated"
        cols = ["course_name"] + fields + ["updated_at"]
        q = f"INSERT INTO courses ({','.join(cols)}) VALUES ({','.join('?' * len(cols))})"
        c.execute(q, [row["course_name"]] + [row.get(f) for f in fields] + [now])
        return "inserted"


def all_courses() -> list:
    with get_conn() as c:
        return [dict(r) for r in c.execute("SELECT * FROM courses ORDER BY course_name")]


def get_course(course_id: int) -> dict:
    with get_conn() as c:
        r = c.execute("SELECT * FROM courses WHERE id=?", (course_id,)).fetchone()
        return dict(r) if r else {}


def delete_course(course_id: int):
    with get_conn() as c:
        c.execute("DELETE FROM courses WHERE id=?", (course_id,))


def save_spec(course_id: int, text: str, filename: str, sections: dict):
    with get_conn() as c:
        c.execute(
            """UPDATE courses SET spec_text=?, spec_filename=?,
               spec_entry_requirements=?, spec_qualification=?, spec_assessment=?,
               spec_other_json=?, updated_at=? WHERE id=?""",
            (text, filename,
             sections.get("entry_requirements"), sections.get("qualification_specification"),
             sections.get("method_of_assessment"),
             json.dumps(sections.get("other", {}), ensure_ascii=False),
             datetime.now().isoformat(timespec="seconds"), course_id),
        )


def save_check_report(course_id: int, checked_by: str, status: str, report: dict, summary: str):
    with get_conn() as c:
        c.execute(
            """INSERT INTO check_reports (course_id, checked_by, checked_at, status, report_json, summary)
               VALUES (?,?,?,?,?,?)""",
            (course_id, checked_by, datetime.now().isoformat(timespec="seconds"),
             status, json.dumps(report, ensure_ascii=False), summary),
        )


def latest_check_reports() -> list:
    q = """
        SELECT r.*, c.course_name, c.course_url, c.level, c.course_type, c.course_number, c.category_id
        FROM check_reports r
        JOIN courses c ON c.id = r.course_id
        WHERE r.id IN (SELECT MAX(id) FROM check_reports GROUP BY course_id)
        ORDER BY c.course_name
    """
    with get_conn() as c:
        return [dict(x) for x in c.execute(q)]


# ═══════════════════════════════════════════════════════════════════
#  FETCHING & TEXT EXTRACTION
# ═══════════════════════════════════════════════════════════════════

def fetch_url(url: str, timeout=30) -> requests.Response:
    return requests.get(url, headers=USER_AGENT, timeout=timeout, allow_redirects=True)


def extract_pdf_text(data: bytes, max_chars=40000) -> str:
    import pdfplumber
    parts = []
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        for page in pdf.pages:
            parts.append(page.extract_text() or "")
            if sum(len(p) for p in parts) > max_chars:
                break
    return "\n".join(parts)[:max_chars]


def extract_docx_text(data: bytes, max_chars=40000) -> str:
    d = Document(io.BytesIO(data))
    parts = [p.text for p in d.paragraphs if p.text.strip()]
    for table in d.tables:
        for row in table.rows:
            parts.append(" | ".join(cell.text.strip() for cell in row.cells))
    return "\n".join(parts)[:max_chars]


def extract_spec_from_url(url: str, max_chars=40000) -> str:
    """Extract text from a specification document URL — PDF or web page."""
    resp = fetch_url(url, timeout=60)
    resp.raise_for_status()
    ctype = resp.headers.get("Content-Type", "").lower()
    if "pdf" in ctype or url.lower().split("?")[0].endswith(".pdf"):
        return extract_pdf_text(resp.content, max_chars)
    soup = BeautifulSoup(resp.text, "html.parser")
    for tag in soup(["script", "style", "nav", "footer", "header", "noscript"]):
        tag.decompose()
    return soup.get_text("\n", strip=True)[:max_chars]


# ── spec section extraction ────────────────────────────────────────
# Pull the required sections out of the raw specification text using
# heading heuristics; the admin can review/edit before saving.

SECTION_PATTERNS = {
    "entry_requirements": [
        r"entry\s+requirements?", r"entry\s+criteria", r"admission\s+requirements?",
        r"who\s+is\s+this\s+qualification\s+for",
    ],
    "method_of_assessment": [
        r"method(?:s)?\s+of\s+assessment", r"assessment\s+method(?:s)?",
        r"assessment\s+and\s+grading", r"how\s+(?:is|will)\s+.{0,40}assessed", r"\bassessment\b",
    ],
    "qualification_specification": [
        r"qualification\s+specification", r"qualification\s+overview",
        r"qualification\s+structure", r"qualification\s+summary", r"about\s+this\s+qualification",
    ],
}

OTHER_SECTION_PATTERNS = {
    "Guided Learning Hours": [r"guided\s+learning\s+hours?", r"\bGLH\b"],
    "Total Qualification Time": [r"total\s+qualification\s+time", r"\bTQT\b"],
    "Progression": [r"\bprogression\b"],
    "Units": [r"unit\s+structure", r"mandatory\s+units?", r"\bunits?\b\s*:"],
    "Grading": [r"\bgrading\b"],
}


def _find_heading_positions(text: str) -> list:
    """Return (pos, line) for every line that looks like a heading."""
    out = []
    for m in re.finditer(r"^(.{2,90})$", text, flags=re.M):
        line = m.group(1).strip()
        if not line:
            continue
        words = line.split()
        # headings: short lines, no ending full stop, mostly title/upper case or numbered
        looks_heading = (
            len(words) <= 10 and not line.endswith((".", ",", ";"))
            and (line.isupper() or line.istitle()
                 or re.match(r"^\d+(\.\d+)*\s+\w", line)
                 or re.match(r"^[A-Z][\w\s&/()\-']+$", line))
        )
        if looks_heading:
            out.append((m.start(1), line))
    return out


def extract_spec_sections(text: str) -> dict:
    """Heuristically split the spec text into the required sections."""
    sections = {"entry_requirements": "", "qualification_specification": "",
                "method_of_assessment": "", "other": {}}
    if not text:
        return sections
    headings = _find_heading_positions(text)

    def grab(patterns, headings_only=False) -> str:
        for pat in patterns:
            rx = re.compile(pat, flags=re.I)
            for i, (pos, line) in enumerate(headings):
                if rx.search(line):
                    start = pos + len(line)
                    end = headings[i + 1][0] if i + 1 < len(headings) else min(len(text), start + 4000)
                    body = text[start:end].strip()
                    if len(body) > 20:
                        return body[:4000]
            if headings_only:
                continue
            # fall back to an inline match anywhere in the text
            m = rx.search(text)
            if m:
                chunk = text[m.end(): m.end() + 1500].strip()
                if len(chunk) > 20:
                    return chunk
        return ""

    for key, pats in SECTION_PATTERNS.items():
        sections[key] = grab(pats)
    for label, pats in OTHER_SECTION_PATTERNS.items():
        val = grab(pats, headings_only=True)
        if val:
            sections["other"][label] = val[:1200]
    # if no explicit "qualification specification" heading, use the document opening
    if not sections["qualification_specification"]:
        sections["qualification_specification"] = text[:2500].strip()
    return sections


# ═══════════════════════════════════════════════════════════════════
#  OPENROUTER (LLM comparison)
# ═══════════════════════════════════════════════════════════════════

def call_openrouter(prompt: str, system: str, api_key: str, model: str, temperature=0.0) -> str:
    resp = requests.post(
        "https://openrouter.ai/api/v1/chat/completions",
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
            "HTTP-Referer": "http://localhost:8501",
            "X-Title": "SLC Course Content Checker",
        },
        json={
            "model": model,
            "temperature": temperature,
            "messages": [
                {"role": "system", "content": system},
                {"role": "user", "content": prompt},
            ],
        },
        timeout=180,
    )
    resp.raise_for_status()
    return resp.json()["choices"][0]["message"]["content"]


def parse_json_reply(raw: str) -> dict:
    raw = raw.strip()
    raw = re.sub(r"^```(?:json)?\s*|\s*```$", "", raw, flags=re.S)
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        m = re.search(r"\{.*\}", raw, flags=re.S)
        if m:
            return json.loads(m.group(0))
        raise


# ═══════════════════════════════════════════════════════════════════
#  RUN CHECK — comparison engine
# ═══════════════════════════════════════════════════════════════════

CHECK_SYSTEM = (
    "You are a meticulous content auditor for South London College. You compare the COURSE "
    "content (from the internal database) against the OFFICIAL QUALIFICATION SPECIFICATION "
    "document, section by section. You reply ONLY with valid JSON — no markdown, no commentary."
)

CHECK_PROMPT = """Audit the course "{name}" (Level: {level} · Type: {ctype} · Number: {number}).

Compare ONLY these sections, COURSE content vs the SPECIFICATION extract for the same section:
1. Qualification Specification
2. Entry Requirements
3. Method of Assessment — compare the WORDING ONLY: flag phrasing that misdescribes the
   assessment method; ignore layout/format differences that do not change the described method.

For each section decide exactly ONE status:
- "correct"           → the course content is accurate and consistent with the specification
- "incorrect_wording" → present but the wording is wrong / misdescribes the specification
- "missing"           → required information from the specification is absent in the course content
- "mismatch"          → the course states requirements that CONTRADICT the specification

For every section that is not "correct", list each distinct error with a matching recommended action.

Reply with EXACTLY this JSON shape:
{{
  "summary": "one or two sentence overall verdict",
  "sections": [
    {{
      "section": "Qualification Specification | Entry Requirements | Method of Assessment",
      "status": "correct | incorrect_wording | missing | mismatch",
      "course_content": "short quote of the course text audited (or 'missing')",
      "errors": [
        {{
          "error": "clear description of the problem (Errors identified)",
          "recommended_action": "the exact fix or action to take (Recommend Action)"
        }}
      ]
    }}
  ]
}}

Sections with status "correct" must have an empty "errors" list.

=== COURSE CONTENT (database) ===
{course_block}

=== SPECIFICATION EXTRACTS ===
{spec_block}
"""


def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip().lower())


def fallback_compare(course: dict) -> dict:
    """Deterministic comparison used when no API key is configured."""
    sections = []
    n_err = 0
    pairs = [
        ("Qualification Specification", course.get("qualification_specification"),
         course.get("spec_qualification")),
        ("Entry Requirements", course.get("entry_requirements"),
         course.get("spec_entry_requirements")),
        ("Method of Assessment", course.get("method_of_assessment"),
         course.get("spec_assessment")),
    ]
    for label, course_val, spec_val in pairs:
        c, s = _norm(course_val), _norm(spec_val)
        errors = []
        if not c and s:
            status = "missing"
            errors.append({"error": f"{label} is missing from the course content but is defined in the specification.",
                           "recommended_action": f"Add the {label} from the specification to the course content."})
        elif not s:
            status = "correct" if c else "missing"
            if status == "missing":
                errors.append({"error": f"{label} is missing in both the course content and the specification extract.",
                               "recommended_action": f"Upload/extract the specification and populate the {label}."})
        elif c == s or c in s or s in c:
            status = "correct"
        else:
            # word-level overlap as a rough wording check
            cw, sw = set(c.split()), set(s.split())
            overlap = len(cw & sw) / max(1, len(sw))
            if overlap >= 0.75:
                status = "incorrect_wording"
                errors.append({"error": f"The {label} wording differs from the specification.",
                               "recommended_action": "Align the wording with the specification text."})
            else:
                status = "mismatch"
                errors.append({"error": f"The {label} does not match the specification requirements.",
                               "recommended_action": f"Replace the course {label} with the specification version and review for accuracy."})
        n_err += len(errors)
        sections.append({"section": label, "status": status,
                         "course_content": (course_val or "missing")[:300], "errors": errors})
    summary = ("All compared sections match the specification."
               if n_err == 0 else
               f"{n_err} issue(s) found across the compared sections (offline comparison — configure an API key for a deeper AI check).")
    return {"summary": summary, "sections": sections}


def run_check(course: dict, api_key: str, model: str) -> dict:
    """Compare the course DB content vs the extracted spec sections.
    Returns {'status', 'summary', 'sections': [...], 'level','type','number'}"""
    header = {"level": course.get("level") or "—",
              "type": course.get("course_type") or "—",
              "number": course.get("course_number") or str(course.get("id", "—")),
              "category_id": course.get("category_id") or "—",
              "course_name": course.get("course_name", "")}

    if not api_key:
        data = fallback_compare(course)
    else:
        course_block = "\n\n".join(
            f"[{label}]\n{course.get(key) or '(not provided)'}"
            for key, label in CHECK_SECTIONS
        )
        spec_map = {"qualification_specification": course.get("spec_qualification"),
                    "entry_requirements": course.get("spec_entry_requirements"),
                    "method_of_assessment": course.get("spec_assessment")}
        spec_block = "\n\n".join(
            f"[{label}]\n{spec_map.get(key) or '(not found in specification)'}"
            for key, label in CHECK_SECTIONS
        )
        prompt = CHECK_PROMPT.format(
            name=course["course_name"], level=header["level"], ctype=header["type"],
            number=header["number"], course_block=course_block, spec_block=spec_block,
        )
        try:
            data = parse_json_reply(call_openrouter(prompt, CHECK_SYSTEM, api_key, model))
        except Exception as e:
            return {**header, "status": "failed",
                    "summary": f"AI check failed: {e}", "sections": []}

    sections = data.get("sections", [])
    has_err = any(s.get("errors") for s in sections) or any(
        s.get("status") != "correct" for s in sections)
    return {**header,
            "status": "errors" if has_err else "pass",
            "summary": data.get("summary", ""),
            "sections": sections}


# ═══════════════════════════════════════════════════════════════════
#  WORD REPORT — same layout as the reference screenshot
# ═══════════════════════════════════════════════════════════════════

RED = RGBColor(0xC6, 0x28, 0x28)
GREEN = RGBColor(0x1D, 0x7A, 0x46)
ORANGE = RGBColor(0xE6, 0x7E, 0x22)
GREY = RGBColor(0x66, 0x66, 0x66)


def _set_cell_border(cell, color="C62828", size="18"):
    """Give a table cell a thick coloured border (mirrors the red boxes in the screenshot)."""
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    tcPr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), size)
        el.set(qn("w:color"), color)
        borders.append(el)
    tcPr.append(borders)


def build_check_word_report(reports: list) -> bytes:
    """reports: list of dicts as produced by run_check() (+ course_name)."""
    doc = Document()

    title = doc.add_heading("SLC Course Check Report", level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub = doc.add_paragraph(f"Generated {datetime.now().strftime('%d %B %Y, %H:%M')}  ·  "
                            f"{len(reports)} course(s)")
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub.runs[0].font.size = Pt(10)
    sub.runs[0].font.color.rgb = GREY

    for rep in reports:
        doc.add_paragraph()
        doc.add_heading(rep.get("course_name", ""), level=1)

        # ── header row: Level | Type | Number (like the screenshot) ──
        hdr = doc.add_table(rows=2, cols=3)
        hdr.style = "Table Grid"
        for i, label in enumerate(["Level", "Type", "Number"]):
            cell = hdr.rows[0].cells[i]
            cell.text = label
            run = cell.paragraphs[0].runs[0]
            run.bold = True
            run.font.color.rgb = RED
            run.font.size = Pt(13)
        for i, key in enumerate(["level", "type", "number"]):
            hdr.rows[1].cells[i].text = str(rep.get(key, "—"))

        if rep.get("summary"):
            p = doc.add_paragraph(rep["summary"])
            p.runs[0].italic = True
        doc.add_paragraph()

        # ── one boxed block per section ──
        for sec in rep.get("sections", []):
            label = sec.get("section", "")
            status = str(sec.get("status", "correct"))
            meta = STATUS_META.get(status, STATUS_META["correct"])

            # section title (red, like "Current Method of Assessment")
            p = doc.add_paragraph()
            r = p.add_run(f"Current {label}" if label == "Method of Assessment"
                          else label)
            r.bold = True
            r.font.size = Pt(12)
            r.font.color.rgb = RED

            # status marker
            p = doc.add_paragraph()
            r = p.add_run(f"{meta['icon']} {meta['label']}")
            r.bold = True
            r.font.color.rgb = GREEN if status == "correct" else (
                ORANGE if status == "mismatch" else RED)

            # two-column layout: boxed course content | errors & actions
            body = doc.add_table(rows=1, cols=2)
            body.autofit = False
            body.columns[0].width = Inches(3.4)
            body.columns[1].width = Inches(3.2)
            left, right = body.rows[0].cells

            left.text = str(sec.get("course_content") or "—")[:1200]
            _set_cell_border(left, color="C62828", size="24")

            errors = sec.get("errors") or []
            if not errors:
                pr = right.paragraphs[0]
                run = pr.add_run("✅ No errors identified")
                run.font.color.rgb = GREEN
                run.bold = True
            else:
                first = True
                for n, err in enumerate(errors, start=1):
                    pe = right.paragraphs[0] if first else right.add_paragraph()
                    first = False
                    re_run = pe.add_run(f"Errors identified {n:02d}: ")
                    re_run.bold = True
                    re_run.font.color.rgb = RED
                    pe.add_run(str(err.get("error", ""))).font.color.rgb = RED

                    pa = right.add_paragraph()
                    ra = pa.add_run(f"Recommend Action {n:02d}: ")
                    ra.bold = True
                    ra.font.color.rgb = GREEN
                    pa.add_run(str(err.get("recommended_action", ""))).font.color.rgb = GREEN
                    right.add_paragraph()
            doc.add_paragraph()

        # summary of detected issues
        total = sum(len(s.get("errors") or []) for s in rep.get("sections", []))
        p = doc.add_paragraph()
        r = p.add_run(f"Summary: {total} issue(s) detected across "
                      f"{len(rep.get('sections', []))} compared section(s).")
        r.bold = True
        doc.add_page_break()

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ═══════════════════════════════════════════════════════════════════
#  UI — THEME & HELPERS
# ═══════════════════════════════════════════════════════════════════

st.set_page_config(page_title="SLC Course Content Checker", page_icon="🎓",
                   layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,600;9..144,700&family=Inter:wght@400;500;600;700&display=swap');
html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

.slc-hero {
  background: linear-gradient(120deg, #10243E 0%, #1D4E89 55%, #2E7D6B 100%);
  border-radius: 18px; padding: 26px 34px; margin-bottom: 6px;
  position: relative; overflow: hidden;
}
.slc-hero:before { content:""; position:absolute; right:-40px; top:-40px; width:220px; height:220px;
  background: radial-gradient(circle, rgba(255,229,92,.35), transparent 70%); }
.slc-hero h1 { font-family:'Fraunces', serif; color:#fff; margin:0; font-size:2rem; }
.slc-hero h1 .hl { background:#FFE55C; color:#10243E; padding:0 .35rem; border-radius:6px; }
.slc-hero p { color:#CFE3F5; margin:.45rem 0 0; font-size:.95rem; }

.stat-card { border-radius:14px; padding:16px 18px; border:1px solid #E6E9EF;
  background:#fff; box-shadow:0 2px 10px rgba(16,36,62,.06); }
.stat-card .num { font-family:'Fraunces',serif; font-size:1.9rem; font-weight:700; color:#10243E; line-height:1.1;}
.stat-card .lbl { font-size:.78rem; text-transform:uppercase; letter-spacing:.09em; color:#6B7688; font-weight:600;}
.stat-card.ok { border-top:4px solid #30A46C; } .stat-card.err { border-top:4px solid #E5484D; }
.stat-card.info { border-top:4px solid #1D4E89; } .stat-card.warn { border-top:4px solid #F5B300; }

/* ── screenshot-style report ── */
.rep-header { display:flex; gap:60px; margin:10px 0 22px; }
.rep-header .cell b { color:#C62828; font-size:1.15rem; }
.rep-header .cell span { display:block; color:#2A2F3A; font-weight:600; margin-top:2px; }
.rep-section-title { color:#C62828; font-weight:700; font-size:1.05rem; margin:18px 0 6px; }
.rep-box { border:3px solid #C62828; border-radius:2px; min-height:90px; padding:12px 14px;
  background:#fff; color:#2A2F3A; font-size:.92rem; white-space:pre-wrap; }
.rep-err  { color:#C62828; font-weight:700; margin:4px 0 0; }
.rep-act  { color:#1D9E4B; font-weight:700; margin:0 0 10px 10px; }
.rep-status { font-weight:700; margin:4px 0; }
.issue-card { border-radius:12px; border:1px solid #E6E9EF; border-left-width:5px;
  padding:12px 16px; margin-bottom:10px; background:#fff; }

.stButton>button[kind="primary"] { background:linear-gradient(120deg,#1D4E89,#2E7D6B);
  border:none; border-radius:10px; font-weight:600; }
div[data-testid="stSidebar"] { background:#F6F8FB; }
</style>
""", unsafe_allow_html=True)


def hero():
    st.markdown(
        """
        <div class="slc-hero">
          <h1>🎓 SLC <span class="hl">Course Content</span> Checker</h1>
          <p>Import courses · upload specifications · run filtered checks · export reports.</p>
        </div>
        """, unsafe_allow_html=True)


def stat(col, value, label, kind="info"):
    col.markdown(
        f'<div class="stat-card {kind}"><div class="num">{value}</div>'
        f'<div class="lbl">{label}</div></div>', unsafe_allow_html=True)


def render_report_html(rep: dict):
    """Render the check report on-screen in the screenshot layout."""
    st.markdown(
        f"""<div class="rep-header">
              <div class="cell"><b>Level</b><span>{html.escape(str(rep.get('level','—')))}</span></div>
              <div class="cell"><b>Type</b><span>{html.escape(str(rep.get('type','—')))}</span></div>
              <div class="cell"><b>Number</b><span>{html.escape(str(rep.get('number','—')))}</span></div>
            </div>""", unsafe_allow_html=True)
    if rep.get("summary"):
        st.caption(rep["summary"])

    for sec in rep.get("sections", []):
        label = sec.get("section", "")
        title = f"Current {label}" if label == "Method of Assessment" else label
        status = str(sec.get("status", "correct"))
        meta = STATUS_META.get(status, STATUS_META["correct"])

        st.markdown(f'<div class="rep-section-title">{html.escape(title)}</div>',
                    unsafe_allow_html=True)
        st.markdown(f'<div class="rep-status" style="color:{meta["color"]}">'
                    f'{meta["icon"]} {meta["label"]}</div>', unsafe_allow_html=True)

        left, right = st.columns([1.1, 1], gap="large")
        with left:
            st.markdown(f'<div class="rep-box">{html.escape(str(sec.get("course_content") or "—")[:1500])}</div>',
                        unsafe_allow_html=True)
        with right:
            errors = sec.get("errors") or []
            if not errors:
                st.markdown('<div class="rep-act" style="margin-left:0">✅ No errors identified</div>',
                            unsafe_allow_html=True)
            for n, err in enumerate(errors, start=1):
                st.markdown(
                    f'<div class="rep-err">❌ Errors identified {n:02d}: '
                    f'{html.escape(str(err.get("error","")))}</div>'
                    f'<div class="rep-act">✅ Recommend Action {n:02d}: '
                    f'{html.escape(str(err.get("recommended_action","")))}</div>',
                    unsafe_allow_html=True)
        st.write("")

    total = sum(len(s.get("errors") or []) for s in rep.get("sections", []))
    n_bad = sum(1 for s in rep.get("sections", []) if s.get("status") != "correct")
    st.markdown("#### 📋 Summary of detected issues")
    c1, c2, c3 = st.columns(3)
    stat(c1, len(rep.get("sections", [])), "Sections compared", "info")
    stat(c2, n_bad, "Sections with issues", "err" if n_bad else "ok")
    stat(c3, total, "Total issues", "err" if total else "ok")


# ═══════════════════════════════════════════════════════════════════
#  AUTHENTICATION GATE
# ═══════════════════════════════════════════════════════════════════

init_db()
seed_default_users()

if "auth" not in st.session_state:
    st.session_state["auth"] = None

if st.session_state["auth"] is None:
    hero()
    st.write("")
    _, mid, _ = st.columns([1, 1.2, 1])
    with mid:
        st.markdown("### 🔐 Sign in")
        with st.form("login_form"):
            u = st.text_input("Username")
            p = st.text_input("Password", type="password")
            ok = st.form_submit_button("Sign in", type="primary", use_container_width=True)
        if ok:
            user = verify_user(u, p)
            if user:
                st.session_state["auth"] = user
                st.rerun()
            else:
                st.error("Invalid username or password.")
        st.caption("Default accounts — admin / admin123 (Admin) · user / user123 (User). "
                   "Change these passwords after first login (Admin → Manage Users).")
    st.stop()

AUTH = st.session_state["auth"]
IS_ADMIN = AUTH["role"] == "admin"

# API key from Streamlit secrets or environment
api_key = ""
try:
    api_key = st.secrets.get("OPENROUTER_API_KEY", "")
except Exception:
    pass
api_key = api_key or os.environ.get("OPENROUTER_API_KEY", "")
model = MODEL

hero()

with st.sidebar:
    st.markdown(f"### 👤 {AUTH['username']}")
    st.caption(f"Role: **{'Admin' if IS_ADMIN else 'User'}**")
    if st.button("🚪 Log out", use_container_width=True):
        st.session_state["auth"] = None
        st.session_state.pop("last_check", None)
        st.rerun()
    st.divider()
    if IS_ADMIN:
        if api_key:
            st.success("🔑 OpenRouter key loaded")
        else:
            st.warning("No API key — checks will use the built-in offline comparison. "
                       "Add `OPENROUTER_API_KEY` to `.streamlit/secrets.toml` for AI checks.")
    with st.expander("🔒 Change my password"):
        np1 = st.text_input("New password", type="password", key="np1")
        np2 = st.text_input("Repeat new password", type="password", key="np2")
        if st.button("Update password"):
            if len(np1) < 6:
                st.error("Use at least 6 characters.")
            elif np1 != np2:
                st.error("Passwords do not match.")
            else:
                change_password(AUTH["username"], np1)
                st.success("Password updated.")


# ═══════════════════════════════════════════════════════════════════
#  PAGES
# ═══════════════════════════════════════════════════════════════════
# Admin  → Import Courses · Specifications · Run Check · Reports · Manage Users
# User   → Run Check (with report view/download) only

if IS_ADMIN:
    tabs = st.tabs(["📥 Import Courses", "📑 Specifications", "🔍 Run Check",
                    "📊 Reports", "👥 Manage Users"])
    tab_import, tab_spec, tab_check, tab_reports, tab_users = tabs
else:
    (tab_check,) = st.tabs(["🔍 Run Check"])
    tab_import = tab_spec = tab_reports = tab_users = None


# ───────────────────────────────────────────────
# ADMIN · IMPORT COURSES
# ───────────────────────────────────────────────
if IS_ADMIN:
    with tab_import:
        st.subheader("Upload the course Excel file")
        st.caption("Expected columns: course name, course URL, spec URL, category ID, level, "
                   "type (Award/Certificate/Diploma), course number, entry requirements, "
                   "method of assessment, qualification specification, course overview. "
                   "You can remap columns below — only Course Name is mandatory.")

        up = st.file_uploader("Course file (.xlsx / .xls / .csv)", type=["xlsx", "xls", "csv"])
        if up:
            df = pd.read_csv(up) if up.name.lower().endswith(".csv") else pd.read_excel(up)
            df.columns = [str(c).strip() for c in df.columns]
            st.dataframe(df.head(10), use_container_width=True)

            def guess(*needles):
                for col in df.columns:
                    low = col.lower()
                    if any(n in low for n in needles):
                        return col
                return None

            cols = ["— none —"] + list(df.columns)

            def pick(label, guessed, key):
                idx = cols.index(guessed) if guessed in cols else 0
                return st.selectbox(label, cols, index=idx, key=key)

            st.markdown("#### Map your columns")
            c1, c2, c3 = st.columns(3)
            with c1:
                m_name = pick("Course name *", guess("course name", "title", "name"), "m1")
                m_url = pick("Course page URL", guess("course url", "page url", "link"), "m2")
                m_spec = pick("Specification URL", guess("spec url", "specification url", "spec"), "m3")
                m_cat = pick("Category ID", guess("category"), "m4")
            with c2:
                m_level = pick("Level", guess("level"), "m5")
                m_type = pick("Type (Award/Certificate/Diploma)", guess("type"), "m6")
                m_num = pick("Course number", guess("number", "code", "ref"), "m7")
                m_entry = pick("Entry Requirements", guess("entry"), "m8")
            with c3:
                m_assess = pick("Method of Assessment", guess("assess"), "m9")
                m_qual = pick("Qualification Specification", guess("qualification spec", "qualification"), "m10")
                m_over = pick("Course Overview", guess("overview", "description"), "m11")

            if st.button("📥 Import / update courses", type="primary",
                         disabled=(m_name == "— none —")):
                def val(row, col):
                    if col == "— none —":
                        return None
                    v = row.get(col)
                    return None if pd.isna(v) else str(v).strip()

                inserted = updated = skipped = 0
                for _, row in df.iterrows():
                    name = val(row, m_name)
                    if not name:
                        skipped += 1
                        continue
                    ctype = val(row, m_type)
                    if ctype:  # normalise to Award/Certificate/Diploma
                        for t in TYPE_OPTIONS:
                            if t.lower() in ctype.lower():
                                ctype = t
                                break
                    result = upsert_course({
                        "course_name": name,
                        "course_url": val(row, m_url),
                        "spec_url": val(row, m_spec),
                        "category_id": val(row, m_cat),
                        "level": val(row, m_level),
                        "course_type": ctype,
                        "course_number": val(row, m_num),
                        "entry_requirements": val(row, m_entry),
                        "method_of_assessment": val(row, m_assess),
                        "qualification_specification": val(row, m_qual),
                        "course_overview": val(row, m_over),
                    })
                    inserted += result == "inserted"
                    updated += result == "updated"

                s1, s2, s3 = st.columns(3)
                stat(s1, inserted, "New courses", "ok")
                stat(s2, updated, "Updated", "info")
                stat(s3, skipped, "Skipped (no name)", "warn")
                st.success("Courses imported into the database ✅")

        st.divider()
        st.markdown("#### 📚 Manage course data")
        courses_now = all_courses()
        if not courses_now:
            st.info("No courses yet — upload the Excel file above.")
        else:
            st.dataframe(
                pd.DataFrame(courses_now)[["id", "course_name", "category_id", "level",
                                           "course_type", "course_number", "spec_filename",
                                           "updated_at"]],
                use_container_width=True, hide_index=True)

            with st.expander("✏️ Edit a course"):
                sel = st.selectbox("Course", [c["course_name"] for c in courses_now], key="edit_sel")
                course = next(c for c in courses_now if c["course_name"] == sel)
                e1, e2, e3, e4 = st.columns(4)
                new_cat = e1.text_input("Category ID", course.get("category_id") or "", key="e_cat")
                new_lvl = e2.text_input("Level", course.get("level") or "", key="e_lvl")
                idx = TYPE_OPTIONS.index(course["course_type"]) + 1 if course.get("course_type") in TYPE_OPTIONS else 0
                new_type = e3.selectbox("Type", ["—"] + TYPE_OPTIONS, index=idx, key="e_type")
                new_num = e4.text_input("Number", course.get("course_number") or "", key="e_num")
                new_entry = st.text_area("Entry Requirements", course.get("entry_requirements") or "", key="e_entry", height=100)
                new_assess = st.text_area("Method of Assessment", course.get("method_of_assessment") or "", key="e_assess", height=100)
                new_qual = st.text_area("Qualification Specification", course.get("qualification_specification") or "", key="e_qual", height=100)
                b1, b2 = st.columns(2)
                if b1.button("💾 Save changes", type="primary"):
                    upsert_course({**course,
                                   "category_id": new_cat or None, "level": new_lvl or None,
                                   "course_type": None if new_type == "—" else new_type,
                                   "course_number": new_num or None,
                                   "entry_requirements": new_entry or None,
                                   "method_of_assessment": new_assess or None,
                                   "qualification_specification": new_qual or None})
                    st.success("Course updated."); st.rerun()
                if b2.button("🗑️ Delete this course"):
                    delete_course(course["id"])
                    st.warning("Course deleted."); st.rerun()


# ───────────────────────────────────────────────
# ADMIN · SPECIFICATION DOCUMENTS
# ───────────────────────────────────────────────
if IS_ADMIN:
    with tab_spec:
        st.subheader("Upload & manage specification documents")
        st.caption("Upload a PDF/DOCX specification per course (or extract it from the spec URL). "
                   "The tool extracts and stores: Entry Requirements · Qualification "
                   "Specification · Method of Assessment · other spec fields.")

        courses = all_courses()
        have = [c for c in courses if c.get("spec_text")]
        missing_url = [c for c in courses if c.get("spec_url") and not c.get("spec_text")]

        s1, s2, s3 = st.columns(3)
        stat(s1, len(courses), "Total courses", "info")
        stat(s2, len(have), "Specs stored", "ok")
        stat(s3, len(courses) - len(have), "Specs missing", "err" if len(courses) != len(have) else "ok")
        st.write("")

        if not courses:
            st.info("Import courses first (📥 Import Courses tab).")
        else:
            sel = st.selectbox("Course", [c["course_name"] for c in courses], key="spec_course")
            course = next(c for c in courses if c["course_name"] == sel)

            colA, colB = st.columns(2, gap="large")
            with colA:
                st.markdown("#### 📤 Upload specification document")
                f = st.file_uploader("Specification (.pdf / .docx / .txt)",
                                     type=["pdf", "docx", "txt"], key="spec_upload")
                if f and st.button("Extract & store from file", type="primary"):
                    with st.spinner("Extracting …"):
                        try:
                            data = f.read()
                            if f.name.lower().endswith(".pdf"):
                                text = extract_pdf_text(data)
                            elif f.name.lower().endswith(".docx"):
                                text = extract_docx_text(data)
                            else:
                                text = data.decode("utf-8", errors="ignore")
                            sections = extract_spec_sections(text)
                            save_spec(course["id"], text, f.name, sections)
                            st.success(f"Specification stored for **{sel}** ✅")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Extraction failed: {e}")
            with colB:
                st.markdown("#### 🌐 Extract from spec URL")
                url = st.text_input("Specification URL", course.get("spec_url") or "", key="spec_url_in")
                if url and st.button("Extract & store from URL"):
                    with st.spinner("Fetching & extracting …"):
                        try:
                            text = extract_spec_from_url(url)
                            sections = extract_spec_sections(text)
                            save_spec(course["id"], text, url.split("/")[-1] or "spec-from-url", sections)
                            st.success(f"Specification stored for **{sel}** ✅")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Extraction failed: {e}")

            if missing_url:
                st.divider()
                if st.button(f"📑 Bulk-extract {len(missing_url)} spec(s) from stored URLs"):
                    bar = st.progress(0.0); ok = fail = 0
                    for i, c in enumerate(missing_url):
                        try:
                            text = extract_spec_from_url(c["spec_url"])
                            save_spec(c["id"], text, c["spec_url"].split("/")[-1] or "spec",
                                      extract_spec_sections(text))
                            ok += 1
                        except Exception as e:
                            fail += 1
                            st.warning(f"❌ {c['course_name']}: {e}")
                        bar.progress((i + 1) / len(missing_url))
                    st.success(f"Done — {ok} extracted, {fail} failed."); st.rerun()

            # review / edit the extracted sections
            course = get_course(course["id"])  # refresh
            if course.get("spec_text"):
                st.divider()
                st.markdown(f"#### 🔎 Extracted requirements — *{course.get('spec_filename') or 'stored spec'}*")
                se = st.text_area("Entry Requirements (from spec)",
                                  course.get("spec_entry_requirements") or "", height=120, key="se")
                sq = st.text_area("Qualification Specification (from spec)",
                                  course.get("spec_qualification") or "", height=120, key="sq")
                sa = st.text_area("Method of Assessment (from spec)",
                                  course.get("spec_assessment") or "", height=120, key="sa")
                other = json.loads(course.get("spec_other_json") or "{}")
                if other:
                    with st.expander("Other extracted spec fields"):
                        for k, v in other.items():
                            st.markdown(f"**{k}**")
                            st.caption(v[:800])
                if st.button("💾 Save edited sections"):
                    save_spec(course["id"], course["spec_text"], course.get("spec_filename") or "",
                              {"entry_requirements": se, "qualification_specification": sq,
                               "method_of_assessment": sa, "other": other})
                    st.success("Sections saved.")
                with st.expander("Full specification text"):
                    st.text_area("Raw text", course["spec_text"], height=260, key="raw_spec")


# ───────────────────────────────────────────────
# RUN CHECK  (available to both roles)
# ───────────────────────────────────────────────
with tab_check:
    st.subheader("Run Check")
    courses = all_courses()

    if not courses:
        st.info("No courses in the database yet. "
                + ("Import your course Excel file first (📥 Import Courses)."
                   if IS_ADMIN else "Ask an admin to import the course data."))
    else:
        # ── 1) filters ──
        st.markdown("#### 1️⃣ Filters")
        f1, f2, f3 = st.columns(3)
        cats = sorted({c["category_id"] for c in courses if c.get("category_id")})
        levels = sorted({c["level"] for c in courses if c.get("level")})
        with f1:
            f_cat = st.selectbox("Category ID", ["All"] + cats)
        with f2:
            f_level = st.selectbox("Level", ["All"] + levels)
        with f3:
            f_type = st.selectbox("Type", ["All"] + TYPE_OPTIONS)

        filtered = [c for c in courses
                    if (f_cat == "All" or c.get("category_id") == f_cat)
                    and (f_level == "All" or c.get("level") == f_level)
                    and (f_type == "All" or c.get("course_type") == f_type)]

        st.caption(f"{len(filtered)} course(s) match the selected filters")

        if not filtered:
            st.warning("No courses match these filters.")
        else:
            sel = st.selectbox("2️⃣ Select the course to check",
                               [c["course_name"] for c in filtered])
            course = get_course(next(c["id"] for c in filtered if c["course_name"] == sel))

            # ── 2) two panels: course details | spec + extracted requirements ──
            left, right = st.columns(2, gap="large")
            with left:
                st.markdown("#### 📘 Course details")
                st.markdown(f"**{course['course_name']}**")
                m1, m2, m3, m4 = st.columns(4)
                m1.metric("Category", course.get("category_id") or "—")
                m2.metric("Level", course.get("level") or "—")
                m3.metric("Type", course.get("course_type") or "—")
                m4.metric("Number", course.get("course_number") or str(course["id"]))
                for key, label in CHECK_SECTIONS:
                    with st.expander(label, expanded=(key == "entry_requirements")):
                        st.write(course.get(key) or "_(not provided)_")
            with right:
                st.markdown("#### 📑 Specification document & extracted requirements")
                if not course.get("spec_text"):
                    st.warning("No specification stored for this course"
                               + (" — upload it in the 📑 Specifications tab." if IS_ADMIN
                                  else " — ask an admin to upload it."))
                else:
                    st.caption(f"Source: {course.get('spec_filename') or 'stored specification'}")
                    spec_map = {
                        "qualification_specification": course.get("spec_qualification"),
                        "entry_requirements": course.get("spec_entry_requirements"),
                        "method_of_assessment": course.get("spec_assessment"),
                    }
                    for key, label in CHECK_SECTIONS:
                        with st.expander(f"{label} (spec)", expanded=(key == "entry_requirements")):
                            st.write(spec_map.get(key) or "_(not found in specification)_")
                    with st.expander("Full specification text"):
                        st.text_area("Spec", course["spec_text"], height=200,
                                     key=f"specview_{course['id']}", disabled=True)

            st.divider()

            # ── 3) run the comparison ──
            st.markdown("#### 3️⃣ Compare required sections")
            st.caption("Compared sections: **Qualification Specification · Entry Requirements · "
                       "Method of Assessment (wording only)**")
            if st.button("🔍 Run check", type="primary",
                         disabled=not course.get("spec_text")):
                with st.spinner(f"Checking {course['course_name']} …"):
                    rep = run_check(course, api_key, model)
                rep["course_name"] = course["course_name"]
                if rep["status"] == "failed":
                    st.error(rep["summary"])
                else:
                    save_check_report(course["id"], AUTH["username"], rep["status"],
                                      rep, rep.get("summary", ""))
                    st.session_state["last_check"] = rep
                    st.success("Check complete — report saved.")

            # ── 4) report in the screenshot format + download ──
            rep = st.session_state.get("last_check")
            if rep and rep.get("course_name") == course["course_name"]:
                st.divider()
                st.markdown(f"### 📄 Validation report — {rep['course_name']}")
                render_report_html(rep)
                st.write("")
                docx_bytes = build_check_word_report([rep])
                st.download_button(
                    "⬇️ Download Report (Word)",
                    data=docx_bytes,
                    file_name=f"{re.sub(r'[^A-Za-z0-9]+', '_', rep['course_name'])}_check_report_{datetime.now():%Y-%m-%d}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary",
                )

        # ── previously generated reports (Users may view & download) ──
        st.divider()
        st.markdown("#### 🗂️ Previously generated reports")
        reports = latest_check_reports()
        if not reports:
            st.info("No saved reports yet.")
        else:
            table = pd.DataFrame([{
                "Course": r["course_name"],
                "Level": r.get("level") or "—",
                "Type": r.get("course_type") or "—",
                "Status": {"pass": "✅ Pass", "errors": "⚠️ Errors", "failed": "❌ Failed"}.get(r["status"], r["status"]),
                "Checked by": r.get("checked_by") or "—",
                "Checked at": r["checked_at"],
            } for r in reports])
            st.dataframe(table, use_container_width=True, hide_index=True)

            pick_rep = st.selectbox("View / download a saved report",
                                    [r["course_name"] for r in reports], key="saved_rep")
            r = next(x for x in reports if x["course_name"] == pick_rep)
            saved = json.loads(r["report_json"] or "{}")
            saved.setdefault("course_name", r["course_name"])
            with st.expander("📄 View report", expanded=False):
                render_report_html(saved)
            st.download_button(
                f"⬇️ Download report for {pick_rep} (Word)",
                data=build_check_word_report([saved]),
                file_name=f"{re.sub(r'[^A-Za-z0-9]+', '_', pick_rep)}_check_report.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key="dl_saved",
            )


# ───────────────────────────────────────────────
# ADMIN · REPORTS DASHBOARD
# ───────────────────────────────────────────────
if IS_ADMIN:
    with tab_reports:
        st.subheader("All saved reports")
        reports = latest_check_reports()
        if not reports:
            st.info("No reports yet — run some checks first.")
        else:
            n_pass = sum(r["status"] == "pass" for r in reports)
            n_err = sum(r["status"] == "errors" for r in reports)
            c1, c2, c3 = st.columns(3)
            stat(c1, len(reports), "Courses checked", "info")
            stat(c2, n_pass, "Passed", "ok")
            stat(c3, n_err, "With errors", "err")
            st.write("")

            st.dataframe(pd.DataFrame([{
                "Course": r["course_name"], "Category": r.get("category_id") or "—",
                "Level": r.get("level") or "—", "Type": r.get("course_type") or "—",
                "Status": {"pass": "✅ Pass", "errors": "⚠️ Errors", "failed": "❌ Failed"}.get(r["status"], r["status"]),
                "Checked by": r.get("checked_by") or "—", "Checked at": r["checked_at"],
                "Summary": r["summary"],
            } for r in reports]), use_container_width=True, hide_index=True)

            err_reports = [r for r in reports if r["status"] == "errors"]
            if err_reports:
                combined = []
                for r in err_reports:
                    d = json.loads(r["report_json"] or "{}")
                    d.setdefault("course_name", r["course_name"])
                    combined.append(d)
                st.download_button(
                    f"⬇️ Download combined Word report ({len(err_reports)} courses with errors)",
                    data=build_check_word_report(combined),
                    file_name=f"SLC_Check_Report_{datetime.now():%Y-%m-%d}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary",
                )


# ───────────────────────────────────────────────
# ADMIN · MANAGE USERS
# ───────────────────────────────────────────────
if IS_ADMIN:
    with tab_users:
        st.subheader("Manage user accounts")
        users = all_users()
        st.dataframe(pd.DataFrame(users), use_container_width=True, hide_index=True)

        c1, c2 = st.columns(2, gap="large")
        with c1:
            st.markdown("#### ➕ Create account")
            with st.form("new_user"):
                nu = st.text_input("Username")
                npw = st.text_input("Password", type="password")
                nrole = st.selectbox("Role", ["user", "admin"])
                sub = st.form_submit_button("Create", type="primary")
            if sub:
                if not nu.strip() or len(npw) < 6:
                    st.error("Provide a username and a password of at least 6 characters.")
                elif create_user(nu, npw, nrole):
                    st.success(f"Account **{nu}** created ({nrole})."); st.rerun()
                else:
                    st.error("That username already exists.")
        with c2:
            st.markdown("#### 🗑️ Remove account")
            removable = [u for u in users if u["username"] != AUTH["username"]]
            if not removable:
                st.caption("No other accounts to remove.")
            else:
                du = st.selectbox("Account", [u["username"] for u in removable])
                if st.button("Delete account"):
                    target = next(u for u in removable if u["username"] == du)
                    admins_left = sum(1 for u in users if u["role"] == "admin" and u["id"] != target["id"])
                    if target["role"] == "admin" and admins_left == 0:
                        st.error("Cannot delete the last admin account.")
                    else:
                        delete_user(target["id"])
                        st.warning(f"Account **{du}** deleted."); st.rerun()
