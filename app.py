"""
SLC Course Content Checker
==========================
Internal tool for the South London College course content team.

Features
--------
1.  Upload an Excel tracker sheet
2.  Import / update courses into a local SQLite database
3.  Extract missing qualification specification text (PDF or web page)
4.  Select which fields to check (Entry Requirements / Method of Assessment / Course Overview)
5.  Check courses one by one
6.  Optional testing mode (check only the first N courses)
7.  Bulk check with progress
8.  Save check reports to the database
9.  Generate a Word document containing ONLY the courses with errors + suggested solutions
10. Content Quality Review — paste or upload text and get a colour-highlighted
    grammar / punctuation / capitalisation / proper-noun review (like a proofreader's markup)

Stack: Python · Streamlit · SQLite · OpenRouter API
Run:   streamlit run app.py
"""

import html
import io
import json
import re
import sqlite3
import time
from datetime import datetime

import pandas as pd
import requests
import streamlit as st
from bs4 import BeautifulSoup
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, RGBColor

# ═══════════════════════════════════════════════════════════════════
#  CONSTANTS
# ═══════════════════════════════════════════════════════════════════

DB_PATH = "slc_checker.db"

FIELD_OPTIONS = {
    "entry_requirements": "Entry Requirements",
    "method_of_assessment": "Method of Assessment",
    "course_overview": "Course Overview",
}

DEFAULT_MODEL = "anthropic/claude-sonnet-4.5"
MODEL_CHOICES = [
    "anthropic/claude-sonnet-4.5",
    "anthropic/claude-3.5-haiku",
    "openai/gpt-4o-mini",
    "openai/gpt-4o",
    "google/gemini-2.0-flash-001",
    "meta-llama/llama-3.3-70b-instruct",
]

# Colour system for the Quality Review markup (mirrors the reference screenshots)
QR_CATEGORIES = {
    "grammar":            {"label": "Grammar",            "bg": "#FDD8D6", "border": "#E5484D"},
    "article":            {"label": "Articles (a/an/the)","bg": "#FFE3C7", "border": "#F76B15"},
    "spelling":           {"label": "Spelling",           "bg": "#E9DDFB", "border": "#8E4EC6"},
    "punctuation":        {"label": "Punctuation & Commas","bg": "#D5E7FB", "border": "#0090FF"},
    "capitalisation":     {"label": "Capitalisation",     "bg": "#D8F3DE", "border": "#30A46C"},
    "proper_noun":        {"label": "Proper Nouns",       "bg": "#FBDCEF", "border": "#D6409F"},
    "sentence_structure": {"label": "Sentence Structure", "bg": "#FBE8B4", "border": "#B58A00"},
    "consistency":        {"label": "Consistency",        "bg": "#D9F0F4", "border": "#0894B3"},
}

USER_AGENT = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                  "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
}

# ═══════════════════════════════════════════════════════════════════
#  DATABASE
# ═══════════════════════════════════════════════════════════════════

def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    with get_conn() as c:
        c.executescript(
            """
            CREATE TABLE IF NOT EXISTS courses (
                id                   INTEGER PRIMARY KEY AUTOINCREMENT,
                course_name          TEXT UNIQUE NOT NULL,
                course_url           TEXT,
                spec_url             TEXT,
                entry_requirements   TEXT,
                method_of_assessment TEXT,
                course_overview      TEXT,
                spec_text            TEXT,
                updated_at           TEXT
            );

            CREATE TABLE IF NOT EXISTS reports (
                id             INTEGER PRIMARY KEY AUTOINCREMENT,
                course_id      INTEGER NOT NULL REFERENCES courses(id) ON DELETE CASCADE,
                checked_at     TEXT,
                fields_checked TEXT,
                status         TEXT,           -- 'pass' | 'errors' | 'failed'
                errors_json    TEXT,
                rewrites_json  TEXT,
                summary        TEXT
            );
            """
        )
        # migration for databases created before wording suggestions existed
        cols = [r[1] for r in c.execute("PRAGMA table_info(reports)")]
        if "rewrites_json" not in cols:
            c.execute("ALTER TABLE reports ADD COLUMN rewrites_json TEXT")


def upsert_course(row: dict) -> str:
    """Insert or update a course by name. Returns 'inserted' or 'updated'."""
    now = datetime.now().isoformat(timespec="seconds")
    with get_conn() as c:
        cur = c.execute("SELECT id FROM courses WHERE course_name = ?", (row["course_name"],))
        existing = cur.fetchone()
        if existing:
            c.execute(
                """UPDATE courses SET course_url=?, spec_url=?, entry_requirements=?,
                   method_of_assessment=?, course_overview=?, updated_at=?
                   WHERE id=?""",
                (row.get("course_url"), row.get("spec_url"), row.get("entry_requirements"),
                 row.get("method_of_assessment"), row.get("course_overview"), now, existing["id"]),
            )
            return "updated"
        c.execute(
            """INSERT INTO courses (course_name, course_url, spec_url, entry_requirements,
               method_of_assessment, course_overview, updated_at)
               VALUES (?,?,?,?,?,?,?)""",
            (row["course_name"], row.get("course_url"), row.get("spec_url"),
             row.get("entry_requirements"), row.get("method_of_assessment"),
             row.get("course_overview"), now),
        )
        return "inserted"


def all_courses() -> list:
    with get_conn() as c:
        return [dict(r) for r in c.execute("SELECT * FROM courses ORDER BY course_name")]


def save_spec_text(course_id: int, text: str):
    with get_conn() as c:
        c.execute("UPDATE courses SET spec_text=?, updated_at=? WHERE id=?",
                  (text, datetime.now().isoformat(timespec="seconds"), course_id))


def save_report(course_id, fields, status, errors, summary, rewrites=None):
    with get_conn() as c:
        c.execute(
            """INSERT INTO reports (course_id, checked_at, fields_checked, status,
               errors_json, rewrites_json, summary)
               VALUES (?,?,?,?,?,?,?)""",
            (course_id, datetime.now().isoformat(timespec="seconds"),
             json.dumps(fields), status, json.dumps(errors, ensure_ascii=False),
             json.dumps(rewrites or [], ensure_ascii=False), summary),
        )


def latest_reports() -> list:
    """Latest report per course, joined with the course name."""
    q = """
        SELECT r.*, c.course_name, c.course_url
        FROM reports r
        JOIN courses c ON c.id = r.course_id
        WHERE r.id IN (SELECT MAX(id) FROM reports GROUP BY course_id)
        ORDER BY c.course_name
    """
    with get_conn() as c:
        return [dict(x) for x in c.execute(q)]


# ═══════════════════════════════════════════════════════════════════
#  FETCHING & EXTRACTION
# ═══════════════════════════════════════════════════════════════════

def fetch_url(url: str, timeout=30) -> requests.Response:
    return requests.get(url, headers=USER_AGENT, timeout=timeout, allow_redirects=True)


def extract_page_text(url: str, max_chars=12000) -> str:
    """Fetch a live course page and return its readable text."""
    resp = fetch_url(url)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, "html.parser")
    for tag in soup(["script", "style", "nav", "footer", "header", "noscript", "form"]):
        tag.decompose()
    main = soup.find("main") or soup.find("article") or soup.body or soup
    text = re.sub(r"\n{3,}", "\n\n", main.get_text("\n", strip=True))
    return text[:max_chars]


def extract_spec_text(url: str, max_chars=15000) -> str:
    """Extract text from a specification document — PDF or web page."""
    resp = fetch_url(url, timeout=60)
    resp.raise_for_status()
    ctype = resp.headers.get("Content-Type", "").lower()
    if "pdf" in ctype or url.lower().split("?")[0].endswith(".pdf"):
        import pdfplumber
        parts = []
        with pdfplumber.open(io.BytesIO(resp.content)) as pdf:
            for page in pdf.pages:
                parts.append(page.extract_text() or "")
                if sum(len(p) for p in parts) > max_chars:
                    break
        return "\n".join(parts)[:max_chars]
    soup = BeautifulSoup(resp.text, "html.parser")
    for tag in soup(["script", "style", "nav", "footer", "header", "noscript"]):
        tag.decompose()
    return soup.get_text("\n", strip=True)[:max_chars]


# ═══════════════════════════════════════════════════════════════════
#  OPENROUTER
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
    data = resp.json()
    return data["choices"][0]["message"]["content"]


def parse_json_reply(raw: str) -> dict:
    """Robustly pull a JSON object out of an LLM reply."""
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
#  COURSE CHECKING
# ═══════════════════════════════════════════════════════════════════

CHECK_SYSTEM = (
    "You are a meticulous content auditor AND professional copywriter for South London College. "
    "You compare the LIVE course web page against the internal TRACKER content and the OFFICIAL "
    "qualification SPECIFICATION. You flag factual mismatches, missing information, outdated "
    "details and misleading statements. You also write replacement wording that is 100% original: "
    "factually faithful to the specification but never copying its sentences or distinctive "
    "phrases. You reply ONLY with valid JSON — no markdown, no commentary."
)

CHECK_PROMPT = """Compare the three sources below for the course "{name}" and audit ONLY these fields: {fields}.

For each audited field decide whether the LIVE PAGE content is accurate and consistent with the TRACKER and the SPECIFICATION. Ignore styling/wording differences that do not change meaning. Flag:
- factual mismatches (grades, units, hours, assessment methods, requirements)
- content present in tracker/spec but missing on the live page
- content on the live page contradicted by the specification

Reply with EXACTLY this JSON shape:
{{
  "status": "pass" or "errors",
  "summary": "one-sentence overall verdict",
  "errors": [
    {{
      "field": "Entry Requirements | Method of Assessment | Course Overview",
      "severity": "high | medium | low",
      "issue": "clear description of the problem",
      "live_content": "the problematic text on the live page (short quote or 'missing')",
      "expected_content": "what tracker/spec says it should be",
      "suggested_fix": "the exact corrected text or action to take"
    }}
  ],
  "rewrites": [
    {{
      "field": "Entry Requirements | Method of Assessment | Course Overview",
      "suggested_wording": "a complete, publish-ready rewrite of this field for the website"
    }}
  ]
}}

RULES FOR "rewrites" (very important):
- Provide one rewrite for EVERY field that has at least one error (empty list if none).
- The rewrite must be publish-ready website copy in UK English: complete, clear, professional.
- It must be factually accurate to the SPECIFICATION and TRACKER (grades, units, hours, methods).
- ZERO PLAGIARISM: write in completely original words. Do NOT reuse sentences, clauses, or
  distinctive phrases of 4+ consecutive words from the specification, the live page, or the
  tracker. Restructure sentences and use your own vocabulary while keeping every fact identical.
- Do not invent facts that are not supported by the specification or tracker.

If everything is accurate return status "pass" and empty errors and rewrites lists.

=== TRACKER CONTENT ===
{tracker}

=== LIVE PAGE CONTENT ===
{live}

=== OFFICIAL SPECIFICATION EXTRACT ===
{spec}
"""


def _shingles(text: str, n: int = 4) -> set:
    words = re.findall(r"[a-z0-9']+", (text or "").lower())
    return {tuple(words[i:i + n]) for i in range(len(words) - n + 1)}


def originality_score(candidate: str, sources: list, n: int = 4) -> float:
    """% of the candidate's 4-word phrases that do NOT appear in any source.
    100 = fully original wording, 0 = copied verbatim."""
    cand = _shingles(candidate, n)
    if not cand:
        return 100.0
    src = set()
    for s in sources:
        src |= _shingles(s, n)
    overlap = len(cand & src) / len(cand)
    return round((1 - overlap) * 100, 1)


REWRITE_RETRY_PROMPT = """Rewrite the text below so that it keeps EXACTLY the same facts but shares NO phrase of 4 or more consecutive words with any of the reference texts. Use different sentence structures and vocabulary. UK English, publish-ready website copy. Reply with ONLY the rewritten text — no JSON, no quotes, no commentary.

=== TEXT TO REWRITE ===
{text}

=== REFERENCE TEXTS IT MUST NOT COPY FROM ===
{refs}
"""


def ensure_original(wording: str, sources: list, api_key: str, model: str,
                    threshold: float = 90.0, max_retries: int = 2) -> tuple:
    """Verify a rewrite is plagiarism-free; if not, ask the model to re-word it.
    Returns (final_wording, originality_score)."""
    score = originality_score(wording, sources)
    tries = 0
    while score < threshold and tries < max_retries:
        try:
            refs = "\n\n---\n\n".join(s[:4000] for s in sources if s)
            wording = call_openrouter(
                REWRITE_RETRY_PROMPT.format(text=wording, refs=refs),
                "You are a professional copywriter. You paraphrase with zero plagiarism.",
                api_key, model, temperature=0.7,
            ).strip()
            score = originality_score(wording, sources)
        except Exception:
            break
        tries += 1
    return wording, score


def check_course(course: dict, fields: list, api_key: str, model: str) -> dict:
    """Run a full 3-way check for one course. Returns a report dict."""
    report = {"course_id": course["id"], "course_name": course["course_name"],
              "status": "failed", "summary": "", "errors": [], "rewrites": []}

    # 1) live page
    try:
        live = extract_page_text(course["course_url"]) if course.get("course_url") else ""
        if not live:
            raise ValueError("No course URL / empty page")
    except Exception as e:
        report["summary"] = f"Could not fetch live page: {e}"
        return report

    # 2) spec text (extract on the fly if missing)
    spec = course.get("spec_text") or ""
    if not spec and course.get("spec_url"):
        try:
            spec = extract_spec_text(course["spec_url"])
            save_spec_text(course["id"], spec)
        except Exception:
            spec = "(specification unavailable)"
    spec = spec or "(no specification provided)"

    # 3) tracker content for the selected fields
    tracker_parts = []
    for key in fields:
        label = FIELD_OPTIONS[key]
        tracker_parts.append(f"[{label}]\n{course.get(key) or '(not provided in tracker)'}")
    tracker = "\n\n".join(tracker_parts)

    prompt = CHECK_PROMPT.format(
        name=course["course_name"],
        fields=", ".join(FIELD_OPTIONS[k] for k in fields),
        tracker=tracker, live=live, spec=spec,
    )

    try:
        raw = call_openrouter(prompt, CHECK_SYSTEM, api_key, model)
        data = parse_json_reply(raw)
        report["status"] = "errors" if data.get("errors") else "pass"
        if data.get("status") == "pass":
            report["status"] = "pass"
        report["summary"] = data.get("summary", "")
        report["errors"] = data.get("errors", [])

        # ── wording suggestions with zero-plagiarism verification ──
        sources = [spec, live, tracker]
        rewrites = []
        for rw in data.get("rewrites", []) or []:
            wording = str(rw.get("suggested_wording", "")).strip()
            if not wording:
                continue
            wording, score = ensure_original(wording, sources, api_key, model)
            rewrites.append({
                "field": rw.get("field", ""),
                "suggested_wording": wording,
                "originality": score,
            })
        report["rewrites"] = rewrites
    except Exception as e:
        report["summary"] = f"AI check failed: {e}"
    return report


# ═══════════════════════════════════════════════════════════════════
#  WORD REPORT (errors only)
# ═══════════════════════════════════════════════════════════════════

SEV_COLORS = {"high": RGBColor(0xC6, 0x28, 0x28),
              "medium": RGBColor(0xE6, 0x7E, 0x22),
              "low": RGBColor(0x2E, 0x7D, 0x32)}


def build_word_report(error_reports: list) -> bytes:
    doc = Document()

    title = doc.add_heading("SLC Course Content Error Report", level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub = doc.add_paragraph(
        f"Generated {datetime.now().strftime('%d %B %Y, %H:%M')}  ·  "
        f"{len(error_reports)} course(s) with errors"
    )
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub.runs[0].font.size = Pt(10)
    sub.runs[0].font.color.rgb = RGBColor(0x66, 0x66, 0x66)
    doc.add_paragraph()

    for rep in error_reports:
        doc.add_heading(rep["course_name"], level=1)
        if rep.get("course_url"):
            p = doc.add_paragraph()
            r = p.add_run(rep["course_url"])
            r.font.size = Pt(9)
            r.font.color.rgb = RGBColor(0x1A, 0x56, 0xDB)
        if rep.get("summary"):
            p = doc.add_paragraph(rep["summary"])
            p.runs[0].italic = True

        errors = rep.get("errors") or []
        if not errors:
            doc.add_paragraph("Check failed — see summary above.")
            continue

        table = doc.add_table(rows=1, cols=5)
        table.style = "Light Grid Accent 1"
        hdr = table.rows[0].cells
        for i, h in enumerate(["Field", "Severity", "Issue", "Live page says", "Suggested solution"]):
            hdr[i].text = h
            for run in hdr[i].paragraphs[0].runs:
                run.bold = True

        for err in errors:
            cells = table.add_row().cells
            cells[0].text = str(err.get("field", ""))
            sev = str(err.get("severity", "medium")).lower()
            cells[1].text = sev.upper()
            for run in cells[1].paragraphs[0].runs:
                run.bold = True
                run.font.color.rgb = SEV_COLORS.get(sev, SEV_COLORS["medium"])
            cells[2].text = str(err.get("issue", ""))
            cells[3].text = str(err.get("live_content", ""))
            fix_cell = cells[4]
            fix_cell.text = str(err.get("suggested_fix", ""))
            if err.get("expected_content"):
                p = fix_cell.add_paragraph()
                r = p.add_run(f"Expected: {err['expected_content']}")
                r.font.size = Pt(8)
                r.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

        rewrites = rep.get("rewrites") or []
        if rewrites:
            doc.add_heading("Suggested replacement wording (plagiarism-checked)", level=2)
            for rw in rewrites:
                p = doc.add_paragraph()
                r = p.add_run(f"{rw.get('field','')}  ·  originality {rw.get('originality','—')}%")
                r.bold = True
                r.font.color.rgb = RGBColor(0x1D, 0x4E, 0x89)
                doc.add_paragraph(str(rw.get("suggested_wording", "")))

        doc.add_paragraph()

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ═══════════════════════════════════════════════════════════════════
#  CONTENT QUALITY REVIEW
# ═══════════════════════════════════════════════════════════════════

QR_SYSTEM = (
    "You are a professional UK-English proofreader and copy editor for a college website. "
    "You review course content for grammar, articles (a/an/the), sentence structure, "
    "capitalisation, proper nouns, spelling, commas and punctuation consistency. "
    "You reply ONLY with valid JSON."
)

QR_PROMPT = """Proofread the text below. Find every issue and classify it into EXACTLY one of these categories:
grammar, article, spelling, punctuation, capitalisation, proper_noun, sentence_structure, consistency

Rules:
- "original" must be an EXACT substring copied verbatim from the text (short — the smallest span that contains the problem).
- "correction" is the fixed version of that span.
- "explanation" is one short sentence.
- Also produce the fully corrected version of the whole text.

Reply with EXACTLY this JSON:
{{
  "issues": [
    {{"category": "...", "original": "...", "correction": "...", "explanation": "..."}}
  ],
  "corrected_text": "..."
}}

=== TEXT TO REVIEW ===
{text}
"""


def run_quality_review(text: str, api_key: str, model: str) -> dict:
    raw = call_openrouter(QR_PROMPT.format(text=text), QR_SYSTEM, api_key, model)
    return parse_json_reply(raw)


def annotate_text_html(text: str, issues: list) -> str:
    """Return HTML with each issue wrapped in a coloured <mark>, numbered like a proofreader's markup."""
    escaped = html.escape(text)
    for n, issue in enumerate(issues, start=1):
        original = html.escape(str(issue.get("original", "")))
        if not original:
            continue
        cat = issue.get("category", "grammar")
        style = QR_CATEGORIES.get(cat, QR_CATEGORIES["grammar"])
        tip = html.escape(f"{style['label']}: {issue.get('correction','')} — {issue.get('explanation','')}")
        mark = (
            f'<mark class="qr-mark" style="background:{style["bg"]};'
            f'border-bottom:2px solid {style["border"]};" title="{tip}">'
            f'<sup class="qr-num" style="background:{style["border"]};">{n}</sup>{original}</mark>'
        )
        escaped = escaped.replace(original, mark, 1)
    return escaped.replace("\n", "<br>")


# ═══════════════════════════════════════════════════════════════════
#  UI — THEME
# ═══════════════════════════════════════════════════════════════════

st.set_page_config(page_title="SLC Course Content Checker", page_icon="🎓",
                   layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,600;9..144,700&family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@500&display=swap');

html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

/* ── Hero banner ─────────────────────────────── */
.slc-hero {
  background: linear-gradient(120deg, #10243E 0%, #1D4E89 55%, #2E7D6B 100%);
  border-radius: 18px; padding: 26px 34px; margin-bottom: 6px;
  position: relative; overflow: hidden;
}
.slc-hero:before {
  content:""; position:absolute; right:-40px; top:-40px; width:220px; height:220px;
  background: radial-gradient(circle, rgba(255,229,92,.35), transparent 70%);
}
.slc-hero h1 {
  font-family:'Fraunces', serif; color:#fff; margin:0; font-size:2rem; letter-spacing:.3px;
}
.slc-hero h1 .hl { background:#FFE55C; color:#10243E; padding:0 .35rem; border-radius:6px; }
.slc-hero p { color:#CFE3F5; margin:.45rem 0 0; font-size:.95rem; }

/* ── Stat cards ─────────────────────────────── */
.stat-card {
  border-radius:14px; padding:16px 18px; border:1px solid #E6E9EF;
  background:#fff; box-shadow:0 2px 10px rgba(16,36,62,.06);
}
.stat-card .num { font-family:'Fraunces',serif; font-size:1.9rem; font-weight:700; color:#10243E; line-height:1.1;}
.stat-card .lbl { font-size:.78rem; text-transform:uppercase; letter-spacing:.09em; color:#6B7688; font-weight:600;}
.stat-card.ok   { border-top:4px solid #30A46C; }
.stat-card.err  { border-top:4px solid #E5484D; }
.stat-card.info { border-top:4px solid #1D4E89; }
.stat-card.warn { border-top:4px solid #F5B300; }

/* ── Quality-review markup ───────────────────── */
.qr-paper {
  background:#FFFDF6; border:1px solid #EDE6D2; border-radius:14px;
  padding:26px 30px; line-height:2.05; font-size:1.0rem; color:#2A2F3A;
  box-shadow: 0 3px 14px rgba(16,36,62,.07);
}
.qr-mark { border-radius:4px; padding:1px 3px; cursor:help; position:relative; }
.qr-num {
  color:#fff; font-size:.62rem; font-weight:700; border-radius:999px;
  padding:0 4px; margin-right:2px; position:relative; top:-7px;
}
.qr-legend span {
  display:inline-block; margin:3px 8px 3px 0; padding:3px 10px; border-radius:999px;
  font-size:.78rem; font-weight:600;
}
.issue-card {
  border-radius:12px; border:1px solid #E6E9EF; border-left-width:5px;
  padding:12px 16px; margin-bottom:10px; background:#fff;
}
.issue-card .cat { font-size:.72rem; font-weight:700; text-transform:uppercase; letter-spacing:.08em;}
.issue-card .orig { text-decoration:line-through; color:#B0341F; }
.issue-card .corr { color:#1D7A46; font-weight:600; }

/* ── Buttons & tabs polish ───────────────────── */
.stButton>button[kind="primary"] {
  background:linear-gradient(120deg,#1D4E89,#2E7D6B); border:none; border-radius:10px;
  font-weight:600;
}
.stTabs [data-baseweb="tab"] { font-weight:600; }
div[data-testid="stSidebar"] { background:#F6F8FB; }
</style>
""", unsafe_allow_html=True)


def hero():
    st.markdown(
        """
        <div class="slc-hero">
          <h1>🎓 SLC <span class="hl">Course Content</span> Checker</h1>
          <p>Compare live course pages · tracker sheet · official specifications — and proofread like a pro.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


def stat(col, value, label, kind="info"):
    col.markdown(
        f'<div class="stat-card {kind}"><div class="num">{value}</div>'
        f'<div class="lbl">{label}</div></div>',
        unsafe_allow_html=True,
    )


# ═══════════════════════════════════════════════════════════════════
#  UI — SIDEBAR
# ═══════════════════════════════════════════════════════════════════

init_db()
hero()

# API key comes from Streamlit secrets (.streamlit/secrets.toml), falling back to an env var
import os
api_key = st.secrets.get("OPENROUTER_API_KEY", os.environ.get("OPENROUTER_API_KEY", ""))

with st.sidebar:
    st.markdown("### ⚙️ Settings")
    if api_key:
        st.success("🔑 OpenRouter key loaded from secrets")
    else:
        st.error("No API key found. Add `OPENROUTER_API_KEY` to `.streamlit/secrets.toml`.")
    model = st.selectbox("AI model", MODEL_CHOICES, index=0)
    st.divider()
    st.markdown("### 🧭 Fields to check")
    selected_fields = [k for k, label in FIELD_OPTIONS.items()
                       if st.checkbox(label, value=True, key=f"field_{k}")]


# ═══════════════════════════════════════════════════════════════════
#  UI — TABS
# ═══════════════════════════════════════════════════════════════════

tab_import, tab_spec, tab_check, tab_reports, tab_quality = st.tabs(
    ["📥 Import Courses", "📑 Spec Extraction", "🔍 Run Checks", "📊 Reports & Word Export", "✍️ Content Quality Review"]
)

# ───────────────────────────────────────────────
# TAB 1 · IMPORT
# ───────────────────────────────────────────────
with tab_import:
    st.subheader("Upload the Excel tracker sheet")
    st.caption("Expected columns: course name, course URL, spec URL, entry requirements, "
               "method of assessment, course overview. You can remap columns below.")

    up = st.file_uploader("Tracker sheet (.xlsx / .xls / .csv)", type=["xlsx", "xls", "csv"])

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
            m_url = pick("Course page URL", guess("course url", "page url", "link", "url"), "m2")
        with c2:
            m_spec = pick("Specification URL", guess("spec"), "m3")
            m_entry = pick("Entry Requirements", guess("entry"), "m4")
        with c3:
            m_assess = pick("Method of Assessment", guess("assess"), "m5")
            m_over = pick("Course Overview", guess("overview", "description"), "m6")

        if st.button("📥 Import / update courses", type="primary", disabled=(m_name == "— none —")):
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
                result = upsert_course({
                    "course_name": name,
                    "course_url": val(row, m_url),
                    "spec_url": val(row, m_spec),
                    "entry_requirements": val(row, m_entry),
                    "method_of_assessment": val(row, m_assess),
                    "course_overview": val(row, m_over),
                })
                inserted += result == "inserted"
                updated += result == "updated"

            s1, s2, s3 = st.columns(3)
            stat(s1, inserted, "New courses", "ok")
            stat(s2, updated, "Updated", "info")
            stat(s3, skipped, "Skipped (no name)", "warn")
            st.success("Tracker imported into the database ✅")

    st.divider()
    if courses_now := all_courses():
        st.markdown("#### 📚 Courses in database")
        st.dataframe(
            pd.DataFrame(courses_now)[["id", "course_name", "course_url", "spec_url", "updated_at"]],
            use_container_width=True, hide_index=True,
        )

# ───────────────────────────────────────────────
# TAB 2 · SPEC EXTRACTION
# ───────────────────────────────────────────────
with tab_spec:
    st.subheader("Extract missing specification text")
    courses = all_courses()
    missing = [c for c in courses if c.get("spec_url") and not c.get("spec_text")]
    have = [c for c in courses if c.get("spec_text")]

    s1, s2, s3 = st.columns(3)
    stat(s1, len(courses), "Total courses", "info")
    stat(s2, len(have), "Specs extracted", "ok")
    stat(s3, len(missing), "Specs missing", "err" if missing else "ok")
    st.write("")

    if not missing:
        st.info("No courses are missing specification text. 🎉")
    else:
        st.dataframe(pd.DataFrame(missing)[["course_name", "spec_url"]],
                     use_container_width=True, hide_index=True)
        targets = missing
        if st.button(f"📑 Extract {len(targets)} missing spec(s)", type="primary"):
            bar = st.progress(0.0)
            status = st.empty()
            ok = fail = 0
            for i, c in enumerate(targets):
                status.markdown(f"Extracting **{c['course_name']}** …")
                try:
                    txt = extract_spec_text(c["spec_url"])
                    save_spec_text(c["id"], txt)
                    ok += 1
                except Exception as e:
                    fail += 1
                    st.warning(f"❌ {c['course_name']}: {e}")
                bar.progress((i + 1) / len(targets))
            status.empty()
            st.success(f"Done — {ok} extracted, {fail} failed.")
            st.rerun()

    if have:
        with st.expander("🔎 Preview an extracted specification"):
            pick_name = st.selectbox("Course", [c["course_name"] for c in have])
            spec = next(c for c in have if c["course_name"] == pick_name)
            st.text_area("Specification text", spec["spec_text"], height=280)

# ───────────────────────────────────────────────
# TAB 3 · RUN CHECKS
# ───────────────────────────────────────────────
with tab_check:
    st.subheader("Check course content against the live site & specification")
    courses = all_courses()

    if not courses:
        st.info("Import your tracker sheet first (📥 Import Courses tab).")
    elif not selected_fields:
        st.warning("Select at least one field to check in the sidebar.")
    else:
        st.caption(f"Checking fields: **{', '.join(FIELD_OPTIONS[k] for k in selected_fields)}**")

        colA, colB = st.columns([1, 1], gap="large")

        # ── Single course check ──
        with colA:
            st.markdown("#### 🎯 Check one course")
            name = st.selectbox("Course", [c["course_name"] for c in courses])
            if st.button("🔍 Check this course", type="primary", key="single"):
                if not api_key:
                    st.error("No API key found — add OPENROUTER_API_KEY to .streamlit/secrets.toml and restart.")
                else:
                    course = next(c for c in courses if c["course_name"] == name)
                    with st.spinner(f"Checking {name} …"):
                        rep = check_course(course, selected_fields, api_key, model)
                    save_report(course["id"], selected_fields, rep["status"],
                                rep["errors"], rep["summary"], rep.get("rewrites"))
                    if rep["status"] == "pass":
                        st.success(f"✅ **{name}** — no errors found. {rep['summary']}")
                    elif rep["status"] == "errors":
                        st.error(f"⚠️ **{name}** — {len(rep['errors'])} issue(s). {rep['summary']}")
                        for err in rep["errors"]:
                            with st.expander(f"❌ {err.get('field')} · {str(err.get('severity','')).upper()} — {err.get('issue','')[:70]}"):
                                st.markdown(f"**Issue:** {err.get('issue')}")
                                st.markdown(f"**Live page says:** {err.get('live_content')}")
                                st.markdown(f"**Expected:** {err.get('expected_content')}")
                                st.markdown(f"**💡 Suggested fix:** {err.get('suggested_fix')}")
                        for rw in rep.get("rewrites", []):
                            orig = rw.get("originality", 0)
                            badge = "🟢" if orig >= 90 else ("🟡" if orig >= 75 else "🔴")
                            with st.expander(f"✨ Suggested wording — {rw.get('field')} "
                                             f"({badge} {orig}% original)", expanded=True):
                                st.text_area("Copy-ready replacement text",
                                             rw.get("suggested_wording", ""),
                                             height=160, key=f"rw_{name}_{rw.get('field')}")
                                st.caption(f"Originality {orig}% — share of 4-word phrases NOT "
                                           "found in the specification, live page or tracker. "
                                           "Verified plagiarism-safe before display.")
                    else:
                        st.warning(f"Check failed — {rep['summary']}")

        # ── Bulk check ──
        with colB:
            st.markdown("#### 🚀 Bulk check")
            targets = courses
            st.caption(f"{len(targets)} course(s) will be checked")
            if st.button(f"🚀 Run bulk check on {len(targets)} courses", type="primary", key="bulk"):
                if not api_key:
                    st.error("No API key found — add OPENROUTER_API_KEY to .streamlit/secrets.toml and restart.")
                else:
                    bar = st.progress(0.0)
                    status = st.empty()
                    results = {"pass": 0, "errors": 0, "failed": 0}
                    log = st.container()
                    for i, course in enumerate(targets):
                        status.markdown(f"⏳ **{i+1}/{len(targets)}** — {course['course_name']}")
                        rep = check_course(course, selected_fields, api_key, model)
                        save_report(course["id"], selected_fields, rep["status"],
                                    rep["errors"], rep["summary"], rep.get("rewrites"))
                        results[rep["status"]] += 1
                        icon = {"pass": "✅", "errors": "⚠️", "failed": "❌"}[rep["status"]]
                        log.markdown(f"{icon} **{course['course_name']}** — "
                                     f"{len(rep['errors'])} issue(s). {rep['summary'][:120]}")
                        bar.progress((i + 1) / len(targets))
                        time.sleep(0.4)  # be gentle to the site & API
                    status.empty()
                    r1, r2, r3 = st.columns(3)
                    stat(r1, results["pass"], "Passed", "ok")
                    stat(r2, results["errors"], "With errors", "err")
                    stat(r3, results["failed"], "Failed to check", "warn")
                    st.success("Bulk check complete — reports saved to the database. "
                               "Head to 📊 Reports to export the Word document.")

# ───────────────────────────────────────────────
# TAB 4 · REPORTS & WORD EXPORT
# ───────────────────────────────────────────────
with tab_reports:
    st.subheader("Saved reports")
    reports = latest_reports()

    if not reports:
        st.info("No reports yet — run some checks first.")
    else:
        n_pass = sum(r["status"] == "pass" for r in reports)
        n_err = sum(r["status"] == "errors" for r in reports)
        n_fail = sum(r["status"] == "failed" for r in reports)
        c1, c2, c3, c4 = st.columns(4)
        stat(c1, len(reports), "Courses checked", "info")
        stat(c2, n_pass, "Passed", "ok")
        stat(c3, n_err, "With errors", "err")
        stat(c4, n_fail, "Check failed", "warn")
        st.write("")

        table = pd.DataFrame([{
            "Course": r["course_name"],
            "Status": {"pass": "✅ Pass", "errors": "⚠️ Errors", "failed": "❌ Failed"}[r["status"]],
            "Issues": len(json.loads(r["errors_json"] or "[]")),
            "Checked at": r["checked_at"],
            "Summary": r["summary"],
        } for r in reports])
        st.dataframe(table, use_container_width=True, hide_index=True)

        # detail viewer
        err_reports_db = [r for r in reports if r["status"] == "errors"]
        if err_reports_db:
            with st.expander("🔎 View error details"):
                sel = st.selectbox("Course with errors", [r["course_name"] for r in err_reports_db])
                r = next(x for x in err_reports_db if x["course_name"] == sel)
                for err in json.loads(r["errors_json"]):
                    style = {"high": "#E5484D", "medium": "#F76B15", "low": "#30A46C"}.get(
                        str(err.get("severity", "medium")).lower(), "#F76B15")
                    st.markdown(
                        f'<div class="issue-card" style="border-left-color:{style}">'
                        f'<div class="cat" style="color:{style}">{err.get("field","")} · {err.get("severity","")}</div>'
                        f'<b>{html.escape(str(err.get("issue","")))}</b><br>'
                        f'<span class="orig">{html.escape(str(err.get("live_content","")))}</span> → '
                        f'<span class="corr">{html.escape(str(err.get("suggested_fix","")))}</span>'
                        f'</div>', unsafe_allow_html=True)
                for rw in json.loads(r.get("rewrites_json") or "[]"):
                    orig = rw.get("originality", 0)
                    badge = "🟢" if orig >= 90 else ("🟡" if orig >= 75 else "🔴")
                    st.markdown(f"**✨ Suggested wording — {rw.get('field')} "
                                f"({badge} {orig}% original)**")
                    st.text_area("Copy-ready replacement text", rw.get("suggested_wording", ""),
                                 height=150, key=f"rep_rw_{r['course_id']}_{rw.get('field')}")

        st.divider()
        st.markdown("#### 📝 Word document — courses with errors only")
        if not err_reports_db:
            st.success("No courses with errors — nothing to export. 🎉")
        else:
            word_input = [{
                "course_name": r["course_name"],
                "course_url": r["course_url"],
                "summary": r["summary"],
                "errors": json.loads(r["errors_json"] or "[]"),
                "rewrites": json.loads(r.get("rewrites_json") or "[]"),
            } for r in err_reports_db]
            docx_bytes = build_word_report(word_input)
            st.download_button(
                f"⬇️ Download Word report ({len(err_reports_db)} courses with errors)",
                data=docx_bytes,
                file_name=f"SLC_Course_Error_Report_{datetime.now():%Y-%m-%d}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary",
            )

# ───────────────────────────────────────────────
# TAB 5 · CONTENT QUALITY REVIEW
# ───────────────────────────────────────────────
with tab_quality:
    st.subheader("✍️ Content Quality Review")
    st.caption("Paste or upload course content. The reviewer highlights grammar, articles, "
               "sentence structure, capitalisation, proper nouns, spelling, commas and "
               "punctuation consistency — colour-coded like a proofreader's markup.")

    # legend
    legend = "".join(
        f'<span style="background:{v["bg"]};border:1px solid {v["border"]};color:#2A2F3A;">{v["label"]}</span>'
        for v in QR_CATEGORIES.values()
    )
    st.markdown(f'<div class="qr-legend">{legend}</div>', unsafe_allow_html=True)
    st.write("")

    src = st.radio("Input", ["Paste text", "Upload file (.txt / .docx)"], horizontal=True)
    text = ""
    if src == "Paste text":
        text = st.text_area("Course content to review", height=220,
                            placeholder="Paste the course description, overview or any page copy here…")
    else:
        f = st.file_uploader("Upload content", type=["txt", "docx"], key="qr_up")
        if f:
            if f.name.lower().endswith(".docx"):
                d = Document(io.BytesIO(f.read()))
                text = "\n".join(p.text for p in d.paragraphs if p.text.strip())
            else:
                text = f.read().decode("utf-8", errors="ignore")
            st.text_area("Loaded content", text, height=180)

    if st.button("✍️ Review content quality", type="primary", disabled=not text.strip()):
        if not api_key:
            st.error("No API key found — add OPENROUTER_API_KEY to .streamlit/secrets.toml and restart.")
        else:
            with st.spinner("Proofreading …"):
                try:
                    result = run_quality_review(text, api_key, model)
                    st.session_state["qr_result"] = result
                    st.session_state["qr_text"] = text
                except Exception as e:
                    st.error(f"Review failed: {e}")

    if "qr_result" in st.session_state:
        result = st.session_state["qr_result"]
        text = st.session_state["qr_text"]
        issues = result.get("issues", [])

        c1, c2, c3 = st.columns(3)
        stat(c1, len(issues), "Issues found", "err" if issues else "ok")
        top_cat = max({i.get("category") for i in issues},
                      key=lambda c: sum(1 for i in issues if i.get("category") == c),
                      default="—")
        stat(c2, QR_CATEGORIES.get(top_cat, {}).get("label", "—"), "Most common issue", "warn")
        stat(c3, f"{max(0, 100 - len(issues) * 4)}%", "Quality score", "info")
        st.write("")

        view_marked, view_fixed, view_list = st.tabs(
            ["🖍️ Marked-up text", "✅ Corrected text", "📋 Issue list"])

        with view_marked:
            if issues:
                st.markdown(f'<div class="qr-paper">{annotate_text_html(text, issues)}</div>',
                            unsafe_allow_html=True)
                st.caption("Hover a highlight to see the correction and explanation.")
            else:
                st.success("No issues found — this content is clean. 🎉")

        with view_fixed:
            corrected = result.get("corrected_text", text)
            st.text_area("Corrected version (copy-ready)", corrected, height=260)
            st.download_button("⬇️ Download corrected text", corrected,
                               file_name="corrected_content.txt")

        with view_list:
            if not issues:
                st.success("Nothing to list — no issues found.")
            for n, issue in enumerate(issues, start=1):
                cat = issue.get("category", "grammar")
                sty = QR_CATEGORIES.get(cat, QR_CATEGORIES["grammar"])
                st.markdown(
                    f'<div class="issue-card" style="border-left-color:{sty["border"]}">'
                    f'<div class="cat" style="color:{sty["border"]}">#{n} · {sty["label"]}</div>'
                    f'<span class="orig">{html.escape(str(issue.get("original","")))}</span> → '
                    f'<span class="corr">{html.escape(str(issue.get("correction","")))}</span><br>'
                    f'<small>{html.escape(str(issue.get("explanation","")))}</small>'
                    f'</div>', unsafe_allow_html=True)