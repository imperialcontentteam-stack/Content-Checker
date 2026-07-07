# 🎓 SLC Course Content Checker — v5

Internal Streamlit tool for the South London College content team: authentication,
tracker-driven filters, **one-time AI extraction of qualification specification
documents** (stored as structured JSON and reused for every check), live-page
validation with downloadable reports, plus separate Grammar Check and Other
Checks tabs.

Stack: **Python · Streamlit · SQLite · OpenRouter API · reportlab**. One file, one local database.

## Setup

```bash
pip install -r requirements.txt
streamlit run app.py
```

Optionally get an OpenRouter API key at https://openrouter.ai/keys and add it to
`.streamlit/secrets.toml` as `OPENROUTER_API_KEY = "..."` (or set it as an
environment variable). Without a key, document extraction and validation fall
back to built-in heuristics. The SQLite database (`slc_checker.db`) is created
automatically next to `app.py`.

## Accounts & permissions

Default logins on first run (change them on the 👥 Users page!):

| Account | Password | Role |
|---|---|---|
| `admin` | `admin123` | Administrator |
| `user`  | `user123`  | User |

**Admin** — import the course Excel tracker, upload & manage specification
documents, reprocess documents when they change, manage course records and user
accounts, plus everything a user can do.

**User** — Run Check, Grammar Check, Other Checks and Reports: run checks, view
and download reports. No permission to upload or modify any data.

Passwords are stored as salted PBKDF2-SHA256 hashes (200k iterations).

## The one-time extraction architecture

1. **📥 Import Courses (admin)** — upload the Excel tracker (Course ID/Number,
   Course Name, Level, Type, Course URL, Qualification Specification Document
   URL; columns are auto-guessed and remappable; Level/Type are derived from
   the course name when not provided). All rows are read automatically. Every
   **distinct** specification URL is registered as a document and linked to its
   courses — when several courses share the same document, this is detected and
   they share one extraction.
2. **📑 Spec Documents (admin)** — process documents **once**: the document is
   fetched from its URL (or uploaded as PDF/Word), read, and AI-extracted into
   structured JSON stored in the database: Qualification Name, Level, Type,
   Entry Requirements, Method of Assessment, Qualification Specification
   Requirements, Learning Outcomes, Mandatory Units and Other Relevant
   Information. Every document and its stored data is listed on the page at all
   times (with a filter box) — no picking required. A **Process all unprocessed**
   button handles new imports in one go. Already-processed documents are skipped (reused); re-fetching an
   unchanged document is detected by content hash and skipped too. **Reprocess**
   is available for when a document has been updated, and every extraction is
   editable.
3. **🔍 Run Check** — filter by **Level**, **Type** and **Course**. The page
   shows the course details (left) and the stored specification data (right),
   loads the **live course page**, and compares it against the **stored**
   extraction — the specification document is never re-read or re-extracted at
   check time. The AI flags correct content, incorrect wording, incorrect
   information, missing information, mismatched requirements and grammar issues,
   with suggested corrections. Reports use the agreed layout (red section boxes,
   numbered red *Errors identified* with green *Recommend Action*, a green *No
   Errors* badge on a clean pass, and a summary table) and can be downloaded as
   **PDF or Word** from the same page.
4. **✍️ Grammar Check** — its own tab (moved out of Run Check): the colour-coded
   proofreading review for pasted text, uploaded files or a course's overview.
5. **🧪 Other Checks** — its own tab: the original live-page 3-way audit
   (single and bulk) with plagiarism-safe rewrite suggestions.
6. **📊 Reports** — validation reports (view + PDF/Word download per course) and
   the live-page audit dashboard with the errors-only Word export.
7. **👥 Users (admin)** — create accounts, reset passwords, delete accounts.

## Performance

- Each specification document is extracted **once** and the structured JSON is
  reused across all courses and all checks — no duplicate processing.
- Shared documents are detected by URL, so 10 courses on one specification cost
  one extraction.
- Validation prompts use the compact stored data instead of raw documents,
  cutting AI token usage substantially and speeding up report generation.
- Content-hash checks mean re-fetching an unchanged document never triggers
  re-extraction; reprocessing happens only when a document has actually been
  updated (or is forced by an admin).
