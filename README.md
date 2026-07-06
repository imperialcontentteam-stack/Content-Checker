# 🎓 SLC Course Content Checker — v4

Internal Streamlit tool for the South London College content team. It compares imported course data against uploaded qualification specification documents and produces a validation report (on-screen and as a downloadable Word document) in the standard SLC error-report format.

Stack: **Python · Streamlit · SQLite · OpenRouter API** — one file, one local database.

## Setup

```bash
pip install -r requirements.txt
streamlit run app.py
```

The SQLite database (`slc_checker.db`) is created automatically next to `app.py`.

Optional: add an OpenRouter key to `.streamlit/secrets.toml` for AI-powered comparison:

```toml
OPENROUTER_API_KEY = "sk-or-..."
```

Without a key the tool falls back to a built-in offline text comparison.

## 🔐 Authentication

Two roles. Default accounts are created on first run — **change these passwords immediately** (sidebar → Change my password, or Admin → Manage Users):

| Account | Password | Role |
|---|---|---|
| `admin` | `admin123` | Admin |
| `user` | `user123` | User |

**Admin can:** import the course Excel file · upload & manage specification documents · manage course data · run checks · view all reports · manage user accounts.

**User can:** access the **Run Check** page only — select filters, run the validation, and view/download generated reports. Users cannot upload or modify any data.

## Admin workflow

1. **📥 Import Courses** — upload the Excel file (.xlsx/.csv). Columns are auto-guessed and can be remapped: Course Name (mandatory), Course URL, Spec URL, **Category ID**, **Level**, **Type** (Award/Certificate/Diploma), Course Number, Entry Requirements, Method of Assessment, Qualification Specification, Course Overview. Courses are inserted/updated by name and can be edited or deleted afterwards.
2. **📑 Specifications** — per course, upload the specification document (**PDF / DOCX / TXT**) or extract it from the stored spec URL. The tool extracts and stores: **Entry Requirements**, **Qualification Specification**, **Method of Assessment**, plus other spec fields (GLH, TQT, progression, units, grading). Extracted sections can be reviewed and edited before saving.

## Run Check (User)

1. **Filters** — Category ID · Level · Type (Award / Certificate / Diploma).
2. Select a course: the **left panel** shows the course details; the **right panel** shows the specification document and its extracted requirements.
3. **Run check** compares only the required sections:
   - Qualification Specification
   - Entry Requirements
   - Method of Assessment (**wording only**)
4. The validation report is displayed in the standard format — a **Level | Type | Number** header, then one boxed block per section with:
   - ✅ Correct · ❌ Incorrect wording · ❌ Missing information · ⚠ Mismatched requirements
   - **Errors identified 01..N** (red) each paired with a **Recommend Action 01..N** (green)
   - a summary of all detected issues.
5. **⬇️ Download Report (Word)** — available on the same page immediately after the check, and for any previously saved report. Admins can also download a combined Word report of all courses with errors from the 📊 Reports tab.

## Notes

- All data (users, courses, specs, reports) persists in `slc_checker.db`; re-importing the Excel file updates existing courses without losing uploaded specs.
- Passwords are stored salted + hashed (SHA-256); the last admin account cannot be deleted.
- Default model is `deepseek/deepseek-v4-pro` via OpenRouter; edit `MODEL` in `app.py` to change it.
