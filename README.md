# 🎓 SLC Course Content Checker — v4

Internal Streamlit tool for the South London College content team, now with
authentication, tracker-driven filters, specification-document extraction and
downloadable validation reports.

Stack: **Python · Streamlit · SQLite · OpenRouter API · reportlab**. One file, one local database.

## Setup

```bash
pip install -r requirements.txt
streamlit run app.py
```

Optionally get an OpenRouter API key at https://openrouter.ai/keys and add it to
`.streamlit/secrets.toml` as `OPENROUTER_API_KEY = "..."` (or set it as an
environment variable). Without a key the AI checks are disabled and validation
falls back to a built-in wording comparison. The SQLite database
(`slc_checker.db`) is created automatically next to `app.py`.

## Accounts & permissions

Default logins on first run (change them on the 👥 Users page!):

| Account | Password | Role |
|---|---|---|
| `admin` | `admin123` | Administrator |
| `user`  | `user123`  | User |

**Admin** — import the course Excel tracker, upload & manage specification
documents, manage course records, manage user accounts, plus everything a user
can do.

**User** — the 🔍 Run Check page only: run validations/checks, view generated
reports and download them. No permission to upload or modify any data.

Passwords are stored as salted PBKDF2-SHA256 hashes (200k iterations).

## Workflow

1. **📥 Import Courses (admin)** — upload the Excel tracker (.xlsx/.csv). Column
   names are auto-guessed and can be remapped. The **Category ID (Number)**,
   **Level** and **Type** (Award / Certificate / Diploma) are read from the sheet —
   Level/Type are derived from the course name when not present as columns —
   and these values populate the Run Check filters automatically.
2. **📑 Spec Documents (admin)** — upload the qualification specification
   (PDF/Word) or extract it from the spec URL. The **Entry Requirements**,
   **Qualification Specification** and **Method of Assessment** sections are
   extracted (heuristically, refined with AI when a key is configured), are
   editable, and are stored per course.
3. **🔍 Run Check** — pick **Category ID / Level / Type**, then the course. The
   left panel shows the course details from the tracker; the right panel shows
   the specification document and its extracted requirements. Three tabs:
   * **✅ Content Check** — validates *Qualification Specification*, *Entry
     Requirements* and *Method of Assessment (wording only)* against the
     specification. The report uses the agreed layout: red section boxes,
     numbered **Errors identified** (red) with **Recommend Action** (green), a
     prominent green **No Errors** badge when the course passes, and a summary
     table of all detected issues. **Download Report** (PDF and Word) sits on
     the same page.
   * **✍️ Grammar Check** — the existing colour-coded Content Quality Review
     (grammar, articles, spelling, punctuation, capitalisation, proper nouns,
     sentence structure, consistency), unchanged.
   * **🧪 Other Checks** — the existing 3-way live-page audit (live page vs
     tracker vs specification) with single and bulk modes and plagiarism-safe
     rewrite suggestions, unchanged.
4. **📊 Reports** — validation reports (view + PDF/Word download per course)
   and the existing live-page audit dashboard with the errors-only Word export.
5. **👥 Users (admin)** — create accounts, reset passwords, delete accounts.

## Tracker sheet columns

The importer looks for (and lets you remap): Number/Category ID, Course Name,
Course URL, Specification URL, Entry Requirements, Method of Assessment,
Course Overview, and optional Level / Type columns. Only **Course Name** is
mandatory.

## Notes

- Run Check filters are generated **dynamically** from the imported tracker —
  re-importing the sheet refreshes them.
- Spec text, extracted sections and reports persist in `slc_checker.db`;
  re-importing the tracker updates courses without losing extracted specs.
- Validation reports are saved per run (with who ran them) and the latest per
  course is available on 📊 Reports.
- Default model is set in `MODEL` at the top of `app.py`; any OpenRouter model
  string works.
