# 🎓 SLC Course Content Checker

Internal Streamlit tool for the South London College content team. It verifies that live course pages are accurate by comparing three sources — the live page, the uploaded tracker sheet, and the official qualification specification (PDF or URL) — then exports a Word report of only the courses with errors. It also includes a colour-coded Content Quality Review (proofreading) mode.

Stack: **Python · Streamlit · SQLite · OpenRouter API**. No Supabase, no Next.js, no edge functions — one file, one local database.

## Setup

```bash
pip install -r requirements.txt
streamlit run app.py
```

Get an OpenRouter API key at https://openrouter.ai/keys and paste it into the sidebar. The SQLite database (`slc_checker.db`) is created automatically next to `app.py`.

## Workflow

1. **📥 Import Courses** — upload the Excel tracker (.xlsx/.csv). Column names are auto-guessed and can be remapped. Courses are inserted or updated by name.
2. **📑 Spec Extraction** — extracts text from each course's specification URL (handles both PDFs and web pages) and stores it in the database. Only courses missing spec text are processed.
3. **🔍 Run Checks** — pick the fields to audit in the sidebar (Entry Requirements, Method of Assessment, Course Overview). Check a single course, or run the bulk check. Enable **Testing mode** in the sidebar to limit the run to the first N courses.
4. **📊 Reports & Word Export** — every check is saved to SQLite. The dashboard shows pass/error counts, per-course details, and a one-click download of a Word document containing **only** the courses with errors, with severity, the live-page text, the expected content, and a suggested fix for each issue.
5. **✍️ Content Quality Review** — paste or upload (.txt/.docx) course copy. The tool highlights grammar, articles, spelling, punctuation & commas, capitalisation, proper nouns, sentence structure and consistency issues — each category in its own colour with numbered marks (hover for the fix), plus a corrected copy-ready version and a full issue list.

## Tracker sheet columns

The importer looks for (and lets you remap): Course Name, Course URL, Specification URL, Entry Requirements, Method of Assessment, Course Overview. Only **Course Name** is mandatory.

## Notes

- Spec text and reports persist in `slc_checker.db`; re-importing the tracker updates existing courses without losing extracted specs.
- The bulk check pauses briefly between courses to be gentle to the website and the API.
- Default model is `anthropic/claude-sonnet-4.5`; any OpenRouter model can be selected in the sidebar.
