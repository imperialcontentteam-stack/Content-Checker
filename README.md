# SLC Course Content Checker - Streamlit Version

This is a simpler Python/Streamlit version of the SLC Course Content Checker.
It avoids Supabase Edge Functions and keeps the whole workflow inside one app.

## Main features

- User login and role-based access
- Admin user account creation
- Tracker sheet upload and import
- Specification PDF text extraction
- Extraction batch size up to 300, default 300
- Dashboard course checks using OpenRouter AI
- Course Overview plagiarism/similarity percentage only
- Download individual check answer as a Word document
- Download all reports as Excel or Word document

## Install

```powershell
pip install -r requirements.txt
```

## Setup environment

Copy `.env.example` and rename it to `.env`.

```env
OPENROUTER_API_KEY=your_openrouter_key_here
OPENROUTER_MODEL=deepseek/deepseek-chat-v3.1
DATABASE_PATH=data/slc_checker.db
ADMIN_USERNAME=admin
ADMIN_PASSWORD=change_this_password
```

## Run

```powershell
streamlit run app.py
```

Open:

```text
http://localhost:8501
```

## Login

The first admin account is created automatically from `.env`.

If you do not set custom admin details, the temporary default is:

```text
Username: admin
Password: admin123
```

After logging in, go to **Admin > User accounts** to create staff users.

Roles:

- `admin`: can import tracker sheet, extract specifications and create users.
- `user`: can run checks and view reports.

## Correct workflow

1. Log in as admin.
2. Open **Admin**.
3. Upload the Course Information Tracker Sheet.
4. Click **Import / Update Courses**.
5. Use **Extract Pending Specifications**. Batch size is set to 300 by default.
6. Open **Dashboard**.
7. Select a ready course.
8. Select fields to check.
9. Click **Run Check**.
10. Download the answer as a Word document if needed.
11. Open **Reports** to download Excel or Word reports.

## Course Overview plagiarism percentage

The app only calculates a plagiarism/similarity percentage for **Course Overview**.
Other fields still show decision, priority and wording similarity risk, but they do not show a plagiarism percentage.

This percentage is an internal wording-similarity estimate. It is not a formal Turnitin-style plagiarism result.
