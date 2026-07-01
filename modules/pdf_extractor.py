import re
from datetime import datetime
import requests
import fitz  # PyMuPDF
from modules.database import get_conn


def convert_google_drive_url(url):
    if not url:
        return url

    match1 = re.search(r"/file/d/([^/]+)", url)
    if match1:
        return f"https://drive.google.com/uc?export=download&id={match1.group(1)}"

    match2 = re.search(r"[?&]id=([^&]+)", url)
    if "drive.google.com" in url and match2:
        return f"https://drive.google.com/uc?export=download&id={match2.group(1)}"

    return url


def download_pdf(url, timeout=60):
    direct_url = convert_google_drive_url(url)
    response = requests.get(
        direct_url,
        headers={"User-Agent": "Mozilla/5.0"},
        timeout=timeout,
        allow_redirects=True,
    )
    response.raise_for_status()

    content = response.content
    content_type = response.headers.get("content-type", "")

    if not content.startswith(b"%PDF"):
        raise ValueError(f"Downloaded file is not a readable PDF. Content-Type: {content_type}")

    return content


def extract_text_from_pdf_bytes(pdf_bytes):
    text_parts = []
    with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
        for page in doc:
            text_parts.append(page.get_text("text"))
    return "\n".join(text_parts).strip()


def get_pending_specs(limit=10):
    with get_conn() as conn:
        rows = conn.execute(
            """
            select id, source_url
            from specifications
            where source_url is not null
              and (specification_text is null or specification_text = '')
              and coalesce(extraction_status, 'pending') != 'extracted'
            order by created_at asc
            limit ?
            """,
            (limit,),
        ).fetchall()
        return [dict(row) for row in rows]


def extract_one_spec(spec):
    spec_id = spec["id"]
    source_url = spec["source_url"]
    try:
        pdf_bytes = download_pdf(source_url)
        text = extract_text_from_pdf_bytes(pdf_bytes)

        if len(text) < 100:
            raise ValueError("PDF text extraction returned very little or no text.")

        with get_conn() as conn:
            conn.execute(
                """
                update specifications
                set specification_text = ?, extraction_status = 'extracted',
                    extraction_error = null, extracted_at = ?
                where id = ?
                """,
                (text, datetime.utcnow().isoformat(), spec_id),
            )

        return {"ok": True, "id": spec_id, "url": source_url, "characters": len(text)}
    except Exception as exc:
        with get_conn() as conn:
            conn.execute(
                """
                update specifications
                set extraction_status = 'failed', extraction_error = ?, extracted_at = ?
                where id = ?
                """,
                (str(exc), datetime.utcnow().isoformat(), spec_id),
            )
        return {"ok": False, "id": spec_id, "url": source_url, "error": str(exc)}


def extract_pending_specs(limit=10):
    specs = get_pending_specs(limit)
    results = []
    for spec in specs:
        results.append(extract_one_spec(spec))
    return results


def reset_failed_specs():
    with get_conn() as conn:
        cur = conn.execute(
            """
            update specifications
            set extraction_status = 'pending', extraction_error = null
            where extraction_status = 'failed'
            """
        )
        return cur.rowcount
