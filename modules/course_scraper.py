import re
import requests
from bs4 import BeautifulSoup


def clean_text(text):
    text = re.sub(r"\s+", " ", text or "").strip()
    return text


def fetch_course_page_text(url, timeout=45):
    if not url:
        raise ValueError("Course URL is missing.")

    response = requests.get(
        url,
        headers={"User-Agent": "Mozilla/5.0"},
        timeout=timeout,
        allow_redirects=True,
    )
    response.raise_for_status()

    soup = BeautifulSoup(response.text, "html.parser")

    for tag in soup(["script", "style", "noscript", "svg", "iframe"]):
        tag.decompose()

    text = soup.get_text(" ")
    return clean_text(text)


def find_relevant_snippet(page_text, field_name, max_chars=3500):
    text = page_text or ""
    lowered = text.lower()

    keywords_by_field = {
        "Entry Requirements": ["entry requirements", "requirements", "entry requirement"],
        "Method of Assessment": ["method of assessment", "assessment", "assessed", "assignment"],
        "Course Overview": ["course overview", "overview", "course description", "about this course"],
        "Awarding Body Details": ["awarding body", "awarded by", "qualification"],
    }

    keywords = keywords_by_field.get(field_name, [field_name.lower()])

    for keyword in keywords:
        idx = lowered.find(keyword)
        if idx != -1:
            start = max(0, idx - 500)
            end = min(len(text), idx + max_chars)
            return text[start:end]

    return text[:max_chars]
