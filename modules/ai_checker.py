import json
import os
import re
from difflib import SequenceMatcher

import requests
from dotenv import load_dotenv

load_dotenv()

OPENROUTER_API_KEY = os.getenv("OPENROUTER_API_KEY")
OPENROUTER_MODEL = os.getenv("OPENROUTER_MODEL", "deepseek/deepseek-chat-v3.1")


def _normalise_text(text: str) -> str:
    text = text or ""
    text = re.sub(r"\s+", " ", text.lower()).strip()
    return text


def course_overview_plagiarism_percentage(course_overview: str, specification_text: str) -> int | None:
    """Estimate how closely the Course Overview wording matches the specification.

    This is not a Turnitin-style plagiarism engine. It is a practical internal similarity
    score to flag copied wording risk for Course Overview only.
    """
    overview = _normalise_text(course_overview)
    spec = _normalise_text(specification_text)

    if not overview or not spec:
        return None

    # Compare against the most relevant specification window rather than the full PDF,
    # otherwise the score becomes artificially tiny for long documents.
    words = overview.split()
    spec_words = spec.split()
    window_size = max(len(words) + 40, 120)

    if len(spec_words) <= window_size:
        ratio = SequenceMatcher(None, overview, spec).ratio()
        return round(ratio * 100)

    best = 0.0
    step = max(40, window_size // 2)
    for start in range(0, len(spec_words), step):
        window = " ".join(spec_words[start : start + window_size])
        ratio = SequenceMatcher(None, overview, window).ratio()
        if ratio > best:
            best = ratio
        if start + window_size >= len(spec_words):
            break

    return round(best * 100)


def build_prompt(course, field_name, live_snippet, specification_text):
    tracker_value_map = {
        "Entry Requirements": course.get("entry_requirements"),
        "Method of Assessment": course.get("method_of_assessment"),
        "Course Overview": course.get("course_overview"),
        "Awarding Body Details": None,
    }
    tracker_value = tracker_value_map.get(field_name) or "Not provided in tracker sheet."

    plagiarism_instruction = ""
    if field_name == "Course Overview":
        plagiarism_instruction = """
For Course Overview only, also estimate whether the website/tracker wording appears too close to the official specification wording. Return plagiarism_percentage as an integer from 0 to 100. This is a wording-similarity percentage, not a formal academic misconduct judgement.
""".strip()
    else:
        plagiarism_instruction = "For this field, set plagiarism_percentage to null. Only Course Overview needs a plagiarism percentage."

    return f"""
You are a course content quality checker for South London College.

Check one field only: {field_name}

Use the official qualification specification as the primary source of truth.
Use the tracker sheet wording as internal expected wording.
Use the live website text as the published content.

{plagiarism_instruction}

Course name:
{course.get('course_name')}

Course URL:
{course.get('course_url')}

Tracker sheet wording for {field_name}:
{tracker_value}

Relevant live website text:
{live_snippet}

Official specification text excerpt/full text:
{specification_text[:20000]}

Return valid JSON only with this exact structure:
{{
  "field_checked": "{field_name}",
  "decision": "Correct | Incorrect | Missing | Needs Review",
  "priority": "High | Medium | Low",
  "current_website_evidence": "brief evidence from live page or state missing",
  "tracker_sheet_evidence": "brief evidence from tracker wording",
  "specification_evidence": "brief evidence from official specification",
  "explanation": "clear explanation of why the decision was made",
  "suggested_corrected_wording": "safe corrected wording for the website",
  "suggested_action": "what the content team should do",
  "wording_similarity_risk": "High | Medium | Low",
  "plagiarism_percentage": null,
  "low_risk_rewritten_wording": "rewritten wording that keeps meaning but avoids copying specification wording too closely"
}}
""".strip()


def call_openrouter(prompt):
    if not OPENROUTER_API_KEY:
        raise ValueError("OPENROUTER_API_KEY is missing. Add it to your .env file.")

    response = requests.post(
        "https://openrouter.ai/api/v1/chat/completions",
        headers={
            "Authorization": f"Bearer {OPENROUTER_API_KEY}",
            "Content-Type": "application/json",
            "HTTP-Referer": "http://localhost:8501",
            "X-Title": "SLC Course Content Checker",
        },
        json={
            "model": OPENROUTER_MODEL,
            "messages": [
                {"role": "system", "content": "Return valid JSON only. Do not include markdown."},
                {"role": "user", "content": prompt},
            ],
            "temperature": 0.1,
        },
        timeout=120,
    )

    if not response.ok:
        raise RuntimeError(f"OpenRouter error {response.status_code}: {response.text}")

    data = response.json()
    content = data["choices"][0]["message"]["content"].strip()

    if content.startswith("```"):
        content = content.strip("`")
        content = content.replace("json", "", 1).strip()

    try:
        parsed = json.loads(content)
        if "plagiarism_percentage" not in parsed:
            parsed["plagiarism_percentage"] = None
        return parsed
    except json.JSONDecodeError:
        return {
            "field_checked": "Unknown",
            "decision": "Needs Review",
            "priority": "Medium",
            "current_website_evidence": "Could not parse model JSON response.",
            "tracker_sheet_evidence": "",
            "specification_evidence": "",
            "explanation": content,
            "suggested_corrected_wording": "",
            "suggested_action": "Review manually.",
            "wording_similarity_risk": "Needs Review",
            "plagiarism_percentage": None,
            "low_risk_rewritten_wording": "",
        }
