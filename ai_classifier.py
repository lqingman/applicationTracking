"""
ai_classifier.py
OpenRouter AI integration for job email classification and data extraction.

Provides a single public function: classify_email()
Falls back to fast keyword-based logic automatically if AI is unavailable.
"""

import json
import os
import re
from typing import Optional

import requests

# ── Keyword fallback data ─────────────────────────────────────────────────────

# Patterns that indicate a digest, newsletter, or platform notification — NOT
# an actual job-application status email.  Checked BEFORE the keyword filter.
_DIGEST_PATTERNS = [
    "daily digest",
    "weekly digest",
    "digest email",
    "newsletter",
    "unsubscribe",
    "job alert",
    "jobs you may like",
    "recommended jobs",
    "new jobs for you",
    "workday inbox",           # Workday platform digest
    "notification digest",
    "activity digest",
    "your daily summary",
    "your weekly summary",
    "digest for",
]

_JOB_KEYWORDS = [
    "application", "applied", "applicant",
    "interview", "interviewer", "schedule a call",
    "job offer", "offer letter", "offer of employment",
    "hiring", "recruiter", "recruitment",
    "position", "job posting", "opening",
    "resume", "cv", "cover letter",
    "rejection", "unfortunately", "not moving forward",
    "thank you for applying", "we received your application",
    "background check", "onboarding",
    "salary", "compensation", "start date",
    "assessment", "coding challenge", "take-home",
    "technical screen", "phone screen",
    "final round", "next steps",
]

_STATUS_RULES = [
    ("Offer",     ["offer letter", "offer of employment", "pleased to offer",
                   "job offer", "onboarding", "start date", "background check",
                   "congratulations"]),
    ("Interview", ["interview", "schedule a call", "phone screen",
                   "technical screen", "coding challenge", "take-home",
                   "assessment", "next round", "final round", "next steps",
                   "meet with"]),
    ("Rejected",  ["unfortunately", "not moving forward", "will not be moving",
                   "other candidates", "not selected", "no longer considering",
                   "position has been filled", "regret to inform",
                   "decided to move", "not a fit", "unable to offer"]),
    ("Applied",   ["we received your application", "thank you for applying",
                   "application received", "application submitted",
                   "application has been", "confirm your application",
                   "application for", "applied for"]),
]

_GENERIC_DOMAINS = {
    # ATS / recruiting platforms — NOT real company names
    "greenhouse.io", "lever.co",
    "workday.com", "myworkday.com", "wd1.myworkday.com", "wd3.myworkday.com",
    "taleo.net", "icims.com", "jobvite.com", "smartrecruiters.com",
    "successfactors.com", "brassring.com", "recruitingsite.com",
    "ultipro.com", "bamboohr.com", "paylocity.com", "adp.com",
    "oracle.com", "sap.com",
    # Generic mail providers
    "gmail.com", "outlook.com", "yahoo.com", "hotmail.com",
    "noreply.com", "no-reply.com", "mail.com",
    # Job boards
    "notifications.linkedin.com", "linkedin.com",
    "indeed.com", "glassdoor.com", "ziprecruiter.com",
    "monster.com", "careerbuilder.com",
}

# ── Public API ────────────────────────────────────────────────────────────────


def is_job_email_fast(subject: str, body_preview: str) -> bool:
    """
    Fast keyword pre-filter — no API call.
    Returns True if the email shows any job-related signal AND is not a
    digest / newsletter / platform notification.
    Used as Stage 1 to avoid wasting AI calls on obvious non-matches.
    """
    text = (subject + " " + body_preview).lower()
    # Reject digests / newsletters first — they often contain job keywords
    # incidentally but are never actual application-status emails.
    if any(pat in text for pat in _DIGEST_PATTERNS):
        return False
    return any(kw in text for kw in _JOB_KEYWORDS)


def classify_email(
    subject: str,
    body_preview: str,
    sender_name: str,
    sender_email: str,
    api_key: Optional[str] = None,
    model: str = "openai/gpt-4o-mini",
    ai_enabled: bool = True,
) -> dict:
    """
    Classify a job-related email and extract structured data.

    Returns a dict:
    {
        "is_job_email": bool,
        "company":      str,    # clean hiring company name
        "job_title":    str,    # role applied for
        "status":       str,    # Applied | Interview | Offer | Rejected | Unknown
        "confidence":   float,  # 0.0–1.0
        "used_ai":      bool,   # True if OpenRouter was called successfully
    }

    Falls back to keyword-based classification automatically on any error.
    """
    if ai_enabled and api_key:
        result = _ai_classify(subject, body_preview, sender_name, sender_email,
                               api_key, model)
        if result is not None:
            return result

    # Keyword fallback
    return _keyword_classify(subject, body_preview, sender_name, sender_email)


# ── OpenRouter AI ─────────────────────────────────────────────────────────────

_OPENROUTER_URL = "https://openrouter.ai/api/v1/chat/completions"

_SYSTEM_PROMPT = (
    "You are an expert job application email classifier. "
    "Analyze the given email and extract structured information. "
    "Always respond with valid JSON only — no markdown, no explanation."
)

_USER_PROMPT_TPL = """\
Analyze this email and classify it:

Subject: {subject}
From: {sender_name} <{sender_email}>
Body:
{body}

Return a JSON object with EXACTLY these fields:
{{
  "is_job_email": true or false,
  "company": "Hiring company name, or Unknown",
  "job_title": "Role the candidate applied for, or Unknown",
  "status": "Applied" | "Interview" | "Offer" | "Rejected" | "Unknown",
  "confidence": 0.0 to 1.0
}}

Classification rules:
- is_job_email: true ONLY if this email is a direct status update for a
  specific job application the candidate submitted.
  Set is_job_email to FALSE for:
    * Digest or summary emails (e.g. "Your Daily Digest", "Workday Inbox",
      "Weekly Digest", "Notification Digest")
    * Newsletter or marketing emails
    * Job-board recommendation emails ("Jobs you may like", "New jobs for you")
    * Platform notification emails that are not tied to a specific application
    * Any email where the subject or body indicates it is a bulk/aggregated
      summary rather than a direct response to a specific application
- company: the ACTUAL hiring employer's name.
  IMPORTANT: Workday, Greenhouse, Lever, iCIMS, Taleo, Jobvite, SmartRecruiters
  and similar are ATS PLATFORMS — never use them as the company.
  Look for the real employer in the email subject (e.g. "TD Careers"),
  body text, or sender display name instead.
  If truly unknown, use "Unknown".
- job_title: the specific position the candidate applied for.
  Look in the subject and body. If unclear, use "Unknown".
- status — choose ONLY based on what this specific email is about:
    Applied   = application received/confirmed ("thank you for applying",
                "we received your application", "your application has been submitted")
    Interview = interview explicitly scheduled or invited (phone screen,
                technical screen, coding challenge, take-home assignment,
                "we'd like to schedule", "next round")
                NOTE: a confirmation email that merely MENTIONS the word
                "interview" as a future possibility is still Applied.
    Offer     = job offer extended, background check initiated, onboarding,
                start date discussed
    Rejected  = rejection, "not moving forward", "other candidates selected",
                "regret to inform"
    Unknown   = cannot determine
- confidence: 0.0 = unsure, 1.0 = very certain
"""


def _ai_classify(subject, body_preview, sender_name, sender_email, api_key, model):
    """
    Call OpenRouter. Returns parsed dict on success, None on any error.
    Retries once on network/timeout errors before giving up.
    """
    prompt = _USER_PROMPT_TPL.format(
        subject=subject,
        sender_name=sender_name,
        sender_email=sender_email,
        body=body_preview,
    )
    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": _SYSTEM_PROMPT},
            {"role": "user",   "content": prompt},
        ],
        "temperature": 0.1,
        "max_tokens": 200,
        "response_format": {"type": "json_object"},
    }
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
        "HTTP-Referer": "https://github.com/local/applicationTracking",
        "X-Title": "Job Application Tracker",
    }
    # (connect_timeout, read_timeout) — read timeout per chunk, not total
    TIMEOUT = (8, 20)

    for attempt in range(2):   # try once, retry once on transient failure
        try:
            resp = requests.post(
                _OPENROUTER_URL,
                headers=headers,
                json=payload,
                timeout=TIMEOUT,
                stream=False,   # read full response at once
            )
            resp.raise_for_status()
            content = resp.json()["choices"][0]["message"]["content"]
            data = json.loads(content)

            valid_statuses = {"Applied", "Interview", "Offer", "Rejected", "Unknown"}
            status = data.get("status", "Unknown")
            if status not in valid_statuses:
                status = "Unknown"

            return {
                "is_job_email": bool(data.get("is_job_email", False)),
                "company":      str(data.get("company", "Unknown") or "Unknown")[:200],
                "job_title":    str(data.get("job_title", "Unknown") or "Unknown")[:200],
                "status":       status,
                "confidence":   float(data.get("confidence", 0.5)),
                "used_ai":      True,
            }

        except (requests.exceptions.Timeout,
                requests.exceptions.ConnectionError) as e:
            if attempt == 0:
                continue   # retry once
            return None    # give up — caller uses keyword fallback
        except Exception:
            return None    # non-retryable error


# ── Keyword-based fallback ────────────────────────────────────────────────────

def _keyword_classify(subject, body_preview, sender_name, sender_email):
    text = (subject + " " + body_preview).lower()
    return {
        "is_job_email": any(kw in text for kw in _JOB_KEYWORDS),
        "company":      _kw_company(sender_name, sender_email),
        "job_title":    _kw_job_title(subject),
        "status":       _kw_status(subject, body_preview),
        "confidence":   0.5,
        "used_ai":      False,
    }


def _kw_company(sender_name: str, sender_email: str) -> str:
    if "@" in sender_email:
        domain = sender_email.split("@")[-1].lower()
        if domain not in _GENERIC_DOMAINS:
            root = domain.split(".")[0]
            return root.capitalize()
    if sender_name:
        parts = sender_name.strip().split()[:2]
        return " ".join(parts)
    return "Unknown"


def _kw_job_title(subject: str) -> str:
    patterns = [
        r"(?:application for|applying for|position[:\s]+|role[:\s]+)(.+)",
        r"(?:your application|application received)[^\-–—]*[-–—]\s*(.+)",
    ]
    for p in patterns:
        m = re.search(p, subject, re.IGNORECASE)
        if m:
            title = m.group(1).strip(" .,;:")
            if 3 < len(title) < 120:
                return title[:100]
    return subject[:100]


def _kw_status(subject: str, body_preview: str) -> str:
    text = (subject + " " + body_preview).lower()
    for status, keywords in _STATUS_RULES:
        if any(kw in text for kw in keywords):
            return status
    return "Applied"
