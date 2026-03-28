"""
email_scraper.py
Reads emails directly from your locally installed Outlook desktop app
using Windows COM automation (pywin32).

✅ No Azure portal, no app registration, no tokens.
   Works with any account already signed into Outlook — including school SSO.
"""

import sqlite3
import re
import os
from datetime import datetime, timedelta

# ── COM / Outlook ─────────────────────────────────────────────────────────────

try:
    import win32com.client
except ImportError:
    print("✗  pywin32 is not installed.")
    print("   Run:  pip install pywin32")
    raise SystemExit(1)

# ── Configuration ─────────────────────────────────────────────────────────────

try:
    import config
    OUTLOOK_FOLDER = getattr(config, "OUTLOOK_FOLDER", "Inbox")
    SCAN_JUNK      = getattr(config, "SCAN_JUNK", True)
    SCAN_DAYS      = getattr(config, "SCAN_DAYS", 730)
except ImportError:
    print("✗  config.py not found.")
    raise SystemExit(1)

DB_PATH = os.path.join(os.path.dirname(__file__), "applications.db")

# ── Keywords ──────────────────────────────────────────────────────────────────

JOB_KEYWORDS = [
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

STATUS_RULES = [
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

# ── Database ───────────────────────────────────────────────────────────────────

def init_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS applications (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            email_id    TEXT    UNIQUE,
            date        TEXT,
            company     TEXT,
            job_title   TEXT,
            status      TEXT,
            sender      TEXT,
            subject     TEXT,
            created_at  TEXT DEFAULT (datetime('now'))
        )
    """)
    conn.commit()
    conn.close()

def save_application(record: dict):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("""
        INSERT OR IGNORE INTO applications
            (email_id, date, company, job_title, status, sender, subject)
        VALUES
            (:email_id, :date, :company, :job_title, :status, :sender, :subject)
    """, record)
    conn.commit()
    conn.close()

# ── Classification helpers ────────────────────────────────────────────────────

def is_job_email(subject: str, body_preview: str) -> bool:
    text = (subject + " " + body_preview).lower()
    return any(kw in text for kw in JOB_KEYWORDS)

def classify_status(subject: str, body_preview: str) -> str:
    text = (subject + " " + body_preview).lower()
    for status, keywords in STATUS_RULES:
        if any(kw in text for kw in keywords):
            return status
    return "Applied"

def extract_company(sender_name: str, sender_email: str) -> str:
    generic = {"greenhouse.io", "lever.co", "workday.com", "taleo.net",
               "icims.com", "jobvite.com", "smartrecruiters.com",
               "successfactors.com", "brassring.com", "gmail.com",
               "outlook.com", "yahoo.com", "hotmail.com", "noreply.com",
               "mail.com", "notifications.linkedin.com", "linkedin.com",
               "indeed.com", "glassdoor.com", "ziprecruiter.com"}
    if "@" in sender_email:
        domain = sender_email.split("@")[-1].lower()
        if domain not in generic:
            root = domain.split(".")[0]
            return root.capitalize()
    if sender_name:
        parts = sender_name.strip().split()[:2]
        return " ".join(parts)
    return "Unknown"

def extract_job_title(subject: str) -> str:
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

# ── Outlook COM helpers ───────────────────────────────────────────────────────

def _get_outlook_folder(namespace, folder_name: str):
    """Find a top-level mail folder by name across all accounts."""
    # Try each account's inbox (handles multiple email accounts in Outlook)
    for store in namespace.Stores:
        try:
            root = store.GetRootFolder()
            for folder in root.Folders:
                if folder.Name.lower() == folder_name.lower():
                    return folder
                # Also check one level down (e.g. account root → Inbox)
                try:
                    sub = folder.Folders[folder_name]
                    return sub
                except Exception:
                    pass
        except Exception:
            continue

    # Fallback: use the default Inbox
    try:
        return namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
    except Exception:
        return None

def _get_junk_folder(namespace):
    """Return the Junk Email folder (olFolderJunk = 23)."""
    try:
        return namespace.GetDefaultFolder(23)
    except Exception:
        return None

def _iter_folder(folder, cutoff: datetime):
    """
    Yield Outlook MailItem objects from a folder that are newer than cutoff.
    Uses Outlook's built-in filter for performance.
    """
    if folder is None:
        return

    # Outlook DASL filter — much faster than iterating all items
    cutoff_str = cutoff.strftime("%m/%d/%Y %I:%M %p")
    restriction = f"[ReceivedTime] >= '{cutoff_str}'"

    try:
        items = folder.Items
        items.Sort("[ReceivedTime]", True)   # newest first
        filtered = items.Restrict(restriction)
        count = filtered.Count
        for i in range(1, count + 1):
            try:
                item = filtered[i]
                # Only process mail items (class 43), skip calendar/tasks etc.
                if hasattr(item, "Class") and item.Class == 43:
                    yield item
            except Exception:
                continue
    except Exception as e:
        print(f"  Warning: could not read folder '{folder.Name}': {e}")

# ── Main Scanner ───────────────────────────────────────────────────────────────

def scan_inbox() -> dict:
    print("Connecting to Outlook…")
    try:
        outlook  = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        namespace.Logon()
        print("  ✓ Connected to Outlook")
    except Exception as e:
        print(f"  ✗ Could not connect to Outlook: {e}")
        print("    Make sure Microsoft Outlook is installed and open.")
        return {"error": str(e), "new": 0, "total": 0}

    cutoff = datetime.now() - timedelta(days=SCAN_DAYS)
    print(f"  Scanning emails since {cutoff.date()}…\n")

    init_db()
    new_count   = 0
    match_count = 0
    total_seen  = 0

    # Determine which folders to scan
    folders_to_scan = []

    inbox = _get_outlook_folder(namespace, OUTLOOK_FOLDER)
    if inbox:
        folders_to_scan.append(("Inbox", inbox))
    else:
        print("  ✗ Could not find Inbox folder")

    if SCAN_JUNK:
        junk = _get_junk_folder(namespace)
        if junk:
            folders_to_scan.append(("Junk", junk))

    for folder_label, folder in folders_to_scan:
        print(f"  📁 Scanning {folder_label}…")
        folder_count = 0

        for item in _iter_folder(folder, cutoff):
            total_seen += 1
            folder_count += 1

            try:
                subject  = str(item.Subject or "")
                sender   = str(item.SenderName or "")
                sender_email = ""
                try:
                    sender_email = str(item.SenderEmailAddress or "")
                    # Exchange internal addresses look like /O=.../CN=... — skip
                    if sender_email.startswith("/O="):
                        sender_email = ""
                except Exception:
                    pass

                received = item.ReceivedTime  # datetime object from COM
                # COM datetime → Python datetime
                try:
                    date_str = received.strftime("%Y-%m-%d")
                except Exception:
                    date_str = datetime.now().strftime("%Y-%m-%d")

                # Body preview (first 1000 chars)
                try:
                    body_preview = str(item.Body or "")[:1000]
                except Exception:
                    body_preview = ""

                if not is_job_email(subject, body_preview):
                    continue

                match_count += 1

                # Unique email ID using EntryID (stable across restarts)
                try:
                    email_id = str(item.EntryID)[:200]
                except Exception:
                    email_id = f"{sender_email}_{date_str}_{hash(subject)}"

                company  = extract_company(sender, sender_email)
                title    = extract_job_title(subject)
                status_v = classify_status(subject, body_preview)
                sender_full = f"{sender} <{sender_email}>" if sender_email else sender

                # Skip if already in DB
                conn = sqlite3.connect(DB_PATH)
                c = conn.cursor()
                c.execute("SELECT 1 FROM applications WHERE email_id = ?", (email_id,))
                exists = c.fetchone()
                conn.close()

                if not exists:
                    save_application({
                        "email_id":  email_id,
                        "date":      date_str,
                        "company":   company,
                        "job_title": title,
                        "status":    status_v,
                        "sender":    sender_full[:255],
                        "subject":   subject[:255],
                    })
                    new_count += 1
                    print(f"    [{status_v:9s}] {company:<20s} — {subject[:50]}")

                if folder_count % 100 == 0:
                    print(f"    … {folder_count} emails checked in {folder_label}")

            except Exception as e:
                print(f"    Warning: skipping one email — {e}")
                continue

        print(f"  ✓ {folder_label}: checked {folder_count} emails\n")

    conn = sqlite3.connect(DB_PATH)
    total = conn.execute("SELECT COUNT(*) FROM applications").fetchone()[0]
    conn.close()

    print(f"✓ Done. Checked {total_seen} emails total, matched {match_count} job emails,")
    print(f"  added {new_count} new ({total} total in DB).")
    return {"new": new_count, "total": total}


if __name__ == "__main__":
    result = scan_inbox()
    if "error" in result:
        print(f"\n✗ Error: {result['error']}")
