"""
email_scraper.py
Reads emails from your locally installed Outlook desktop app via COM automation.

Pipeline:
  1. Stage 1 — fast keyword pre-filter (no API cost)
  2. Stage 2 — OpenRouter AI classification (company, job title, status)
  3. Stage 3 — fuzzy-match to existing application, or create new
              → status only advances (Offer > Interview > Applied; Rejected wins)

Database schema (two tables):
  applications  — one row per unique job application
  email_events  — one row per email (status timeline for each application)
"""

import os
import re
import sqlite3
from datetime import datetime, timedelta
from difflib import SequenceMatcher

# ── dotenv ────────────────────────────────────────────────────────────────────
try:
    from dotenv import load_dotenv
    load_dotenv(os.path.join(os.path.dirname(__file__), ".env"))
except ImportError:
    pass  # python-dotenv not installed; key must be in environment already

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
    OUTLOOK_FOLDER     = getattr(config, "OUTLOOK_FOLDER", "Inbox")
    SCAN_JUNK          = getattr(config, "SCAN_JUNK", True)
    SCAN_DAYS          = getattr(config, "SCAN_DAYS", 730)
    OPENROUTER_MODEL   = getattr(config, "OPENROUTER_MODEL", "openai/gpt-4o-mini")
    AI_ENABLED         = getattr(config, "AI_ENABLED", True)
    AI_MATCH_THRESHOLD = getattr(config, "AI_MATCH_THRESHOLD", 0.75)
except ImportError:
    print("✗  config.py not found.")
    raise SystemExit(1)

# Load API key from environment (.env file)
OPENROUTER_API_KEY = os.environ.get("OPENROUTER_API_KEY", "")

from ai_classifier import classify_email, is_job_email_fast

DB_PATH = os.path.join(os.path.dirname(__file__), "applications.db")

# ── Status priority ───────────────────────────────────────────────────────────
# Rejected always wins (terminal state). Otherwise status can only advance.
STATUS_PRIORITY = {
    "Unknown":   0,
    "Applied":   1,
    "Interview": 2,
    "Offer":     3,
    "Rejected":  99,
}


# ── Database ──────────────────────────────────────────────────────────────────

def init_db():
    """
    Create or migrate the database.

    New schema (v2):
      applications  — one row per unique job application
      email_events  — one row per email / status event

    If the old single-table schema is detected (email_id column in applications),
    it is automatically migrated to the new schema.
    """
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()

    # ── Detect old schema ─────────────────────────────────────────────────────
    c.execute("PRAGMA table_info(applications)")
    cols = {row["name"] for row in c.fetchall()}

    if "email_id" in cols:
        # Old schema detected → migrate
        _migrate_old_schema(conn)
    else:
        # Fresh install — create new tables
        _create_tables(conn)

    conn.close()


def _create_tables(conn):
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS applications (
            id           INTEGER PRIMARY KEY AUTOINCREMENT,
            company      TEXT,
            job_title    TEXT,
            status       TEXT,
            applied_date TEXT,
            last_update  TEXT,
            sender       TEXT,
            subject      TEXT,
            ai_used      INTEGER DEFAULT 0,
            created_at   TEXT DEFAULT (datetime('now'))
        );

        CREATE TABLE IF NOT EXISTS email_events (
            id             INTEGER PRIMARY KEY AUTOINCREMENT,
            application_id INTEGER REFERENCES applications(id),
            email_id       TEXT UNIQUE,
            date           TEXT,
            status         TEXT,
            subject        TEXT,
            sender         TEXT,
            created_at     TEXT DEFAULT (datetime('now'))
        );
    """)
    conn.commit()


def _migrate_old_schema(conn):
    """Migrate old single-table schema to new two-table schema."""
    print("  ℹ️  Migrating database to new schema (v2)…")
    c = conn.cursor()

    # Rename old table
    c.execute("ALTER TABLE applications RENAME TO applications_v1")

    # Create new tables
    _create_tables(conn)

    # Migrate: each old row becomes one application + one email_event
    rows = c.execute("""
        SELECT email_id, date, company, job_title, status, sender, subject, created_at
        FROM applications_v1
    """).fetchall()

    migrated = 0
    for row in rows:
        try:
            c.execute("""
                INSERT INTO applications
                    (company, job_title, status, applied_date, last_update,
                     sender, subject, ai_used, created_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, 0, ?)
            """, (
                row["company"], row["job_title"], row["status"],
                row["date"], row["date"],
                row["sender"], row["subject"], row["created_at"],
            ))
            app_id = c.lastrowid
            c.execute("""
                INSERT OR IGNORE INTO email_events
                    (application_id, email_id, date, status, subject, sender)
                VALUES (?, ?, ?, ?, ?, ?)
            """, (
                app_id, row["email_id"], row["date"],
                row["status"], row["subject"], row["sender"],
            ))
            migrated += 1
        except Exception as e:
            print(f"    Warning: could not migrate row — {e}")

    conn.commit()
    print(f"  ✓ Migrated {migrated} records to new schema")


# ── Application grouping ──────────────────────────────────────────────────────

def _normalize(text: str) -> str:
    """Lowercase, strip punctuation for fuzzy comparison."""
    return re.sub(r"[^\w\s]", "", text.lower()).strip()


def _similarity(a: str, b: str) -> float:
    return SequenceMatcher(None, _normalize(a), _normalize(b)).ratio()


def _compute_status(conn, application_id: int) -> str:
    """
    Derive the application's current status from all its email_events.
    Rejected always wins. Otherwise return highest-priority status.
    """
    rows = conn.execute(
        "SELECT status FROM email_events WHERE application_id = ?",
        (application_id,)
    ).fetchall()
    if not rows:
        return "Applied"
    statuses = [r["status"] for r in rows]
    if "Rejected" in statuses:
        return "Rejected"
    return max(statuses, key=lambda s: STATUS_PRIORITY.get(s, 0))


def find_or_create_application(conn, company: str, job_title: str,
                                date_str: str, sender: str,
                                subject: str, ai_used: bool) -> tuple:
    """
    Fuzzy-match company + job_title against existing applications.
    Returns (application_id: int, is_new: bool).
    """
    threshold = AI_MATCH_THRESHOLD

    # Skip fuzzy matching if company or job_title is too generic
    skip_match = company.lower() in ("unknown", "") or job_title.lower() in ("unknown", "")

    if not skip_match:
        rows = conn.execute(
            "SELECT id, company, job_title FROM applications"
        ).fetchall()

        best_id    = None
        best_score = 0.0
        for row in rows:
            co_sim    = _similarity(company,   row["company"])
            ti_sim    = _similarity(job_title, row["job_title"])
            # Both must independently meet threshold
            if co_sim >= threshold and ti_sim >= threshold:
                score = (co_sim + ti_sim) / 2
                if score > best_score:
                    best_score = score
                    best_id    = row["id"]

        if best_id is not None:
            # Update last_update timestamp
            conn.execute(
                "UPDATE applications SET last_update = ? WHERE id = ?",
                (date_str, best_id)
            )
            return best_id, False

    # Create a new application record
    conn.execute("""
        INSERT INTO applications
            (company, job_title, status, applied_date, last_update,
             sender, subject, ai_used)
        VALUES (?, ?, 'Applied', ?, ?, ?, ?, ?)
    """, (company, job_title, date_str, date_str,
          sender, subject, int(ai_used)))
    conn.commit()
    app_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
    return app_id, True


def add_email_event(conn, application_id: int, email_id: str, date_str: str,
                    status: str, subject: str, sender: str) -> bool:
    """
    Insert a new email_event row.
    Returns True if inserted, False if already exists (duplicate).
    """
    try:
        conn.execute("""
            INSERT INTO email_events
                (application_id, email_id, date, status, subject, sender)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (application_id, email_id, date_str, status,
              subject[:255], sender[:255]))
        conn.commit()
        return True
    except sqlite3.IntegrityError:
        return False   # duplicate email_id


def refresh_application_status(conn, application_id: int):
    """Recompute and write the application's current status from all its events."""
    new_status = _compute_status(conn, application_id)
    conn.execute(
        "UPDATE applications SET status = ? WHERE id = ?",
        (new_status, application_id)
    )
    conn.commit()


# ── Outlook COM helpers ───────────────────────────────────────────────────────

def _get_outlook_folder(namespace, folder_name: str):
    for store in namespace.Stores:
        try:
            root = store.GetRootFolder()
            for folder in root.Folders:
                if folder.Name.lower() == folder_name.lower():
                    return folder
                try:
                    sub = folder.Folders[folder_name]
                    return sub
                except Exception:
                    pass
        except Exception:
            continue
    try:
        return namespace.GetDefaultFolder(6)  # olFolderInbox
    except Exception:
        return None


def _get_junk_folder(namespace):
    try:
        return namespace.GetDefaultFolder(23)  # olFolderJunk
    except Exception:
        return None


def _iter_folder(folder, cutoff: datetime):
    if folder is None:
        return
    cutoff_str  = cutoff.strftime("%m/%d/%Y %I:%M %p")
    restriction = f"[ReceivedTime] >= '{cutoff_str}'"
    try:
        items    = folder.Items
        items.Sort("[ReceivedTime]", False)  # oldest first → correct status progression
        filtered = items.Restrict(restriction)
        count    = filtered.Count
        for i in range(1, count + 1):
            try:
                item = filtered[i]
                if hasattr(item, "Class") and item.Class == 43:
                    yield item
            except Exception:
                continue
    except Exception as e:
        print(f"  Warning: could not read folder '{folder.Name}': {e}")


# ── Main Scanner ──────────────────────────────────────────────────────────────

def scan_inbox() -> dict:
    if AI_ENABLED and not OPENROUTER_API_KEY:
        print("  ⚠️  AI_ENABLED=True but no OPENROUTER_API_KEY found in .env")
        print("      Falling back to keyword-only classification.")

    print("Connecting to Outlook…")
    try:
        outlook   = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        namespace.Logon()
        print("  ✓ Connected to Outlook")
    except Exception as e:
        print(f"  ✗ Could not connect to Outlook: {e}")
        print("    Make sure Microsoft Outlook is installed and open.")
        return {"error": str(e), "new": 0, "total": 0}

    cutoff = datetime.now() - timedelta(days=SCAN_DAYS)
    print(f"  Scanning emails since {cutoff.date()}…")
    ai_mode = "AI + keyword fallback" if (AI_ENABLED and OPENROUTER_API_KEY) else "keyword-only"
    print(f"  Classification mode: {ai_mode}\n")

    init_db()

    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row

    new_apps    = 0   # brand-new applications created
    updated     = 0   # existing applications whose status advanced
    skipped     = 0   # emails already in DB
    total_seen  = 0
    match_count = 0

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
            total_seen  += 1
            folder_count += 1

            try:
                subject      = str(item.Subject or "")
                sender       = str(item.SenderName or "")
                sender_email = ""
                try:
                    sender_email = str(item.SenderEmailAddress or "")
                    if sender_email.startswith("/O="):
                        sender_email = ""
                except Exception:
                    pass

                try:
                    date_str = item.ReceivedTime.strftime("%Y-%m-%d")
                except Exception:
                    date_str = datetime.now().strftime("%Y-%m-%d")

                try:
                    body_preview = str(item.Body or "")[:1000]
                except Exception:
                    body_preview = ""

                # Stage 1 — fast keyword pre-filter (free)
                if not is_job_email_fast(subject, body_preview):
                    continue
                match_count += 1

                # Stage 2 — AI classification
                ai_result = classify_email(
                    subject      = subject,
                    body_preview = body_preview,
                    sender_name  = sender,
                    sender_email = sender_email,
                    api_key      = OPENROUTER_API_KEY if AI_ENABLED else None,
                    model        = OPENROUTER_MODEL,
                    ai_enabled   = AI_ENABLED,
                )

                # Double-check: AI may reject false positives
                if not ai_result["is_job_email"]:
                    continue

                company   = ai_result["company"]
                job_title = ai_result["job_title"]
                status    = ai_result["status"] if ai_result["status"] != "Unknown" else "Applied"
                ai_used   = ai_result["used_ai"]

                sender_full = f"{sender} <{sender_email}>" if sender_email else sender

                try:
                    email_id = str(item.EntryID)[:200]
                except Exception:
                    email_id = f"{sender_email}_{date_str}_{hash(subject)}"

                # Stage 3 — find or create application
                app_id, is_new = find_or_create_application(
                    conn, company, job_title, date_str,
                    sender_full, subject, ai_used
                )

                # Add the email event (skip if duplicate)
                inserted = add_email_event(
                    conn, app_id, email_id, date_str,
                    status, subject, sender_full
                )

                if not inserted:
                    skipped += 1
                    continue

                # Recompute application status from all events
                old_status = conn.execute(
                    "SELECT status FROM applications WHERE id = ?", (app_id,)
                ).fetchone()["status"]

                refresh_application_status(conn, app_id)

                new_status = conn.execute(
                    "SELECT status FROM applications WHERE id = ?", (app_id,)
                ).fetchone()["status"]

                if is_new:
                    new_apps += 1
                    tag = "🆕"
                elif new_status != old_status:
                    updated += 1
                    tag = f"📈 {old_status}→{new_status}"
                else:
                    tag = "  "

                ai_tag = "🤖" if ai_used else "🔤"
                print(f"    {ai_tag}[{new_status:9s}] {company:<22} — {subject[:45]}  {tag}")

                if folder_count % 100 == 0:
                    print(f"    … {folder_count} emails checked in {folder_label}")

            except Exception as e:
                print(f"    Warning: skipping one email — {e}")
                continue

        print(f"  ✓ {folder_label}: checked {folder_count} emails\n")

    conn.close()

    total_apps    = sqlite3.connect(DB_PATH).execute(
        "SELECT COUNT(*) FROM applications"
    ).fetchone()[0]
    total_events  = sqlite3.connect(DB_PATH).execute(
        "SELECT COUNT(*) FROM email_events"
    ).fetchone()[0]

    print(f"✓ Done. Scanned {total_seen} emails, matched {match_count} job emails.")
    print(f"  New applications: {new_apps}  |  Status updates: {updated}  |"
          f"  Duplicates skipped: {skipped}")
    print(f"  DB: {total_apps} applications, {total_events} email events total.")

    return {"new": new_apps, "updated": updated, "total": total_apps}


if __name__ == "__main__":
    result = scan_inbox()
    if "error" in result:
        print(f"\n✗ Error: {result['error']}")
