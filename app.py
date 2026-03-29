"""
app.py
Flask server — serves the dashboard and exposes the JSON API.
Works with the new two-table schema (applications + email_events).
"""

from flask import Flask, jsonify, render_template, request
import sqlite3
import os
from datetime import datetime

app = Flask(__name__)
DB_PATH = os.path.join(os.path.dirname(__file__), "applications.db")


# ── Helpers ───────────────────────────────────────────────────────────────────

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def ensure_db():
    from email_scraper import init_db
    init_db()


# ── API Routes ────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("dashboard.html")


@app.route("/api/stats")
def api_stats():
    ensure_db()
    conn = get_db()

    total    = conn.execute("SELECT COUNT(*) FROM applications").fetchone()[0]
    by_status = conn.execute(
        "SELECT status, COUNT(*) as cnt FROM applications GROUP BY status"
    ).fetchall()
    status_map = {row["status"]: row["cnt"] for row in by_status}

    # Applications per month (last 12 months), keyed on applied_date
    monthly = conn.execute("""
        SELECT strftime('%Y-%m', applied_date) as month, COUNT(*) as cnt
        FROM applications
        WHERE applied_date >= date('now', '-12 months')
        GROUP BY month
        ORDER BY month
    """).fetchall()

    # Top 10 companies
    top_companies = conn.execute("""
        SELECT company, COUNT(*) as cnt
        FROM applications
        WHERE company NOT IN ('Unknown', '')
        GROUP BY company
        ORDER BY cnt DESC
        LIMIT 10
    """).fetchall()

    # AI usage stat
    ai_count = conn.execute(
        "SELECT COUNT(*) FROM applications WHERE ai_used = 1"
    ).fetchone()[0]

    conn.close()
    return jsonify({
        "total":         total,
        "applied":       status_map.get("Applied",   0),
        "interview":     status_map.get("Interview", 0),
        "offer":         status_map.get("Offer",     0),
        "rejected":      status_map.get("Rejected",  0),
        "ai_classified": ai_count,
        "status_chart":  [{"status": k, "count": v} for k, v in status_map.items()],
        "monthly":       [{"month": r["month"], "count": r["cnt"]} for r in monthly],
        "top_companies": [{"company": r["company"], "count": r["cnt"]} for r in top_companies],
    })


@app.route("/api/applications")
def api_applications():
    ensure_db()
    conn = get_db()

    search   = request.args.get("q", "").strip()
    status   = request.args.get("status", "").strip()
    sort     = request.args.get("sort", "last_update")
    order    = request.args.get("order", "desc").upper()
    page     = max(1, int(request.args.get("page", 1)))
    per_page = 25

    allowed_sorts = {"last_update", "applied_date", "company", "status", "job_title"}
    if sort not in allowed_sorts:
        sort = "last_update"
    if order not in ("ASC", "DESC"):
        order = "DESC"

    conditions: list[str] = []
    params:     list      = []

    if search:
        conditions.append(
            "(company LIKE ? OR job_title LIKE ? OR subject LIKE ? OR sender LIKE ?)"
        )
        q = f"%{search}%"
        params.extend([q, q, q, q])
    if status:
        conditions.append("status = ?")
        params.append(status)

    where = ("WHERE " + " AND ".join(conditions)) if conditions else ""

    total_rows = conn.execute(
        f"SELECT COUNT(*) FROM applications {where}", params
    ).fetchone()[0]

    rows = conn.execute(
        f"""SELECT
                a.*,
                (SELECT COUNT(*) FROM email_events e WHERE e.application_id = a.id) AS event_count
            FROM applications a
            {where}
            ORDER BY {sort} {order}
            LIMIT ? OFFSET ?""",
        params + [per_page, (page - 1) * per_page],
    ).fetchall()

    conn.close()
    return jsonify({
        "total": total_rows,
        "page":  page,
        "pages": max(1, -(-total_rows // per_page)),
        "items": [dict(r) for r in rows],
    })


@app.route("/api/applications/<int:app_id>")
def api_application_detail(app_id):
    """Return a single application with its full email event timeline."""
    ensure_db()
    conn = get_db()

    app_row = conn.execute(
        "SELECT * FROM applications WHERE id = ?", (app_id,)
    ).fetchone()
    if not app_row:
        conn.close()
        return jsonify({"error": "Not found"}), 404

    events = conn.execute("""
        SELECT id, date, status, subject, sender, created_at
        FROM email_events
        WHERE application_id = ?
        ORDER BY date ASC
    """, (app_id,)).fetchall()

    conn.close()
    return jsonify({
        **dict(app_row),
        "events": [dict(e) for e in events],
    })


@app.route("/api/applications/<int:app_id>/events")
def api_application_events(app_id):
    """Return only the email event timeline for one application."""
    ensure_db()
    conn = get_db()
    events = conn.execute("""
        SELECT id, date, status, subject, sender, created_at
        FROM email_events
        WHERE application_id = ?
        ORDER BY date ASC
    """, (app_id,)).fetchall()
    conn.close()
    return jsonify([dict(e) for e in events])


@app.route("/api/scan", methods=["POST"])
def api_scan():
    """Trigger a fresh Outlook scan (runs synchronously)."""
    try:
        from email_scraper import scan_inbox
        result = scan_inbox()
        return jsonify({"ok": True, **result})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


if __name__ == "__main__":
    ensure_db()
    print("Dashboard running at http://localhost:5000")
    app.run(debug=False, port=5000)
