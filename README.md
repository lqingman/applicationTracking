# 📬 Job Application Tracker

A local tool that scans your **Outlook Inbox** for job application emails and displays a beautiful analysis dashboard in your browser — entirely offline, no cloud required.

![Dashboard Preview](docs/dashboard_preview.png)

---

## ✨ Features

- 🔍 **Automatic email scanning** — reads your Outlook Inbox via Windows COM (no API keys needed)
- 🗂️ **Smart classification** — categorises emails as *Applied*, *Interview*, *Offer*, or *Rejected* using keyword heuristics
- 📊 **Live dashboard** — KPI cards, donut chart, monthly timeline, top companies, and a searchable/sortable data table
- 💾 **Local SQLite storage** — all data stays on your machine
- 🔄 **On-demand re-scan** — one click to pull in new emails from the browser dashboard

---

## 🚀 Quick Start

### Prerequisites

- Windows with **Microsoft Outlook** desktop app installed and open
- Python 3.9+

### 1. Clone & set up virtual environment

```bash
git clone <repo-url>
cd applicationTracking

python -m venv venv
venv\Scripts\activate
```

### 2. Install dependencies

```bash
pip install -r requirements.txt
```

### 3. Run the tracker

**Option A — one-click launcher (recommended):**

```bash
scan_and_update.bat
```

**Option B — manual steps:**

```bash
# Step 1: Scan your Outlook inbox
python email_scraper.py

# Step 2: Start the dashboard server
python app.py
```

Then open **http://localhost:5000** in your browser.

---

## 📁 Project Structure

```
applicationTracking/
├── email_scraper.py       # Reads Outlook Inbox via win32com, saves to SQLite
├── app.py                 # Flask API server + dashboard host
├── templates/
│   └── dashboard.html     # Dark-mode dashboard (Chart.js + Vanilla JS)
├── requirements.txt       # Python dependencies
├── scan_and_update.bat    # One-click launcher
├── applications.db        # SQLite database (auto-created, git-ignored)
└── venv/                  # Virtual environment (git-ignored)
```

---

## ⚙️ Configuration

Open `email_scraper.py` and adjust these constants at the top:

| Variable | Default | Description |
|---|---|---|
| `SCAN_DAYS` | `730` | How many days back to scan (~2 years) |
| `JOB_KEYWORDS` | *(list)* | Keywords to identify job-related emails |
| `STATUS_RULES` | *(list)* | Rules for classifying email status |

---

## 🛡️ Privacy

- All data is stored in a local `applications.db` SQLite file
- No data is sent to any external server
- The database is excluded from git via `.gitignore`
- Outlook is accessed read-only via the Windows COM interface

---

## 🐛 Troubleshooting

**"Could not connect to Outlook"**
→ Make sure Outlook is open and you're logged in before running the script.

**No emails found after scanning**
→ The keywords may not match your inbox style. Add company names or custom phrases to `JOB_KEYWORDS` in `email_scraper.py`.

**Dashboard shows blank charts**
→ Run a scan first — the DB starts empty. Click **"🔄 Scan Outlook"** in the dashboard header.

---

## 📦 Dependencies

| Package | Purpose |
|---|---|
| `pywin32` | Access Outlook via Windows COM/MAPI |
| `flask` | Lightweight web server for the dashboard |
| `pandas` | Data manipulation (optional helpers) |

---

## 📄 License

MIT — for personal use.
