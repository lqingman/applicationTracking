"""
test_ai.py
Quick sanity-check for the OpenRouter AI integration.
Run this BEFORE doing a full inbox scan to verify your API key works.

Usage:
    python test_ai.py
"""

import os

try:
    from dotenv import load_dotenv
    load_dotenv(os.path.join(os.path.dirname(__file__), ".env"))
except ImportError:
    pass

from ai_classifier import classify_email, is_job_email_fast

API_KEY = os.environ.get("OPENROUTER_API_KEY", "")

# ── Sample emails ─────────────────────────────────────────────────────────────

SAMPLES = [
    {
        "label":        "Application confirmation",
        "subject":      "Your application for Software Engineer at Acme Corp",
        "body_preview": "Thank you for applying for the Software Engineer position at Acme Corp. "
                        "We have received your application and will be in touch shortly.",
        "sender_name":  "Acme Recruiting",
        "sender_email": "no-reply@greenhouse.io",
        "expect_status": "Applied",
    },
    {
        "label":        "Interview invite",
        "subject":      "Interview Invitation – Data Scientist Role",
        "body_preview": "We'd like to schedule a technical phone screen with you for the Data Scientist "
                        "position. Please use the link below to pick a time that works for you.",
        "sender_name":  "Sarah from TechCo",
        "sender_email": "sarah@techco.com",
        "expect_status": "Interview",
    },
    {
        "label":        "Rejection",
        "subject":      "Your application with Global Bank",
        "body_preview": "Thank you for your interest in the Analyst position. After careful consideration, "
                        "we have decided to move forward with other candidates. We wish you the best.",
        "sender_name":  "Global Bank Talent",
        "sender_email": "talent@globalbank.com",
        "expect_status": "Rejected",
    },
    {
        "label":        "Non-job email (should be filtered)",
        "subject":      "Your Amazon order has shipped",
        "body_preview": "Your order #123-456 has been shipped and will arrive by Friday.",
        "sender_name":  "Amazon",
        "sender_email": "ship-confirm@amazon.com",
        "expect_status": None,   # should not be a job email
    },
]


def run_tests():
    if not API_KEY:
        print("⚠️  No OPENROUTER_API_KEY found in .env — testing keyword fallback only.\n")
    else:
        print(f"✓ API key found ({API_KEY[:8]}…)\n")

    passed = 0
    total  = len(SAMPLES)

    for s in SAMPLES:
        print(f"── {s['label']} ──")

        # Stage 1: keyword pre-filter
        fast = is_job_email_fast(s["subject"], s["body_preview"])
        print(f"   Keyword filter:  {'✓ job email' if fast else '✗ filtered out'}")

        if not fast and s["expect_status"] is None:
            print("   ✅ Correctly filtered as non-job email\n")
            passed += 1
            continue

        # Stage 2: AI classification
        result = classify_email(
            subject      = s["subject"],
            body_preview = s["body_preview"],
            sender_name  = s["sender_name"],
            sender_email = s["sender_email"],
            api_key      = API_KEY,
            ai_enabled   = bool(API_KEY),
        )

        mode = "🤖 AI" if result["used_ai"] else "🔤 keyword"
        print(f"   Mode:            {mode}")
        print(f"   is_job_email:    {result['is_job_email']}")
        print(f"   company:         {result['company']}")
        print(f"   job_title:       {result['job_title']}")
        print(f"   status:          {result['status']}  (expected: {s['expect_status']})")
        print(f"   confidence:      {result['confidence']:.2f}")

        ok = (result["status"] == s["expect_status"]) and result["is_job_email"]
        if s["expect_status"] is None:
            ok = not result["is_job_email"]
        print(f"   {'✅ PASS' if ok else '❌ FAIL'}\n")
        if ok:
            passed += 1

    print(f"Results: {passed}/{total} passed")


if __name__ == "__main__":
    run_tests()
