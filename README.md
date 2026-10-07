# Virtual Tax Auditor — Phase 3 Closed-Loop Test

Extends the tested Phase 2 workflow with Taxpayer Response & Evidence, AI Response Analysis, Second Human Validation, Further Action, Outcome & Closure, and an expanded audit trail.

This is a controlled deterministic test engine, not a production LLM/RAG tax-law system.

Render build: `pip install -r requirements.txt`
Render start: `gunicorn app:app`

Environment variables:
- GOOGLE_CLIENT_SECRETS=/etc/secrets/client_secret.json
- BASE_URL=https://virtual-tax-auditor-phase3.onrender.com
- COOKIE_SECURE=1
- FLASK_SECRET_KEY=<secret>

Google OAuth redirect URI:
https://virtual-tax-auditor-phase3.onrender.com/oauth2callback
