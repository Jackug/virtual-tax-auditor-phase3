# Virtual Tax Auditor — Phase 3

This package is the **closed-loop Phase 3 test application** for the Virtual Tax Auditor.

## Phase 3 definition of done

1. Risk Assessment
2. Risk Universe / taxpayer prioritisation
3. AI Audit Analysis
4. Human Finding Validation
5. Human Communication / Email approval and Gmail sending
6. Taxpayer Response & Evidence
7. AI Response Analysis
8. Second Human Validation
9. Further Action / Outcome / Closure
10. Audit trail throughout the workflow

AI-generated findings and response analysis remain preliminary. Human users approve findings, approve taxpayer communications, validate taxpayer responses and determine consequential outcomes.

## Dashboard

The four KPI cards are functional navigation controls:

- Taxpayers → Risk Universe
- High / Critical → filtered Risk Universe
- Open Cases → Virtual Audit case queue
- AI Findings → AI findings validation queue

## Closed-loop behaviour

After Second Human Validation, unresolved cases can create further actions. A completed further action returns the case to the taxpayer-response stage so the response/evidence → AI analysis → human validation loop can continue.

## Test environment

The application seeds one synthetic taxpayer and synthetic risk rules. It does not connect to URA production data.

The AI audit and response-analysis components in this package are controlled deterministic test engines. They are deliberately labelled as such; they are not represented as a production LLM/RAG service.

## Google/Gmail

Set the same environment variables used by the Render service:

- GOOGLE_CLIENT_SECRETS=/etc/secrets/client_secret.json
- BASE_URL=https://<your-render-service>.onrender.com
- FLASK_SECRET_KEY=<random-secret>
- COOKIE_SECURE=1

The Google OAuth redirect URI must exactly match `<BASE_URL>/oauth2callback`.

Do **not** commit `client_secret.json`, OAuth tokens, or the SQLite database to GitHub.

## Render

Build command:

`pip install -r requirements.txt`

Start command:

`gunicorn app:app`
