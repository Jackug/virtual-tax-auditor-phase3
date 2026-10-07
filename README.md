# Virtual Tax Auditor — Phase 2 Build

This build implements a functional synthetic-data Phase 2 workflow:

1. Google OAuth login (same Gmail sender capability as the previous build)
2. Synthetic taxpayer evidence base
3. Natural-language risk rules stored with structured logic
4. Risk assessment and risk drivers
5. Risk Universe filtering
6. Taxpayer 360
7. Human selection for Virtual Audit
8. Deterministic Phase-2 audit analysis engine over synthetic data
9. Structured AI-labelled draft findings
10. Human validation: Approve / Reject / Needs Review
11. Decision history and system audit events

## Important honesty note
The Phase 2 analysis engine is deliberately deterministic. It demonstrates the DTD workflow and data lineage without pretending that a live LLM/RAG tax-law agent has already been connected. A production AI/RAG layer should replace or augment `run_ai_analysis()` after approved tax-law knowledge sources and model governance are configured.

## Google OAuth
Use the same Google OAuth client as the existing deployment. Keep `client_secret.json` outside source control.

Required environment variables:
- `GOOGLE_CLIENT_SECRETS=/etc/secrets/client_secret.json`
- `BASE_URL=https://vta-8qs7.onrender.com`
- `FLASK_SECRET_KEY=<random secret>`
- `COOKIE_SECURE=1`

Redirect URI:
`https://vta-8qs7.onrender.com/oauth2callback`

## Local test
```bash
pip install -r requirements.txt
set GOOGLE_CLIENT_SECRETS=C:\path\to\client_secret.json
set BASE_URL=http://localhost:5000
python app.py
```

For production use Gunicorn/Render rather than Flask's development server.

## Reset demo database
```bash
flask --app app reset-demo
```

This deletes the local Phase 2 SQLite database and recreates the synthetic ABC Traders Ltd case.

## Phase 2 test path
Login → My Tasks → Data & Rules → Risk Assessment → Risk Universe → View 360 → Select for Virtual Audit → Run AI Audit Analysis → Open finding → Human Validation.
