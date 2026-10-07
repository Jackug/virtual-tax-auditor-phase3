import os
import json
import sqlite3
import base64
from datetime import datetime, timezone
from functools import wraps
from email.mime.text import MIMEText

from flask import Flask, render_template, redirect, url_for, session, request, abort, flash
from google_auth_oauthlib.flow import Flow
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from google.auth.transport.requests import Request

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, os.getenv("DB_PATH", "vta_phase2.sqlite3"))
CLIENT_SECRETS = os.getenv("GOOGLE_CLIENT_SECRETS", os.path.join(BASE_DIR, "client_secret.json"))
BASE_URL = os.getenv("BASE_URL", "http://localhost:5000").rstrip("/")
SECRET_KEY = os.getenv("FLASK_SECRET_KEY", "change-me-in-production")
COOKIE_SECURE = os.getenv("COOKIE_SECURE", "0") == "1"

SCOPES = [
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/userinfo.email",
    "openid",
]

app = Flask(__name__)
app.secret_key = SECRET_KEY
app.config.update(
    SESSION_COOKIE_SECURE=COOKIE_SECURE,
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
)


def now():
    return datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")


def db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def init_db():
    conn = db()
    conn.executescript(
        """
        CREATE TABLE IF NOT EXISTS oauth_credentials (
            email TEXT PRIMARY KEY,
            token TEXT NOT NULL,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS users (
            email TEXT PRIMARY KEY,
            role TEXT NOT NULL DEFAULT 'AUDITOR',
            active INTEGER NOT NULL DEFAULT 1,
            created_at TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS taxpayers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            tin TEXT UNIQUE NOT NULL,
            name TEXT NOT NULL,
            sector TEXT NOT NULL,
            station TEXT NOT NULL,
            financial_year TEXT NOT NULL,
            status TEXT NOT NULL DEFAULT 'Active'
        );

        CREATE TABLE IF NOT EXISTS taxpayer_metrics (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            taxpayer_id INTEGER NOT NULL,
            metric_name TEXT NOT NULL,
            declared_value REAL NOT NULL,
            observed_value REAL,
            source TEXT NOT NULL,
            period TEXT NOT NULL,
            FOREIGN KEY(taxpayer_id) REFERENCES taxpayers(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS risk_rules (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            natural_language TEXT NOT NULL,
            structured_logic TEXT NOT NULL,
            category TEXT NOT NULL,
            approved INTEGER NOT NULL DEFAULT 1,
            created_at TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS risk_assessments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            taxpayer_id INTEGER NOT NULL,
            score REAL NOT NULL,
            band TEXT NOT NULL,
            exposure REAL,
            assessed_at TEXT NOT NULL,
            engine_version TEXT NOT NULL,
            FOREIGN KEY(taxpayer_id) REFERENCES taxpayers(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS risk_drivers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            assessment_id INTEGER NOT NULL,
            rule_id INTEGER NOT NULL,
            description TEXT NOT NULL,
            variance REAL,
            variance_pct REAL,
            FOREIGN KEY(assessment_id) REFERENCES risk_assessments(id) ON DELETE CASCADE,
            FOREIGN KEY(rule_id) REFERENCES risk_rules(id)
        );

        CREATE TABLE IF NOT EXISTS audit_cases (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            case_ref TEXT UNIQUE NOT NULL,
            taxpayer_id INTEGER NOT NULL,
            status TEXT NOT NULL,
            selected_by TEXT,
            selected_at TEXT,
            assigned_to TEXT,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            FOREIGN KEY(taxpayer_id) REFERENCES taxpayers(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS audit_analyses (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            case_id INTEGER NOT NULL,
            analysis_status TEXT NOT NULL,
            summary TEXT NOT NULL,
            methodology TEXT NOT NULL,
            limitations TEXT NOT NULL,
            created_at TEXT NOT NULL,
            FOREIGN KEY(case_id) REFERENCES audit_cases(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS findings (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            finding_ref TEXT UNIQUE NOT NULL,
            case_id INTEGER NOT NULL,
            risk_category TEXT NOT NULL,
            description TEXT NOT NULL,
            financial_year TEXT NOT NULL,
            expected_value REAL,
            observed_value REAL,
            variance REAL,
            variance_pct REAL,
            potential_exposure REAL,
            source TEXT NOT NULL,
            audit_test TEXT NOT NULL,
            explanation TEXT NOT NULL,
            ai_confidence TEXT NOT NULL,
            limitations TEXT NOT NULL,
            status TEXT NOT NULL DEFAULT 'AI Generated',
            created_at TEXT NOT NULL,
            FOREIGN KEY(case_id) REFERENCES audit_cases(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS finding_decisions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            finding_id INTEGER NOT NULL,
            decision TEXT NOT NULL,
            reason TEXT NOT NULL,
            decided_by TEXT NOT NULL,
            decided_at TEXT NOT NULL,
            FOREIGN KEY(finding_id) REFERENCES findings(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS audit_events (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            actor_email TEXT,
            case_ref TEXT,
            event_type TEXT NOT NULL,
            detail TEXT NOT NULL,
            created_at TEXT NOT NULL
        );
        """
    )
    seed_demo_data(conn)
    conn.commit()
    conn.close()


def seed_demo_data(conn):
    cur = conn.cursor()
    existing = cur.execute("SELECT COUNT(*) AS n FROM taxpayers").fetchone()["n"]
    if existing:
        return

    cur.execute(
        "INSERT INTO taxpayers(tin,name,sector,station,financial_year) VALUES (?,?,?,?,?)",
        ("1000000001", "ABC Traders Ltd", "Wholesale & Retail Trade", "Kampala", "2025/26"),
    )
    taxpayer_id = cur.lastrowid

    metrics = [
        (taxpayer_id, "Sales", 1_000_000_000, 1_500_000_000, "Third-party / observed sales", "2025/26"),
        (taxpayer_id, "Purchases", 700_000_000, 700_000_000, "Declared purchases", "2025/26"),
        (taxpayer_id, "Imports", 700_000_000, 800_000_000, "Customs / import data", "2025/26"),
        (taxpayer_id, "PAYE", 50_000_000, 52_000_000, "PAYE / employment data", "2025/26"),
    ]
    cur.executemany(
        "INSERT INTO taxpayer_metrics(taxpayer_id,metric_name,declared_value,observed_value,source,period) VALUES (?,?,?,?,?,?)",
        metrics,
    )

    rules = [
        (
            "Observed Sales Reconciliation",
            "I want taxpayers where observed sales are higher than declared sales",
            "observed_sales > declared_sales",
            "Sales reconciliation",
        ),
        (
            "Imports versus Purchases",
            "I want taxpayers where imports are higher than declared purchases",
            "imports > declared_purchases",
            "Imports / purchases reconciliation",
        ),
    ]
    for name, nl, logic, category in rules:
        cur.execute(
            "INSERT INTO risk_rules(name,natural_language,structured_logic,category,approved,created_at) VALUES (?,?,?,?,1,?)",
            (name, nl, logic, category, now()),
        )


def log_event(event_type, detail, case_ref=None, actor=None):
    conn = db()
    conn.execute(
        "INSERT INTO audit_events(actor_email,case_ref,event_type,detail,created_at) VALUES (?,?,?,?,?)",
        (actor or session.get("email"), case_ref, event_type, detail, now()),
    )
    conn.commit()
    conn.close()


def current_email():
    return session.get("email")


def login_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if not current_email():
            return redirect(url_for("login"))
        return fn(*args, **kwargs)
    return wrapper


def save_credentials(email, creds):
    payload = json.dumps({
        "token": creds.token,
        "refresh_token": creds.refresh_token,
        "token_uri": creds.token_uri,
        "client_id": creds.client_id,
        "client_secret": creds.client_secret,
        "scopes": creds.scopes,
    })
    conn = db()
    conn.execute(
        "INSERT INTO oauth_credentials(email,token,created_at,updated_at) VALUES (?,?,?,?) "
        "ON CONFLICT(email) DO UPDATE SET token=excluded.token,updated_at=excluded.updated_at",
        (email, payload, now(), now()),
    )
    conn.execute(
        "INSERT INTO users(email,role,created_at) VALUES (?, 'AUDITOR', ?) ON CONFLICT(email) DO NOTHING",
        (email, now()),
    )
    conn.commit()
    conn.close()


def get_credentials(email):
    conn = db()
    row = conn.execute("SELECT token FROM oauth_credentials WHERE email=?", (email,)).fetchone()
    conn.close()
    if not row:
        return None
    data = json.loads(row["token"])
    return Credentials.from_authorized_user_info(data, SCOPES)


def oauth_flow(state=None, code_verifier=None):
    if not os.path.exists(CLIENT_SECRETS):
        raise RuntimeError(f"Google client secret not found: {CLIENT_SECRETS}")

    # Persist the PKCE verifier between the authorization request and callback.
    # The previous Phase 2 build generated a verifier during /authorize but
    # created a new flow during /oauth2callback without restoring it, causing:
    # "(invalid_grant) Missing code verifier."
    flow = Flow.from_client_secrets_file(
        CLIENT_SECRETS,
        scopes=SCOPES,
        state=state,
        code_verifier=code_verifier,
    )
    flow.redirect_uri = f"{BASE_URL}/oauth2callback"
    return flow


def get_metric_map(conn, taxpayer_id):
    rows = conn.execute("SELECT * FROM taxpayer_metrics WHERE taxpayer_id=?", (taxpayer_id,)).fetchall()
    return {r["metric_name"].lower(): r for r in rows}


def run_risk_assessment(taxpayer_id):
    conn = db()
    taxpayer = conn.execute("SELECT * FROM taxpayers WHERE id=?", (taxpayer_id,)).fetchone()
    if not taxpayer:
        conn.close()
        abort(404)
    metrics = get_metric_map(conn, taxpayer_id)
    rules = conn.execute("SELECT * FROM risk_rules WHERE approved=1 ORDER BY id").fetchall()

    drivers = []
    # Rule 1: observed sales > declared sales
    sales = metrics.get("sales")
    if sales and sales["observed_value"] is not None and sales["observed_value"] > sales["declared_value"]:
        variance = sales["observed_value"] - sales["declared_value"]
        pct = (variance / sales["declared_value"] * 100) if sales["declared_value"] else None
        rule = next((r for r in rules if r["structured_logic"] == "observed_sales > declared_sales"), None)
        if rule:
            drivers.append((rule, "Observed sales exceed declared sales by UGX {:,.0f}.".format(variance), variance, pct))

    # Rule 2: imports > declared purchases
    imports = metrics.get("imports")
    purchases = metrics.get("purchases")
    if imports and purchases and imports["observed_value"] is not None and imports["observed_value"] > purchases["declared_value"]:
        variance = imports["observed_value"] - purchases["declared_value"]
        pct = (variance / purchases["declared_value"] * 100) if purchases["declared_value"] else None
        rule = next((r for r in rules if r["structured_logic"] == "imports > declared_purchases"), None)
        if rule:
            drivers.append((rule, "Imports exceed declared purchases by UGX {:,.0f}.".format(variance), variance, pct))

    score = min(100, 25 * len(drivers) + (25 if any((d[3] or 0) >= 30 for d in drivers) else 0))
    if len(drivers) >= 2:
        score = max(score, 75)
    band = "Critical" if score >= 90 else "High" if score >= 70 else "Medium" if score >= 40 else "Low"
    exposure = sum(d[2] for d in drivers) if drivers else 0

    old = conn.execute("SELECT id FROM risk_assessments WHERE taxpayer_id=?", (taxpayer_id,)).fetchall()
    for row in old:
        conn.execute("DELETE FROM risk_assessments WHERE id=?", (row["id"],))

    conn.execute(
        "INSERT INTO risk_assessments(taxpayer_id,score,band,exposure,assessed_at,engine_version) VALUES (?,?,?,?,?,?)",
        (taxpayer_id, score, band, exposure, now(), "Phase-2-Risk-Engine-1.0"),
    )
    assessment_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
    for rule, desc, variance, pct in drivers:
        conn.execute(
            "INSERT INTO risk_drivers(assessment_id,rule_id,description,variance,variance_pct) VALUES (?,?,?,?,?)",
            (assessment_id, rule["id"], desc, variance, pct),
        )
    conn.commit()
    conn.close()
    log_event("RISK_ASSESSMENT_RUN", f"Risk assessment generated for TIN {taxpayer['tin']}; score={score}; band={band}")
    return assessment_id


def get_assessment(conn, taxpayer_id):
    return conn.execute("SELECT * FROM risk_assessments WHERE taxpayer_id=? ORDER BY id DESC LIMIT 1", (taxpayer_id,)).fetchone()


def create_audit_case(taxpayer_id, email):
    conn = db()
    taxpayer = conn.execute("SELECT * FROM taxpayers WHERE id=?", (taxpayer_id,)).fetchone()
    existing = conn.execute("SELECT * FROM audit_cases WHERE taxpayer_id=? ORDER BY id DESC LIMIT 1", (taxpayer_id,)).fetchone()
    if existing and existing["status"] not in ("Closed", "Rejected"):
        conn.close()
        return existing["id"]
    case_ref = f"VTA-TEST-{datetime.now(timezone.utc).strftime('%Y%m%d%H%M%S')}"
    conn.execute(
        "INSERT INTO audit_cases(case_ref,taxpayer_id,status,selected_by,selected_at,assigned_to,created_at,updated_at) VALUES (?,?,?,?,?,?,?,?)",
        (case_ref, taxpayer_id, "Selected for Virtual Audit", email, now(), email, now(), now()),
    )
    case_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
    conn.commit()
    conn.close()
    log_event("CASE_SELECTED", f"Taxpayer {taxpayer['tin']} selected for Virtual Audit", case_ref=case_ref, actor=email)
    return case_id


def run_ai_analysis(case_id, email):
    conn = db()
    case = conn.execute("SELECT * FROM audit_cases WHERE id=?", (case_id,)).fetchone()
    taxpayer = conn.execute("SELECT * FROM taxpayers WHERE id=?", (case["taxpayer_id"],)).fetchone()
    metrics = get_metric_map(conn, taxpayer["id"])
    # Deterministic analysis for the synthetic test case. This is intentionally labelled as a demo analysis engine,
    # not a claim of live LLM reasoning.
    sales = metrics["sales"]
    imports = metrics["imports"]
    purchases = metrics["purchases"]
    summary = (
        "The Phase 2 analysis engine identified two potential reconciliation issues: "
        "observed sales exceed declared sales, and imports exceed declared purchases. "
        "These are risk indicators requiring human review and taxpayer clarification; they are not final findings of non-compliance."
    )
    methodology = (
        "Compared declared sales with observed sales; compared import records with declared purchases; "
        "calculated absolute and percentage variances; linked each result to its source metric."
    )
    limitations = (
        "This Phase 2 build uses a deterministic test analysis engine over synthetic data. "
        "It is not a production LLM/RAG tax-law conclusion and does not make a final tax decision."
    )
    conn.execute(
        "INSERT INTO audit_analyses(case_id,analysis_status,summary,methodology,limitations,created_at) VALUES (?,?,?,?,?,?)",
        (case_id, "Completed", summary, methodology, limitations, now()),
    )
    # Clear previous demo findings for repeatable testing.
    conn.execute("DELETE FROM findings WHERE case_id=?", (case_id,))
    f1_var = sales["observed_value"] - sales["declared_value"]
    f1_pct = f1_var / sales["declared_value"] * 100
    f2_var = imports["observed_value"] - purchases["declared_value"]
    f2_pct = f2_var / purchases["declared_value"] * 100
    findings = [
        (
            f"F-{datetime.now(timezone.utc).strftime('%Y%m%d%H%M%S')}-01", case_id,
            "Sales reconciliation",
            "Observed sales exceed declared sales.", taxpayer["financial_year"],
            sales["declared_value"], sales["observed_value"], f1_var, f1_pct, None,
            sales["source"], "Observed Sales - Declared Sales",
            "The observed sales amount is higher than the declared sales amount. The difference may have explanations that require taxpayer evidence and reconciliation.",
            "Medium", limitations,
        ),
        (
            f"F-{datetime.now(timezone.utc).strftime('%Y%m%d%H%M%S')}-02", case_id,
            "Imports / purchases reconciliation",
            "Imports exceed declared purchases.", taxpayer["financial_year"],
            purchases["declared_value"], imports["observed_value"], f2_var, f2_pct, None,
            imports["source"], "Imports - Declared Purchases",
            "Recorded imports are higher than declared purchases. Timing, inventory, classification or other explanations should be considered before treating the variance as confirmed non-compliance.",
            "Medium", limitations,
        ),
    ]
    conn.executemany(
        "INSERT INTO findings(finding_ref,case_id,risk_category,description,financial_year,expected_value,observed_value,variance,variance_pct,potential_exposure,source,audit_test,explanation,ai_confidence,limitations,status,created_at) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
        [f + ("AI Generated", now()) for f in findings],
    )
    conn.execute("UPDATE audit_cases SET status='AI Analysis Completed',updated_at=? WHERE id=?", (now(), case_id))
    conn.commit()
    conn.close()
    log_event("AI_ANALYSIS_COMPLETED", "Phase 2 analysis and draft findings generated; findings remain AI Generated pending human validation.", case_ref=case["case_ref"], actor=email)


@app.route("/")
def index():
    if current_email():
        return redirect(url_for("tasks"))
    return redirect(url_for("login"))


@app.route("/login")
def login():
    return render_template("login.html")


@app.route("/authorize")
def authorize():
    try:
        flow = oauth_flow()
        authorization_url, state = flow.authorization_url(
            access_type="offline",
            include_granted_scopes="true",
            prompt="consent",
        )

        # Store both values because the callback creates a new Flow object.
        session["oauth_state"] = state
        session["oauth_code_verifier"] = flow.code_verifier

        return redirect(authorization_url)
    except Exception as exc:
        return render_template("error.html", message=str(exc)), 500


@app.route("/oauth2callback")
def oauth2callback():
    try:
        state = session.get("oauth_state")
        code_verifier = session.get("oauth_code_verifier")

        if not state:
            raise RuntimeError(
                "OAuth session state is missing or expired. Please start the login again."
            )

        if not code_verifier:
            raise RuntimeError(
                "OAuth PKCE verifier is missing. Please start the login again."
            )

        flow = oauth_flow(
            state=state,
            code_verifier=code_verifier,
        )

        flow.fetch_token(authorization_response=request.url)
        creds = flow.credentials

        oauth_service = build("oauth2", "v2", credentials=creds)
        profile = oauth_service.userinfo().get().execute()
        email = profile["email"]

        save_credentials(email, creds)

        session["email"] = email
        session.pop("oauth_state", None)
        session.pop("oauth_code_verifier", None)

        log_event("LOGIN", "Google OAuth login successful", actor=email)
        return redirect(url_for("tasks"))

    except Exception as exc:
        return render_template("error.html", message=str(exc)), 500


@app.route("/logout")
@login_required
def logout():
    email = current_email()
    log_event("LOGOUT", "User logged out", actor=email)
    session.clear()
    return redirect(url_for("login"))


@app.route("/tasks")
@login_required
def tasks():
    conn = db()
    counts = {
        "taxpayers": conn.execute("SELECT COUNT(*) FROM taxpayers").fetchone()[0],
        "assessed": conn.execute("SELECT COUNT(*) FROM risk_assessments").fetchone()[0],
        "high": conn.execute("SELECT COUNT(*) FROM risk_assessments WHERE band='High'").fetchone()[0],
        "cases": conn.execute("SELECT COUNT(*) FROM audit_cases WHERE status NOT IN ('Closed','Rejected')").fetchone()[0],
        "findings": conn.execute("SELECT COUNT(*) FROM findings WHERE status='AI Generated'").fetchone()[0],
    }
    conn.close()
    return render_template("tasks.html", email=current_email(), counts=counts)


@app.route("/data-sources")
@login_required
def data_sources():
    conn = db()
    taxpayer = conn.execute("SELECT * FROM taxpayers ORDER BY id LIMIT 1").fetchone()
    metrics = conn.execute("SELECT * FROM taxpayer_metrics WHERE taxpayer_id=? ORDER BY id", (taxpayer["id"],)).fetchall() if taxpayer else []
    rules = conn.execute("SELECT * FROM risk_rules ORDER BY id").fetchall()
    conn.close()
    return render_template("data_sources.html", taxpayer=taxpayer, metrics=metrics, rules=rules)


@app.route("/risk-assessment", methods=["GET", "POST"])
@login_required
def risk_assessment():
    conn = db()
    taxpayers = conn.execute("SELECT * FROM taxpayers ORDER BY name").fetchall()
    if request.method == "POST":
        taxpayer_id = int(request.form["taxpayer_id"])
        run_risk_assessment(taxpayer_id)
        flash("Risk assessment completed.", "success")
        return redirect(url_for("risk_assessment", taxpayer_id=taxpayer_id))
    taxpayer_id = request.args.get("taxpayer_id", type=int) or (taxpayers[0]["id"] if taxpayers else None)
    selected = conn.execute("SELECT * FROM taxpayers WHERE id=?", (taxpayer_id,)).fetchone() if taxpayer_id else None
    assessment = get_assessment(conn, taxpayer_id) if taxpayer_id else None
    drivers = conn.execute("SELECT rd.*, rr.name, rr.natural_language, rr.structured_logic FROM risk_drivers rd JOIN risk_rules rr ON rr.id=rd.rule_id WHERE rd.assessment_id=?", (assessment["id"],)).fetchall() if assessment else []
    conn.close()
    return render_template("risk_assessment.html", taxpayers=taxpayers, selected=selected, assessment=assessment, drivers=drivers)


@app.route("/risk-universe")
@login_required
def risk_universe():
    band = request.args.get("band", "All")
    conn = db()
    query = """
        SELECT t.*, ra.id AS assessment_id, ra.score, ra.band, ra.exposure, ra.assessed_at,
               (SELECT COUNT(*) FROM risk_drivers rd WHERE rd.assessment_id=ra.id) AS driver_count,
               (SELECT GROUP_CONCAT(rd.description, ' | ') FROM risk_drivers rd WHERE rd.assessment_id=ra.id) AS drivers
        FROM taxpayers t JOIN risk_assessments ra ON ra.taxpayer_id=t.id
        WHERE ra.id IN (SELECT MAX(id) FROM risk_assessments GROUP BY taxpayer_id)
    """
    params = []
    if band != "All":
        query += " AND ra.band=?"
        params.append(band)
    query += " ORDER BY ra.score DESC, ra.exposure DESC"
    rows = conn.execute(query, params).fetchall()
    conn.close()
    return render_template("risk_universe.html", rows=rows, band=band)


@app.route("/taxpayer/<int:taxpayer_id>")
@login_required
def taxpayer_360(taxpayer_id):
    conn = db()
    taxpayer = conn.execute("SELECT * FROM taxpayers WHERE id=?", (taxpayer_id,)).fetchone()
    if not taxpayer:
        conn.close(); abort(404)
    metrics = conn.execute("SELECT * FROM taxpayer_metrics WHERE taxpayer_id=? ORDER BY id", (taxpayer_id,)).fetchall()
    assessment = get_assessment(conn, taxpayer_id)
    drivers = conn.execute("SELECT rd.*, rr.name, rr.natural_language, rr.structured_logic FROM risk_drivers rd JOIN risk_rules rr ON rr.id=rd.rule_id WHERE rd.assessment_id=?", (assessment["id"],)).fetchall() if assessment else []
    cases = conn.execute("SELECT * FROM audit_cases WHERE taxpayer_id=? ORDER BY id DESC", (taxpayer_id,)).fetchall()
    conn.close()
    return render_template("taxpayer_360.html", taxpayer=taxpayer, metrics=metrics, assessment=assessment, drivers=drivers, cases=cases)


@app.route("/case/<int:case_id>/select", methods=["POST"])
@login_required
def select_case(case_id):
    conn = db()
    case = conn.execute("SELECT * FROM audit_cases WHERE id=?", (case_id,)).fetchone()
    conn.close()
    if not case:
        abort(404)
    flash("Case is already selected.", "info")
    return redirect(url_for("case_detail", case_id=case_id))


@app.route("/taxpayer/<int:taxpayer_id>/select", methods=["POST"])
@login_required
def select_taxpayer(taxpayer_id):
    conn = db()
    assessment = get_assessment(conn, taxpayer_id)
    taxpayer = conn.execute("SELECT * FROM taxpayers WHERE id=?", (taxpayer_id,)).fetchone()
    conn.close()
    if not taxpayer or not assessment:
        abort(400)
    if assessment["band"] not in ("High", "Critical"):
        flash("Only High/Critical cases are selectable in this Phase 2 test.", "error")
        return redirect(url_for("taxpayer_360", taxpayer_id=taxpayer_id))
    case_id = create_audit_case(taxpayer_id, current_email())
    flash("Taxpayer selected for Virtual Audit.", "success")
    return redirect(url_for("case_detail", case_id=case_id))


@app.route("/case/<int:case_id>")
@login_required
def case_detail(case_id):
    conn = db()
    case = conn.execute("SELECT ac.*, t.name AS taxpayer_name, t.tin, t.sector, t.financial_year FROM audit_cases ac JOIN taxpayers t ON t.id=ac.taxpayer_id WHERE ac.id=?", (case_id,)).fetchone()
    if not case:
        conn.close(); abort(404)
    analysis = conn.execute("SELECT * FROM audit_analyses WHERE case_id=? ORDER BY id DESC LIMIT 1", (case_id,)).fetchone()
    findings = conn.execute("SELECT * FROM findings WHERE case_id=? ORDER BY id", (case_id,)).fetchall()
    conn.close()
    return render_template("case.html", case=case, analysis=analysis, findings=findings)


@app.route("/case/<int:case_id>/analyze", methods=["POST"])
@login_required
def analyze_case(case_id):
    conn = db()
    case = conn.execute("SELECT * FROM audit_cases WHERE id=?", (case_id,)).fetchone()
    conn.close()
    if not case:
        abort(404)
    run_ai_analysis(case_id, current_email())
    flash("Analysis completed and findings placed in the Human Validation Queue.", "success")
    return redirect(url_for("case_detail", case_id=case_id))


@app.route("/finding/<int:finding_id>")
@login_required
def finding_detail(finding_id):
    conn = db()
    finding = conn.execute("SELECT f.*, ac.case_ref, t.name AS taxpayer_name, t.tin FROM findings f JOIN audit_cases ac ON ac.id=f.case_id JOIN taxpayers t ON t.id=ac.taxpayer_id WHERE f.id=?", (finding_id,)).fetchone()
    decisions = conn.execute("SELECT * FROM finding_decisions WHERE finding_id=? ORDER BY id DESC", (finding_id,)).fetchall()
    conn.close()
    if not finding:
        abort(404)
    return render_template("finding.html", finding=finding, decisions=decisions)


@app.route("/finding/<int:finding_id>/decision", methods=["POST"])
@login_required
def finding_decision(finding_id):
    decision = request.form.get("decision", "").strip().upper()
    reason = request.form.get("reason", "").strip()
    if decision not in {"APPROVE", "REJECT", "NEEDS REVIEW"} or not reason:
        flash("Decision and reason are required.", "error")
        return redirect(url_for("finding_detail", finding_id=finding_id))
    status = {"APPROVE": "Human Validated", "REJECT": "Rejected", "NEEDS REVIEW": "Needs Review"}[decision]
    conn = db()
    finding = conn.execute("SELECT f.*, ac.case_ref FROM findings f JOIN audit_cases ac ON ac.id=f.case_id WHERE f.id=?", (finding_id,)).fetchone()
    if not finding:
        conn.close(); abort(404)
    conn.execute("UPDATE findings SET status=? WHERE id=?", (status, finding_id))
    conn.execute(
        "INSERT INTO finding_decisions(finding_id,decision,reason,decided_by,decided_at) VALUES (?,?,?,?,?)",
        (finding_id, decision, reason, current_email(), now()),
    )
    conn.execute("UPDATE audit_cases SET status=?,updated_at=? WHERE id=?", ("Finding Validation Completed" if decision != "NEEDS REVIEW" else "Needs Further Review", now(), finding["case_id"]))
    conn.commit(); conn.close()
    log_event("FINDING_HUMAN_DECISION", f"Finding {finding['finding_ref']} decision={decision}; reason={reason}", case_ref=finding["case_ref"], actor=current_email())
    flash(f"Finding marked: {status}.", "success")
    return redirect(url_for("finding_detail", finding_id=finding_id))


@app.route("/communication", methods=["GET", "POST"])
@login_required
def communication():
    default_recipient = "jacksonakampurira@gmail.com"
    if request.method == "POST":
        recipient = request.form.get("recipient", "").strip()
        subject = request.form.get("subject", "").strip()
        body = request.form.get("body", "").strip()
        if not recipient or not subject or not body:
            flash("Recipient, subject and message are required.", "error")
            return render_template("communication.html", email=current_email(), recipient=recipient, subject=subject, body=body)
        try:
            creds = get_credentials(current_email())
            if not creds:
                raise RuntimeError("No stored Google credentials for the signed-in user. Sign in again.")
            if creds.expired and creds.refresh_token:
                creds.refresh(Request())
                save_credentials(current_email(), creds)
            service = build("gmail", "v1", credentials=creds)
            msg = MIMEText(body, "plain", "utf-8")
            msg["to"] = recipient
            msg["subject"] = subject
            raw = base64.urlsafe_b64encode(msg.as_bytes()).decode("utf-8")
            sent = service.users().messages().send(userId="me", body={"raw": raw}).execute()
            message_id = sent.get("id", "")
            log_event("EMAIL_SENT", f"Email sent to {recipient}; Gmail message ID={message_id}")
            return render_template("sent.html", recipient=recipient, subject=subject, message_id=message_id, sender=current_email())
        except Exception as exc:
            log_event("EMAIL_SEND_FAILED", f"Email to {recipient} failed: {exc}")
            return render_template("error.html", message=f"Email send failed: {exc}"), 500
    return render_template("communication.html", email=current_email(), recipient=default_recipient, subject="Virtual Tax Auditor Phase 2 Test", body="This is a controlled Phase 2 test communication from the Virtual Tax Auditor.")


@app.route("/audit")
@login_required
def audit_log():
    conn = db()
    events = conn.execute("SELECT * FROM audit_events ORDER BY id DESC LIMIT 200").fetchall()
    conn.close()
    return render_template("audit.html", events=events)


@app.route("/health")
def health():
    return {"status": "ok", "phase": 2}


@app.cli.command("reset-demo")
def reset_demo():
    if os.path.exists(DB_PATH):
        os.remove(DB_PATH)
    init_db()
    print(f"Reset Phase 2 demo database: {DB_PATH}")


init_db()

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.getenv("PORT", "5000")))
