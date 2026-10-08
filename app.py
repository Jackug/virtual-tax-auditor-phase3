import os, json, sqlite3, base64, re, csv, uuid
from pathlib import Path
from datetime import datetime, timezone
from functools import wraps
from email.mime.text import MIMEText
from werkzeug.utils import secure_filename
from flask import Flask, render_template, redirect, url_for, session, request, abort, flash, send_from_directory
from google_auth_oauthlib.flow import Flow
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from google.auth.transport.requests import Request

try:
    import pypdf
except Exception:
    pypdf = None

try:
    from docx import Document as DocxDocument
except Exception:
    DocxDocument = None

try:
    import pandas as pd
except Exception:
    pd = None

try:
    from openpyxl import load_workbook
    from openpyxl.styles import Alignment as OpenpyxlAlignment
except Exception:
    load_workbook = None
    OpenpyxlAlignment = None

try:
    import xlrd
except Exception:
    xlrd = None

BASE=Path(__file__).resolve().parent
DB=Path(os.getenv('DB_PATH', BASE/'vta_phase3.sqlite3'))
CLIENT=os.getenv('GOOGLE_CLIENT_SECRETS', str(BASE/'client_secret.json'))
BASE_URL=os.getenv('BASE_URL','http://localhost:5000').rstrip('/')
UPLOAD=Path(os.getenv('UPLOAD_DIR', BASE/'uploads')); UPLOAD.mkdir(exist_ok=True)
SCOPES=['https://www.googleapis.com/auth/gmail.send','https://www.googleapis.com/auth/userinfo.email','openid']
app=Flask(__name__); app.secret_key=os.getenv('FLASK_SECRET_KEY','change-me'); app.config.update(SESSION_COOKIE_SECURE=os.getenv('COOKIE_SECURE','0')=='1',SESSION_COOKIE_HTTPONLY=True,SESSION_COOKIE_SAMESITE='Lax',MAX_CONTENT_LENGTH=10*1024*1024)

def now(): return datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M:%S UTC')
@app.template_filter('fromjson')
def fromjson_filter(value):
    try: return json.loads(value) if value else {}
    except Exception: return {}
def db():
    c=sqlite3.connect(DB); c.row_factory=sqlite3.Row; c.execute('PRAGMA foreign_keys=ON'); return c

def getcase(c,cid):
    return c.execute(
        'SELECT ac.*,t.name taxpayer_name,t.tin,t.sector,t.financial_year '
        'FROM audit_cases ac JOIN taxpayers t ON t.id=ac.taxpayer_id WHERE ac.id=?',
        (cid,)
    ).fetchone()
def email(): return session.get('email')
def login_required(f):
    @wraps(f)
    def w(*a,**k): return f(*a,**k) if email() else redirect(url_for('login'))
    return w
def log(event,detail,case_ref=None,actor=None):
    c=db(); c.execute('INSERT INTO audit_events(actor_email,case_ref,event_type,detail,created_at) VALUES(?,?,?,?,?)',(actor or email(),case_ref,event,detail,now())); c.commit(); c.close()

def init_db():
    c=db(); c.executescript('''
CREATE TABLE IF NOT EXISTS oauth_credentials(email TEXT PRIMARY KEY,token TEXT NOT NULL,created_at TEXT,updated_at TEXT);
CREATE TABLE IF NOT EXISTS users(email TEXT PRIMARY KEY,role TEXT DEFAULT 'AUDITOR',active INTEGER DEFAULT 1,created_at TEXT);
CREATE TABLE IF NOT EXISTS taxpayers(id INTEGER PRIMARY KEY AUTOINCREMENT,tin TEXT UNIQUE,name TEXT,sector TEXT,station TEXT,financial_year TEXT,status TEXT DEFAULT 'Active');
CREATE TABLE IF NOT EXISTS taxpayer_metrics(id INTEGER PRIMARY KEY AUTOINCREMENT,taxpayer_id INTEGER,metric_name TEXT,declared_value REAL,observed_value REAL,source TEXT,period TEXT,FOREIGN KEY(taxpayer_id) REFERENCES taxpayers(id));
CREATE TABLE IF NOT EXISTS risk_rules(id INTEGER PRIMARY KEY AUTOINCREMENT,name TEXT,natural_language TEXT,structured_logic TEXT,category TEXT,approved INTEGER DEFAULT 1,created_at TEXT);
CREATE TABLE IF NOT EXISTS risk_assessments(id INTEGER PRIMARY KEY AUTOINCREMENT,taxpayer_id INTEGER,score REAL,band TEXT,exposure REAL,assessed_at TEXT,engine_version TEXT,FOREIGN KEY(taxpayer_id) REFERENCES taxpayers(id));
CREATE TABLE IF NOT EXISTS risk_drivers(id INTEGER PRIMARY KEY AUTOINCREMENT,assessment_id INTEGER,rule_id INTEGER,description TEXT,variance REAL,variance_pct REAL);
CREATE TABLE IF NOT EXISTS risk_rule_versions(id INTEGER PRIMARY KEY AUTOINCREMENT,rule_id INTEGER NOT NULL,version_number INTEGER NOT NULL,rule_name TEXT NOT NULL,risk_category TEXT NOT NULL,tax_type TEXT,taxpayer_scope TEXT,applicable_period TEXT,detection_logic_json TEXT NOT NULL DEFAULT '{}',legal_basis_json TEXT NOT NULL DEFAULT '[]',information_requests_json TEXT NOT NULL DEFAULT '[]',exposure_config_json TEXT NOT NULL DEFAULT '{}',ranking_config_json TEXT NOT NULL DEFAULT '{}',status TEXT NOT NULL DEFAULT 'Draft',created_by TEXT,created_at TEXT,approved_by TEXT,approved_at TEXT,FOREIGN KEY(rule_id) REFERENCES risk_rules(id) ON DELETE CASCADE,UNIQUE(rule_id,version_number));
CREATE TABLE IF NOT EXISTS risk_rule_legal_basis(id INTEGER PRIMARY KEY AUTOINCREMENT,rule_id INTEGER NOT NULL,version_number INTEGER NOT NULL,document_id INTEGER NOT NULL,document_version INTEGER NOT NULL,chunk_id INTEGER,legal_passage TEXT,created_at TEXT NOT NULL,FOREIGN KEY(rule_id) REFERENCES risk_rules(id) ON DELETE CASCADE,FOREIGN KEY(document_id) REFERENCES knowledge_documents(id) ON DELETE CASCADE);
CREATE TABLE IF NOT EXISTS risk_rule_information_requests(id INTEGER PRIMARY KEY AUTOINCREMENT,rule_id INTEGER NOT NULL,version_number INTEGER NOT NULL,request_title TEXT NOT NULL,purpose TEXT,period TEXT,mandatory INTEGER NOT NULL DEFAULT 1,evidence_type TEXT,sort_order INTEGER NOT NULL DEFAULT 1,created_at TEXT NOT NULL,FOREIGN KEY(rule_id) REFERENCES risk_rules(id) ON DELETE CASCADE);
CREATE TABLE IF NOT EXISTS risk_rule_runs(id INTEGER PRIMARY KEY AUTOINCREMENT,rule_id INTEGER NOT NULL,version_number INTEGER NOT NULL,run_status TEXT NOT NULL DEFAULT 'Pending',records_evaluated INTEGER DEFAULT 0,records_triggered INTEGER DEFAULT 0,total_exposure REAL DEFAULT 0,executed_at TEXT,executed_by TEXT);
CREATE INDEX IF NOT EXISTS idx_rule_versions_rule ON risk_rule_versions(rule_id,version_number);
CREATE INDEX IF NOT EXISTS idx_rule_legal_rule ON risk_rule_legal_basis(rule_id,version_number);
CREATE INDEX IF NOT EXISTS idx_rule_requests_rule ON risk_rule_information_requests(rule_id,version_number);

CREATE TABLE IF NOT EXISTS audit_cases(id INTEGER PRIMARY KEY AUTOINCREMENT,case_ref TEXT UNIQUE,taxpayer_id INTEGER,status TEXT,selected_by TEXT,selected_at TEXT,assigned_to TEXT,created_at TEXT,updated_at TEXT);
CREATE TABLE IF NOT EXISTS audit_analyses(id INTEGER PRIMARY KEY AUTOINCREMENT,case_id INTEGER,analysis_status TEXT,summary TEXT,methodology TEXT,limitations TEXT,created_at TEXT);
CREATE TABLE IF NOT EXISTS findings(id INTEGER PRIMARY KEY AUTOINCREMENT,finding_ref TEXT UNIQUE,case_id INTEGER,risk_category TEXT,description TEXT,financial_year TEXT,expected_value REAL,observed_value REAL,variance REAL,variance_pct REAL,potential_exposure REAL,source TEXT,audit_test TEXT,explanation TEXT,ai_confidence TEXT,limitations TEXT,status TEXT DEFAULT 'AI Generated',created_at TEXT);
CREATE TABLE IF NOT EXISTS finding_decisions(id INTEGER PRIMARY KEY AUTOINCREMENT,finding_id INTEGER,decision TEXT,reason TEXT,decided_by TEXT,decided_at TEXT);
CREATE TABLE IF NOT EXISTS communications(id INTEGER PRIMARY KEY AUTOINCREMENT,case_id INTEGER,finding_id INTEGER,recipient TEXT,subject TEXT,ai_draft TEXT,human_body TEXT,status TEXT DEFAULT 'Draft',approved_by TEXT,approved_at TEXT,sent_at TEXT,gmail_message_id TEXT,created_at TEXT);
CREATE TABLE IF NOT EXISTS taxpayer_responses(id INTEGER PRIMARY KEY AUTOINCREMENT,case_id INTEGER,response_text TEXT,submitted_at TEXT,submitted_by TEXT,status TEXT DEFAULT 'Received');
CREATE TABLE IF NOT EXISTS evidence(id INTEGER PRIMARY KEY AUTOINCREMENT,case_id INTEGER,finding_id INTEGER,response_id INTEGER,original_filename TEXT,stored_filename TEXT,mime_type TEXT,size_bytes INTEGER,uploaded_at TEXT,uploaded_by TEXT);
CREATE TABLE IF NOT EXISTS response_analyses(id INTEGER PRIMARY KEY AUTOINCREMENT,case_id INTEGER,response_id INTEGER,analysis_status TEXT,overall_assessment TEXT,findings_supported TEXT,contradictions TEXT,missing_evidence TEXT,recommended_action TEXT,limitations TEXT,created_at TEXT);
CREATE TABLE IF NOT EXISTS second_validations(id INTEGER PRIMARY KEY AUTOINCREMENT,case_id INTEGER,finding_id INTEGER,decision TEXT,reason TEXT,decided_by TEXT,decided_at TEXT);
CREATE TABLE IF NOT EXISTS further_actions(id INTEGER PRIMARY KEY AUTOINCREMENT,case_id INTEGER,action_type TEXT,instruction TEXT,status TEXT DEFAULT 'Open',created_by TEXT,created_at TEXT,completed_at TEXT);
CREATE TABLE IF NOT EXISTS outcomes(id INTEGER PRIMARY KEY AUTOINCREMENT,case_id INTEGER UNIQUE,outcome_type TEXT,rationale TEXT,decided_by TEXT,decided_at TEXT,closed_at TEXT);
CREATE TABLE IF NOT EXISTS audit_events(id INTEGER PRIMARY KEY AUTOINCREMENT,actor_email TEXT,case_ref TEXT,event_type TEXT,detail TEXT,created_at TEXT);
CREATE TABLE IF NOT EXISTS knowledge_documents(
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    document_ref TEXT UNIQUE NOT NULL,
    title TEXT NOT NULL,
    category TEXT NOT NULL,
    document_type TEXT NOT NULL,
    tax_type TEXT,
    issuing_authority TEXT,
    description TEXT,
    effective_date TEXT,
    expiry_date TEXT,
    current_version INTEGER DEFAULT 1,
    status TEXT DEFAULT 'Draft',
    created_by TEXT,
    created_at TEXT,
    updated_at TEXT
);
CREATE TABLE IF NOT EXISTS knowledge_document_versions(
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    document_id INTEGER NOT NULL,
    version_number INTEGER NOT NULL,
    original_filename TEXT NOT NULL,
    stored_filename TEXT UNIQUE NOT NULL,
    mime_type TEXT,
    size_bytes INTEGER,
    file_path TEXT NOT NULL,
    uploaded_by TEXT,
    uploaded_at TEXT,
    version_status TEXT DEFAULT 'Draft',
    change_summary TEXT,
    FOREIGN KEY(document_id) REFERENCES knowledge_documents(id) ON DELETE CASCADE,
    UNIQUE(document_id, version_number)
);
CREATE TABLE IF NOT EXISTS knowledge_reviews(
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    document_id INTEGER NOT NULL,
    version_number INTEGER NOT NULL,
    decision TEXT NOT NULL,
    comments TEXT,
    reviewed_by TEXT NOT NULL,
    reviewed_at TEXT NOT NULL,
    FOREIGN KEY(document_id) REFERENCES knowledge_documents(id) ON DELETE CASCADE
);
CREATE TABLE IF NOT EXISTS knowledge_chunks(
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    document_id INTEGER NOT NULL,
    version_number INTEGER NOT NULL,
    chunk_number INTEGER NOT NULL,
    page_number INTEGER,
    source_filename TEXT NOT NULL,
    text_content TEXT NOT NULL,
    character_count INTEGER NOT NULL,
    extracted_at TEXT NOT NULL,
    FOREIGN KEY(document_id) REFERENCES knowledge_documents(id) ON DELETE CASCADE,
    UNIQUE(document_id, version_number, chunk_number)
);
CREATE INDEX IF NOT EXISTS idx_knowledge_chunks_doc_version
    ON knowledge_chunks(document_id, version_number);
CREATE TABLE IF NOT EXISTS knowledge_retrievals(
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    case_id INTEGER,
    analysis_id INTEGER,
    chunk_id INTEGER NOT NULL,
    query_text TEXT NOT NULL,
    rank_order INTEGER NOT NULL,
    retrieved_at TEXT NOT NULL,
    FOREIGN KEY(case_id) REFERENCES audit_cases(id) ON DELETE SET NULL,
    FOREIGN KEY(analysis_id) REFERENCES audit_analyses(id) ON DELETE SET NULL,
    FOREIGN KEY(chunk_id) REFERENCES knowledge_chunks(id) ON DELETE CASCADE
);
CREATE TABLE IF NOT EXISTS knowledge_chunk_staging(
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    batch_token TEXT NOT NULL,
    document_id INTEGER NOT NULL,
    version_number INTEGER NOT NULL,
    chunk_number INTEGER NOT NULL,
    page_number INTEGER,
    source_filename TEXT NOT NULL,
    text_content TEXT NOT NULL,
    character_count INTEGER NOT NULL,
    extracted_at TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_kb_staging_batch
    ON knowledge_chunk_staging(batch_token);
CREATE TABLE IF NOT EXISTS data_sources(
    id INTEGER PRIMARY KEY AUTOINCREMENT, source_ref TEXT UNIQUE NOT NULL, source_name TEXT NOT NULL,
    source_type TEXT NOT NULL, description TEXT, owner_department TEXT, reporting_period TEXT, tax_type TEXT,
    current_version INTEGER DEFAULT 1, status TEXT DEFAULT 'Draft', created_by TEXT, created_at TEXT, updated_at TEXT
);
CREATE TABLE IF NOT EXISTS data_source_versions(
    id INTEGER PRIMARY KEY AUTOINCREMENT, source_id INTEGER NOT NULL, version_number INTEGER NOT NULL,
    original_filename TEXT NOT NULL, stored_filename TEXT UNIQUE NOT NULL, mime_type TEXT, size_bytes INTEGER,
    file_path TEXT NOT NULL, uploaded_by TEXT, uploaded_at TEXT, version_status TEXT DEFAULT 'Draft',
    record_count INTEGER DEFAULT 0, column_count INTEGER DEFAULT 0, columns_json TEXT,
    validation_status TEXT DEFAULT 'Pending', validation_summary TEXT, validation_errors INTEGER DEFAULT 0,
    validation_warnings INTEGER DEFAULT 0, change_summary TEXT, approved_by TEXT, approved_at TEXT,
    FOREIGN KEY(source_id) REFERENCES data_sources(id) ON DELETE CASCADE, UNIQUE(source_id, version_number)
);
CREATE TABLE IF NOT EXISTS data_source_reviews(
    id INTEGER PRIMARY KEY AUTOINCREMENT, source_id INTEGER NOT NULL, version_number INTEGER NOT NULL,
    decision TEXT NOT NULL, comments TEXT, reviewed_by TEXT NOT NULL, reviewed_at TEXT NOT NULL,
    FOREIGN KEY(source_id) REFERENCES data_sources(id) ON DELETE CASCADE
);
CREATE TABLE IF NOT EXISTS data_source_issues(
    id INTEGER PRIMARY KEY AUTOINCREMENT, source_id INTEGER NOT NULL, version_number INTEGER NOT NULL,
    issue_type TEXT NOT NULL, severity TEXT NOT NULL, column_name TEXT, row_reference TEXT,
    description TEXT NOT NULL, created_at TEXT NOT NULL,
    FOREIGN KEY(source_id) REFERENCES data_sources(id) ON DELETE CASCADE
);
CREATE INDEX IF NOT EXISTS idx_data_source_versions_status ON data_source_versions(version_status);
CREATE INDEX IF NOT EXISTS idx_data_source_issues_source_version ON data_source_issues(source_id, version_number);
CREATE TABLE IF NOT EXISTS data_transform_recipes(
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    source_id INTEGER NOT NULL,
    version_number INTEGER NOT NULL,
    recipe_name TEXT,
    steps_json TEXT NOT NULL DEFAULT '[]',
    status TEXT NOT NULL DEFAULT 'Draft',
    created_by TEXT,
    created_at TEXT,
    updated_at TEXT,
    FOREIGN KEY(source_id) REFERENCES data_sources(id) ON DELETE CASCADE
);
CREATE TABLE IF NOT EXISTS data_transform_runs(
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    recipe_id INTEGER NOT NULL,
    source_id INTEGER NOT NULL,
    version_number INTEGER NOT NULL,
    output_filename TEXT,
    output_path TEXT,
    output_rows INTEGER,
    output_columns INTEGER,
    transformation_count INTEGER DEFAULT 0,
    status TEXT NOT NULL DEFAULT 'Preview',
    created_by TEXT,
    created_at TEXT,
    FOREIGN KEY(recipe_id) REFERENCES data_transform_recipes(id) ON DELETE CASCADE,
    FOREIGN KEY(source_id) REFERENCES data_sources(id) ON DELETE CASCADE
);
CREATE INDEX IF NOT EXISTS idx_transform_recipe_source ON data_transform_recipes(source_id, version_number);
CREATE INDEX IF NOT EXISTS idx_transform_run_source ON data_transform_runs(source_id, version_number);
CREATE TABLE IF NOT EXISTS risk_rule_run_results(
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    run_id INTEGER NOT NULL,
    rule_id INTEGER NOT NULL,
    version_number INTEGER NOT NULL,
    taxpayer_tin TEXT,
    exposure REAL DEFAULT 0,
    result_json TEXT NOT NULL DEFAULT '{}',
    created_at TEXT NOT NULL,
    FOREIGN KEY(run_id) REFERENCES risk_rule_runs(id) ON DELETE CASCADE,
    FOREIGN KEY(rule_id) REFERENCES risk_rules(id) ON DELETE CASCADE
);
CREATE INDEX IF NOT EXISTS idx_rule_run_results_tin ON risk_rule_run_results(taxpayer_tin);
''')
    # Step 3A migration: extend the existing risk_rules table without resetting data.
    existing_cols={row['name'] for row in c.execute('PRAGMA table_info(risk_rules)').fetchall()}
    additions={'rule_ref':'TEXT','version_number':'INTEGER DEFAULT 1','risk_category':'TEXT','tax_type':'TEXT','taxpayer_scope':'TEXT','applicable_period':'TEXT','detection_logic_json':"TEXT DEFAULT '{}'",'legal_basis_json':"TEXT DEFAULT '[]'",'information_requests_json':"TEXT DEFAULT '[]'",'exposure_config_json':"TEXT DEFAULT '{}'",'ranking_config_json':"TEXT DEFAULT '{}'",'status':"TEXT DEFAULT 'Draft'",'created_by':'TEXT','updated_at':'TEXT','approved_by':'TEXT','approved_at':'TEXT'}
    for col,definition in additions.items():
        if col not in existing_cols: c.execute(f'ALTER TABLE risk_rules ADD COLUMN {col} {definition}')
    c.execute("UPDATE risk_rules SET risk_category=COALESCE(NULLIF(risk_category,''),category), status=CASE WHEN approved=1 THEN 'Active' ELSE COALESCE(status,'Draft') END, version_number=COALESCE(version_number,1), updated_at=COALESCE(updated_at,created_at)")
    # Data transformation / risk-engine tables are additive migrations; existing data is retained.
    c.executescript('''
        CREATE TABLE IF NOT EXISTS data_transform_recipes(
            id INTEGER PRIMARY KEY AUTOINCREMENT, source_id INTEGER NOT NULL, version_number INTEGER NOT NULL,
            recipe_name TEXT, steps_json TEXT NOT NULL DEFAULT '[]', status TEXT NOT NULL DEFAULT 'Draft',
            created_by TEXT, created_at TEXT, updated_at TEXT,
            FOREIGN KEY(source_id) REFERENCES data_sources(id) ON DELETE CASCADE
        );
        CREATE TABLE IF NOT EXISTS data_transform_runs(
            id INTEGER PRIMARY KEY AUTOINCREMENT, recipe_id INTEGER NOT NULL, source_id INTEGER NOT NULL, version_number INTEGER NOT NULL,
            output_filename TEXT, output_path TEXT, output_rows INTEGER, output_columns INTEGER,
            transformation_count INTEGER DEFAULT 0, status TEXT NOT NULL DEFAULT 'Preview', created_by TEXT, created_at TEXT,
            FOREIGN KEY(recipe_id) REFERENCES data_transform_recipes(id) ON DELETE CASCADE,
            FOREIGN KEY(source_id) REFERENCES data_sources(id) ON DELETE CASCADE
        );
        CREATE INDEX IF NOT EXISTS idx_transform_recipe_source ON data_transform_recipes(source_id, version_number);
        CREATE INDEX IF NOT EXISTS idx_transform_run_source ON data_transform_runs(source_id, version_number);
        CREATE TABLE IF NOT EXISTS risk_rule_run_results(
            id INTEGER PRIMARY KEY AUTOINCREMENT, run_id INTEGER NOT NULL, rule_id INTEGER NOT NULL, version_number INTEGER NOT NULL,
            taxpayer_tin TEXT, exposure REAL DEFAULT 0, result_json TEXT NOT NULL DEFAULT '{}', created_at TEXT NOT NULL,
            FOREIGN KEY(run_id) REFERENCES risk_rule_runs(id) ON DELETE CASCADE,
            FOREIGN KEY(rule_id) REFERENCES risk_rules(id) ON DELETE CASCADE
        );
        CREATE INDEX IF NOT EXISTS idx_rule_run_results_tin ON risk_rule_run_results(taxpayer_tin);
    ''')

    # Step 3B migration: execution-plan governance and preview results.
    rule_version_cols={row['name'] for row in c.execute('PRAGMA table_info(risk_rule_versions)').fetchall()}
    rule_version_additions={'execution_plan_json':"TEXT DEFAULT '{}'",'generated_natural_language':'TEXT','preview_summary_json':"TEXT DEFAULT '{}'",'preview_status':"TEXT DEFAULT 'Not Run'",'preview_run_at':'TEXT','preview_approved':'INTEGER DEFAULT 0'}
    for col,definition in rule_version_additions.items():
        if col not in rule_version_cols:c.execute(f'ALTER TABLE risk_rule_versions ADD COLUMN {col} {definition}')
    c.executescript("""CREATE TABLE IF NOT EXISTS risk_rule_preview_results(
        id INTEGER PRIMARY KEY AUTOINCREMENT,rule_id INTEGER NOT NULL,version_number INTEGER NOT NULL,
        sample_rank INTEGER NOT NULL,result_json TEXT NOT NULL,created_at TEXT NOT NULL,
        FOREIGN KEY(rule_id) REFERENCES risk_rules(id) ON DELETE CASCADE,UNIQUE(rule_id,version_number,sample_rank));
        CREATE INDEX IF NOT EXISTS idx_rule_preview_rule ON risk_rule_preview_results(rule_id,version_number);""")
    seed(c); c.commit(); c.close()

def seed(c):
    if c.execute('SELECT COUNT(*) FROM taxpayers').fetchone()[0]: return
    c.execute("INSERT INTO taxpayers(tin,name,sector,station,financial_year) VALUES(?,?,?,?,?)",('1000000001','ABC Traders Ltd','Wholesale & Retail Trade','Kampala','2025/26')); tid=c.execute('SELECT last_insert_rowid()').fetchone()[0]
    c.executemany('INSERT INTO taxpayer_metrics(taxpayer_id,metric_name,declared_value,observed_value,source,period) VALUES(?,?,?,?,?,?)',[(tid,'Sales',1e9,1.5e9,'Third-party / observed sales','2025/26'),(tid,'Purchases',7e8,7e8,'Declared purchases','2025/26'),(tid,'Imports',7e8,8e8,'Customs / import data','2025/26'),(tid,'PAYE',5e7,5.2e7,'PAYE / employment data','2025/26')])
    c.executemany('INSERT INTO risk_rules(name,natural_language,structured_logic,category,created_at) VALUES(?,?,?,?,?)',[('Observed Sales Reconciliation','I want taxpayers where observed sales are higher than declared sales','observed_sales > declared_sales','Sales reconciliation',now()),('Imports versus Purchases','I want taxpayers where imports are higher than declared purchases','imports > declared_purchases','Imports / purchases reconciliation',now())])

def oauth_flow(state=None,verifier=None):
    if not os.path.exists(CLIENT): raise RuntimeError('Google client secret not found: '+CLIENT)
    f=Flow.from_client_secrets_file(CLIENT,scopes=SCOPES,state=state,code_verifier=verifier); f.redirect_uri=BASE_URL+'/oauth2callback'; return f
def save_creds(e,creds):
    payload=json.dumps({'token':creds.token,'refresh_token':creds.refresh_token,'token_uri':creds.token_uri,'client_id':creds.client_id,'client_secret':creds.client_secret,'scopes':creds.scopes}); c=db(); c.execute("INSERT INTO oauth_credentials VALUES(?,?,?,?) ON CONFLICT(email) DO UPDATE SET token=excluded.token,updated_at=excluded.updated_at",(e,payload,now(),now())); c.execute("INSERT OR IGNORE INTO users(email,created_at) VALUES(?,?)",(e,now())); c.commit(); c.close()
def get_creds(e):
    if not e:
        return None
    c = db()
    r = c.execute(
        'SELECT token FROM oauth_credentials WHERE email=?',
        (e,)
    ).fetchone()
    c.close()
    if not r:
        return None
    try:
        return Credentials.from_authorized_user_info(
            json.loads(r['token']),
            SCOPES
        )
    except Exception:
        # Do not allow a corrupt/stale credential record to crash the
        # communication workflow.
        try:
            c = db()
            c.execute('DELETE FROM oauth_credentials WHERE email=?', (e,))
            c.commit()
            c.close()
        except Exception:
            pass
        return None

def metrics(c,tid): return {r['metric_name'].lower():r for r in c.execute('SELECT * FROM taxpayer_metrics WHERE taxpayer_id=?',(tid,)).fetchall()}
def _risk_band(score):
    score=float(score or 0)
    return 'Critical' if score>=90 else ('High' if score>=70 else ('Medium' if score>=40 else 'Low'))

def _normalise_header(value):
    return re.sub(r'[^a-z0-9]+','',str(value or '').lower())

def _find_tin_column(df):
    candidates={'tin','taxpayertin','taxpayeridentificationnumber','taxpayeridentificationno','taxpayeridentificationnumbertin'}
    for col in df.columns:
        if _normalise_header(col) in candidates:
            return col
    return None

def _rule_exposure_from_row(row, plan):
    exposure=plan.get('exposure',{}) or {}
    field=exposure.get('field')
    if field and field in row:
        value=_parse_number(row.get(field))
        if value is not None:return value
    minimum=_parse_number(exposure.get('minimum'))
    return minimum or 0.0

def _execute_approved_rule_for_engine(c, rule, version, plan):
    summary=_execute_rule_plan(c,plan,sample_only=5,return_records=True)
    run_stamp=now()
    c.execute('INSERT INTO risk_rule_runs(rule_id,version_number,run_status,records_evaluated,records_triggered,total_exposure,executed_at,executed_by) VALUES(?,?,?,?,?,?,?,?)',
              (rule['id'],version['version_number'],'Completed',summary['records_evaluated'],summary['records_triggered'],summary['total_exposure'],run_stamp,email()))
    run_id=c.execute('SELECT last_insert_rowid()').fetchone()[0]
    result=summary.get('_result_frame')
    source_rows=summary.get('_source_rows',{})
    if result is None or len(result)==0:
        return {'run_id':run_id,'triggered_tins':set(),'summary':summary}
    # Identify a TIN from any source participating in the rule.
    tin_series=None
    for sid,srcrow in source_rows.items():
        dfcols=summary.get('_loaded_original_columns',{}).get(int(sid),[])
        tin_col=None
        for col in dfcols:
            if _normalise_header(col) in {'tin','taxpayertin','taxpayeridentificationnumber','taxpayeridentificationno','taxpayeridentificationnumbertin'}:
                tin_col=col; break
        if tin_col:
            key=_field_key(int(sid),tin_col)
            if key in result.columns:
                tin_series=result[key]
                break
    triggered_tins=set()
    if tin_series is not None:
        exposure_series=result.get('__exposure__')
        for idx in result.index:
            tin=_clean_value(tin_series.loc[idx])
            if tin is None or str(tin).strip()=='':
                continue
            tin=str(tin).strip()
            exposure=_clean_value(exposure_series.loc[idx]) if exposure_series is not None else 0
            try: exposure=float(exposure or 0)
            except Exception: exposure=0.0
            obj={}
            # Store a compact auditable row from preview fields / result columns.
            for col in result.columns[:30]:
                if col=='__exposure__': continue
                try: obj[col]=_clean_value(result.loc[idx,col])
                except Exception: pass
            c.execute('INSERT INTO risk_rule_run_results(run_id,rule_id,version_number,taxpayer_tin,exposure,result_json,created_at) VALUES(?,?,?,?,?,?,?)',
                      (run_id,rule['id'],version['version_number'],tin,exposure,json.dumps(obj,default=str,ensure_ascii=False),run_stamp))
            triggered_tins.add(tin)
    return {'run_id':run_id,'triggered_tins':triggered_tins,'summary':summary}

def _refresh_risk_universe_from_active_rules(c):
    active=c.execute('''SELECT r.*,v.* FROM risk_rules r JOIN risk_rule_versions v
                        ON v.rule_id=r.id AND v.version_number=r.version_number
                        WHERE r.status IN ('Active','Approved') AND v.status IN ('Active','Approved')
                          AND v.execution_plan_json IS NOT NULL AND v.execution_plan_json!='{}' ''').fetchall()
    if not active:
        return 0
    # Remove only assessments generated by the generic rule engine; legacy/manual assessments are retained
    # when no generic rules are active. For an active generic rule universe, rebuild the current assessment set.
    c.execute('DELETE FROM risk_drivers')
    c.execute('DELETE FROM risk_assessments')
    taxpayer_rows={str(r['tin']).strip():r for r in c.execute('SELECT * FROM taxpayers WHERE tin IS NOT NULL').fetchall()}
    triggered={}
    for rule in active:
        plan=_json_or_default(rule['execution_plan_json'],{})
        if not plan: continue
        try:
            result=_execute_approved_rule_for_engine(c,rule,rule,plan)
        except Exception as exc:
            c.execute("INSERT INTO risk_rule_runs(rule_id,version_number,run_status,records_evaluated,records_triggered,total_exposure,executed_at,executed_by) VALUES(?,?,?,?,?,?,?,?)",
                      (rule['id'],rule['version_number'],'Failed',0,0,0,now(),email()))
            log('RISK_ENGINE_RULE_FAILED',f"{rule['rule_ref']} v{rule['version_number']}: {type(exc).__name__}: {exc}")
            continue
        for tin in result['triggered_tins']:
            triggered.setdefault(tin,[]).append((rule,result))
    count=0
    for tin,items in triggered.items():
        taxpayer=taxpayer_rows.get(tin)
        if not taxpayer: continue
        exposures=[]
        for rule,run_result in items:
            exp=0.0
            run_id=run_result['run_id']
            row=c.execute('SELECT COALESCE(SUM(exposure),0) FROM risk_rule_run_results WHERE run_id=? AND taxpayer_tin=?',(run_id,tin)).fetchone()
            exp=float(row[0] or 0)
            pct=0.0
            desc=rule['name']
            exposures.append((rule,exp,pct,desc))
        score=min(100.0,25.0*len(exposures)+sum(10.0 for x in exposures if x[1]>0))
        score=min(100.0,score)
        band=_risk_band(score)
        exposure=sum(x[1] for x in exposures)
        c.execute('INSERT INTO risk_assessments(taxpayer_id,score,band,exposure,assessed_at,engine_version) VALUES(?,?,?,?,?,?)',(taxpayer['id'],score,band,exposure,now(),'Phase-3-Risk-Engine-Rule-Library-1.0'))
        aid=c.execute('SELECT last_insert_rowid()').fetchone()[0]
        for rule,exp,pct,desc in exposures:
            c.execute('INSERT INTO risk_drivers(assessment_id,rule_id,description,variance,variance_pct) VALUES(?,?,?,?,?)',(aid,rule['id'],f"{desc}: approved rule triggered for TIN {tin}; estimated exposure UGX {exp:,.0f}.",exp,pct))
        count+=1
    return count

def assess(tid):
    """Assess a taxpayer using only the approved executable Risk Rule Library.

    Synthetic/demo fallback assessment has been removed. A taxpayer can only enter
    the Risk Universe from approved rules executed against approved data-source versions.
    """
    c=db()
    t=c.execute('SELECT * FROM taxpayers WHERE id=?',(tid,)).fetchone()
    if not t:
        c.close(); raise ValueError('Taxpayer not found.')
    active=c.execute("""SELECT r.*,v.execution_plan_json FROM risk_rules r JOIN risk_rule_versions v
                        ON v.rule_id=r.id AND v.version_number=r.version_number
                        WHERE r.status IN ('Active','Approved') AND v.status IN ('Active','Approved')
                          AND v.execution_plan_json IS NOT NULL AND v.execution_plan_json!='{}' """).fetchall()
    if not active:
        c.close(); raise ValueError('No Active/Approved executable risk rules are available. Approve a rule before running taxpayer risk assessment.')
    _refresh_risk_universe_from_active_rules(c)
    c.commit()
    c.close()
    log('RISK_ASSESSMENT_RUN',f'TIN {t["tin"]}; assessment rebuilt from approved executable rule library.')
    return True


def case_for(tid):
    c=db(); r=c.execute("SELECT * FROM audit_cases WHERE taxpayer_id=? AND status NOT IN('Closed','Rejected') ORDER BY id DESC LIMIT 1",(tid,)).fetchone()
    if r: c.close(); return r['id']
    t=c.execute('SELECT * FROM taxpayers WHERE id=?',(tid,)).fetchone()
    if not t:
        c.close(); raise ValueError('Taxpayer not found.')
    ref='VTA-'+datetime.now(timezone.utc).strftime('%Y%m%d%H%M%S%f')
    c.execute('INSERT INTO audit_cases(case_ref,taxpayer_id,status,selected_by,selected_at,assigned_to,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?)',(ref,tid,'Selected for Virtual Audit',email(),now(),email(),now(),now()))
    cid=c.execute('SELECT last_insert_rowid()').fetchone()[0]
    c.commit(); c.close(); log('CASE_SELECTED',f'TIN {t["tin"]} selected for Virtual Audit',ref); return cid

def normalize_kb_text(text):
    text = text or ''
    text = text.replace('\x00', ' ')
    text = re.sub(r'\s+', ' ', text)
    return text.strip()


def chunk_knowledge_text(text, chunk_size=1800, overlap=250):
    text = normalize_kb_text(text)
    if not text:
        return []
    if len(text) <= chunk_size:
        return [text]
    chunks = []
    start = 0
    length = len(text)
    while start < length:
        end = min(start + chunk_size, length)
        if end < length:
            boundary = max(text.rfind('. ', start, end), text.rfind('; ', start, end), text.rfind(' ', start, end))
            if boundary > start + int(chunk_size * 0.60):
                end = boundary + 1
        part = text[start:end].strip()
        if part:
            chunks.append(part)
        if end >= length:
            break
        start = max(end - overlap, start + 1)
    return chunks


def iter_knowledge_document(file_path, original_filename):
    'Yield (page_or_sheet_number, text) incrementally without building the document in RAM.'
    path = Path(file_path)
    ext = path.suffix.lower().lstrip('.')

    if ext == 'pdf':
        if pypdf is None:
            raise RuntimeError('PDF extraction requires pypdf. Add pypdf to requirements.txt and redeploy.')
        reader = pypdf.PdfReader(str(path))
        for number, page in enumerate(reader.pages, start=1):
            text = normalize_kb_text(page.extract_text() or '')
            if text:
                yield number, text
        return

    if ext == 'docx':
        if DocxDocument is None:
            raise RuntimeError('DOCX extraction requires python-docx. Add python-docx to requirements.txt and redeploy.')
        doc = DocxDocument(str(path))
        buffer = []
        buffer_chars = 0
        logical_page = 1

        def flush_buffer():
            nonlocal buffer, buffer_chars, logical_page
            if not buffer:
                return None
            text = normalize_kb_text('\n'.join(buffer))
            buffer = []
            buffer_chars = 0
            if not text:
                return None
            result = (logical_page, text)
            logical_page += 1
            return result

        for paragraph in doc.paragraphs:
            value = (paragraph.text or '').strip()
            if not value:
                continue
            buffer.append(value)
            buffer_chars += len(value)
            if buffer_chars >= 12000:
                result = flush_buffer()
                if result:
                    yield result

        for table in doc.tables:
            for row in table.rows:
                values = [(cell.text or '').strip() for cell in row.cells]
                values = [v for v in values if v]
                if not values:
                    continue
                line = ' | '.join(values)
                buffer.append(line)
                buffer_chars += len(line)
                if buffer_chars >= 12000:
                    result = flush_buffer()
                    if result:
                        yield result

        result = flush_buffer()
        if result:
            yield result
        return

    if ext == 'txt':
        logical_page = 1
        buffer = []
        chars = 0
        with path.open('r', encoding='utf-8', errors='replace') as fh:
            for line in fh:
                line = line.strip()
                if not line:
                    continue
                buffer.append(line)
                chars += len(line)
                if chars >= 12000:
                    text = normalize_kb_text('\n'.join(buffer))
                    if text:
                        yield logical_page, text
                        logical_page += 1
                    buffer = []
                    chars = 0
        if buffer:
            text = normalize_kb_text('\n'.join(buffer))
            if text:
                yield logical_page, text
        return

    if ext == 'csv':
        logical_page = 1
        buffer = []
        chars = 0
        with path.open('r', encoding='utf-8-sig', errors='replace', newline='') as fh:
            reader = csv.reader(fh)
            for row in reader:
                values = [str(v).strip() for v in row if str(v).strip()]
                if not values:
                    continue
                line = ' | '.join(values)
                buffer.append(line)
                chars += len(line)
                if chars >= 12000:
                    text = normalize_kb_text('\n'.join(buffer))
                    if text:
                        yield logical_page, text
                        logical_page += 1
                    buffer = []
                    chars = 0
        if buffer:
            text = normalize_kb_text('\n'.join(buffer))
            if text:
                yield logical_page, text
        return

    if ext == 'xlsx':
        if load_workbook is None:
            raise RuntimeError('XLSX extraction requires openpyxl. Add openpyxl to requirements.txt and redeploy.')
        wb = load_workbook(filename=str(path), read_only=True, data_only=True)
        try:
            for number, ws in enumerate(wb.worksheets, start=1):
                buffer = [f'Sheet: {ws.title}']
                chars = len(buffer[0])
                for row in ws.iter_rows(values_only=True):
                    values = []
                    for value in row:
                        if value is None:
                            continue
                        value = str(value).strip()
                        if value:
                            values.append(value)
                    if not values:
                        continue
                    line = ' | '.join(values)
                    buffer.append(line)
                    chars += len(line)
                    if chars >= 12000:
                        text = normalize_kb_text('\n'.join(buffer))
                        if text:
                            yield number, text
                        buffer = [f'Sheet: {ws.title} (continued)']
                        chars = len(buffer[0])
                if len(buffer) > 1:
                    text = normalize_kb_text('\n'.join(buffer))
                    if text:
                        yield number, text
        finally:
            wb.close()
        return

    if ext == 'xls':
        if xlrd is None:
            raise RuntimeError('XLS extraction requires xlrd. Add xlrd to requirements.txt and redeploy.')
        book = xlrd.open_workbook(str(path), on_demand=True)
        try:
            for number, sheet in enumerate(book.sheets(), start=1):
                buffer = [f'Sheet: {sheet.name}']
                chars = len(buffer[0])
                for row_index in range(sheet.nrows):
                    values = []
                    for col_index in range(sheet.ncols):
                        value = sheet.cell_value(row_index, col_index)
                        if value is None:
                            continue
                        value = str(value).strip()
                        if value:
                            values.append(value)
                    if not values:
                        continue
                    line = ' | '.join(values)
                    buffer.append(line)
                    chars += len(line)
                    if chars >= 12000:
                        text = normalize_kb_text('\n'.join(buffer))
                        if text:
                            yield number, text
                        buffer = [f'Sheet: {sheet.name} (continued)']
                        chars = len(buffer[0])
                if len(buffer) > 1:
                    text = normalize_kb_text('\n'.join(buffer))
                    if text:
                        yield number, text
        finally:
            book.release_resources()
        return

    if ext == 'doc':
        raise RuntimeError('Legacy .doc files are accepted for registration but cannot be indexed automatically. Convert to .docx or PDF and upload a new version.')

    raise RuntimeError(f'Unsupported knowledge file type: {ext or "unknown"}')


def extract_knowledge_document(file_path, original_filename):
    'Compatibility wrapper. New indexing code uses the streaming iterator.'
    return list(iter_knowledge_document(file_path, original_filename))


def index_knowledge_version(document_id, version_number, batch_size=50):
    'Memory-safe KB indexing using streamed extraction and a staging table.'
    c = db()
    v = c.execute(
        'SELECT * FROM knowledge_document_versions WHERE document_id=? AND version_number=?',
        (document_id, version_number)
    ).fetchone()
    if not v:
        c.close()
        raise ValueError('Knowledge document version not found.')

    file_path = Path(v['file_path'])
    if not file_path.exists():
        c.close()
        raise FileNotFoundError(f'Knowledge file is missing: {file_path}')

    batch_token = uuid.uuid4().hex
    extracted_at = now()
    chunk_no = 1
    total = 0
    pending = []

    try:
        for page_number, page_text in iter_knowledge_document(file_path, v['original_filename']):
            for chunk in chunk_knowledge_text(page_text):
                pending.append((batch_token, document_id, version_number, chunk_no,
                                page_number, v['original_filename'], chunk, len(chunk), extracted_at))
                chunk_no += 1
                total += 1

                if len(pending) >= batch_size:
                    c.executemany(
                        '''INSERT INTO knowledge_chunk_staging
                           (batch_token,document_id,version_number,chunk_number,page_number,
                            source_filename,text_content,character_count,extracted_at)
                           VALUES(?,?,?,?,?,?,?,?,?)''',
                        pending
                    )
                    c.commit()
                    pending.clear()

        if pending:
            c.executemany(
                '''INSERT INTO knowledge_chunk_staging
                   (batch_token,document_id,version_number,chunk_number,page_number,
                    source_filename,text_content,character_count,extracted_at)
                   VALUES(?,?,?,?,?,?,?,?,?)''',
                pending
            )
            c.commit()
            pending.clear()

        if total == 0:
            raise ValueError('No readable text could be extracted from this knowledge document.')

        # Replace the live index only after complete extraction succeeds.
        c.execute('BEGIN')
        c.execute('DELETE FROM knowledge_chunks WHERE document_id=? AND version_number=?',
                  (document_id, version_number))
        c.execute('''
            INSERT INTO knowledge_chunks
            (document_id,version_number,chunk_number,page_number,source_filename,
             text_content,character_count,extracted_at)
            SELECT document_id,version_number,chunk_number,page_number,source_filename,
                   text_content,character_count,extracted_at
            FROM knowledge_chunk_staging
            WHERE batch_token=?
            ORDER BY chunk_number
        ''', (batch_token,))
        c.execute('DELETE FROM knowledge_chunk_staging WHERE batch_token=?', (batch_token,))
        c.commit()
        return total

    except Exception:
        c.rollback()
        try:
            c.execute('DELETE FROM knowledge_chunk_staging WHERE batch_token=?', (batch_token,))
            c.commit()
        except Exception:
            pass
        raise
    finally:
        c.close()


def search_approved_knowledge(query, limit=5, case_id=None, analysis_id=None):
    'Bounded keyword retrieval over APPROVED, already-indexed knowledge only.'
    query = normalize_kb_text(query)
    if not query:
        return []

    terms = [x.lower() for x in re.findall(r'[A-Za-z0-9]{3,}', query)]
    stop = {'the','and','for','with','from','where','this','that','tax','taxes',
            'value','values','period','year','declared','observed','audit','review'}
    terms = list(dict.fromkeys(t for t in terms if t not in stop))[:12]
    if not terms:
        return []

    c = db()
    like_clauses = []
    params = []
    for term in terms:
        pattern = f'%{term}%'
        like_clauses.append(
            '(LOWER(kc.text_content) LIKE ? OR LOWER(kd.title) LIKE ? '
            'OR LOWER(kd.category) LIKE ? OR LOWER(kd.document_type) LIKE ? '
            'OR LOWER(COALESCE(kd.tax_type,\'\')) LIKE ? '
            'OR LOWER(COALESCE(kd.issuing_authority,\'\')) LIKE ?)'
        )
        params.extend([pattern] * 6)

    sql = '''
        SELECT kc.*, kd.document_ref, kd.title, kd.category, kd.document_type,
               kd.tax_type, kd.issuing_authority, kv.version_status
        FROM knowledge_chunks kc
        JOIN knowledge_documents kd ON kd.id=kc.document_id
        JOIN knowledge_document_versions kv
          ON kv.document_id=kc.document_id AND kv.version_number=kc.version_number
        WHERE kv.version_status='Approved'
          AND (''' + ' OR '.join(like_clauses) + ''')
        ORDER BY kc.id DESC
        LIMIT ?
    '''
    params.append(max(100, min(500, limit * 40)))

    scored = []
    cursor = c.execute(sql, params)
    while True:
        batch = cursor.fetchmany(100)
        if not batch:
            break
        for row in batch:
            hay = ' '.join([
                row['title'] or '', row['category'] or '', row['document_type'] or '',
                row['tax_type'] or '', row['issuing_authority'] or '', row['text_content'] or ''
            ]).lower()
            score = sum(hay.count(term) for term in terms)
            if score > 0:
                scored.append((score, row))

    scored.sort(key=lambda x: (-x[0], -int(x[1]['id'])))
    selected = [row for _, row in scored[:max(1, min(limit, 20))]]

    if case_id is not None and selected:
        c.executemany(
            '''INSERT INTO knowledge_retrievals
               (case_id,analysis_id,chunk_id,query_text,rank_order,retrieved_at)
               VALUES(?,?,?,?,?,?)''',
            [(case_id, analysis_id, row['id'], query, rank, now())
             for rank, row in enumerate(selected, start=1)]
        )
        c.commit()
    c.close()
    return selected


def ensure_approved_knowledge_indexed():
    'Report missing indexes without indexing automatically.'
    c = db()
    rows = c.execute('''
        SELECT kv.document_id, kv.version_number, kd.document_ref, kd.title
        FROM knowledge_document_versions kv
        JOIN knowledge_documents kd ON kd.id=kv.document_id
        WHERE kv.version_status='Approved'
        ORDER BY kv.document_id, kv.version_number
    ''').fetchall()
    missing = []
    for row in rows:
        count = c.execute(
            'SELECT COUNT(*) FROM knowledge_chunks WHERE document_id=? AND version_number=?',
            (row['document_id'], row['version_number'])
        ).fetchone()[0]
        if count == 0:
            missing.append(dict(row))
    c.close()
    return missing

def knowledge_context_for_case(c, case_id):
    case = c.execute('SELECT * FROM audit_cases WHERE id=?', (case_id,)).fetchone()
    if not case:
        return [], ''
    taxpayer = c.execute('SELECT * FROM taxpayers WHERE id=?', (case['taxpayer_id'],)).fetchone()
    query = f"tax audit {taxpayer['sector']} {taxpayer['financial_year']} sales imports purchases compliance reconciliation"
    # Search uses its own connection so it can record retrievals cleanly.
    rows = search_approved_knowledge(query, limit=5, case_id=case_id)
    if not rows:
        return [], 'No approved Knowledge Base content matched this case. The analysis therefore did not rely on unapproved or general knowledge.'
    context = []
    for row in rows:
        context.append(
            f"[{row['document_ref']} v{row['version_number']} | {row['title']} | {row['category']} | page/sheet {row['page_number']}] {row['text_content']}"
        )
    return rows, '\n\n'.join(context)


def run_analysis(cid):
    c = db()
    case = c.execute('SELECT * FROM audit_cases WHERE id=?', (cid,)).fetchone()
    if not case:
        c.close()
        raise ValueError('Audit case not found.')
    t = c.execute('SELECT * FROM taxpayers WHERE id=?', (case['taxpayer_id'],)).fetchone()
    m = metrics(c, t['id'])
    s, i, p = m['sales'], m['imports'], m['purchases']

    kb_rows, kb_context = knowledge_context_for_case(c, cid)
    lim = ('Controlled deterministic Risk Rule Engine. It does not make a final tax decision. '
           'Knowledge retrieval is restricted to human-approved Knowledge Base versions.')
    methodology = (
        'Compared declared/observed sales and imports/purchases and calculated variances.\n\n'
        'Approved Knowledge Base context retrieved for this analysis:\n' + kb_context
    )
    c.execute('INSERT INTO audit_analyses(case_id,analysis_status,summary,methodology,limitations,created_at) VALUES(?,?,?,?,?,?)',
              (cid,'Completed','Two potential reconciliation issues identified: observed sales exceed declared sales and imports exceed declared purchases.',methodology,lim,now()))
    analysis_id = c.execute('SELECT last_insert_rowid()').fetchone()[0]
    # Retrievals were recorded before the analysis id existed; link the latest case retrievals now.
    c.execute('UPDATE knowledge_retrievals SET analysis_id=? WHERE case_id=? AND analysis_id IS NULL', (analysis_id, cid))

    c.execute('DELETE FROM findings WHERE case_id=?',(cid,))
    stamp=datetime.now(timezone.utc).strftime('%Y%m%d%H%M%S%f')
    fs=[
        (f'F-P3-{stamp}-01',cid,'Sales reconciliation','Observed sales exceed declared sales.',t['financial_year'],s['declared_value'],s['observed_value'],s['observed_value']-s['declared_value'],(s['observed_value']-s['declared_value'])/s['declared_value']*100,None,s['source'],'Observed Sales - Declared Sales','The difference requires taxpayer explanation and reconciliation.','Medium',lim,'AI Generated',now()),
        (f'F-P3-{stamp}-02',cid,'Imports / purchases reconciliation','Imports exceed declared purchases.',t['financial_year'],p['declared_value'],i['observed_value'],i['observed_value']-p['declared_value'],(i['observed_value']-p['declared_value'])/p['declared_value']*100,None,i['source'],'Imports - Declared Purchases','The difference may have timing, inventory, classification or other explanations.','Medium',lim,'AI Generated',now())
    ]
    c.executemany('INSERT INTO findings(finding_ref,case_id,risk_category,description,financial_year,expected_value,observed_value,variance,variance_pct,potential_exposure,source,audit_test,explanation,ai_confidence,limitations,status,created_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)',fs)
    c.execute("UPDATE audit_cases SET status='AI Analysis Completed',updated_at=? WHERE id=?",(now(),cid))
    c.commit()
    c.close()
    log('AI_ANALYSIS_COMPLETED',f'Controlled AI analysis completed; {len(kb_rows)} approved Knowledge Base chunk(s) retrieved; findings remain AI Generated.',case['case_ref'])


def ai_draft(f): return f'''Dear Taxpayer,\n\nDuring a review of your records for {f["financial_year"]}, a reconciliation issue requiring clarification was identified.\n\nIssue: {f["description"]}\nExpected/declared value: UGX {f["expected_value"]:,.0f}\nObserved value: UGX {f["observed_value"]:,.0f}\nVariance: UGX {f["variance"]:,.0f} ({f["variance_pct"]:.2f}%)\n\nPlease explain the difference and provide supporting records. This is a request for clarification and is not a final finding of non-compliance.\n\nRegards,\nVirtual Tax Auditor'''

@app.route('/')
def index(): return redirect(url_for('tasks' if email() else 'login'))
@app.route('/login')
def login(): return render_template('login.html')
@app.route('/authorize')
def authorize():
    try:
        f=oauth_flow(); u,state=f.authorization_url(access_type='offline',include_granted_scopes='true',prompt='consent'); session['oauth_state']=state; session['oauth_code_verifier']=f.code_verifier; return redirect(u)
    except Exception as e: return render_template('error.html',message=str(e)),500
@app.route('/oauth2callback')
def callback():
    try:
        state = session.get('oauth_state')
        verifier = session.get('oauth_code_verifier')
        f = oauth_flow(state, verifier)
        f.fetch_token(authorization_response=request.url)
        creds = f.credentials
        p = build('oauth2', 'v2', credentials=creds).userinfo().get().execute()
        e = p['email']
        save_creds(e, creds)
        session['email'] = e
        session.pop('oauth_state', None)
        session.pop('oauth_code_verifier', None)

        # If Gmail authorization was requested while sending a communication,
        # continue that send automatically after OAuth succeeds.
        pending_mid = session.pop('pending_send_mid', None)
        if pending_mid:
            log('GMAIL_REAUTHORIZED', 'Google/Gmail authorization completed; resuming pending communication send.', actor=e)
            return redirect(url_for('send_comm', mid=pending_mid))

        log('LOGIN', 'Google OAuth login successful', actor=e)
        return redirect(url_for('tasks'))
    except Exception as ex:
        session.pop('oauth_state', None)
        session.pop('oauth_code_verifier', None)
        return render_template('error.html', message='Google authorization failed: ' + str(ex)), 500
@app.route('/logout')
@login_required
def logout(): session.clear(); return redirect(url_for('login'))
@app.route('/tasks')
@login_required
def tasks():
    c=db()
    counts={k:c.execute(q).fetchone()[0] for k,q in {
        'taxpayers':'SELECT COUNT(*) FROM taxpayers',
        'high':"SELECT COUNT(*) FROM risk_assessments WHERE band IN('High','Critical')",
        'cases':"SELECT COUNT(*) FROM audit_cases WHERE status NOT IN('Closed','Rejected')",
        'findings':"SELECT COUNT(*) FROM findings WHERE status='AI Generated'",
        'responses':'SELECT COUNT(*) FROM taxpayer_responses',
        'knowledge_pending':"SELECT COUNT(*) FROM knowledge_documents d JOIN knowledge_document_versions v ON v.document_id=d.id AND v.version_number=d.current_version WHERE v.version_status='Draft'"
    }.items()}
    case=c.execute("SELECT * FROM audit_cases WHERE status NOT IN ('Closed','Rejected') ORDER BY id DESC LIMIT 1").fetchone()
    analysis_count_case=0
    validated_count=0
    sent_count=0
    response_count=0
    analysis_count=0
    validation_count=0
    latest_finding_id=None
    if case:
        analysis_count_case=c.execute("SELECT COUNT(*) FROM audit_analyses WHERE case_id=? AND analysis_status='Completed'",(case['id'],)).fetchone()[0]
        validated_count=c.execute("SELECT COUNT(*) FROM findings WHERE case_id=? AND status='Human Validated'",(case['id'],)).fetchone()[0]
        sent_count=c.execute("SELECT COUNT(*) FROM communications WHERE case_id=? AND status='Sent'",(case['id'],)).fetchone()[0]
        response_count=c.execute("SELECT COUNT(*) FROM taxpayer_responses WHERE case_id=?",(case['id'],)).fetchone()[0]
        analysis_count=c.execute("SELECT COUNT(*) FROM response_analyses WHERE case_id=?",(case['id'],)).fetchone()[0]
        validation_count=c.execute("SELECT COUNT(*) FROM second_validations WHERE case_id=?",(case['id'],)).fetchone()[0]
    latest_finding=c.execute("SELECT id FROM findings ORDER BY id DESC LIMIT 1").fetchone()
    latest_finding_id=latest_finding['id'] if latest_finding else None
    c.close()
    workflow={
        'case_id': case['id'] if case else None,
        'case_ref': case['case_ref'] if case else None,
        'analyzed': analysis_count_case > 0,
        'validated': validated_count > 0,
        'sent': sent_count > 0,
        'response': response_count > 0,
        'response_analysis': analysis_count > 0,
        'second_validation': validation_count > 0,
        'latest_finding_id': latest_finding_id,
    }
    return render_template('tasks.html',counts=counts,workflow=workflow)
ALLOWED_KB_EXTENSIONS = {'pdf', 'doc', 'docx', 'txt', 'csv', 'xls', 'xlsx'}
KB_CATEGORIES = ['Tax Laws', 'Regulations', 'Procedures', 'Audit Guidance', 'Sector Knowledge', 'Risk Knowledge']
KB_DOCUMENT_TYPES = ['Act', 'Regulation', 'Statutory Instrument', 'Procedure', 'Manual', 'Guideline', 'Circular', 'Directive', 'Sector Guide', 'Risk Note', 'Other']
KB_TAX_TYPES = ['General', 'VAT', 'CIT', 'PAYE', 'WHT', 'Excise Duty', 'Local Excise', 'Customs', 'Other']

def kb_allowed_file(filename):
    return bool(filename and '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_KB_EXTENSIONS)

def kb_ext(filename):
    return filename.rsplit('.', 1)[1].lower() if '.' in filename else ''

@app.route('/knowledge-base')
@login_required
def knowledge_base():
    c = db()
    documents = c.execute('''
        SELECT d.*, v.version_number, v.original_filename, v.size_bytes,
               v.uploaded_by, v.uploaded_at, v.version_status
        FROM knowledge_documents d
        LEFT JOIN knowledge_document_versions v
          ON v.document_id=d.id AND v.version_number=d.current_version
        ORDER BY d.id DESC
    ''').fetchall()
    c.close()
    return render_template(
        'knowledge_base.html',
        documents=documents,
        categories=KB_CATEGORIES,
        document_types=KB_DOCUMENT_TYPES,
        tax_types=KB_TAX_TYPES
    )

@app.route('/knowledge-base/register', methods=['POST'])
@login_required
def knowledge_base_register():
    title = request.form.get('title','').strip()
    category = request.form.get('category','').strip()
    document_type = request.form.get('document_type','').strip()
    tax_type = request.form.get('tax_type','').strip()
    authority = request.form.get('issuing_authority','').strip()
    description = request.form.get('description','').strip()
    effective_date = request.form.get('effective_date','').strip() or None
    expiry_date = request.form.get('expiry_date','').strip() or None
    change_summary = request.form.get('change_summary','').strip()
    file = request.files.get('document_file')

    if not title or category not in KB_CATEGORIES or document_type not in KB_DOCUMENT_TYPES:
        flash('Title, knowledge category and document type are required.', 'error')
        return redirect(url_for('knowledge_base'))
    if not file or not file.filename:
        flash('Please select a knowledge document to upload.', 'error')
        return redirect(url_for('knowledge_base'))
    if not kb_allowed_file(file.filename):
        flash('Unsupported file type. Allowed: PDF, DOC, DOCX, TXT, CSV, XLS and XLSX.', 'error')
        return redirect(url_for('knowledge_base'))

    safe_original = secure_filename(file.filename)
    if not safe_original:
        flash('The selected filename is not valid.', 'error')
        return redirect(url_for('knowledge_base'))

    stamp = datetime.now(timezone.utc).strftime('%Y%m%d%H%M%S%f')
    document_ref = 'KB-' + stamp
    stored_filename = document_ref + '_' + safe_original
    target = UPLOAD / 'knowledge_base'
    target.mkdir(parents=True, exist_ok=True)
    file_path = target / stored_filename
    file.save(file_path)

    c = db()
    try:
        c.execute('''
            INSERT INTO knowledge_documents
            (document_ref,title,category,document_type,tax_type,issuing_authority,description,
             effective_date,expiry_date,current_version,status,created_by,created_at,updated_at)
            VALUES(?,?,?,?,?,?,?,?,?,1,'Draft',?,?,?)
        ''', (document_ref,title,category,document_type,tax_type,authority,description,
              effective_date,expiry_date,email(),now(),now()))
        document_id = c.execute('SELECT last_insert_rowid()').fetchone()[0]
        c.execute('''
            INSERT INTO knowledge_document_versions
            (document_id,version_number,original_filename,stored_filename,mime_type,size_bytes,
             file_path,uploaded_by,uploaded_at,version_status,change_summary)
            VALUES(?,?,?,?,?,?,?,?,?,'Draft',?)
        ''', (document_id,1,safe_original,stored_filename,file.mimetype,file_path.stat().st_size,
              str(file_path),email(),now(),change_summary or 'Initial document registration.'))
        c.commit()
    except Exception:
        c.rollback()
        try:
            file_path.unlink(missing_ok=True)
        except Exception:
            pass
        c.close()
        raise
    c.close()
    log('KNOWLEDGE_DOCUMENT_REGISTERED', f'{document_ref}: {title}; version 1; status Draft')
    flash(f'Knowledge document {document_ref} registered successfully as version 1 (Draft).', 'success')
    return redirect(url_for('knowledge_base'))

@app.route('/knowledge-base/<int:document_id>/version', methods=['POST'])
@login_required
def knowledge_base_new_version(document_id):
    file = request.files.get('document_file')
    change_summary = request.form.get('change_summary','').strip()
    if not file or not file.filename:
        flash('Please select the replacement/version file.', 'error')
        return redirect(url_for('knowledge_base'))
    if not kb_allowed_file(file.filename):
        flash('Unsupported file type. Allowed: PDF, DOC, DOCX, TXT, CSV, XLS and XLSX.', 'error')
        return redirect(url_for('knowledge_base'))

    safe_original = secure_filename(file.filename)
    c = db()
    doc = c.execute('SELECT * FROM knowledge_documents WHERE id=?',(document_id,)).fetchone()
    if not doc:
        c.close()
        abort(404)
    next_version = int(doc['current_version']) + 1
    document_ref = doc['document_ref']
    stored_filename = f'{document_ref}_v{next_version}_{safe_original}'
    target = UPLOAD / 'knowledge_base'
    target.mkdir(parents=True, exist_ok=True)
    file_path = target / stored_filename
    file.save(file_path)
    try:
        c.execute('''
            INSERT INTO knowledge_document_versions
            (document_id,version_number,original_filename,stored_filename,mime_type,size_bytes,
             file_path,uploaded_by,uploaded_at,version_status,change_summary)
            VALUES(?,?,?,?,?,?,?,?,?,'Draft',?)
        ''', (document_id,next_version,safe_original,stored_filename,file.mimetype,file_path.stat().st_size,
              str(file_path),email(),now(),change_summary or f'New version {next_version}.'))
        c.execute("UPDATE knowledge_documents SET current_version=?, status='Draft', updated_at=? WHERE id=?", (next_version, now(), document_id))
        c.commit()
    except Exception:
        c.rollback()
        try:
            file_path.unlink(missing_ok=True)
        except Exception:
            pass
        c.close()
        raise
    c.close()
    log('KNOWLEDGE_DOCUMENT_VERSION_REGISTERED', f'{document_ref}: version {next_version}; status Draft')
    flash(f'{document_ref} version {next_version} uploaded successfully. The previous version remains retained.', 'success')
    return redirect(url_for('knowledge_base'))

@app.route('/knowledge-base/<int:document_id>/download')
@login_required
def knowledge_base_download(document_id):
    c = db()
    doc = c.execute('''
        SELECT d.document_ref, v.original_filename, v.stored_filename
        FROM knowledge_documents d
        JOIN knowledge_document_versions v
          ON v.document_id=d.id AND v.version_number=d.current_version
        WHERE d.id=?
    ''',(document_id,)).fetchone()
    c.close()
    if not doc:
        abort(404)
    return send_from_directory(str(UPLOAD / 'knowledge_base'), doc['stored_filename'], as_attachment=True, download_name=doc['original_filename'])

@app.route('/knowledge-base/<int:document_id>/versions')
@login_required
def knowledge_base_versions(document_id):
    c=db()
    doc=c.execute('SELECT * FROM knowledge_documents WHERE id=?',(document_id,)).fetchone()
    versions=c.execute('SELECT * FROM knowledge_document_versions WHERE document_id=? ORDER BY version_number DESC',(document_id,)).fetchall()
    c.close()
    if not doc:
        abort(404)
    return render_template('knowledge_base_versions.html', document=doc, versions=versions)

@app.route('/knowledge-base/<int:document_id>/version/<int:version_number>/download')
@login_required
def knowledge_base_version_download(document_id, version_number):
    c = db()
    v = c.execute('''
        SELECT d.document_ref, v.original_filename, v.stored_filename
        FROM knowledge_documents d
        JOIN knowledge_document_versions v ON v.document_id=d.id
        WHERE d.id=? AND v.version_number=?
    ''', (document_id, version_number)).fetchone()
    c.close()
    if not v:
        abort(404)
    return send_from_directory(str(UPLOAD / 'knowledge_base'), v['stored_filename'],
                               as_attachment=True, download_name=v['original_filename'])

@app.route('/knowledge-base/<int:document_id>/edit', methods=['GET', 'POST'])
@login_required
def knowledge_base_edit(document_id):
    # Edit the current Draft KB version without overwriting its file.
    c = db()

    doc = c.execute(
        'SELECT * FROM knowledge_documents WHERE id=?',
        (document_id,)
    ).fetchone()

    if not doc:
        c.close()
        abort(404)

    current = c.execute(
        'SELECT * FROM knowledge_document_versions WHERE document_id=? AND version_number=?',
        (document_id, doc['current_version'])
    ).fetchone()

    if not current:
        c.close()
        abort(404)

    if current['version_status'] != 'Draft':
        c.close()
        flash(
            'Only the current Draft version can be edited before approval.',
            'error'
        )
        return redirect(
            url_for('knowledge_base_versions', document_id=document_id)
        )

    if request.method == 'GET':
        c.close()
        return render_template(
            'knowledge_base_edit.html',
            document=doc,
            version=current,
            categories=KB_CATEGORIES,
            document_types=KB_DOCUMENT_TYPES,
            tax_types=KB_TAX_TYPES
        )

    title = request.form.get('title', '').strip()
    category = request.form.get('category', '').strip()
    document_type = request.form.get('document_type', '').strip()
    tax_type = request.form.get('tax_type', '').strip() or 'General'
    authority = request.form.get('issuing_authority', '').strip()
    description = request.form.get('description', '').strip()
    effective_date = request.form.get('effective_date', '').strip() or None
    expiry_date = request.form.get('expiry_date', '').strip() or None
    change_summary = request.form.get('change_summary', '').strip()
    file = request.files.get('document_file')

    if not title or category not in KB_CATEGORIES or document_type not in KB_DOCUMENT_TYPES:
        c.close()
        flash(
            'Title, knowledge category and document type are required.',
            'error'
        )
        return redirect(
            url_for('knowledge_base_edit', document_id=document_id)
        )

    if not change_summary:
        c.close()
        flash(
            'Please provide an amendment/change summary before saving.',
            'error'
        )
        return redirect(
            url_for('knowledge_base_edit', document_id=document_id)
        )

    if file and file.filename and not kb_allowed_file(file.filename):
        c.close()
        flash(
            'Unsupported file type. Allowed: PDF, DOC, DOCX, TXT, CSV, XLS and XLSX.',
            'error'
        )
        return redirect(
            url_for('knowledge_base_edit', document_id=document_id)
        )

    next_version = int(doc['current_version']) + 1
    document_ref = doc['document_ref']
    target = UPLOAD / 'knowledge_base'
    target.mkdir(parents=True, exist_ok=True)
    new_file_path = None

    try:
        if file and file.filename:
            safe_original = secure_filename(file.filename)
            if not safe_original:
                raise ValueError('The selected filename is not valid.')

            stored_filename = f'{document_ref}_v{next_version}_{safe_original}'
            new_file_path = target / stored_filename
            file.save(new_file_path)
            original_filename = safe_original
            mime_type = file.mimetype or current['mime_type']
        else:
            source_path = Path(current['file_path'])
            if not source_path.exists():
                raise FileNotFoundError(
                    'The current Knowledge Base file could not be found. '
                    'Please upload the amended document.'
                )

            original_filename = current['original_filename']
            ext = kb_ext(original_filename)
            stored_filename = (
                f'{document_ref}_v{next_version}_amended.{ext}'
                if ext else
                f'{document_ref}_v{next_version}_amended'
            )
            new_file_path = target / stored_filename

            # Copy the old file into a NEW version file. The old file is never overwritten.
            new_file_path.write_bytes(source_path.read_bytes())
            mime_type = current['mime_type']

        c.execute(
            "UPDATE knowledge_document_versions SET version_status='Superseded' "
            "WHERE document_id=? AND version_number=?",
            (document_id, doc['current_version'])
        )

        c.execute(
            '''
            INSERT INTO knowledge_document_versions
            (document_id, version_number, original_filename, stored_filename,
             mime_type, size_bytes, file_path, uploaded_by, uploaded_at,
             version_status, change_summary)
            VALUES(?,?,?,?,?,?,?,?,?,'Draft',?)
            ''',
            (
                document_id,
                next_version,
                original_filename,
                stored_filename,
                mime_type,
                new_file_path.stat().st_size,
                str(new_file_path),
                email(),
                now(),
                change_summary
            )
        )

        c.execute(
            '''
            UPDATE knowledge_documents
            SET title=?, category=?, document_type=?, tax_type=?,
                issuing_authority=?, description=?, effective_date=?,
                expiry_date=?, current_version=?, status='Draft', updated_at=?
            WHERE id=?
            ''',
            (
                title,
                category,
                document_type,
                tax_type,
                authority,
                description,
                effective_date,
                expiry_date,
                next_version,
                now(),
                document_id
            )
        )

        c.commit()

    except Exception as exc:
        c.rollback()
        if new_file_path is not None:
            try:
                new_file_path.unlink(missing_ok=True)
            except Exception:
                pass
        c.close()
        flash(f'Knowledge Base amendment could not be saved: {exc}', 'error')
        return redirect(
            url_for('knowledge_base_edit', document_id=document_id)
        )

    c.close()

    log(
        'KNOWLEDGE_DOCUMENT_AMENDED',
        f'{document_ref}: v{doc["current_version"]} amended to v{next_version}; '
        f'new status Draft; change={change_summary}'
    )

    flash(
        f'{document_ref} amended successfully. Version {next_version} is now Draft and requires human approval.',
        'success'
    )
    return redirect(
        url_for('knowledge_base_versions', document_id=document_id)
    )

@app.route('/knowledge-base/<int:document_id>/review', methods=['POST'])
@login_required
def knowledge_base_review(document_id):
    decision = request.form.get('decision','').strip()
    comments = request.form.get('comments','').strip()
    if decision not in ('Approve','Reject','Retire'):
        flash('Invalid Knowledge Base review decision.', 'error')
        return redirect(url_for('knowledge_base_versions', document_id=document_id))

    c = db()
    doc = c.execute('SELECT * FROM knowledge_documents WHERE id=?', (document_id,)).fetchone()
    if not doc:
        c.close()
        abort(404)
    v = c.execute('SELECT * FROM knowledge_document_versions WHERE document_id=? AND version_number=?',
                  (document_id, doc['current_version'])).fetchone()
    if not v:
        c.close()
        abort(404)

    current_status = v['version_status']
    if decision in ('Approve','Reject') and current_status != 'Draft':
        c.close()
        flash('Only a Draft current version can be approved or rejected.', 'error')
        return redirect(url_for('knowledge_base_versions', document_id=document_id))
    if decision == 'Retire' and current_status != 'Approved':
        c.close()
        flash('Only an Approved current version can be retired.', 'error')
        return redirect(url_for('knowledge_base_versions', document_id=document_id))

    # Approval is the control point at which content becomes usable by VTA.
    # Index first; if extraction fails, the version remains Draft and cannot become authoritative.
    if decision == 'Approve':
        c.close()
        try:
            chunk_count = index_knowledge_version(document_id, v['version_number'])
        except Exception as exc:
            log('KNOWLEDGE_APPROVAL_BLOCKED',
                f"{doc['document_ref']} v{v['version_number']} could not be approved because indexing failed: {type(exc).__name__}: {exc}")
            flash(f'Approval blocked because the document could not be indexed: {exc}', 'error')
            return redirect(url_for('knowledge_base_versions', document_id=document_id))
        c = db()
        # Only after successful extraction do we displace the prior approved version.
        c.execute('''UPDATE knowledge_document_versions
                     SET version_status='Superseded'
                     WHERE document_id=? AND version_status='Approved' AND version_number<>?''',
                  (document_id, v['version_number']))
        new_status = 'Approved'
        c.execute('UPDATE knowledge_document_versions SET version_status=? WHERE id=?', (new_status, v['id']))
        c.execute('UPDATE knowledge_documents SET status=?, updated_at=? WHERE id=?', (new_status, now(), document_id))
        c.execute('''INSERT INTO knowledge_reviews
                     (document_id,version_number,decision,comments,reviewed_by,reviewed_at)
                     VALUES(?,?,?,?,?,?)''',
                  (document_id, v['version_number'], decision,
                   (comments + f' Indexed chunks: {chunk_count}.').strip(), email(), now()))
        c.commit()
        c.close()
        log('KNOWLEDGE_GOVERNANCE_DECISION',
            f"{doc['document_ref']}: v{v['version_number']} Approve; status=Approved; indexed_chunks={chunk_count}; comments={comments or 'None'}")
        flash(f"{doc['document_ref']} v{v['version_number']} is now Approved and usable by VTA ({chunk_count} indexed chunks).", 'success')
        return redirect(url_for('knowledge_base_versions', document_id=document_id))

    new_status = 'Rejected' if decision == 'Reject' else 'Retired'
    c.execute('UPDATE knowledge_document_versions SET version_status=? WHERE id=?', (new_status, v['id']))
    if decision == 'Reject':
        prior_approved = c.execute(
            "SELECT 1 FROM knowledge_document_versions WHERE document_id=? AND version_status='Approved' LIMIT 1",
            (document_id,)
        ).fetchone()
        document_status = 'Approved' if prior_approved else 'Rejected'
    else:
        document_status = 'Retired'
    c.execute('UPDATE knowledge_documents SET status=?, updated_at=? WHERE id=?', (document_status, now(), document_id))
    c.execute('''INSERT INTO knowledge_reviews
                 (document_id,version_number,decision,comments,reviewed_by,reviewed_at)
                 VALUES(?,?,?,?,?,?)''',
              (document_id, v['version_number'], decision, comments, email(), now()))
    c.commit()
    c.close()
    log('KNOWLEDGE_GOVERNANCE_DECISION',
        f"{doc['document_ref']}: v{v['version_number']} {decision}; status={new_status}; comments={comments or 'None'}")
    flash(f"{doc['document_ref']} v{v['version_number']} is now {new_status}.", 'success')
    return redirect(url_for('knowledge_base_versions', document_id=document_id))

@app.route('/knowledge-base/<int:document_id>/reviews')
@login_required
def knowledge_base_reviews(document_id):
    c = db()
    doc = c.execute('SELECT * FROM knowledge_documents WHERE id=?', (document_id,)).fetchone()
    reviews = c.execute('''SELECT * FROM knowledge_reviews
                           WHERE document_id=? ORDER BY id DESC''', (document_id,)).fetchall()
    c.close()
    if not doc:
        abort(404)
    return render_template('knowledge_base_reviews.html', document=doc, reviews=reviews)

@app.route('/knowledge-base/approvals')
@login_required
def knowledge_approval_queue():
    c = db()
    pending = c.execute('''
        SELECT d.*, v.version_number, v.original_filename, v.uploaded_by, v.uploaded_at,
               v.change_summary, v.version_status
        FROM knowledge_documents d
        JOIN knowledge_document_versions v
          ON v.document_id=d.id AND v.version_number=d.current_version
        WHERE v.version_status='Draft'
        ORDER BY v.uploaded_at ASC, d.id ASC
    ''').fetchall()
    c.close()
    return render_template('knowledge_approval_queue.html', pending=pending)

@app.route('/knowledge-base/search', methods=['GET'])
@login_required
def knowledge_base_search():
    q = request.args.get('q', '').strip()
    results = search_approved_knowledge(q, limit=20) if q else []
    return render_template('knowledge_base_search.html', query=q, results=results)


@app.route('/knowledge-base/<int:document_id>/version/<int:version_number>/index', methods=['POST'])
@login_required
def knowledge_base_index_version(document_id, version_number):
    c = db()
    v = c.execute('SELECT * FROM knowledge_document_versions WHERE document_id=? AND version_number=?',
                  (document_id, version_number)).fetchone()
    c.close()
    if not v:
        abort(404)
    if v['version_status'] != 'Approved':
        flash('Only an Approved version can be indexed for VTA retrieval.', 'error')
        return redirect(url_for('knowledge_base_versions', document_id=document_id))
    try:
        count = index_knowledge_version(document_id, version_number)
        log('KNOWLEDGE_VERSION_INDEXED', f'Document {document_id} v{version_number}; {count} chunks indexed.')
        flash(f'Knowledge version {version_number} indexed successfully: {count} chunks.', 'success')
    except Exception as exc:
        log('KNOWLEDGE_INDEX_FAILED', f'Document {document_id} v{version_number}; {type(exc).__name__}: {exc}')
        flash(f'Knowledge indexing failed: {exc}', 'error')
    return redirect(url_for('knowledge_base_versions', document_id=document_id))


@app.route('/knowledge-base/reindex-approved', methods=['POST'])
@login_required
def knowledge_base_reindex_approved():
    c = db()
    approved = c.execute('''
        SELECT document_id, version_number
        FROM knowledge_document_versions
        WHERE version_status='Approved'
        ORDER BY document_id, version_number
    ''').fetchall()
    c.close()
    total = 0
    failures = []
    for row in approved:
        try:
            total += index_knowledge_version(row['document_id'], row['version_number'])
        except Exception as exc:
            failures.append(f"document {row['document_id']} v{row['version_number']}: {exc}")
    if failures:
        flash(f'Reindexed {total} chunks, but some approved documents could not be indexed: ' + ' | '.join(failures), 'error')
    else:
        flash(f'Approved Knowledge Base reindexed successfully: {total} chunks.', 'success')
    log('KNOWLEDGE_APPROVED_REINDEX', f'{total} chunks indexed; failures={len(failures)}')
    return redirect(url_for('knowledge_base'))


DATA_SOURCE_EXTENSIONS = {'csv','xlsx','xls'}
DATA_SOURCE_TYPES = ['Taxpayer Master','Payments','Returns','Assessments','VAT','CIT','PAYE','WHT','Customs / Imports','Customs / Exports','EFRIS','Payment Integrator','Third-Party Data','Audit History','Other']
DATA_SOURCE_TAX_TYPES = ['General','VAT','CIT','PAYE','WHT','Excise Duty','Customs','Other']

def data_source_allowed_file(filename):
    return bool(filename and '.' in filename and filename.rsplit('.',1)[1].lower() in DATA_SOURCE_EXTENSIONS)

def data_source_ext(filename):
    return filename.rsplit('.',1)[1].lower() if '.' in filename else ''

def inspect_data_source(path):
    if pd is None: raise RuntimeError('pandas is required for data-source validation.')
    ext=data_source_ext(path.name)
    if ext=='csv':
        df=pd.read_csv(path,nrows=5000,low_memory=False)
        with open(path,'rb') as fh: total_rows=max(sum(1 for _ in fh)-1,0)
    else:
        df=pd.read_excel(path,nrows=5000)
        try: total_rows=int(pd.read_excel(path,usecols=[0]).shape[0])
        except Exception: total_rows=len(df)
    df.columns=[str(x).strip() for x in df.columns]; cols=list(df.columns); errors=[]; warnings=[]
    if not cols: errors.append('No columns were detected.')
    if len(cols)!=len(set(cols)): errors.append('Duplicate column names were detected.')
    blank=[c for c in cols if not str(c).strip() or str(c).lower().startswith('unnamed')]
    if blank: warnings.append('Blank/unnamed columns detected: '+', '.join(blank[:10]))
    blank_rows=len(df)-len(df.dropna(how='all'))
    if blank_rows: warnings.append(f'{blank_rows} blank rows were found in the validation sample.')
    normalized={c.lower().replace(' ','').replace('_','') for c in cols}
    if not normalized.intersection({'tin','taxpayertin','taxpayeridentificationnumber','taxpayeridentificationno'}):
        warnings.append('No obvious TIN column was detected. TIN mapping will be required before Taxpayer 360 ingestion.')
    null_summary={c:int(df[c].isna().sum()) for c in cols if int(df[c].isna().sum())>0}
    if null_summary:
        top=sorted(null_summary.items(),key=lambda x:x[1],reverse=True)[:10]
        warnings.append('Missing values detected in: '+', '.join(f'{k} ({v})' for k,v in top))
    summary={'sample_rows':int(len(df)),'record_count':int(total_rows),'column_count':len(cols),'columns':cols,'nulls_sample':null_summary,'errors':errors,'warnings':warnings}
    return summary, ('Failed' if errors else ('Warnings' if warnings else 'Passed'))

# ========================= STEP 3A: ANALYST RISK RULE LIBRARY =========================
RISK_RULE_STATUSES=['Draft','Under Review','Approved','Active','Suspended','Retired','Rejected']
RISK_RULE_CATEGORIES=['Sales Reconciliation','Purchases / Imports Reconciliation','VAT Risk','CIT Risk','PAYE Risk','WHT Risk','Registration Risk','Payment Behaviour','Assessment Behaviour','Third-Party Mismatch','Sector / Peer Risk','Other']
RISK_RULE_TAX_TYPES=['General','VAT','CIT','PAYE','WHT','Excise Duty','Customs','Other']
RISK_LOGIC_OPERATORS=['=','!=','>','>=','<','<=','CONTAINS','NOT CONTAINS','IN','NOT IN','EXISTS','NOT EXISTS']
RISK_CALC_OPERATIONS=['ADD','SUBTRACT','MULTIPLY','DIVIDE','PERCENTAGE DIFFERENCE','PERCENTAGE OF','SUM','COUNT','AVERAGE','MIN','MAX']
RISK_EXPOSURE_TYPES=['Estimated Revenue Exposure','Potential VAT Exposure','Potential Income Tax Exposure','Potential PAYE Exposure','Potential WHT Exposure','Potential Excise Exposure','No Exposure Calculation']
RISK_CLAUSES=['FROM','TO','WHERE','WHEN','IF','AND','OR','WHILE','EXCEPT','THEN','ELSE','STOP','GROUP BY','JOIN','CALCULATE']
RULE_PREVIEW_SAMPLE_SIZE=5
RULE_ENGINE_VERSION='Phase-3-Risk-Engine-Rule-Library-2.0'

def _json_or_default(value, default):
    try:return json.loads(value) if value else default
    except Exception:return default

def _rule_ref(c): return 'RR-'+datetime.now(timezone.utc).strftime('%Y%m%d%H%M%S%f')

def _approved_legal_basis(c):
    return c.execute("""SELECT d.id document_id,d.document_ref,d.title,d.category,d.document_type,d.tax_type,d.issuing_authority,v.version_number,v.version_status
        FROM knowledge_documents d JOIN knowledge_document_versions v
        ON v.document_id=d.id AND v.version_number=d.current_version
        WHERE v.version_status='Approved' AND d.status IN ('Approved','Active') ORDER BY d.title""").fetchall()

def _approved_data_source_catalog(c):
    """Return only source versions that are actually approved for VTA execution.

    A newer Draft/Rejected version must never hide the last Approved version.
    If an approved clean transformation exists for an approved source version,
    that clean output is the controlled downstream file.
    """
    rows=c.execute("""
        SELECT d.id,d.source_ref,d.source_name,d.source_type,d.reporting_period,d.tax_type,
               d.status AS source_status,
               v.version_number,v.original_filename,v.file_path,v.columns_json,v.record_count,
               v.validation_status,v.version_status,
               tr.id AS transform_run_id,tr.output_path AS transform_output_path,
               tr.output_rows AS transform_rows,tr.output_columns AS transform_columns,
               tr.status AS transform_status
        FROM data_sources d
        JOIN data_source_versions v
          ON v.id=(SELECT av.id FROM data_source_versions av
                  WHERE av.source_id=d.id AND av.version_status='Approved'
                  ORDER BY av.version_number DESC LIMIT 1)
        LEFT JOIN data_transform_runs tr
          ON tr.id=(SELECT MAX(tx.id) FROM data_transform_runs tx
                    WHERE tx.source_id=d.id AND tx.version_number=v.version_number
                      AND tx.status='Approved')
        WHERE d.status='Approved'
        ORDER BY d.source_name
    """).fetchall()
    catalog=[]
    for r in rows:
        cols=_json_or_default(r['columns_json'],[])
        if isinstance(cols,dict): cols=list(cols.keys())
        effective_path=r['transform_output_path'] or r['file_path']
        effective_rows=(r['transform_rows'] if r['transform_output_path'] else (r['record_count'] or 0))
        if r['transform_output_path'] and Path(effective_path).exists() and pd is not None:
            try:
                ext=data_source_ext(Path(effective_path).name)
                if ext=='csv': cols=[str(x).strip() for x in pd.read_csv(effective_path,nrows=0).columns]
                elif ext in ('xlsx','xls'): cols=[str(x).strip() for x in pd.read_excel(effective_path,nrows=0).columns]
            except Exception:
                pass
        catalog.append({
            'id':r['id'],'ref':r['source_ref'],'name':r['source_name'],'type':r['source_type'],
            'period':r['reporting_period'] or 'Not specified','tax_type':r['tax_type'] or 'General',
            'version':r['version_number'],'file_path':effective_path,'columns':cols,
            'record_count':effective_rows,'transformed':bool(r['transform_output_path'])
        })
    return catalog

def _source_lookup(c, source_id):
    """Resolve the latest approved version without allowing Draft versions to shadow it."""
    return c.execute("""
        SELECT d.*,v.version_number,v.file_path,v.columns_json,v.version_status,
               tr.id AS transform_run_id,tr.output_path AS transform_output_path,
               tr.output_rows AS transform_rows,tr.status AS transform_status
        FROM data_sources d
        JOIN data_source_versions v
          ON v.id=(SELECT av.id FROM data_source_versions av
                  WHERE av.source_id=d.id AND av.version_status='Approved'
                  ORDER BY av.version_number DESC LIMIT 1)
        LEFT JOIN data_transform_runs tr
          ON tr.id=(SELECT MAX(tx.id) FROM data_transform_runs tx
                    WHERE tx.source_id=d.id AND tx.version_number=v.version_number
                      AND tx.status='Approved')
        WHERE d.id=? AND d.status='Approved'
    """,(source_id,)).fetchone()

def _safe_name(value): return re.sub(r'[^A-Za-z0-9]+','_',str(value)).strip('_').lower()
def _field_key(source_id,column): return f's{int(source_id)}__{_safe_name(column)}'

def _parse_number(value):
    if value is None:return None
    if isinstance(value,(int,float)) and not isinstance(value,bool):return float(value)
    text=str(value).strip().replace(',',''); pct=text.endswith('%')
    if pct:text=text[:-1]
    if not text:return None
    try:
        n=float(text); return n/100 if pct else n
    except Exception:return None

def _clean_value(v):
    try:
        if pd is not None and pd.isna(v):return None
    except Exception:pass
    try:return v.item()
    except Exception:return v

def _load_approved_source(c, source_id):
    if pd is None:raise RuntimeError('pandas is required to execute a data-source risk rule.')
    row=_source_lookup(c,int(source_id))
    if not row:raise ValueError(f'Approved data source {source_id} was not found.')
    path=Path(row['transform_output_path'] or row['file_path'])
    if not path.exists():raise ValueError(f'Data source {row["source_name"]} file is not available on the server.')
    ext=data_source_ext(path.name)
    if ext=='csv':df=pd.read_csv(path,low_memory=False)
    elif ext in ('xlsx','xls'):df=pd.read_excel(path)
    else:raise ValueError(f'Unsupported source format for {row["source_name"]}.')
    df.columns=[str(x).strip() for x in df.columns]
    original=list(df.columns); df=df.rename(columns={x:_field_key(row['id'],x) for x in original})
    return row,df,original

def _condition_text(cond,source_map):
    left=cond.get('left',{})
    if left.get('type')=='calculation':left_text=left.get('name','calculated value')
    else:
        src=source_map.get(str(left.get('source_id')),{});left_text=f"{src.get('name','Source')}.{left.get('field','field')}"
    op=cond.get('operator','')
    if op in ('EXISTS','NOT EXISTS'):return f"{left_text} {'exists' if op=='EXISTS' else 'does not exist'}"
    if cond.get('right_type')=='field':
        right=cond.get('right',{})
        if right.get('type')=='calculation':right_text=right.get('name','calculated value')
        else:
            src=source_map.get(str(right.get('source_id')),{});right_text=f"{src.get('name','Source')}.{right.get('field','field')}"
    else:right_text=str(cond.get('value',''))
    return f"{left_text} {op} {right_text}"

def _generate_rule_natural_language(plan):
    source_map={str(x['id']):x for x in plan.get('sources',[])}
    parts=[]
    base=plan.get('base_source') or {}
    if base:
        bname=source_map.get(str(base.get('source_id')),{}).get('name','the primary data source')
        parts.append(f"The system will start FROM the approved {bname} data source.")
    joins=plan.get('joins',[])
    if joins:
        js=[]
        for j in joins:
            l=source_map.get(str(j.get('left_source_id')),{}).get('name','source')
            r=source_map.get(str(j.get('right_source_id')),{}).get('name','source')
            jt=j.get('join_type','INNER').upper()
            js.append(f"{jt} JOIN {r} to {l} using {l}.{j.get('left_field','')} = {r}.{j.get('right_field','')}")
        parts.append('The selected sources will be combined as follows: ' + '; '.join(js) + '.')
    per=plan.get('period',{}) or {}
    if per.get('description'): parts.append(per['description'].strip().rstrip('.')+'.')
    if per.get('from') or per.get('to'):
        parts.append(f"The applicable period will run FROM {per.get('from') or 'the beginning of the configured period'} TO {per.get('to') or 'the end of the configured period'}.")
    if per.get('field'): parts.append(f"The period will be evaluated using the field {per['field']}.")
    for f in plan.get('filters',[]):
        text=_condition_text(f,source_map); clause=f.get('clause','WHERE')
        if clause=='EXCEPT': parts.append(f"The system will EXCLUDE records where {text}.")
        elif clause=='WHEN': parts.append(f"The system will apply this rule WHEN {text}.")
        elif clause=='WHILE': parts.append(f"The condition must remain true WHILE {text}.")
        else: parts.append(f"The analysis will be restricted WHERE {text}.")
    for calc in plan.get('calculations',[]):
        expr=calc.get('expression_text') or calc.get('operation','the configured calculation')
        parts.append(f"The system will CALCULATE {calc.get('name','the derived value')} using {expr}.")
    conditions=plan.get('conditions',[])
    if conditions:
        ct=[]
        for i,cond in enumerate(conditions):
            connector='' if i==0 else f" {cond.get('connector','AND')} "
            ct.append(connector+_condition_text(cond,source_map))
        parts.append('A taxpayer will be flagged IF ' + ''.join(ct) + '.')
    if plan.get('else_description'): parts.append('Otherwise, the system will ' + plan['else_description'].strip().rstrip('.') + '.')
    if plan.get('except_description'): parts.append('Additional EXCEPT rules: ' + plan['except_description'].strip().rstrip('.') + '.')
    stop=plan.get('stop',{}) or {}
    if stop.get('description'): parts.append('Evaluation will STOP for the affected record when ' + stop['description'].strip().rstrip('.') + '.')
    if plan.get('result',{}).get('risk_description'): parts.append(f"The resulting risk will be recorded as {plan['result']['risk_description']}.")
    exposure=plan.get('exposure',{}) or {}
    if exposure.get('formula'): parts.append(f"Estimated revenue exposure will be calculated using {exposure['formula']}.")
    if exposure.get('minimum'): parts.append(f"Only exposures at or above the configured minimum of UGX {exposure['minimum']} will be treated as meeting the exposure threshold.")
    ranking=plan.get('ranking',{}) or {}
    if ranking.get('description'): parts.append(ranking['description'].strip().rstrip('.')+'.')
    return ' '.join(parts) or 'The system will execute the configured risk rule against the selected approved data sources.'

def _parse_rule_plan(form,source_catalog):
    raw=form.get('rule_plan','').strip()
    if not raw:raise ValueError('Build the rule logic before generating the system interpretation.')
    try:plan=json.loads(raw)
    except Exception as exc:raise ValueError(f'Invalid rule definition: {exc}')
    if not isinstance(plan,dict):raise ValueError('The rule definition must be a structured object.')
    allowed={int(x['id']):x for x in source_catalog};ids=set()
    base=plan.get('base_source') or {}
    if base.get('source_id'):ids.add(int(base['source_id']))
    for x in plan.get('sources',[]):
        if x.get('id'):ids.add(int(x['id']))
    for j in plan.get('joins',[]):
        for k in ('left_source_id','right_source_id'):
            if j.get(k):ids.add(int(j[k]))
    for obj in plan.get('filters',[])+plan.get('conditions',[]):
        for side in ('left','right'):
            if obj.get(side,{}).get('source_id'):ids.add(int(obj[side]['source_id']))
    for calc in plan.get('calculations',[]):
        for side in ('left','right'):
            if calc.get(side,{}).get('source_id'):ids.add(int(calc[side]['source_id']))
        for x in calc.get('group_by',[]):
            if x.get('source_id'):ids.add(int(x['source_id']))
    missing=[str(x) for x in ids if x not in allowed]
    if missing:raise ValueError('The rule references data sources that are not approved or no longer available: '+', '.join(missing))
    if not base.get('source_id'):raise ValueError('Select a primary data source.')
    plan['sources']=[{'id':sid,'name':allowed[sid]['name'],'version':allowed[sid]['version'],'ref':allowed[sid]['ref'],'columns':allowed[sid]['columns']} for sid in sorted(ids)]
    plan['base_source']={'source_id':int(base['source_id']),'name':allowed[int(base['source_id'])]['name']}
    return plan

def _save_rule_version(c,name,category,tax_type,scope,period,natural,plan,legal,requests,exposure,ranking,existing_rule_id=None):
    stamp=now();plan_json=json.dumps(plan,ensure_ascii=False);exposure_json=json.dumps(exposure,ensure_ascii=False);ranking_json=json.dumps(ranking,ensure_ascii=False);legal_json=json.dumps(legal,ensure_ascii=False);req_json=json.dumps(requests,ensure_ascii=False)
    if existing_rule_id:
        r=c.execute('SELECT * FROM risk_rules WHERE id=?',(existing_rule_id,)).fetchone()
        if not r:raise ValueError('Risk rule not found.')
        version=int(r['version_number'] or 0)+1;rid=r['id'];ref=r['rule_ref']
    else:ref=_rule_ref(c);version=1;rid=None
    if rid is None:
        c.execute("INSERT INTO risk_rules(name,natural_language,structured_logic,category,approved,created_at,rule_ref,version_number,risk_category,tax_type,taxpayer_scope,applicable_period,detection_logic_json,legal_basis_json,information_requests_json,exposure_config_json,ranking_config_json,status,created_by,updated_at) VALUES(?,?,?,?,0,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",(name,natural,json.dumps(plan),category,stamp,ref,version,category,tax_type,scope,period,plan_json,legal_json,req_json,exposure_json,ranking_json,'Draft',email(),stamp));rid=c.execute('SELECT last_insert_rowid()').fetchone()[0]
    else:
        c.execute("UPDATE risk_rules SET name=?,natural_language=?,structured_logic=?,category=?,approved=0,version_number=?,risk_category=?,tax_type=?,taxpayer_scope=?,applicable_period=?,detection_logic_json=?,legal_basis_json=?,information_requests_json=?,exposure_config_json=?,ranking_config_json=?,status='Draft',updated_at=?,approved_by=NULL,approved_at=NULL WHERE id=?",(name,natural,json.dumps(plan),category,version,category,tax_type,scope,period,plan_json,legal_json,req_json,exposure_json,ranking_json,stamp,rid))
    c.execute('INSERT INTO risk_rule_versions(rule_id,version_number,rule_name,risk_category,tax_type,taxpayer_scope,applicable_period,detection_logic_json,legal_basis_json,information_requests_json,exposure_config_json,ranking_config_json,status,created_by,created_at,execution_plan_json,generated_natural_language,preview_summary_json,preview_status,preview_approved) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)',(rid,version,name,category,tax_type,scope,period,plan_json,legal_json,req_json,exposure_json,ranking_json,'Draft',email(),stamp,plan_json,natural,'{}','Not Run',0))
    for docid in legal:
        doc=c.execute("SELECT d.*,v.version_number FROM knowledge_documents d JOIN knowledge_document_versions v ON v.document_id=d.id AND v.version_number=d.current_version WHERE d.id=? AND v.version_status='Approved' AND d.status IN ('Approved','Active')",(int(docid),)).fetchone()
        if doc:c.execute('INSERT INTO risk_rule_legal_basis(rule_id,version_number,document_id,document_version,legal_passage,created_at) VALUES(?,?,?,?,?,?)',(rid,version,docid,doc['version_number'],'',stamp))
    for idx,req in enumerate(requests,1):c.execute('INSERT INTO risk_rule_information_requests(rule_id,version_number,request_title,purpose,period,mandatory,evidence_type,sort_order,created_at) VALUES(?,?,?,?,?,?,?,?,?)',(rid,version,req.get('title',''),req.get('purpose',''),req.get('period',''),1 if req.get('mandatory') else 0,req.get('evidence_type',''),idx,stamp))
    return rid,version,ref

def _evaluate_operator(left,op,right=None):
    if op in ('EXISTS','NOT EXISTS'):
        mask=left.notna() if hasattr(left,'notna') else left is not None
        return mask if op=='EXISTS' else ~mask
    if op in ('IN','NOT IN'):
        vals=right if isinstance(right,list) else [right];mask=left.isin(vals) if hasattr(left,'isin') else left in vals;return mask if op=='IN' else ~mask
    if op in ('CONTAINS','NOT CONTAINS'):
        mask=left.astype(str).str.contains(str(right),case=False,na=False) if hasattr(left,'astype') else str(right).lower() in str(left).lower();return mask if op=='CONTAINS' else ~mask
    if op=='=':return left==right
    if op=='!=':return left!=right
    if op=='>':return left>right
    if op=='>=':return left>=right
    if op=='<':return left<right
    if op=='<=':return left<=right
    raise ValueError(f'Unsupported operator: {op}')

def _coerce_series(series,value):
    n=_parse_number(value)
    if n is not None:
        numeric=pd.to_numeric(series,errors='coerce')
        if numeric.notna().sum()>0:return numeric,n
    return series.astype(str),str(value)

def _execute_rule_plan(c,plan,sample_only=5,return_records=False):
    """Execute a deterministic rule against approved source versions only.

    Business controls:
      - Every source must resolve to an approved version.
      - Join types are restricted to inner/left/right/outer.
      - Join keys are checked for accidental many-to-many multiplication.
      - Conditions are evaluated after calculations so calculated business logic works.
      - Exposure may reference either a calculated field or a source field.
      - Non-finite numeric results are treated as null, never as exposure.
    """
    if pd is None:
        raise RuntimeError('pandas is required for rule execution.')

    loaded={}; source_rows={}; source_originals={}
    for src in plan.get('sources',[]):
        sid=int(src['id'])
        row,df,original=_load_approved_source(c,sid)
        loaded[sid]=df
        source_rows[sid]=row
        source_originals[sid]=original

    base_id=int(plan['base_source']['source_id'])
    if base_id not in loaded:
        raise ValueError('The selected primary source is not part of the approved source set.')
    df=loaded[base_id].copy()
    calc_series={}

    def series_for(ref):
        if not ref:
            return None
        if ref.get('type')=='calculation':
            name=ref.get('name')
            if name not in calc_series:
                raise ValueError(f"Calculated field '{name}' is not available at this stage.")
            return calc_series[name]
        sid=int(ref.get('source_id'))
        field=ref.get('field')
        key=_field_key(sid,field)
        if key not in df.columns:
            raise ValueError(f"Field '{field}' from source {sid} is not available after source combination.")
        return df[key]

    allowed_join_types={'inner','left','right','outer'}
    for j in plan.get('joins',[]):
        rid=int(j['right_source_id'])
        if rid not in loaded:
            raise ValueError(f'Join source {rid} is not in the approved source set.')
        right=loaded[rid].copy()
        lk=_field_key(int(j['left_source_id']),j['left_field'])
        rk=_field_key(rid,j['right_field'])
        if lk not in df.columns or rk not in right.columns:
            raise ValueError(f"Join field not found: {j.get('left_field')} or {j.get('right_field')}.")
        how=str(j.get('join_type','inner')).lower()
        if how not in allowed_join_types:
            raise ValueError(f'Unsupported join type: {how}')

        # A many-to-many join silently inflates rows and can overstate risk/exposure.
        # Allow it only when explicitly requested by the rule designer.
        allow_m2m=bool(j.get('allow_many_to_many',False))
        if not allow_m2m:
            left_dupes=df[lk].notna() & df[lk].duplicated(keep=False)
            right_dupes=right[rk].notna() & right[rk].duplicated(keep=False)
            if left_dupes.any() and right_dupes.any():
                raise ValueError(
                    f"Join {j.get('left_field')} = {j.get('right_field')} would create a many-to-many row multiplication. "
                    'Aggregate or deduplicate the source, or explicitly allow many-to-many joins in the rule.'
                )
        df=df.merge(right,left_on=lk,right_on=rk,how=how,suffixes=('','__dup'))

    for f in plan.get('filters',[]):
        left=series_for(f.get('left'))
        op=f.get('operator','')
        right=series_for(f.get('right')) if f.get('right_type')=='field' else (
            [x.strip() for x in str(f.get('value','')).split(',')] if op in ('IN','NOT IN') else f.get('value','')
        )
        if hasattr(left,'dtype') and op not in ('CONTAINS','NOT CONTAINS','IN','NOT IN','EXISTS','NOT EXISTS'):
            left,right=_coerce_series(left,right)
        mask=_evaluate_operator(left,op,right)
        df=df.loc[~mask].copy() if f.get('clause')=='EXCEPT' else df.loc[mask].copy()

    for calc in plan.get('calculations',[]):
        name=str(calc.get('name','')).strip()
        if not name:
            raise ValueError('Every calculation must have a name.')
        op=calc.get('operation')
        a=pd.to_numeric(series_for(calc.get('left')),errors='coerce')
        b=pd.to_numeric(series_for(calc.get('right')),errors='coerce') if calc.get('right') else None
        if op=='ADD': v=a+b
        elif op=='SUBTRACT': v=a-b
        elif op=='MULTIPLY': v=a*b
        elif op=='DIVIDE': v=a/b.replace(0,float('nan'))
        elif op=='PERCENTAGE DIFFERENCE': v=(a-b)/a.replace(0,float('nan'))*100
        elif op=='PERCENTAGE OF': v=a/b.replace(0,float('nan'))*100
        elif op in ('SUM','COUNT','AVERAGE','MIN','MAX'):
            if calc.get('group_by'):
                keys=[_field_key(int(x['source_id']),x['field']) for x in calc['group_by']]
                base_ref=calc.get('left') or {}
                basecol=_field_key(int(base_ref['source_id']),base_ref['field'])
                missing=[k for k in keys+[basecol] if k not in df.columns]
                if missing: raise ValueError(f'Aggregation field(s) not available: {missing}')
                g=df.groupby(keys,dropna=False)[basecol]
                v={'SUM':g.transform('sum'),'COUNT':g.transform('count'),'AVERAGE':g.transform('mean'),
                   'MIN':g.transform('min'),'MAX':g.transform('max')}[op]
            else:
                v=a
        else:
            raise ValueError(f'Unsupported calculation: {op}')
        calc_series[name]=v

    combined_mask=None
    for idx,cond in enumerate(plan.get('conditions',[])):
        left=series_for(cond.get('left'))
        op=cond.get('operator')
        right=series_for(cond.get('right')) if cond.get('right_type')=='field' else (
            [x.strip() for x in str(cond.get('value','')).split(',')] if op in ('IN','NOT IN') else cond.get('value','')
        )
        if hasattr(left,'dtype') and op not in ('CONTAINS','NOT CONTAINS','IN','NOT IN','EXISTS','NOT EXISTS'):
            left,right=_coerce_series(left,right)
        mask=_evaluate_operator(left,op,right)
        connector=str(cond.get('connector','AND')).upper()
        if combined_mask is None:
            combined_mask=mask
        elif connector=='OR':
            combined_mask=combined_mask | mask
        else:
            combined_mask=combined_mask & mask

    if combined_mask is None:
        combined_mask=pd.Series(True,index=df.index)
    result=df.loc[combined_mask].copy()

    exposure=plan.get('exposure',{}) or {}
    exposure_field=str(exposure.get('field') or '').strip()
    if exposure_field:
        if exposure_field in calc_series:
            exposure_series=calc_series[exposure_field].loc[result.index]
        else:
            # Permit a direct source field by exact field name or fully-qualified field key.
            direct_key=None
            for src in plan.get('sources',[]):
                sid=int(src['id'])
                for original in source_originals.get(sid,[]):
                    if original==exposure_field or _field_key(sid,original)==exposure_field:
                        direct_key=_field_key(sid,original); break
                if direct_key: break
            if direct_key and direct_key in result.columns:
                exposure_series=result[direct_key]
            else:
                raise ValueError(f"Exposure field '{exposure_field}' is not a calculated field or approved source field.")
    else:
        minimum=_parse_number(exposure.get('minimum')) or 0.0
        exposure_series=pd.Series(minimum,index=result.index,dtype='float64')

    exposure_series=pd.to_numeric(exposure_series,errors='coerce').replace([float('inf'),float('-inf')],float('nan')).fillna(0.0)
    result['__exposure__']=exposure_series
    summary={
        'records_evaluated':int(len(df)),
        'records_triggered':int(len(result)),
        'total_exposure':float(result['__exposure__'].sum()) if len(result) else 0.0
    }

    cols=[]
    for ref in plan.get('preview_fields',[]):
        if ref.get('type')=='calculation' and ref.get('name') in calc_series:
            cols.append((ref['name'],calc_series[ref['name']]))
        else:
            sid=int(ref.get('source_id'))
            key=_field_key(sid,ref.get('field'))
            if key in result.columns:
                cols.append((f"{source_rows[sid]['source_name']}.{ref.get('field')}",result[key]))
    if not cols:
        for col in list(result.columns)[:8]:
            cols.append((col,result[col]))

    preview=[]
    for _,row in result.head(sample_only).iterrows():
        obj={name:_clean_value(series.loc[row.name]) for name,series in cols}
        obj['Risk Result']=plan.get('result',{}).get('risk_description','Potential Risk')
        obj['Estimated Exposure']=_clean_value(row.get('__exposure__',0))
        preview.append(obj)
    summary['preview']=preview

    if return_records:
        summary['_result_frame']=result
        summary['_source_rows']=source_rows
        summary['_loaded_original_columns']=source_originals
    return summary

def _save_preview(c,rule_id,version,summary):
    stamp=now()
    c.execute('DELETE FROM risk_rule_preview_results WHERE rule_id=? AND version_number=?',(rule_id,version))
    for i,row in enumerate(summary.get('preview',[]),1):
        c.execute('INSERT INTO risk_rule_preview_results(rule_id,version_number,sample_rank,result_json,created_at) VALUES(?,?,?,?,?)',(rule_id,version,i,json.dumps(row,default=str,ensure_ascii=False),stamp))
    compact={k:v for k,v in summary.items() if k not in ('preview','_result_frame','_source_rows','_loaded_original_columns')}
    compact['preview_sample_size']=len(summary.get('preview',[]))
    compact['preview_sample_is_display_only']=True
    c.execute("UPDATE risk_rule_versions SET preview_summary_json=?,preview_status='Completed',preview_run_at=?,preview_approved=0 WHERE rule_id=? AND version_number=?",(json.dumps(compact,default=str),stamp,rule_id,version))
    c.execute("UPDATE risk_rules SET status='Under Review',approved=0,updated_at=? WHERE id=?",(stamp,rule_id))


def _validate_rule_sources_current(c,plan):
    current={int(x['id']):x for x in _approved_data_source_catalog(c)}
    used=plan.get('sources',[]) or []
    if not used:
        raise ValueError('The rule has no approved data sources.')
    for src in used:
        sid=int(src['id'])
        if sid not in current:
            raise ValueError(f"Data source {sid} is no longer approved. Rebuild or re-run the rule against the current approved source.")
        expected=int(src.get('version',current[sid]['version']))
        actual=int(current[sid]['version'])
        if expected!=actual:
            raise ValueError(f"Data source {current[sid]['name']} changed from approved version {expected} to version {actual}. Re-run the rule before approval.")
    return current


def _run_rule_preview(c,rule_id,version,plan,sample_size=RULE_PREVIEW_SAMPLE_SIZE):
    _validate_rule_sources_current(c,plan)
    summary=_execute_rule_plan(c,plan,sample_only=int(sample_size))
    _save_preview(c,rule_id,version,summary)
    return summary

def _rule_form_context(c,rule=None):
    requests=[];selected=[]
    if rule:
        requests=c.execute('SELECT * FROM risk_rule_information_requests WHERE rule_id=? AND version_number=? ORDER BY sort_order',(rule['id'],rule['version_number'])).fetchall();selected=[x['document_id'] for x in c.execute('SELECT * FROM risk_rule_legal_basis WHERE rule_id=? AND version_number=?',(rule['id'],rule['version_number'])).fetchall()]
    return dict(rule=rule,legal_basis=_approved_legal_basis(c),selected_basis=selected,requests=requests,data_sources=_approved_data_source_catalog(c),categories=RISK_RULE_CATEGORIES,tax_types=RISK_RULE_TAX_TYPES,logic_operators=RISK_LOGIC_OPERATORS,calc_operations=RISK_CALC_OPERATIONS,clauses=RISK_CLAUSES,exposure_types=RISK_EXPOSURE_TYPES)

@app.route('/business')
@login_required
def business_rules():
    c=db();rows=c.execute('''SELECT r.*,(SELECT MAX(v.version_number) FROM risk_rule_versions v WHERE v.rule_id=r.id) AS latest_version,(SELECT v.status FROM risk_rule_versions v WHERE v.rule_id=r.id ORDER BY v.version_number DESC LIMIT 1) AS latest_status FROM risk_rules r ORDER BY r.id DESC''').fetchall();c.close();return render_template('business_rules.html',rules=rows)

@app.route('/business/rules/create')
@login_required
def business_rule_create():return redirect(url_for('risk_rule_new'))

@app.route('/business/rules/<int:rule_id>')
@login_required
def business_rule_view(rule_id):return redirect(url_for('risk_rule_detail',rule_id=rule_id))

@app.route('/business/rules/<int:rule_id>/edit')
@login_required
def business_rule_edit(rule_id):return redirect(url_for('risk_rule_edit',rule_id=rule_id))

@app.route('/business/rules/<int:rule_id>/versions')
@login_required
def business_rule_versions(rule_id):return redirect(url_for('risk_rule_detail',rule_id=rule_id))

@app.route('/risk-rules')
@login_required
def risk_rules():
    q=request.args.get('q','').strip();status=request.args.get('status','').strip();category=request.args.get('category','').strip();c=db();sql='SELECT * FROM risk_rules WHERE 1=1';params=[]
    if q:sql+=' AND (name LIKE ? OR rule_ref LIKE ? OR category LIKE ?)';like='%'+q+'%';params += [like,like,like]
    if status:sql+=' AND status=?';params.append(status)
    if category:sql+=' AND COALESCE(risk_category,category)=?';params.append(category)
    sql+=" ORDER BY CASE status WHEN 'Active' THEN 1 WHEN 'Approved' THEN 2 WHEN 'Under Review' THEN 3 WHEN 'Draft' THEN 4 ELSE 5 END,id DESC";rows=c.execute(sql,params).fetchall();c.close();return render_template('risk_rules.html',rows=rows,q=q,status=status,category=category,statuses=RISK_RULE_STATUSES,categories=RISK_RULE_CATEGORIES)

def _extract_common_rule_form(c):
    name=request.form.get('rule_name','').strip();category=request.form.get('risk_category','').strip();tax_type=request.form.get('tax_type','General').strip() or 'General';scope=request.form.get('taxpayer_scope','').strip();period=request.form.get('applicable_period','').strip()
    if not name or category not in RISK_RULE_CATEGORIES:raise ValueError('Rule name and Risk Category are required.')
    plan=_parse_rule_plan(request.form,_approved_data_source_catalog(c));plan['period']={'description':request.form.get('period_description','').strip()};plan['result']={'risk_description':request.form.get('risk_description','').strip() or name};plan['except_description']=request.form.get('except_description','').strip()
    exposure={'type':request.form.get('exposure_type','Estimated Revenue Exposure'),'formula':request.form.get('exposure_formula','').strip(),'minimum':request.form.get('minimum_exposure','').strip(),'field':request.form.get('exposure_field','').strip()};ranking={'base_score':request.form.get('base_score','0'),'exposure_weight':request.form.get('exposure_weight','0'),'frequency_weight':request.form.get('frequency_weight','0'),'description':request.form.get('ranking_description','').strip()}
    legal=[int(x) for x in request.form.getlist('legal_basis') if str(x).isdigit()];requests=[];titles=request.form.getlist('request_title');purposes=request.form.getlist('request_purpose');periods=request.form.getlist('request_period');mandatory=request.form.getlist('request_mandatory');evidence=request.form.getlist('request_evidence')
    for i,title in enumerate(titles):
        title=title.strip()
        if title:requests.append({'title':title,'purpose':purposes[i].strip() if i<len(purposes) else '','period':periods[i].strip() if i<len(periods) else '','mandatory':bool(i<len(mandatory) and mandatory[i] in ('1','on','true')),'evidence_type':evidence[i].strip() if i<len(evidence) else ''})
    return name,category,tax_type,scope,period,plan,legal,requests,exposure,ranking

@app.route('/risk-rules/new',methods=['GET','POST'])
@login_required
def risk_rule_new():
    c=db()
    if request.method=='POST':
        try:
            name,category,tax_type,scope,period,plan,legal,requests,exposure,ranking=_extract_common_rule_form(c)
            natural=_generate_rule_natural_language(plan)
            rid,version,ref=_save_rule_version(c,name,category,tax_type,scope,period,natural,plan,legal,requests,exposure,ranking)
            c.commit(); c.close()
            # Generate means generate AND execute against approved real data.
            pc=db()
            try:
                summary=_run_rule_preview(pc,rid,version,plan,RULE_PREVIEW_SAMPLE_SIZE)
                pc.commit(); pc.close()
                log('RISK_RULE_INTERPRETED_AND_EXECUTED',f'{ref} v{version}: evaluated={summary["records_evaluated"]}; triggered={summary["records_triggered"]}; exposure={summary["total_exposure"]:,.2f}')
                flash(f'Rule {ref} Version {version} generated and executed against approved data: {summary["records_triggered"]:,} matching record(s) from {summary["records_evaluated"]:,} evaluated.','success')
            except Exception as preview_exc:
                pc.rollback(); pc.execute("UPDATE risk_rule_versions SET preview_status='Failed',preview_summary_json=?,preview_approved=0 WHERE rule_id=? AND version_number=?",(json.dumps({'error':str(preview_exc)}),rid,version)); pc.commit(); pc.close()
                log('RISK_RULE_PREVIEW_FAILED',f'{ref} v{version}: {type(preview_exc).__name__}: {preview_exc}')
                flash(f'Rule {ref} was saved, but execution failed: {preview_exc}','error')
            return redirect(url_for('risk_rule_review',rule_id=rid))
        except Exception as exc:
            c.rollback();flash(str(exc),'error')
    ctx=_rule_form_context(c);c.close();return render_template('risk_rule_form.html',**ctx)

@app.route('/risk-rules/<int:rule_id>/edit',methods=['GET','POST'])
@login_required
def risk_rule_edit(rule_id):
    c=db();rule=c.execute('SELECT * FROM risk_rules WHERE id=?',(rule_id,)).fetchone()
    if not rule:c.close();abort(404)
    if request.method=='POST':
        try:
            name,category,tax_type,scope,period,plan,legal,requests,exposure,ranking=_extract_common_rule_form(c)
            natural=_generate_rule_natural_language(plan)
            rid,version,ref=_save_rule_version(c,name,category,tax_type,scope,period,natural,plan,legal,requests,exposure,ranking,existing_rule_id=rule_id)
            c.commit(); c.close()
            pc=db()
            try:
                summary=_run_rule_preview(pc,rid,version,plan,RULE_PREVIEW_SAMPLE_SIZE)
                pc.commit(); pc.close()
                log('RISK_RULE_VERSION_CREATED_AND_EXECUTED',f'{ref} v{version}: evaluated={summary["records_evaluated"]}; triggered={summary["records_triggered"]}; exposure={summary["total_exposure"]:,.2f}')
                flash(f'Rule {ref} Version {version} updated and executed against approved data: {summary["records_triggered"]:,} matching record(s).','success')
            except Exception as preview_exc:
                pc.rollback(); pc.execute("UPDATE risk_rule_versions SET preview_status='Failed',preview_summary_json=?,preview_approved=0 WHERE rule_id=? AND version_number=?",(json.dumps({'error':str(preview_exc)}),rid,version)); pc.commit(); pc.close()
                log('RISK_RULE_PREVIEW_FAILED',f'{ref} v{version}: {type(preview_exc).__name__}: {preview_exc}')
                flash(f'Rule {ref} was saved, but execution failed: {preview_exc}','error')
            return redirect(url_for('risk_rule_review',rule_id=rid))
        except Exception as exc:
            c.rollback();flash(str(exc),'error')
    ctx=_rule_form_context(c,rule);c.close();return render_template('risk_rule_form.html',**ctx)

@app.route('/risk-rules/<int:rule_id>/review')
@login_required
def risk_rule_review(rule_id):
    c=db();r=c.execute('SELECT * FROM risk_rules WHERE id=?',(rule_id,)).fetchone()
    if not r:c.close();abort(404)
    v=c.execute('SELECT * FROM risk_rule_versions WHERE rule_id=? AND version_number=?',(rule_id,r['version_number'])).fetchone()
    if not v:c.close();abort(404)
    previews=c.execute('SELECT * FROM risk_rule_preview_results WHERE rule_id=? AND version_number=? ORDER BY sample_rank',(rule_id,r['version_number'])).fetchall()
    latest_run=c.execute('SELECT * FROM risk_rule_runs WHERE rule_id=? AND version_number=? ORDER BY id DESC LIMIT 1',(rule_id,r['version_number'])).fetchone()
    legal_count=c.execute('SELECT COUNT(*) FROM risk_rule_legal_basis WHERE rule_id=? AND version_number=?',(rule_id,r['version_number'])).fetchone()[0]
    plan=_json_or_default(v['execution_plan_json'] or v['detection_logic_json'],{})
    source_catalog={int(x['id']):x for x in _approved_data_source_catalog(c)}
    source_state=[]
    for src in plan.get('sources',[]) or []:
        sid=int(src['id']); current=source_catalog.get(sid)
        source_state.append({'id':sid,'name':src.get('name') or (current or {}).get('name','Unknown source'),'rule_version':src.get('version'),'current_version':(current or {}).get('version'),'approved':bool(current),'current_name':(current or {}).get('name')})
    c.close()
    return render_template('risk_rule_review.html',
        rule=r,version=v,plan=plan,
        natural_language=v['generated_natural_language'] or r['natural_language'],
        summary=_json_or_default(v['preview_summary_json'],{}),
        preview_rows=[_json_or_default(x['result_json'],{}) for x in previews],
        latest_run=latest_run,legal_count=legal_count,source_state=source_state)

@app.route('/risk-rules/<int:rule_id>/preview',methods=['POST'])
@login_required
def risk_rule_preview(rule_id):
    c=db();r=c.execute('SELECT * FROM risk_rules WHERE id=?',(rule_id,)).fetchone()
    if not r:c.close();abort(404)
    v=c.execute('SELECT * FROM risk_rule_versions WHERE rule_id=? AND version_number=?',(rule_id,r['version_number'])).fetchone()
    if not v:c.close();abort(404)
    try:
        plan=_json_or_default(v['execution_plan_json'] or v['detection_logic_json'],{})
        summary=_run_rule_preview(c,rule_id,r['version_number'],plan,RULE_PREVIEW_SAMPLE_SIZE)
        c.commit()
        log('RISK_RULE_PREVIEW_EXECUTED',f'{r["rule_ref"]} v{r["version_number"]}: evaluated={summary["records_evaluated"]}; triggered={summary["records_triggered"]}; exposure={summary["total_exposure"]:,.2f}')
        flash(f'Execution completed against approved data: {summary["records_triggered"]:,} matching record(s) from {summary["records_evaluated"]:,} evaluated; estimated exposure UGX {summary["total_exposure"]:,.2f}.','success')
    except Exception as exc:
        c.rollback()
        c.execute("UPDATE risk_rule_versions SET preview_status='Failed',preview_summary_json=?,preview_approved=0 WHERE rule_id=? AND version_number=?",(json.dumps({'error':str(exc)}),rule_id,r['version_number']))
        c.commit()
        log('RISK_RULE_PREVIEW_FAILED',f'{r["rule_ref"]} v{r["version_number"]}: {type(exc).__name__}: {exc}')
        flash(f'Execution failed: {exc}','error')
    c.close();return redirect(url_for('risk_rule_review',rule_id=rule_id))

@app.route('/risk-rules/<int:rule_id>/approve',methods=['POST'])
@login_required
def risk_rule_approve(rule_id):
    c=db();r=c.execute('SELECT * FROM risk_rules WHERE id=?',(rule_id,)).fetchone()
    if not r:c.close();abort(404)
    v=c.execute('SELECT * FROM risk_rule_versions WHERE rule_id=? AND version_number=?',(rule_id,r['version_number'])).fetchone()
    try:
        if not v or v['preview_status']!='Completed':
            raise ValueError('Run the rule preview before approval.')
        if not v['preview_summary_json'] or _json_or_default(v['preview_summary_json'],{}).get('records_evaluated') is None:
            raise ValueError('The rule preview record is incomplete. Run the preview again.')
        plan=_json_or_default(v['execution_plan_json'],{})
        current_catalog={int(x['id']):x for x in _approved_data_source_catalog(c)}
        for s in plan.get('sources',[]):
            sid=int(s['id'])
            if sid not in current_catalog:
                raise ValueError(f"Approved source {sid} used by this rule is no longer approved. Re-run the rule against the current approved source.")
            if int(current_catalog[sid]['version']) != int(s.get('version',current_catalog[sid]['version'])):
                raise ValueError(f"Source {current_catalog[sid]['name']} has changed since the rule was previewed. Re-run the preview before approval.")
        if not c.execute('SELECT COUNT(*) FROM risk_rule_legal_basis WHERE rule_id=? AND version_number=?',(rule_id,r['version_number'])).fetchone()[0]:raise ValueError('A rule cannot be approved without approved Knowledge Base legal backing.')
        stamp=now()
        c.execute("UPDATE risk_rule_versions SET status='Active',approved_by=?,approved_at=?,preview_approved=1 WHERE rule_id=? AND version_number=?",(email(),stamp,rule_id,r['version_number']))
        c.execute("UPDATE risk_rules SET status='Active',approved=1,approved_by=?,approved_at=?,updated_at=? WHERE id=?",(email(),stamp,stamp,rule_id))
        # Immediately execute the newly approved rule against approved data so the Risk Universe is populated.
        affected=_refresh_risk_universe_from_active_rules(c)
        c.commit()
        log('RISK_RULE_APPROVED',f'{r["rule_ref"]} v{r["version_number"]}: approved after interpretation and preview by {email()}; risk universe taxpayers refreshed={affected}')
        flash(f'{r["rule_ref"]} Version {r["version_number"]} is now Active. Risk Universe refreshed for {affected:,} taxpayer(s).','success')
    except Exception as exc:c.rollback();flash(str(exc),'error')
    c.close();return redirect(url_for('risk_rule_review',rule_id=rule_id))

@app.route('/risk-rules/<int:rule_id>/decision',methods=['POST'])
@login_required
def risk_rule_decision(rule_id):
    decision=request.form.get('decision','').strip();comments=request.form.get('comments','').strip()
    if decision=='Approved':return risk_rule_approve(rule_id)
    if decision not in ('Under Review','Rejected','Suspended','Retired'):flash('Invalid rule decision.','error');return redirect(url_for('risk_rules'))
    c=db();r=c.execute('SELECT * FROM risk_rules WHERE id=?',(rule_id,)).fetchone()
    if not r:c.close();abort(404)
    c.execute('UPDATE risk_rules SET status=?,approved=0,updated_at=? WHERE id=?',(decision,now(),rule_id));c.execute('UPDATE risk_rule_versions SET status=?,approved_by=NULL,approved_at=NULL WHERE rule_id=? AND version_number=?',(decision,rule_id,r['version_number']));c.commit();c.close();log('RISK_RULE_DECISION',f'{r["rule_ref"]} v{r["version_number"]}: {decision}; {comments}');flash(f'{r["rule_ref"]} Version {r["version_number"]}: {decision}.','success');return redirect(url_for('risk_rules'))

@app.route('/risk-rules/<int:rule_id>')
@login_required
def risk_rule_detail(rule_id):
    c=db();r=c.execute('SELECT * FROM risk_rules WHERE id=?',(rule_id,)).fetchone()
    if not r:c.close();abort(404)
    basis=c.execute("SELECT b.*,d.document_ref,d.title,d.category,d.document_type,d.tax_type,d.issuing_authority,v.version_status FROM risk_rule_legal_basis b JOIN knowledge_documents d ON d.id=b.document_id JOIN knowledge_document_versions v ON v.document_id=d.id AND v.version_number=b.document_version WHERE b.rule_id=? AND b.version_number=?",(rule_id,r['version_number'])).fetchall();req=c.execute('SELECT * FROM risk_rule_information_requests WHERE rule_id=? AND version_number=? ORDER BY sort_order',(rule_id,r['version_number'])).fetchall();v=c.execute('SELECT * FROM risk_rule_versions WHERE rule_id=? AND version_number=?',(rule_id,r['version_number'])).fetchone();previews=c.execute('SELECT * FROM risk_rule_preview_results WHERE rule_id=? AND version_number=? ORDER BY sample_rank',(rule_id,r['version_number'])).fetchall();c.close();return render_template('risk_rule_detail.html',rule=r,basis=basis,requests=req,detection=_json_or_default(r['detection_logic_json'],{}),exposure=_json_or_default(r['exposure_config_json'],{}),ranking=_json_or_default(r['ranking_config_json'],{}),version=v,preview_rows=[_json_or_default(x['result_json'],{}) for x in previews])

@app.route('/data-sources')
@login_required
def data_sources():
    c=db()
    rows=c.execute("SELECT d.*,v.version_number,v.original_filename,v.record_count,v.column_count,v.validation_status,v.validation_errors,v.validation_warnings,v.version_status,v.uploaded_by,v.uploaded_at,v.approved_by,v.approved_at FROM data_sources d LEFT JOIN data_source_versions v ON v.source_id=d.id AND v.version_number=d.current_version ORDER BY d.id DESC").fetchall()
    pending=c.execute("SELECT d.*,v.version_number,v.original_filename,v.record_count,v.column_count,v.validation_status,v.validation_summary,v.validation_errors,v.validation_warnings,v.uploaded_at,v.uploaded_by FROM data_sources d JOIN data_source_versions v ON v.source_id=d.id AND v.version_number=d.current_version WHERE v.version_status='Draft' ORDER BY d.id DESC").fetchall()
    transform_status={}
    runs=c.execute('SELECT * FROM data_transform_runs ORDER BY id DESC').fetchall()
    for run in runs:
        if run['source_id'] not in transform_status:
            transform_status[run['source_id']]=run
    c.close(); return render_template('data_sources.html',rows=rows,pending=pending,source_types=DATA_SOURCE_TYPES,tax_types=DATA_SOURCE_TAX_TYPES,transform_status=transform_status)

@app.route('/data-sources/register',methods=['POST'])
@login_required
def data_source_register():
    name=request.form.get('source_name','').strip(); source_type=request.form.get('source_type','').strip(); description=request.form.get('description','').strip(); owner=request.form.get('owner_department','').strip(); period=request.form.get('reporting_period','').strip(); tax_type=request.form.get('tax_type','').strip() or 'General'; file=request.files.get('data_file')
    if not name or source_type not in DATA_SOURCE_TYPES: flash('Source name and valid source type are required.','error'); return redirect(url_for('data_sources'))
    if not file or not file.filename: flash('Please select a CSV or Excel data source.','error'); return redirect(url_for('data_sources'))
    if not data_source_allowed_file(file.filename): flash('Unsupported data source. Allowed: CSV, XLSX and XLS.','error'); return redirect(url_for('data_sources'))
    safe=secure_filename(file.filename)
    if not safe: flash('The selected filename is not valid.','error'); return redirect(url_for('data_sources'))
    stamp=datetime.now(timezone.utc).strftime('%Y%m%d%H%M%S%f'); source_ref='DS-'+stamp; target=UPLOAD/'data_sources'; target.mkdir(parents=True,exist_ok=True); stored=f'{source_ref}_v1_{safe}'; path=target/stored; file.save(path)
    try:
        summary,validation=inspect_data_source(path); c=db()
        c.execute("INSERT INTO data_sources(source_ref,source_name,source_type,description,owner_department,reporting_period,tax_type,current_version,status,created_by,created_at,updated_at) VALUES(?,?,?,?,?,?,?,1,'Draft',?,?,?)",(source_ref,name,source_type,description,owner,period,tax_type,email(),now(),now()))
        sid=c.execute('SELECT last_insert_rowid()').fetchone()[0]
        c.execute("INSERT INTO data_source_versions(source_id,version_number,original_filename,stored_filename,mime_type,size_bytes,file_path,uploaded_by,uploaded_at,version_status,record_count,column_count,columns_json,validation_status,validation_summary,validation_errors,validation_warnings,change_summary) VALUES(?,?,?,?,?,?,?,?,?,'Draft',?,?,?,?,?,?,?,?)",(sid,1,safe,stored,file.mimetype,path.stat().st_size,str(path),email(),now(),summary['record_count'],summary['column_count'],json.dumps(summary['columns']),validation,json.dumps(summary),len(summary['errors']),len(summary['warnings']),'Initial data-source registration.'))
        for msg in summary['errors']: c.execute("INSERT INTO data_source_issues(source_id,version_number,issue_type,severity,description,created_at) VALUES(?,1,'Validation','Error',?,?)",(sid,msg,now()))
        for msg in summary['warnings']: c.execute("INSERT INTO data_source_issues(source_id,version_number,issue_type,severity,description,created_at) VALUES(?,1,'Validation','Warning',?,?)",(sid,msg,now()))
        c.commit(); c.close()
    except Exception as exc:
        try: path.unlink(missing_ok=True)
        except Exception: pass
        flash(f'Data source validation failed: {exc}','error'); return redirect(url_for('data_sources'))
    log('DATA_SOURCE_REGISTERED',f'{source_ref}: {name}; v1; validation={validation}'); flash(f'{source_ref} registered as Version 1 ({validation}). It is NOT available to the VTA until approved.','success'); return redirect(url_for('data_sources'))

@app.route('/data-sources/<int:source_id>/review',methods=['POST'])
@login_required
def data_source_review(source_id):
    decision=request.form.get('decision','').strip(); comments=request.form.get('comments','').strip()
    if decision not in ('Approved','Rejected','Needs Revision'): flash('Invalid review decision.','error'); return redirect(url_for('data_sources'))
    c=db(); src=c.execute('SELECT * FROM data_sources WHERE id=?',(source_id,)).fetchone()
    if not src: c.close(); abort(404)
    v=c.execute('SELECT * FROM data_source_versions WHERE source_id=? AND version_number=?',(source_id,src['current_version'])).fetchone()
    if not v: c.close(); abort(404)
    if decision=='Approved' and v['validation_status']=='Failed': c.close(); flash('This source cannot be approved because validation failed. Correct it and upload a new version.','error'); return redirect(url_for('data_sources'))
    c.execute('INSERT INTO data_source_reviews(source_id,version_number,decision,comments,reviewed_by,reviewed_at) VALUES(?,?,?,?,?,?)',(source_id,v['version_number'],decision,comments,email(),now()))
    c.execute('UPDATE data_source_versions SET version_status=?,approved_by=?,approved_at=? WHERE id=?',(decision,email(),now() if decision=='Approved' else None,v['id']))
    c.execute('UPDATE data_sources SET status=?,updated_at=? WHERE id=?',('Approved' if decision=='Approved' else ('Rejected' if decision=='Rejected' else 'Draft'),now(),source_id)); c.commit(); c.close()
    log('DATA_SOURCE_REVIEWED',f'{src["source_ref"]} v{v["version_number"]}: {decision}; {comments}'); flash(f'{src["source_ref"]} Version {v["version_number"]}: {decision}.','success'); return redirect(url_for('data_sources'))

@app.route('/data-sources/<int:source_id>/version',methods=['POST'])
@login_required
def data_source_new_version(source_id):
    file=request.files.get('data_file'); change=request.form.get('change_summary','').strip()
    if not file or not file.filename or not data_source_allowed_file(file.filename): flash('Please upload a CSV, XLSX or XLS file for the new version.','error'); return redirect(url_for('data_sources'))
    safe=secure_filename(file.filename); c=db(); src=c.execute('SELECT * FROM data_sources WHERE id=?',(source_id,)).fetchone()
    if not src: c.close(); abort(404)
    next_v=int(src['current_version'])+1; ref=src['source_ref']; target=UPLOAD/'data_sources'; target.mkdir(parents=True,exist_ok=True); stored=f'{ref}_v{next_v}_{safe}'; path=target/stored; file.save(path)
    try:
        summary,validation=inspect_data_source(path); c.execute("INSERT INTO data_source_versions(source_id,version_number,original_filename,stored_filename,mime_type,size_bytes,file_path,uploaded_by,uploaded_at,version_status,record_count,column_count,columns_json,validation_status,validation_summary,validation_errors,validation_warnings,change_summary) VALUES(?,?,?,?,?,?,?,?,?,'Draft',?,?,?,?,?,?,?,?)",(source_id,next_v,safe,stored,file.mimetype,path.stat().st_size,str(path),email(),now(),summary['record_count'],summary['column_count'],json.dumps(summary['columns']),validation,json.dumps(summary),len(summary['errors']),len(summary['warnings']),change or f'New version {next_v}.'))
        for msg in summary['errors']: c.execute("INSERT INTO data_source_issues(source_id,version_number,issue_type,severity,description,created_at) VALUES(?,?, 'Validation','Error',?,?)",(source_id,next_v,msg,now()))
        for msg in summary['warnings']: c.execute("INSERT INTO data_source_issues(source_id,version_number,issue_type,severity,description,created_at) VALUES(?,?, 'Validation','Warning',?,?)",(source_id,next_v,msg,now()))
        c.execute("UPDATE data_sources SET current_version=?,status='Draft',updated_at=? WHERE id=?",(next_v,now(),source_id)); c.commit()
    except Exception as exc:
        c.rollback()
        try: path.unlink(missing_ok=True)
        except Exception: pass
        c.close(); flash(f'New version validation failed: {exc}','error'); return redirect(url_for('data_sources'))
    c.close(); log('DATA_SOURCE_VERSION_REGISTERED',f'{ref}: v{next_v}; validation={validation}'); flash(f'{ref} Version {next_v} uploaded as Draft ({validation}). The previous approved version remains retained.','success'); return redirect(url_for('data_sources'))

@app.route('/data-sources/<int:source_id>/versions')
@login_required
def data_source_versions(source_id):
    c=db(); src=c.execute('SELECT * FROM data_sources WHERE id=?',(source_id,)).fetchone()
    if not src: c.close(); abort(404)
    versions=c.execute('SELECT * FROM data_source_versions WHERE source_id=? ORDER BY version_number DESC',(source_id,)).fetchall(); reviews=c.execute('SELECT * FROM data_source_reviews WHERE source_id=? ORDER BY id DESC',(source_id,)).fetchall(); c.close(); return render_template('data_source_versions.html',source=src,versions=versions,reviews=reviews)

@app.route('/data-sources/<int:source_id>/download')
@login_required
def data_source_download(source_id):
    c=db(); src=c.execute('SELECT * FROM data_sources WHERE id=?',(source_id,)).fetchone()
    if not src: c.close(); abort(404)
    v=c.execute('SELECT * FROM data_source_versions WHERE source_id=? AND version_number=?',(source_id,src['current_version'])).fetchone(); c.close()
    if not v: abort(404)
    return send_from_directory(str(UPLOAD/'data_sources'),v['stored_filename'],as_attachment=True,download_name=v['original_filename'])

def transform_read_source(path):
    if pd is None:
        raise RuntimeError('pandas is required for data transformation.')
    ext=data_source_ext(path.name)
    if ext=='csv':
        return pd.read_csv(path, low_memory=False)
    if ext in ('xls','xlsx'):
        return pd.read_excel(path)
    raise RuntimeError('Unsupported data source format.')

def transform_apply_step(df, step):
    operation=step.get('operation')

    # Replace the column headers with a user-selected row.
    # row_number is 1-based so that it matches the visible spreadsheet row number.
    # The selected row becomes the header and all rows above it are removed.
    if operation=='replace_headers_with_row':
        try:
            row_number=max(int(step.get('row_number',1) or 1),1)
        except (TypeError,ValueError):
            row_number=1
        position=row_number-1
        if position >= len(df):
            return df
        header_values=df.iloc[position].tolist()
        new_columns=[]
        used={}
        for index,value in enumerate(header_values):
            if pd.isna(value) or str(value).strip()=='':
                base=f'Column{index+1}'
            else:
                base=str(value).strip()
            # Ensure headers are unique because pandas columns must be addressable.
            count=used.get(base,0)+1
            used[base]=count
            new_columns.append(base if count==1 else f'{base}_{count}')
        result=df.iloc[position+1:].copy()
        result.columns=new_columns
        return result.reset_index(drop=True)

    if operation=='remove_columns':
        columns=step.get('columns',[])
        existing=[c for c in columns if c in df.columns]
        if existing: df=df.drop(columns=existing)
        return df

    if operation=='select_columns':
        columns=step.get('columns',[])
        existing=[c for c in columns if c in df.columns]
        if existing: df=df[existing]
        return df

    if operation=='rename_column':
        old_name=step.get('old_name')
        new_name=step.get('new_name')
        if old_name in df.columns and new_name and new_name not in df.columns:
            df=df.rename(columns={old_name:new_name.strip()})
        return df

    if operation=='remove_blank_rows':
        return df.dropna(how='all').reset_index(drop=True)

    if operation=='remove_duplicates':
        columns=step.get('columns',[])
        existing=[c for c in columns if c in df.columns]
        df=df.drop_duplicates(subset=existing) if existing else df.drop_duplicates()
        return df.reset_index(drop=True)

    if operation=='remove_top_rows':
        count=max(int(step.get('count',0) or 0),0)
        return df.iloc[count:].reset_index(drop=True)

    if operation=='remove_bottom_rows':
        count=max(int(step.get('count',0) or 0),0)
        if count==0: return df.reset_index(drop=True)
        if count>=len(df): return df.iloc[0:0].reset_index(drop=True)
        return df.iloc[:-count].reset_index(drop=True)

    if operation=='keep_top_rows':
        count=max(int(step.get('count',0) or 0),0)
        return df.head(count).reset_index(drop=True)

    if operation=='filter':
        column=step.get('column')
        operator=step.get('operator','contains')
        value=str(step.get('value',''))
        if column not in df.columns: return df
        series=df[column].astype('string')
        if operator=='contains': mask=series.str.contains(value,case=False,na=False,regex=False)
        elif operator=='equals': mask=series.str.strip().str.lower()==value.strip().lower()
        elif operator=='not_equals': mask=series.str.strip().str.lower()!=value.strip().lower()
        elif operator=='starts_with': mask=series.str.startswith(value,na=False)
        elif operator=='ends_with': mask=series.str.endswith(value,na=False)
        elif operator=='blank': mask=df[column].isna() | (series.str.strip()=='')
        elif operator=='not_blank': mask=~(df[column].isna() | (series.str.strip()==''))
        else: mask=pd.Series(True,index=df.index)
        return df.loc[mask].reset_index(drop=True)

    if operation=='replace_value':
        column=step.get('column'); old_value=step.get('old_value',''); new_value=step.get('new_value','')
        if column in df.columns:
            df[column]=df[column].replace(old_value,new_value)
            df[column]=df[column].replace(str(old_value),new_value)
        return df

    if operation=='replace_text':
        column=step.get('column'); old_value=str(step.get('old_value','')); new_value=str(step.get('new_value',''))
        if column in df.columns:
            df[column]=df[column].astype('string').str.replace(old_value,new_value,regex=False)
        return df

    if operation=='trim':
        columns=step.get('columns',[]) or list(df.columns)
        for column in columns:
            if column in df.columns:
                df[column]=df[column].apply(lambda x:x.strip() if isinstance(x,str) else x)
        return df

    if operation=='clean_text':
        column=step.get('column')
        if column in df.columns:
            df[column]=df[column].astype('string').str.replace(r'[\\x00-\\x1F\\x7F]','',regex=True).str.strip()
        return df

    if operation=='change_type':
        column=step.get('column'); target_type=step.get('target_type')
        if column not in df.columns: return df
        if target_type=='text': df[column]=df[column].astype('string')
        elif target_type=='whole': df[column]=pd.to_numeric(df[column],errors='coerce').astype('Int64')
        elif target_type=='decimal': df[column]=pd.to_numeric(df[column],errors='coerce')
        elif target_type in ('date','datetime'): df[column]=pd.to_datetime(df[column],errors='coerce')
        elif target_type=='boolean':
            mapping={'true':True,'false':False,'yes':True,'no':False,'1':True,'0':False}
            df[column]=df[column].astype('string').str.strip().str.lower().map(mapping)
        return df

    if operation=='fill_down':
        column=step.get('column')
        if column in df.columns: df[column]=df[column].ffill()
        return df

    if operation=='fill_up':
        column=step.get('column')
        if column in df.columns: df[column]=df[column].bfill()
        return df

    if operation=='sort':
        column=step.get('column'); direction=step.get('direction','ascending')
        if column in df.columns:
            df=df.sort_values(by=column,ascending=(direction=='ascending'))
        return df.reset_index(drop=True)

    return df

def transform_apply_recipe(df,steps):
    for step in steps:
        df=transform_apply_step(df,step)
    return df

def transform_get_recipe(c,source_id,version_number):
    row=c.execute('SELECT * FROM data_transform_recipes WHERE source_id=? AND version_number=? ORDER BY id DESC LIMIT 1',(source_id,version_number)).fetchone()
    if not row: return None
    try: steps=json.loads(row['steps_json']) if row['steps_json'] else []
    except Exception: steps=[]
    # Always return template-safe, JSON-safe recipe data.
    if not isinstance(steps, list):
        steps=[]
    safe_steps=[]
    for item in steps:
        if isinstance(item, dict):
            safe_steps.append(item)
    return {
        'id': row['id'],
        'recipe_name': row['recipe_name'] or '',
        'steps': safe_steps,
        'status': row['status'] or 'Draft'
    }

def transform_get_source(source_id):
    c=db()
    source=c.execute('SELECT * FROM data_sources WHERE id=?',(source_id,)).fetchone()
    if not source:
        c.close(); abort(404)
    version=c.execute('SELECT * FROM data_source_versions WHERE source_id=? AND version_number=?',(source_id,source['current_version'])).fetchone()
    c.close()
    if not version: abort(404)
    return source,version

@app.route('/data-sources/<int:source_id>/transform',methods=['GET','POST'])
@login_required
def data_source_transform(source_id):
    source,version=transform_get_source(source_id)
    source_path=Path(version['file_path'])
    if not source_path.exists():
        flash('The original source file could not be found.','error')
        return redirect(url_for('data_sources'))

    c=db()
    recipe=transform_get_recipe(c,source_id,version['version_number'])
    if recipe is None:
        c.execute('INSERT INTO data_transform_recipes(source_id,version_number,recipe_name,steps_json,status,created_by,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?)',(source_id,version['version_number'],f"Transform {source['source_name']} v{version['version_number']}",'[]','Draft',email(),now(),now()))
        c.commit()
        recipe=transform_get_recipe(c,source_id,version['version_number'])
    # Normalize the recipe before it reaches Jinja.
    # IMPORTANT: never pass a Jinja Undefined value into |tojson.
    steps=recipe.get('steps',[]) if isinstance(recipe,dict) else []
    if not isinstance(steps,list):
        steps=[]

    def _json_safe(value):
        if value is None or isinstance(value,(str,int,float,bool)):
            return value
        if isinstance(value,(list,tuple)):
            return [_json_safe(v) for v in value]
        if isinstance(value,dict):
            return {str(k):_json_safe(v) for k,v in value.items()}
        return str(value)

    steps=[_json_safe(step) for step in steps if isinstance(step,dict)]
    recipe['steps']=steps
    recipe['recipe_name']=str(recipe.get('recipe_name') or '')

    # Serialize in Python, before Jinja renders the page. This completely
    # removes Flask/Jinja Undefined objects from the JavaScript payload.
    try:
        steps_json=json.dumps(steps,ensure_ascii=False,allow_nan=False,default=str)
    except (TypeError,ValueError):
        steps_json='[]'

    if request.method=='POST':
        action=request.form.get('action','')
        if action in ('add_step','edit_step'):
            operation=request.form.get('operation','').strip()
            if not operation:
                c.close(); flash('No transformation operation was selected.','error')
                return redirect(url_for('data_source_transform',source_id=source_id))
            step={'operation':operation}
            if operation in ('remove_columns','select_columns','remove_duplicates','trim','alignment'):
                step['columns']=request.form.getlist('columns')
                if not step['columns'] and request.form.get('column',''):
                    step['columns']=[request.form.get('column')]
                if operation=='alignment':
                    step['alignment']=request.form.get('alignment','left')
            elif operation=='replace_headers_with_row':
                try:
                    step['row_number']=max(int(request.form.get('row_number','1') or 1),1)
                except (TypeError,ValueError):
                    step['row_number']=1
            elif operation=='rename_column':
                step['old_name']=request.form.get('old_name','').strip(); step['new_name']=request.form.get('new_name','').strip()
            elif operation in ('remove_top_rows','remove_bottom_rows','keep_top_rows'):
                try: step['count']=max(int(request.form.get('count',0) or 0),0)
                except (TypeError,ValueError): step['count']=0
            elif operation=='filter':
                step['column']=request.form.get('column',''); step['operator']=request.form.get('operator','contains'); step['value']=request.form.get('value','')
            elif operation in ('replace_value','replace_text'):
                step['column']=request.form.get('column',''); step['old_value']=request.form.get('old_value',''); step['new_value']=request.form.get('new_value','')
            elif operation=='change_type':
                step['column']=request.form.get('column',''); step['target_type']=request.form.get('target_type','text')
            elif operation in ('fill_down','fill_up','clean_text'):
                step['column']=request.form.get('column','')
            elif operation=='sort':
                step['column']=request.form.get('column',''); step['direction']=request.form.get('direction','ascending')

            if action=='edit_step':
                try: idx=int(request.form.get('step_index','-1'))
                except (TypeError,ValueError): idx=-1
                if idx < 0 or idx >= len(steps):
                    c.close(); flash('The selected transformation step could not be found.','error')
                    return redirect(url_for('data_source_transform',source_id=source_id))
                steps[idx]=step
                event='DATA_TRANSFORM_STEP_EDITED'
            else:
                steps.append(step)
                event='DATA_TRANSFORM_STEP_ADDED'
            c.execute('UPDATE data_transform_recipes SET steps_json=?,updated_at=? WHERE id=?',(json.dumps(steps),now(),recipe['id']))
            c.commit(); c.close()
            log(event,f"{source['source_ref']} v{version['version_number']}: {operation}")
            return redirect(url_for('data_source_transform',source_id=source_id))

        if action=='delete_step':
            try: idx=int(request.form.get('step_index','-1'))
            except (TypeError,ValueError): idx=-1
            if idx < 0 or idx >= len(steps):
                c.close(); flash('The selected transformation step could not be found.','error')
                return redirect(url_for('data_source_transform',source_id=source_id))
            removed=steps.pop(idx)
            c.execute('UPDATE data_transform_recipes SET steps_json=?,updated_at=? WHERE id=?',(json.dumps(steps),now(),recipe['id']))
            c.commit(); c.close()
            log('DATA_TRANSFORM_STEP_DELETED',f"{source['source_ref']} v{version['version_number']}: step={idx+1}; {removed.get('operation')}")
            return redirect(url_for('data_source_transform',source_id=source_id))

        if action=='undo':
            if steps:
                removed=steps.pop()
                c.execute('UPDATE data_transform_recipes SET steps_json=?,updated_at=? WHERE id=?',(json.dumps(steps),now(),recipe['id']))
                c.commit(); log('DATA_TRANSFORM_STEP_UNDONE',f"{source['source_ref']} v{version['version_number']}: {removed.get('operation')}")
            c.close(); return redirect(url_for('data_source_transform',source_id=source_id))

        if action=='clear':
            c.execute("UPDATE data_transform_recipes SET steps_json='[]',updated_at=? WHERE id=?",(now(),recipe['id']))
            c.commit(); c.close(); log('DATA_TRANSFORM_CLEARED',f"{source['source_ref']} v{version['version_number']}")
            return redirect(url_for('data_source_transform',source_id=source_id))

        if action=='save_recipe':
            recipe_name=request.form.get('recipe_name','').strip() or f"Transform {source['source_name']} v{version['version_number']}"
            c.execute("UPDATE data_transform_recipes SET recipe_name=?,status='Saved',updated_at=? WHERE id=?",(recipe_name,now(),recipe['id']))
            c.commit(); c.close()

            # Save means: save the recipe AND create the controlled clean version.
            # The user is then returned to the main Data Sources page where the
            # clean version can be Approved or Rejected. The original source is
            # never overwritten.
            try:
                df=transform_read_source(source_path)
                cleaned=transform_apply_recipe(df.copy(),steps)
                target=UPLOAD/'data_sources'/'cleaned'; target.mkdir(parents=True,exist_ok=True)
                stamp=datetime.now(timezone.utc).strftime('%Y%m%d%H%M%S%f')
                output_filename=f"{source['source_ref']}_v{version['version_number']}_CLEAN_{stamp}.xlsx"
                output_path=target/output_filename
                cleaned.to_excel(output_path,index=False)

                alignment_steps=[step for step in steps if step.get('operation')=='alignment' and step.get('columns')]
                if alignment_steps:
                    if load_workbook is None or OpenpyxlAlignment is None:
                        raise RuntimeError('Excel alignment requires openpyxl. Add openpyxl to requirements.txt and redeploy.')
                    wb=load_workbook(output_path)
                    ws=wb.active
                    header_map={str(cell.value):cell.column for cell in ws[1]}
                    for step in alignment_steps:
                        horizontal=step.get('alignment','left')
                        for column_name in step.get('columns',[]):
                            col_idx=header_map.get(str(column_name))
                            if not col_idx: continue
                            for row in ws.iter_rows(min_row=2,min_col=col_idx,max_col=col_idx):
                                row[0].alignment=OpenpyxlAlignment(horizontal=horizontal,vertical='center')
                    wb.save(output_path)

                c=db()
                c.execute('INSERT INTO data_transform_runs(recipe_id,source_id,version_number,output_filename,output_path,output_rows,output_columns,transformation_count,status,created_by,created_at) VALUES(?,?,?,?,?,?,?,?,?,?,?)',(recipe['id'],source_id,version['version_number'],output_filename,str(output_path),len(cleaned),len(cleaned.columns),len(steps),'Pending Approval',email(),now()))
                c.execute("UPDATE data_transform_recipes SET status='Applied',updated_at=? WHERE id=?",(now(),recipe['id']))
                c.commit(); run_id=c.execute('SELECT last_insert_rowid()').fetchone()[0]; c.close()
                log('DATA_TRANSFORM_RECIPE_SAVED',f"{source['source_ref']} v{version['version_number']}: {recipe_name}; clean_run={run_id}")
                flash('Transformation saved and clean version created. It is now awaiting approval.','success')
                return redirect(url_for('data_sources'))
            except Exception as exc:
                try: c.close()
                except Exception: pass
                flash(f'Transformation was saved, but the clean version could not be created: {exc}','error')
                return redirect(url_for('data_source_transform',source_id=source_id))

        c.close()

    try:
        original_df=transform_read_source(source_path)
        original_rows=len(original_df); original_columns=[str(c) for c in original_df.columns]
        preview_df=transform_apply_recipe(original_df.copy(),steps).head(100)
        preview_columns=[str(c) for c in preview_df.columns]
        preview_rows=len(preview_df)
        preview_records=preview_df.fillna('').astype(str).to_dict(orient='records')
    except Exception as exc:
        flash(f'Transformation preview failed: {exc}','error')
        return redirect(url_for('data_sources'))

    return render_template(
        'data_transform.html',
        source=source,
        version=version,
        recipe=recipe or {'recipe_name':'','steps':[],'status':'Draft'},
        steps=steps or [],
        steps_json=steps_json,
        original_rows=original_rows,
        original_columns=original_columns,
        preview_rows=preview_rows,
        preview_columns=preview_columns,
        preview_records=preview_records
    )

@app.route('/data-sources/<int:source_id>/decision/<decision>',methods=['POST'])
@login_required
def data_source_decision(source_id,decision):
    if decision not in ('Approved','Rejected'):
        flash('Invalid data-source decision.','error')
        return redirect(url_for('data_sources'))
    c=db()
    src=c.execute('SELECT * FROM data_sources WHERE id=?',(source_id,)).fetchone()
    if not src:
        c.close(); abort(404)

    # Prefer the latest clean transformation run when one exists.
    run=c.execute('SELECT * FROM data_transform_runs WHERE source_id=? ORDER BY id DESC LIMIT 1',(source_id,)).fetchone()
    if run and run['status'] in ('Pending Approval','Created - Pending Approval'):
        if decision=='Approved':
            # The clean output can only become executable if its underlying source version
            # is itself approved. Approve the underlying version first, then supersede the
            # older approved version atomically.
            clean_version=c.execute('SELECT * FROM data_source_versions WHERE source_id=? AND version_number=?',(source_id,run['version_number'])).fetchone()
            if not clean_version or clean_version['validation_status']=='Failed':
                c.close(); flash('The clean version cannot be approved because its underlying source version failed validation.','error'); return redirect(url_for('data_sources'))
            c.execute("UPDATE data_source_versions SET version_status='Superseded' WHERE source_id=? AND version_status='Approved' AND version_number<>?",(source_id,run['version_number']))
            c.execute("UPDATE data_source_versions SET version_status='Approved',approved_by=?,approved_at=? WHERE source_id=? AND version_number=?",(email(),now(),source_id,run['version_number']))
        c.execute('UPDATE data_transform_runs SET status=? WHERE id=?',(decision,run['id']))
        c.execute('UPDATE data_sources SET current_version=?,status=?,updated_at=? WHERE id=?',(run['version_number'],decision,now(),source_id))
        c.commit(); c.close()
        log('DATA_TRANSFORM_REVIEWED',f"{src['source_ref']} clean run {run['id']}: {decision}")
        flash(f"{src['source_ref']} clean version: {decision}.",'success')
        return redirect(url_for('data_sources'))

    # If no clean version is pending, retain the original source approval flow.
    v=c.execute('SELECT * FROM data_source_versions WHERE source_id=? AND version_number=?',(source_id,src['current_version'])).fetchone()
    if not v:
        c.close(); abort(404)
    if decision=='Approved' and v['validation_status']=='Failed':
        c.close(); flash('This source cannot be approved because validation failed.','error')
        return redirect(url_for('data_sources'))
    c.execute('INSERT INTO data_source_reviews(source_id,version_number,decision,comments,reviewed_by,reviewed_at) VALUES(?,?,?,?,?,?)',(source_id,v['version_number'],decision,'Decision made from Data Sources action bar.',email(),now()))
    if decision=='Approved':
        c.execute("UPDATE data_source_versions SET version_status='Superseded' WHERE source_id=? AND version_status='Approved' AND id<>?",(source_id,v['id']))
    c.execute('UPDATE data_source_versions SET version_status=?,approved_by=?,approved_at=? WHERE id=?',(decision,email(),now() if decision=='Approved' else None,v['id']))
    c.execute('UPDATE data_sources SET status=?,updated_at=? WHERE id=?',(decision,now(),source_id))
    c.commit(); c.close()
    log('DATA_SOURCE_REVIEWED',f"{src['source_ref']} v{v['version_number']}: {decision}")
    flash(f"{src['source_ref']} Version {v['version_number']}: {decision}.",'success')
    return redirect(url_for('data_sources'))

@app.route('/data-sources/<int:source_id>/transform/apply',methods=['POST'])
@login_required
def data_source_transform_apply(source_id):
    source,version=transform_get_source(source_id)
    source_path=Path(version['file_path'])
    if not source_path.exists():
        flash('Original source file could not be found.','error')
        return redirect(url_for('data_sources'))
    c=db(); recipe=transform_get_recipe(c,source_id,version['version_number'])
    if not recipe:
        c.close(); flash('No transformation recipe exists.','error')
        return redirect(url_for('data_source_transform',source_id=source_id))
    steps=recipe['steps']
    if not steps:
        c.close(); flash('Add at least one transformation before creating a clean version.','error')
        return redirect(url_for('data_source_transform',source_id=source_id))
    try:
        df=transform_read_source(source_path)
        cleaned=transform_apply_recipe(df.copy(),steps)
        target=UPLOAD/'data_sources'/'cleaned'; target.mkdir(parents=True,exist_ok=True)
        stamp=datetime.now(timezone.utc).strftime('%Y%m%d%H%M%S%f')
        output_filename=f"{source['source_ref']}_v{version['version_number']}_CLEAN_{stamp}.xlsx"
        output_path=target/output_filename
        cleaned.to_excel(output_path,index=False)

        # Apply presentation-only alignment steps to the generated Excel file.
        # These steps do not alter the underlying data values.
        alignment_steps=[step for step in steps if step.get('operation')=='alignment' and step.get('columns')]
        if alignment_steps:
            if load_workbook is None or OpenpyxlAlignment is None:
                raise RuntimeError('Excel alignment requires openpyxl. Add openpyxl to requirements.txt and redeploy.')
            wb=load_workbook(output_path)
            ws=wb.active
            header_map={str(cell.value):cell.column for cell in ws[1]}
            for step in alignment_steps:
                horizontal=step.get('alignment','left')
                for column_name in step.get('columns',[]):
                    col_idx=header_map.get(str(column_name))
                    if not col_idx:
                        continue
                    for row in ws.iter_rows(min_row=2,min_col=col_idx,max_col=col_idx):
                        row[0].alignment=OpenpyxlAlignment(horizontal=horizontal,vertical='center')
            wb.save(output_path)

        c.execute('INSERT INTO data_transform_runs(recipe_id,source_id,version_number,output_filename,output_path,output_rows,output_columns,transformation_count,status,created_by,created_at) VALUES(?,?,?,?,?,?,?,?,?,?,?)',(recipe['id'],source_id,version['version_number'],output_filename,str(output_path),len(cleaned),len(cleaned.columns),len(steps),'Created - Pending Approval',email(),now()))
        c.execute("UPDATE data_transform_recipes SET status='Applied',updated_at=? WHERE id=?",(now(),recipe['id']))
        c.commit(); run_id=c.execute('SELECT last_insert_rowid()').fetchone()[0]; c.close()
        log('DATA_TRANSFORM_APPLIED',f"{source['source_ref']} v{version['version_number']}; run={run_id}; rows={len(cleaned)}; columns={len(cleaned.columns)}; steps={len(steps)}")
        flash('Clean version created successfully. It remains pending human approval and does not replace the original source.','success')
        return redirect(url_for('data_source_transform',source_id=source_id))
    except Exception as exc:
        c.rollback(); c.close(); flash(f'Could not create clean version: {exc}','error')
        return redirect(url_for('data_source_transform',source_id=source_id))

@app.route('/data-sources/<int:source_id>/transform/<int:run_id>/download')
@login_required
def data_source_transform_download(source_id,run_id):
    c=db()
    run=c.execute('SELECT * FROM data_transform_runs WHERE id=? AND source_id=?',(run_id,source_id)).fetchone()
    c.close()
    if not run: abort(404)
    path=Path(run['output_path'])
    if not path.exists():
        flash('The transformed output file could not be found.','error')
        return redirect(url_for('data_source_transform',source_id=source_id))
    return send_from_directory(str(path.parent),path.name,as_attachment=True,download_name=path.name)


@app.route('/data-sources/<int:source_id>/transform-data')
@login_required
def data_source_transform_legacy(source_id):
    return redirect(url_for('data_source_transform', source_id=source_id))

@app.route('/risk-assessment',methods=['GET','POST'])
@login_required
def risk_assessment():
    c=db(); ts=c.execute('SELECT * FROM taxpayers ORDER BY name').fetchall()
    if request.method=='POST': tid=int(request.form['taxpayer_id']); c.close(); assess(tid); flash('Risk assessment completed.','success'); return redirect(url_for('risk_assessment',taxpayer_id=tid))
    tid=request.args.get('taxpayer_id',type=int) or ts[0]['id']; t=c.execute('SELECT * FROM taxpayers WHERE id=?',(tid,)).fetchone(); a=c.execute('SELECT * FROM risk_assessments WHERE taxpayer_id=? ORDER BY id DESC LIMIT 1',(tid,)).fetchone(); ds=c.execute('SELECT rd.*,rr.name,rr.natural_language FROM risk_drivers rd JOIN risk_rules rr ON rr.id=rd.rule_id WHERE rd.assessment_id=?',(a['id'],)).fetchall() if a else []; c.close(); return render_template('risk_assessment.html',taxpayers=ts,selected=t,assessment=a,drivers=ds)
@app.route('/risk-universe')
@login_required
def risk_universe():
    band=request.args.get('band','').strip()
    q=request.args.get('q','').strip()
    c=db(); sql="SELECT t.*,ra.id assessment_id,ra.score,ra.band,ra.exposure,(SELECT GROUP_CONCAT(description,' | ') FROM risk_drivers WHERE assessment_id=ra.id) drivers FROM taxpayers t JOIN risk_assessments ra ON ra.taxpayer_id=t.id WHERE ra.id IN(SELECT MAX(id) FROM risk_assessments GROUP BY taxpayer_id)"; params=[]
    if band in ('High','Critical'):
        sql += " AND ra.band=?"; params.append(band)
    elif band == 'HighCritical':
        sql += " AND ra.band IN ('High','Critical')"
    if q:
        sql += " AND (t.tin LIKE ? OR t.name LIKE ? OR t.sector LIKE ? OR t.station LIKE ?)"; like='%'+q+'%'; params += [like]*4
    sql += " ORDER BY ra.score DESC,ra.exposure DESC"
    rows=c.execute(sql,params).fetchall(); c.close(); return render_template('risk_universe.html',rows=rows,band=band,q=q)

@app.route('/cases')
@login_required
def cases():
    c=db(); rows=c.execute("SELECT ac.*,t.name taxpayer_name,t.tin,t.sector FROM audit_cases ac JOIN taxpayers t ON t.id=ac.taxpayer_id ORDER BY CASE WHEN ac.status NOT IN ('Closed','Rejected') THEN 0 ELSE 1 END, ac.id DESC").fetchall(); c.close(); return render_template('cases.html',rows=rows)

@app.route('/findings')
@login_required
def findings_queue():
    c=db(); rows=c.execute("SELECT f.*,ac.case_ref,t.name taxpayer_name,t.tin FROM findings f JOIN audit_cases ac ON ac.id=f.case_id JOIN taxpayers t ON t.id=ac.taxpayer_id ORDER BY CASE WHEN f.status='AI Generated' THEN 0 ELSE 1 END,f.id DESC").fetchall(); c.close(); return render_template('findings_queue.html',rows=rows)
@app.route('/taxpayer/<int:tid>')
@login_required
def taxpayer360(tid):
    c=db(); t=c.execute('SELECT * FROM taxpayers WHERE id=?',(tid,)).fetchone(); a=c.execute('SELECT * FROM risk_assessments WHERE taxpayer_id=? ORDER BY id DESC LIMIT 1',(tid,)).fetchone(); ds=c.execute('SELECT rd.*,rr.name FROM risk_drivers rd JOIN risk_rules rr ON rr.id=rd.rule_id WHERE rd.assessment_id=?',(a['id'],)).fetchall() if a else []; cases=c.execute('SELECT * FROM audit_cases WHERE taxpayer_id=? ORDER BY id DESC',(tid,)).fetchall(); c.close(); return render_template('taxpayer.html',taxpayer=t,assessment=a,drivers=ds,cases=cases)
@app.route('/taxpayer/<int:tid>/select',methods=['POST'])
@login_required
def select_taxpayer(tid):
    c=db(); a=c.execute('SELECT * FROM risk_assessments WHERE taxpayer_id=? ORDER BY id DESC LIMIT 1',(tid,)).fetchone(); c.close();
    if not a or a['band'] not in ('High','Critical'): flash('Only High/Critical cases are selectable.','error'); return redirect(url_for('taxpayer360',tid=tid))
    return redirect(url_for('case_detail',cid=case_for(tid)))
@app.route('/case/<int:cid>')
@login_required
def case_detail(cid):
    c=db(); case=getcase(c,cid)
    if not case:
        c.close(); abort(404)
    analysis=c.execute('SELECT * FROM audit_analyses WHERE case_id=? ORDER BY id DESC LIMIT 1',(cid,)).fetchone(); findings=c.execute('SELECT * FROM findings WHERE case_id=? ORDER BY id',(cid,)).fetchall(); comm=c.execute('SELECT * FROM communications WHERE case_id=? ORDER BY id DESC',(cid,)).fetchall(); responses=c.execute('SELECT * FROM taxpayer_responses WHERE case_id=? ORDER BY id DESC',(cid,)).fetchall(); evidence=c.execute('SELECT * FROM evidence WHERE case_id=? ORDER BY id DESC',(cid,)).fetchall(); ra=c.execute('SELECT * FROM response_analyses WHERE case_id=? ORDER BY id DESC',(cid,)).fetchall(); val=c.execute('SELECT * FROM second_validations WHERE case_id=? ORDER BY id DESC',(cid,)).fetchall(); acts=c.execute('SELECT * FROM further_actions WHERE case_id=? ORDER BY id DESC',(cid,)).fetchall(); outcome=c.execute('SELECT * FROM outcomes WHERE case_id=?',(cid,)).fetchone(); open_actions=[a for a in acts if a['status']=='Open']; unresolved_findings=[f for f in findings if f['status'] not in ('Human Validated','Rejected')]; latest_validation=val[0] if val else None; c.close(); return render_template('case.html',case=case,analysis=analysis,findings=findings,communications=comm,responses=responses,evidence=evidence,response_analyses=ra,validations=val,actions=acts,outcome=outcome,open_actions=open_actions,unresolved_findings=unresolved_findings,latest_validation=latest_validation)
@app.route('/case/<int:cid>/analyze',methods=['POST'])
@login_required
def analyze(cid): run_analysis(cid); flash('AI Audit Analysis completed.','success'); return redirect(url_for('case_detail',cid=cid))
@app.route('/finding/<int:fid>')
@login_required
def finding(fid):
    c=db(); f=c.execute('SELECT f.*,ac.case_ref,ac.id case_id,t.name taxpayer_name FROM findings f JOIN audit_cases ac ON ac.id=f.case_id JOIN taxpayers t ON t.id=ac.taxpayer_id WHERE f.id=?',(fid,)).fetchone(); ds=c.execute('SELECT * FROM finding_decisions WHERE finding_id=? ORDER BY id DESC',(fid,)).fetchall(); c.close(); return render_template('finding.html',finding=f,decisions=ds)
@app.route('/finding/<int:fid>/decision',methods=['POST'])
@login_required
def finding_decision(fid):
    d=request.form.get('decision'); reason=request.form.get('reason','').strip(); c=db(); f=c.execute('SELECT f.*,ac.case_ref FROM findings f JOIN audit_cases ac ON ac.id=f.case_id WHERE f.id=?',(fid,)).fetchone(); status={'APPROVE':'Human Validated','REJECT':'Rejected','NEEDS REVIEW':'Needs Review'}.get(d)
    if not f or not status or not reason: c.close(); flash('Decision and reason are required.','error'); return redirect(url_for('finding',fid=fid))
    c.execute('UPDATE findings SET status=? WHERE id=?',(status,fid)); c.execute('INSERT INTO finding_decisions(finding_id,decision,reason,decided_by,decided_at) VALUES(?,?,?,?,?)',(fid,d,reason,email(),now())); c.execute("UPDATE audit_cases SET status='Finding Validation Completed',updated_at=? WHERE id=?",(now(),f['case_id'])); c.commit(); c.close(); log('FINDING_HUMAN_DECISION',f'{f["finding_ref"]}: {d}; {reason}',f['case_ref']); flash('Human finding validation recorded.','success')
    if d == 'APPROVE':
        return redirect(url_for('case_communication',cid=f['case_id'],finding_id=fid))
    return redirect(url_for('finding',fid=fid))
@app.route('/case/<int:cid>/communication',methods=['GET','POST'])
@login_required
def case_communication(cid):
    c=db(); case=getcase(c,cid); fs=c.execute("SELECT * FROM findings WHERE case_id=? AND status='Human Validated'",(cid,)).fetchall(); fid=request.args.get('finding_id',type=int) or (fs[0]['id'] if fs else None); f=c.execute('SELECT * FROM findings WHERE id=?',(fid,)).fetchone() if fid else None
    if request.method=='POST':
        fid=int(request.form['finding_id']); f=c.execute('SELECT * FROM findings WHERE id=? AND status="Human Validated"',(fid,)).fetchone()
        if not f:
            c.close(); flash('The selected finding is no longer Human Validated and cannot be communicated.','error'); return redirect(url_for('case_communication',cid=cid))
        recipient=request.form['recipient'].strip(); subject=request.form['subject'].strip(); body=request.form['body'].strip(); approved=request.form.get('approve_send')=='1'; draft=ai_draft(f)
        c.execute('INSERT INTO communications(case_id,finding_id,recipient,subject,ai_draft,human_body,status,created_at) VALUES(?,?,?,?,?,?,?,?)',(cid,fid,recipient,subject,draft,body,'Approved' if approved else 'Draft',now())); mid=c.execute('SELECT last_insert_rowid()').fetchone()[0]; c.commit(); c.close()
        if approved:
            return redirect(url_for('send_comm', mid=mid))
        flash('Draft saved. Human approval is still required before sending.','success'); return redirect(url_for('case_communication',cid=cid,finding_id=fid))
    original_ai_draft=ai_draft(f) if f else ''
    c.close(); return render_template(
        'communication.html',
        case=case,
        findings=fs,
        selected=f,
        ai_draft=original_ai_draft,
        recipient='jacksonakampurira@gmail.com',
        subject='Virtual Tax Auditor – Request for Clarification',
        body=original_ai_draft
    )
@app.route('/send-communication/<int:mid>')
@login_required
def send_comm(mid):
    c = db()
    x = c.execute(
        'SELECT c.*,ac.case_ref FROM communications c '
        'JOIN audit_cases ac ON ac.id=c.case_id WHERE c.id=?',
        (mid,)
    ).fetchone()
    c.close()

    if not x:
        abort(404)

    # Only an explicitly approved communication can be transmitted.
    if x['status'] not in ('Approved', 'Sending'):
        flash('This communication has not been approved for sending.', 'error')
        return redirect(url_for('case_communication', cid=x['case_id'], finding_id=x['finding_id']))

    try:
        cr = get_creds(email())

        # The local SQLite database may not contain the Gmail credential after
        # a Render restart/redeployment. In that case, request Google
        # authorization again rather than failing with a blank page.
        if not cr:
            session['pending_send_mid'] = mid
            flash('Gmail authorization is required. Please authorize Google again to send this communication.', 'error')
            return redirect(url_for('authorize'))

        # Refresh an expired credential when a refresh token is available.
        if cr.expired:
            if not cr.refresh_token:
                session['pending_send_mid'] = mid
                flash('Your Gmail authorization has expired. Please authorize Google again.', 'error')
                return redirect(url_for('authorize'))

            try:
                cr.refresh(Request())
                save_creds(email(), cr)
            except Exception:
                # Refresh token is no longer usable. Re-authorize instead of
                # returning a blank/error page.
                session['pending_send_mid'] = mid
                c = db()
                c.execute('DELETE FROM oauth_credentials WHERE email=?', (email(),))
                c.commit()
                c.close()
                flash('Your Gmail authorization could not be refreshed. Please authorize Google again.', 'error')
                return redirect(url_for('authorize'))

        # Mark as Sending so the audit trail records that transmission started.
        c = db()
        c.execute("UPDATE communications SET status='Sending' WHERE id=?", (mid,))
        c.commit()
        c.close()

        service = build('gmail', 'v1', credentials=cr)
        msg = MIMEText(x['human_body'], 'plain', 'utf-8')
        msg['to'] = x['recipient']
        msg['subject'] = x['subject']
        raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()

        sent = service.users().messages().send(
            userId='me',
            body={'raw': raw}
        ).execute()

        gid = sent.get('id', '')

        c = db()
        c.execute(
            "UPDATE communications SET status='Sent',approved_by=?,approved_at=?,sent_at=?,gmail_message_id=? WHERE id=?",
            (email(), now(), now(), gid, mid)
        )
        c.execute(
            "UPDATE audit_cases SET status='Communication Sent',updated_at=? WHERE id=?",
            (now(), x['case_id'])
        )
        c.commit()
        c.close()

        log(
            'EMAIL_SENT',
            f'Gmail message {gid} sent to {x["recipient"]}',
            x['case_ref']
        )
        flash(
            f'Email sent successfully to {x["recipient"]}. Gmail message ID: {gid}',
            'success'
        )
        return redirect(url_for('response', cid=x['case_id']))

    except Exception as e:
        # Restore the communication to Approved so the human can retry after
        # correcting the Gmail problem. Do not silently lose the approved draft.
        try:
            c = db()
            c.execute(
                "UPDATE communications SET status='Approved' WHERE id=? AND status='Sending'",
                (mid,)
            )
            c.commit()
            c.close()
        except Exception:
            pass

        log(
            'EMAIL_SEND_FAILED',
            f'Gmail send failed: {type(e).__name__}: {str(e)}',
            x['case_ref']
        )

        return render_template(
            'error.html',
            message=(
                'Gmail send failed. The approved communication was NOT sent. '
                'No tax workflow data was deleted.\n\n'
                f'Google/Gmail error: {type(e).__name__}: {str(e)}'
            )
        ), 500


@app.route('/communication')
@login_required
def communication():
    c=db(); x=c.execute("SELECT id,status FROM audit_cases WHERE status IN ('Finding Validation Completed','Communication Sent') ORDER BY id DESC LIMIT 1").fetchone(); c.close()
    if not x:
        return redirect(url_for('tasks'))
    if x['status'] == 'Communication Sent':
        return redirect(url_for('response',cid=x['id']))
    return redirect(url_for('case_communication',cid=x['id']))
@app.route('/case/<int:cid>/response',methods=['GET','POST'])
@login_required
def response(cid):
    c=db(); case=getcase(c,cid)
    if not case:
        c.close(); abort(404)
    sent=c.execute("SELECT COUNT(*) FROM communications WHERE case_id=? AND status='Sent'",(cid,)).fetchone()[0]
    if sent == 0:
        c.close(); flash('Taxpayer response is available only after an approved communication has been sent.','error'); return redirect(url_for('case_detail',cid=cid))
    if request.method=='POST':
        text=request.form.get('response_text','').strip();
        if not text: c.close(); flash('Response text is required.','error'); return redirect(url_for('response',cid=cid))
        c.execute('INSERT INTO taxpayer_responses(case_id,response_text,submitted_at,submitted_by) VALUES(?,?,?,?)',(cid,text,now(),email())); rid=c.execute('SELECT last_insert_rowid()').fetchone()[0]; files=request.files.getlist('evidence'); n=0
        for f in files:
            if f and f.filename:
                safe=secure_filename(f.filename); stored=f'{cid}_{datetime.now(timezone.utc).strftime("%Y%m%d%H%M%S%f")}_{safe}'; path=UPLOAD/stored; f.save(path); c.execute('INSERT INTO evidence(case_id,response_id,original_filename,stored_filename,mime_type,size_bytes,uploaded_at,uploaded_by) VALUES(?,?,?,?,?,?,?,?)',(cid,rid,f.filename,stored,f.mimetype,path.stat().st_size,now(),email())); n+=1
        c.execute("UPDATE audit_cases SET status='Taxpayer Response Received',updated_at=? WHERE id=?",(now(),cid)); c.commit(); c.close(); log('TAXPAYER_RESPONSE_RECEIVED',f'Response {rid}; {n} evidence file(s).',case['case_ref']); flash('Response and evidence recorded.','success'); return redirect(url_for('case_detail',cid=cid))
    rs=c.execute('SELECT * FROM taxpayer_responses WHERE case_id=? ORDER BY id DESC',(cid,)).fetchall(); es=c.execute('SELECT * FROM evidence WHERE case_id=? ORDER BY id DESC',(cid,)).fetchall(); c.close(); return render_template('response.html',case=case,responses=rs,evidence=es)
@app.route('/evidence/<path:name>')
@login_required
def evidence(name): return send_from_directory(str(UPLOAD),name,as_attachment=True)
@app.route('/case/<int:cid>/response-analysis',methods=['GET','POST'])
@login_required
def response_analysis(cid):
    c=db(); case=getcase(c,cid); r=c.execute('SELECT * FROM taxpayer_responses WHERE case_id=? ORDER BY id DESC LIMIT 1',(cid,)).fetchone(); fs=c.execute("SELECT * FROM findings WHERE case_id=? AND status='Human Validated'",(cid,)).fetchall(); es=c.execute('SELECT * FROM evidence WHERE response_id=?',(r['id'],)).fetchall() if r else []
    if not r:
        c.close(); flash('Record a taxpayer response first.','error'); return redirect(url_for('case_detail',cid=cid))
    if request.method == 'GET':
        latest=c.execute('SELECT * FROM response_analyses WHERE case_id=? ORDER BY id DESC LIMIT 1',(cid,)).fetchone()
        c.close()
        return render_template('response_analysis.html',case=case,response=r,evidence=es,analysis=latest)
    text=r['response_text'].lower(); support=[]; missing=[]
    for f in fs:
        terms=['timing','credit note','return','invoice','reversal','excluded'] if 'sales' in f['description'].lower() else ['inventory','stock','timing','goods in transit','classification','capital']
        (support if any(x in text for x in terms) else missing).append(f'{f["finding_ref"]}: '+('potential explanation requiring documentary reconciliation.' if any(x in text for x in terms) else 'no specific documentary explanation identified.'))
    if es: support.append(f'{len(es)} evidence file(s) attached for human review.')
    else: missing.append('No supporting evidence files were attached.')
    overall='Potentially supported, subject to human review.' if support and not missing else 'Partially supported; additional reconciliation/evidence appears necessary.' if support else 'Not resolved by the information currently provided.'; action='Second human validation.' if support and not missing else 'Request targeted additional evidence or clarification, then second human validation.' if support else 'Request further clarification/evidence or proceed to further action, subject to human decision.'
    lim='Controlled deterministic response-analysis engine using text matching and evidence metadata. It does not make a final compliance decision.'
    c.execute('INSERT INTO response_analyses(case_id,response_id,analysis_status,overall_assessment,findings_supported,contradictions,missing_evidence,recommended_action,limitations,created_at) VALUES(?,?,?,?,?,?,?,?,?,?)',(cid,r['id'],'Completed',overall,'\n'.join(support) or 'None','No direct contradiction identified by the controlled response-analysis engine.','\n'.join(missing) or 'None',action,lim,now())); c.execute("UPDATE audit_cases SET status='AI Response Analysis Completed',updated_at=? WHERE id=?",(now(),cid)); c.commit(); c.close(); log('AI_RESPONSE_ANALYSIS_COMPLETED',action,case['case_ref']); flash('AI Response Analysis completed.','success'); return redirect(url_for('case_detail',cid=cid))
@app.route('/case/<int:cid>/second-validation',methods=['POST'])
@login_required
def second_validation(cid):
    d=request.form.get('decision'); reason=request.form.get('reason','').strip(); fid=request.form.get('finding_id',type=int); allowed={'RESOLVED','PARTIAL','NOT RESOLVED','MORE EVIDENCE','FURTHER ACTION'}
    if d not in allowed or not reason: flash('Decision and reason are required.','error'); return redirect(url_for('case_detail',cid=cid))
    c=db(); case=getcase(c,cid);
    if not case:
        c.close(); abort(404)
    if c.execute('SELECT COUNT(*) FROM response_analyses WHERE case_id=? AND analysis_status="Completed"',(cid,)).fetchone()[0] == 0:
        c.close(); flash('Second Human Validation requires completed AI Response Analysis first.','error'); return redirect(url_for('case_detail',cid=cid))
    if not fid or not c.execute("SELECT 1 FROM findings WHERE id=? AND case_id=? AND status='Human Validated'",(fid,cid)).fetchone():
        c.close(); flash('Select a human-validated finding.','error'); return redirect(url_for('case_detail',cid=cid))
    c.execute('INSERT INTO second_validations(case_id,finding_id,decision,reason,decided_by,decided_at) VALUES(?,?,?,?,?,?)',(cid,fid,d,reason,email(),now())); status={'RESOLVED':'Finding Resolved - Pending Outcome','PARTIAL':'Partially Resolved - Pending Action','NOT RESOLVED':'Finding Not Resolved','MORE EVIDENCE':'More Evidence Required','FURTHER ACTION':'Further Action Required'}[d]; c.execute('UPDATE audit_cases SET status=?,updated_at=? WHERE id=?',(status,now(),cid)); c.commit(); c.close(); log('SECOND_HUMAN_VALIDATION',f'{d}: {reason}',case['case_ref']); flash('Second Human Validation recorded.','success'); return redirect(url_for('case_detail',cid=cid))
@app.route('/case/<int:cid>/further-action',methods=['POST'])
@login_required
def further_action(cid):
    typ=request.form.get('action_type'); instruction=request.form.get('instruction','').strip(); c=db(); case=getcase(c,cid); second=c.execute('SELECT COUNT(*) FROM second_validations WHERE case_id=?',(cid,)).fetchone()[0]
    if second == 0:
        c.close(); flash('Further Action requires Second Human Validation first.','error'); return redirect(url_for('case_detail',cid=cid))
    c.execute('INSERT INTO further_actions(case_id,action_type,instruction,created_by,created_at) VALUES(?,?,?,?,?)',(cid,typ,instruction,email(),now())); c.execute("UPDATE audit_cases SET status='Further Action Open',updated_at=? WHERE id=?",(now(),cid)); c.commit(); c.close(); log('FURTHER_ACTION_CREATED',f'{typ}: {instruction}',case['case_ref']); flash('Further action recorded.','success'); return redirect(url_for('case_detail',cid=cid))
@app.route('/case/<int:cid>/further-action/<int:aid>/complete',methods=['POST'])
@login_required
def complete_further_action(cid,aid):
    c=db(); case=getcase(c,cid); action=c.execute('SELECT * FROM further_actions WHERE id=? AND case_id=?',(aid,cid)).fetchone()
    if not action:
        c.close(); abort(404)
    c.execute("UPDATE further_actions SET status='Completed',completed_at=? WHERE id=?",(now(),aid)); c.execute("UPDATE audit_cases SET status='Further Action Completed',updated_at=? WHERE id=?",(now(),cid)); c.commit(); c.close(); log('FURTHER_ACTION_COMPLETED',f'{action["action_type"]}: {action["instruction"]}',case['case_ref']); flash('Further action marked completed. The case can continue through the controlled loop.','success'); return redirect(url_for('case_detail',cid=cid))

@app.route('/case/<int:cid>/outcome',methods=['POST'])
@login_required
def outcome(cid):
    typ=request.form.get('outcome_type'); rationale=request.form.get('rationale','').strip(); c=db(); case=getcase(c,cid); open_actions=c.execute("SELECT COUNT(*) FROM further_actions WHERE case_id=? AND status='Open'",(cid,)).fetchone()[0]; second=c.execute('SELECT COUNT(*) FROM second_validations WHERE case_id=?',(cid,)).fetchone()[0]
    unresolved=c.execute("SELECT COUNT(*) FROM findings WHERE case_id=? AND status NOT IN ('Human Validated','Rejected')",(cid,)).fetchone()[0]
    if open_actions or second==0 or unresolved>0:
        c.close(); flash('Case cannot close until further actions are complete, Second Human Validation is recorded, and all findings have a final human status.','error'); return redirect(url_for('case_detail',cid=cid))
    c.execute('INSERT INTO outcomes(case_id,outcome_type,rationale,decided_by,decided_at,closed_at) VALUES(?,?,?,?,?,?)',(cid,typ,rationale,email(),now(),now())); c.execute("UPDATE audit_cases SET status='Closed',updated_at=? WHERE id=?",(now(),cid)); c.commit(); c.close(); log('CASE_CLOSED',f'{typ}: {rationale}',case['case_ref']); flash('Final outcome recorded and case closed.','success'); return redirect(url_for('case_detail',cid=cid))
@app.route('/audit')
@login_required
def audit_log(): c=db(); es=c.execute('SELECT * FROM audit_events ORDER BY id DESC LIMIT 300').fetchall(); c.close(); return render_template('audit.html',events=es)
@app.route('/health')
def health(): return {'status':'ok','phase':3,'workflow':'closed-loop-controlled','knowledge_base':'approved-version-retrieval'}
@app.cli.command('reset-demo')
def reset_demo():
    if DB.exists(): DB.unlink()
    init_db(); print('Reset Phase 3 demo database')
init_db()
if __name__=='__main__': app.run(host='0.0.0.0',port=int(os.getenv('PORT','5000')))
