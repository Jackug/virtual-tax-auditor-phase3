import os, json, sqlite3, base64, shutil
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

BASE=Path(__file__).resolve().parent
DB=Path(os.getenv('DB_PATH', BASE/'vta_phase3.sqlite3'))
CLIENT=os.getenv('GOOGLE_CLIENT_SECRETS', str(BASE/'client_secret.json'))
BASE_URL=os.getenv('BASE_URL','http://localhost:5000').rstrip('/')
UPLOAD=Path(os.getenv('UPLOAD_DIR', BASE/'uploads')); UPLOAD.mkdir(exist_ok=True)
SCOPES=['https://www.googleapis.com/auth/gmail.send','https://www.googleapis.com/auth/userinfo.email','openid']
app=Flask(__name__); app.secret_key=os.getenv('FLASK_SECRET_KEY','change-me'); app.config.update(SESSION_COOKIE_SECURE=os.getenv('COOKIE_SECURE','0')=='1',SESSION_COOKIE_HTTPONLY=True,SESSION_COOKIE_SAMESITE='Lax',MAX_CONTENT_LENGTH=10*1024*1024)

def now(): return datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M:%S UTC')
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
'''); seed(c); c.commit(); c.close()

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
def assess(tid):
    c=db(); t=c.execute('SELECT * FROM taxpayers WHERE id=?',(tid,)).fetchone(); m=metrics(c,tid); rules=c.execute('SELECT * FROM risk_rules WHERE approved=1').fetchall(); drivers=[]
    s=m.get('sales'); i=m.get('imports'); p=m.get('purchases')
    if s and s['observed_value']>s['declared_value']:
        v=s['observed_value']-s['declared_value']; pct=v/s['declared_value']*100; r=next(x for x in rules if x['structured_logic']=='observed_sales > declared_sales'); drivers.append((r,f'Observed sales exceed declared sales by UGX {v:,.0f}.',v,pct))
    if i and p and i['observed_value']>p['declared_value']:
        v=i['observed_value']-p['declared_value']; pct=v/p['declared_value']*100; r=next(x for x in rules if x['structured_logic']=='imports > declared_purchases'); drivers.append((r,f'Imports exceed declared purchases by UGX {v:,.0f}.',v,pct))
    score=min(100,25*len(drivers)+(25 if any(x[3]>=30 for x in drivers) else 0)); score=max(score,75) if len(drivers)>=2 else score; band='Critical' if score>=90 else 'High' if score>=70 else 'Medium' if score>=40 else 'Low'; exp=sum(x[2] for x in drivers)
    c.execute('DELETE FROM risk_assessments WHERE taxpayer_id=?',(tid,)); c.execute('INSERT INTO risk_assessments(taxpayer_id,score,band,exposure,assessed_at,engine_version) VALUES(?,?,?,?,?,?)',(tid,score,band,exp,now(),'Phase-3-Risk-Engine-1.0')); aid=c.execute('SELECT last_insert_rowid()').fetchone()[0]
    c.executemany('INSERT INTO risk_drivers(assessment_id,rule_id,description,variance,variance_pct) VALUES(?,?,?,?,?)',[(aid,*x) for x in [(r['id'],d,v,pct) for r,d,v,pct in drivers]]); c.commit(); c.close(); log('RISK_ASSESSMENT_RUN',f'TIN {t["tin"]}; score={score}; band={band}'); return aid

def case_for(tid):
    c=db(); r=c.execute("SELECT * FROM audit_cases WHERE taxpayer_id=? AND status NOT IN('Closed','Rejected') ORDER BY id DESC LIMIT 1",(tid,)).fetchone()
    if r: c.close(); return r['id']
    t=c.execute('SELECT * FROM taxpayers WHERE id=?',(tid,)).fetchone(); ref='VTA-P3-TEST-'+datetime.now(timezone.utc).strftime('%Y%m%d%H%M%S'); c.execute('INSERT INTO audit_cases(case_ref,taxpayer_id,status,selected_by,selected_at,assigned_to,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?)',(ref,tid,'Selected for Virtual Audit',email(),now(),email(),now(),now())); cid=c.execute('SELECT last_insert_rowid()').fetchone()[0]; c.commit(); c.close(); log('CASE_SELECTED',f'TIN {t["tin"]} selected for Virtual Audit',ref); return cid

def run_analysis(cid):
    c=db(); case=c.execute('SELECT * FROM audit_cases WHERE id=?',(cid,)).fetchone(); t=c.execute('SELECT * FROM taxpayers WHERE id=?',(case['taxpayer_id'],)).fetchone(); m=metrics(c,t['id']); s,i,p=m['sales'],m['imports'],m['purchases']; lim='Controlled deterministic test engine over synthetic data; not production LLM/RAG and no final tax decision.'
    c.execute('INSERT INTO audit_analyses(case_id,analysis_status,summary,methodology,limitations,created_at) VALUES(?,?,?,?,?,?)',(cid,'Completed','Two potential reconciliation issues identified: observed sales exceed declared sales and imports exceed declared purchases.','Compared declared/observed sales and imports/purchases and calculated variances.',lim,now())); c.execute('DELETE FROM findings WHERE case_id=?',(cid,)); stamp=datetime.now(timezone.utc).strftime('%Y%m%d%H%M%S%f'); fs=[(f'F-P3-{stamp}-01',cid,'Sales reconciliation','Observed sales exceed declared sales.',t['financial_year'],s['declared_value'],s['observed_value'],s['observed_value']-s['declared_value'],(s['observed_value']-s['declared_value'])/s['declared_value']*100,None,s['source'],'Observed Sales - Declared Sales','The difference requires taxpayer explanation and reconciliation.','Medium',lim,'AI Generated',now()),(f'F-P3-{stamp}-02',cid,'Imports / purchases reconciliation','Imports exceed declared purchases.',t['financial_year'],p['declared_value'],i['observed_value'],i['observed_value']-p['declared_value'],(i['observed_value']-p['declared_value'])/p['declared_value']*100,None,i['source'],'Imports - Declared Purchases','The difference may have timing, inventory, classification or other explanations.','Medium',lim,'AI Generated',now())]; c.executemany('INSERT INTO findings(finding_ref,case_id,risk_category,description,financial_year,expected_value,observed_value,variance,variance_pct,potential_exposure,source,audit_test,explanation,ai_confidence,limitations,status,created_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)',fs); c.execute("UPDATE audit_cases SET status='AI Analysis Completed',updated_at=? WHERE id=?",(now(),cid)); c.commit(); c.close(); log('AI_ANALYSIS_COMPLETED','Phase 3 controlled AI analysis completed; findings remain AI Generated.',case['case_ref'])


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
            UPDATE knowledge_document_versions
            SET version_status='Superseded'
            WHERE document_id=? AND version_number=? AND version_status='Approved'
        ''', (document_id, doc['current_version']))
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
    c = db()

    doc = c.execute(
        'SELECT * FROM knowledge_documents WHERE id=?',
        (document_id,)
    ).fetchone()

    if not doc:
        c.close()
        abort(404)

    current = c.execute(
        '''
        SELECT *
        FROM knowledge_document_versions
        WHERE document_id=? AND version_number=?
        ''',
        (document_id, doc['current_version'])
    ).fetchone()

    if not current:
        c.close()
        abort(404)

    if current['version_status'] != 'Draft':
        c.close()
        flash(
            'Only a Draft Knowledge Base version can be edited before approval.',
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
    tax_type = request.form.get('tax_type', '').strip()
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

    if file and file.filename:
        safe_original = secure_filename(file.filename)
        if not safe_original:
            c.close()
            flash('The selected filename is not valid.', 'error')
            return redirect(
                url_for('knowledge_base_edit', document_id=document_id)
            )

        stored_filename = f'{document_ref}_v{next_version}_{safe_original}'
        new_file_path = target / stored_filename
        file.save(new_file_path)
        original_filename = safe_original
        mime_type = file.mimetype
    else:
        source_path = Path(current['file_path'])

        if not source_path.exists():
            c.close()
            flash(
                'The current Knowledge Base file could not be found. '
                'Please upload the amended document.',
                'error'
            )
            return redirect(
                url_for('knowledge_base_edit', document_id=document_id)
            )

        original_filename = current['original_filename']
        ext = kb_ext(original_filename)
        stored_filename = (
            f'{document_ref}_v{next_version}_amended.{ext}'
            if ext else
            f'{document_ref}_v{next_version}_amended'
        )
        new_file_path = target / stored_filename
        shutil.copy2(source_path, new_file_path)
        mime_type = current['mime_type']

    try:
        c.execute(
            '''
            UPDATE knowledge_document_versions
            SET version_status='Superseded'
            WHERE document_id=? AND version_number=?
            ''',
            (document_id, doc['current_version'])
        )

        c.execute(
            '''
            INSERT INTO knowledge_document_versions
            (
                document_id,
                version_number,
                original_filename,
                stored_filename,
                mime_type,
                size_bytes,
                file_path,
                uploaded_by,
                uploaded_at,
                version_status,
                change_summary
            )
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
                change_summary or f'Reviewer amendment to version {doc["current_version"]}.'
            )
        )

        c.execute(
            '''
            UPDATE knowledge_documents
            SET
                title=?,
                category=?,
                document_type=?,
                tax_type=?,
                issuing_authority=?,
                description=?,
                effective_date=?,
                expiry_date=?,
                current_version=?,
                status='Draft',
                updated_at=?
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

    except Exception:
        c.rollback()

        try:
            new_file_path.unlink(missing_ok=True)
        except Exception:
            pass

        c.close()
        raise

    c.close()

    log(
        'KNOWLEDGE_DOCUMENT_REVIEWER_AMENDED',
        (
            f'{document_ref}: reviewer amended v{doc["current_version"]} '
            f'and created v{next_version} Draft; '
            f'change={change_summary or "No change summary supplied"}'
        )
    )

    flash(
        f'{document_ref} was amended by the reviewer and saved as '
        f'version {next_version} (Draft). It must be reviewed and approved again.',
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
    if decision == 'Approve' and current_status != 'Draft':
        c.close()
        flash('Only a Draft current version can be approved.', 'error')
        return redirect(url_for('knowledge_base_versions', document_id=document_id))
    if decision == 'Reject' and current_status != 'Draft':
        c.close()
        flash('Only a Draft current version can be rejected.', 'error')
        return redirect(url_for('knowledge_base_versions', document_id=document_id))
    if decision == 'Retire' and current_status != 'Approved':
        c.close()
        flash('Only an Approved current version can be retired.', 'error')
        return redirect(url_for('knowledge_base_versions', document_id=document_id))

    new_status = {'Approve':'Approved', 'Reject':'Rejected', 'Retire':'Retired'}[decision]
    doc_status = new_status
    c.execute('UPDATE knowledge_document_versions SET version_status=? WHERE id=?', (new_status, v['id']))
    c.execute('UPDATE knowledge_documents SET status=?, updated_at=? WHERE id=?', (doc_status, now(), document_id))
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

@app.route('/data-sources')
@login_required
def data_sources():
    c=db(); t=c.execute('SELECT * FROM taxpayers LIMIT 1').fetchone(); ms=c.execute('SELECT * FROM taxpayer_metrics WHERE taxpayer_id=?',(t['id'],)).fetchall(); rules=c.execute('SELECT * FROM risk_rules').fetchall(); c.close(); return render_template('data_sources.html',taxpayer=t,metrics=ms,rules=rules)
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
        subject='Virtual Tax Auditor Phase 3 Test Communication',
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
    c.execute('INSERT INTO response_analyses(case_id,response_id,analysis_status,overall_assessment,findings_supported,contradictions,missing_evidence,recommended_action,limitations,created_at) VALUES(?,?,?,?,?,?,?,?,?,?)',(cid,r['id'],'Completed',overall,'\n'.join(support) or 'None','No direct contradiction identified by test engine.','\n'.join(missing) or 'None',action,lim,now())); c.execute("UPDATE audit_cases SET status='AI Response Analysis Completed',updated_at=? WHERE id=?",(now(),cid)); c.commit(); c.close(); log('AI_RESPONSE_ANALYSIS_COMPLETED',action,case['case_ref']); flash('AI Response Analysis completed.','success'); return redirect(url_for('case_detail',cid=cid))
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
def health(): return {'status':'ok','phase':3,'workflow':'closed-loop-test'}
@app.cli.command('reset-demo')
def reset_demo():
    if DB.exists(): DB.unlink()
    init_db(); print('Reset Phase 3 demo database')
init_db()
if __name__=='__main__': app.run(host='0.0.0.0',port=int(os.getenv('PORT','5000')))
