import os, logging, threading, time, re, json
from datetime import datetime
from dotenv import load_dotenv
import pandas as pd

from flask import Flask, jsonify, request, send_from_directory
from flask_socketio import SocketIO, emit, disconnect

from dashboard.utils.google_sheets import GoogleSheetsManager
from dashboard.utils.data_analyzer import DataAnalyzer
from dashboard.utils.notification_system import NotificationSystem
from dashboard.utils.auth import get_user_from_headers

load_dotenv()
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("dashboard")

app = Flask(__name__, static_folder="static")
app.config["SECRET_KEY"] = os.getenv("SECRET_KEY", "change-this")
socketio = SocketIO(app, async_mode="eventlet", cors_allowed_origins="*")

SHEET_ID = os.getenv("SHEET_ID")
SHEET_RANGE = os.getenv("SHEET_RANGE", "공사 현황!A1:AM10000")
SERVICE_JSON = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON")
OWNER_EMAIL_COLUMN = os.getenv("OWNER_EMAIL_COLUMN", "담당자 이메일")
PROJECT_CODE_COLUMN = os.getenv("PROJECT_CODE_COLUMN", "프로젝트 코드")
POLL_INTERVAL = int(os.getenv("POLL_INTERVAL", "10"))

# load credentials.json
try:
    cred_path = os.getenv("CREDENTIALS_JSON") or os.path.join(os.path.dirname(__file__), "credentials.json")
    with open(cred_path, "r", encoding="utf-8") as _f:
        _creds = json.load(_f)
except Exception:
    _creds = {}

if not SERVICE_JSON:
    SERVICE_JSON = _creds.get("google_service_account_json") or SERVICE_JSON
if not SHEET_ID:
    SHEET_ID = _creds.get("sheet_id") or SHEET_ID
if not SHEET_RANGE or SHEET_RANGE == "공사 현황!A1:AM10000":
    SHEET_RANGE = _creds.get("sheet_range") or SHEET_RANGE
if not os.getenv("SLACK_WEBHOOK_URL") and _creds.get("slack_webhook_url"):
    os.environ["SLACK_WEBHOOK_URL"] = _creds.get("slack_webhook_url","" )

gs = GoogleSheetsManager(SHEET_ID, SHEET_RANGE, SERVICE_JSON)
notifier = NotificationSystem()

_cache_df = pd.DataFrame()
_cache_hash = None
_clients = {}

def _df_hash(df: pd.DataFrame) -> str:
    if df is None or df.empty:
        return "empty"
    return str(hash(tuple(df.astype(str).fillna("").itertuples(index=False))))

def _filter_df_for_user(df: pd.DataFrame, email: str, admin: bool) -> pd.DataFrame:
    if admin:
        return df
    alias_map = (_creds.get("user_alias_map") or {})
    owner_name = alias_map.get(email.lower(), alias_map.get(email, ""))
    if owner_name and '담당자' in df.columns:
        return df[df['담당자'].astype(str).str.strip() == owner_name]
    if OWNER_EMAIL_COLUMN in df.columns:
        return df[df[OWNER_EMAIL_COLUMN].str.lower() == (email or "").lower()]
    return df[df.index < 0]

def _apply_query_filters(df: pd.DataFrame, args) -> pd.DataFrame:
    out = df
    month = args.get("month", "").strip()
    manager = args.get("manager", "").strip()
    created_col = os.getenv("CREATED_DATE_COLUMN", _creds.get("created_date_column", "등록일"))
    if month and created_col in out.columns:
        out = out.copy()
        out["__month"] = out[created_col].astype(str).str.slice(0,7)
        out = out[out["__month"] == month].drop(columns=["__month"])
    if manager:
        if '담당자' in out.columns:
            out = out[out['담당자'].astype(str).str.strip() == manager]
        elif OWNER_EMAIL_COLUMN in out.columns:
            out = out[out[OWNER_EMAIL_COLUMN].astype(str).str.contains(manager, case=False, na=False)]
    return out

def _extract_number(code: str):
    m = re.match(r'[A-Z](\d{4})-', str(code))
    return int(m.group(1)) if m else None

def _suffix_from_code(code: str):
    m = re.match(r'[A-Z]\d{4}-([A-Z]+)$', str(code))
    return m.group(1) if m else None

def _build_company_prefix_map(df: pd.DataFrame):
    m = {}
    if '프로젝트 코드' in df.columns and '사업자' in df.columns:
        for _, row in df.iterrows():
            code = str(row.get('프로젝트 코드',''))
            comp = str(row.get('사업자','')).strip()
            mm = re.match(r'([A-Z])\d{4}-', code)
            if comp and mm and comp not in m:
                m[comp] = mm.group(1)
    for k,v in (_creds.get('company_prefix_map', {}) or {}).items():
        m.setdefault(k, v)
    return m

def _build_owner_suffix_map(df: pd.DataFrame):
    m = {k:str(v).upper() for k,v in (_creds.get('owner_suffix_map', {}) or {}).items()}
    if '프로젝트 코드' in df.columns and '담당자' in df.columns:
        from collections import Counter, defaultdict
        grouped = defaultdict(list)
        for _, row in df.iterrows():
            name = str(row.get('담당자','')).strip()
            code = str(row.get('프로젝트 코드','')).strip()
            suf = _suffix_from_code(code)
            if name and suf:
                grouped[name].append(suf)
        for name, arr in grouped.items():
            common = Counter(arr).most_common(1)[0][0]
            m.setdefault(name, common)
    return m

def _next_running_number(df: pd.DataFrame):
    nums = []
    if '프로젝트 코드' in df.columns:
        for c in df['프로젝트 코드'].astype(str):
            n = _extract_number(c)
            if n is not None:
                nums.append(n)
    return (max(nums) + 1) if nums else 1

def _auto_project_code(df: pd.DataFrame, company: str, owner: str) -> str:
    comp_map = _build_company_prefix_map(df)
    own_map = _build_owner_suffix_map(df)
    prefix = comp_map.get(company.strip())
    suffix = own_map.get(owner.strip())
    if not prefix or not suffix:
        raise ValueError(f'코드 생성 실패: 회사/담당자 매핑을 확인하세요 (company={company}, owner={owner})')
    num = _next_running_number(df)
    return f"{prefix}{num:04d}-{suffix}"

def poller():
    global _cache_df, _cache_hash
    while True:
        try:
            df = gs.fetch_dataframe()
            h = _df_hash(df)
            if h != _cache_hash:
                _cache_df, _cache_hash = df, h
                logger.info("시트 변경 감지 → 구독자에게 전송")
                analyzer = DataAnalyzer(df)
                missing = analyzer.missing_fields()
                if missing:
                    notifier.notify_missing_fields(missing)
                for sid in list(_clients.keys()):
                    try:
                        emit("projects:update", {"rows": df.to_dict(orient="records"), "ts": datetime.now().isoformat()}, to=sid)
                    except Exception as e:
                        logger.warning(f"emit failed: {e}")
        except Exception as e:
            logger.exception(f"poller error: {e}")
        time.sleep(POLL_INTERVAL)

@app.route("/")
def index():
    return send_from_directory("templates", "index.html")

@app.route("/static/<path:path>")
def static_files(path):
    return send_from_directory("static", path)

# --- 목록: 전체 보기 (GUI 기본)
@app.route("/api/projects", methods=["GET"])
def list_projects():
    email, admin = get_user_from_headers(request.headers)
    if not email:
        return jsonify({"error":"Unauthorized"}), 401
    df = _cache_df if not _cache_df.empty else gs.fetch_dataframe()
    df = _apply_query_filters(df, request.args)
    return jsonify({"rows": df.to_dict(orient="records")})

# --- 목록: 정산 전용 (본인 담당만)
@app.route("/api/settlement/projects", methods=["GET"])
def list_projects_settlement():
    email, admin = get_user_from_headers(request.headers)
    if not email:
        return jsonify({"error":"Unauthorized"}), 401
    base_df = _cache_df if not _cache_df.empty else gs.fetch_dataframe()
    df = _filter_df_for_user(base_df, email, admin)
    df = _apply_query_filters(df, request.args)
    return jsonify({"rows": df.to_dict(orient="records")})

# --- 신규 행 추가(수동)
@app.route("/api/projects", methods=["POST"])
def add_project():
    email, admin = get_user_from_headers(request.headers)
    if not email:
        return jsonify({"ok": False, "error":"Unauthorized"}), 401
    payload = request.json or {}
    required = [x.strip() for x in (_creds.get("required_fields") or os.getenv("REQUIRED_FIELDS","" )).split(",") if x.strip()]
    misses = [c for c in required if c not in payload or str(payload.get(c,"")).strip()==""]
    if misses:
        return jsonify({"ok": False, "error": f"누락 필드: {', '.join(misses)}"}), 400
    try:
        gs.append_row(payload)
        return jsonify({"ok": True})
    except Exception as e:
        logger.exception(e)
        return jsonify({"ok": False, "error": str(e)}), 500

# --- 신규 행 추가(자동 코드 생성)
@app.route("/api/projects/auto", methods=["POST"])
def add_project_auto():
    email, admin = get_user_from_headers(request.headers)
    if not email:
        return jsonify({"ok": False, "error":"Unauthorized"}), 401
    payload = request.json or {}
    company = str(payload.get("사업자","" )).strip()
    owner = str(payload.get("담당자","" )).strip()
    if not company or not owner:
        return jsonify({"ok": False, "error": "사업자/담당자는 필수"}), 400

    # 코드 생성
    df = _cache_df if not _cache_df.empty else gs.fetch_dataframe()
    try:
        code = _auto_project_code(df, company, owner)
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 400
    payload["프로젝트 코드"] = code

    # 필수 필드 검증
    required = [x.strip() for x in (_creds.get("required_fields") or os.getenv("REQUIRED_FIELDS","" )).split(",") if x.strip()]
    misses = [c for c in required if c not in payload or str(payload.get(c,"")).strip()==""]
    if misses:
        return jsonify({"ok": False, "error": f"누락 필드: {', '.join(misses)}"}), 400

    try:
        gs.append_row(payload)
        return jsonify({"ok": True, "project_code": code})
    except Exception as e:
        logger.exception(e)
        return jsonify({"ok": False, "error": str(e)}), 500

# --- 드롭다운 옵션: 사업자/담당자 자동 추출
@app.route("/api/meta/options", methods=["GET"])
def meta_options():
    email, admin = get_user_from_headers(request.headers)
    if not email:
        return jsonify({"error":"Unauthorized"}), 401
    df = _cache_df if not _cache_df.empty else gs.fetch_dataframe()
    companies = []
    owners = []
    if "사업자" in df.columns:
        companies = sorted(set(x.strip() for x in df["사업자"].astype(str) if x.strip() and x.strip() not in ("-", "없음", "N/A", "n/a")))
    if "담당자" in df.columns:
        owners = sorted(set(x.strip() for x in df["담당자"].astype(str) if x.strip() and x.strip() not in ("-", "없음", "N/A", "n/a")))
    return jsonify({"companies": companies, "owners": owners})

@socketio.on("auth")
def sock_auth(data):
    api = data.get("apiKey","" )
    userEmail = data.get("userEmail","" ).strip()
    req_headers = {"X-API-Key": api, "X-User-Email": userEmail}
    email, admin = get_user_from_headers(req_headers)
    if not email:
        emit("error", {"message":"Unauthorized"})
        disconnect()
        return
    _clients[request.sid] = {"email": email, "admin": admin}

@socketio.on("disconnect")
def sock_disconn():
    _clients.pop(request.sid, None)

@socketio.on("projects:subscribe")
def sock_subscribe(data):
    from flask import request
    ctx = _clients.get(request.sid)
    if not ctx:
        emit("error", {"message":"Unauthorized"})
        disconnect()
        return
    df = _cache_df if not _cache_df.empty else gs.fetch_dataframe()
    emit("projects:update", {"rows": df.to_dict(orient="records"), "ts": datetime.now().isoformat()})

def boot():
    global _cache_df, _cache_hash
    df = gs.fetch_dataframe()
    _cache_df = df
    _cache_hash = _df_hash(df)
    th = threading.Thread(target=poller, daemon=True)
    th.start()

if __name__ == "__main__":
    boot()
    port = int(os.getenv("PORT","5000"))
    debug = os.getenv("DEBUG","True").lower()=="true"
    logger.info(f"Dashboard http://localhost:{port}")
    socketio.run(app, host="0.0.0.0", port=port, debug=debug)
