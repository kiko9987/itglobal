from flask import Flask, render_template, jsonify, request, send_from_directory, session, redirect, url_for
from flask_socketio import SocketIO, emit
import os
import sys
import logging
import time
from datetime import datetime, timedelta
import json
import re
from collections import Counter, defaultdict
from dotenv import load_dotenv

# 인증 시스템 import
from auth import user_manager, login_required, admin_required, get_user_role
from google_oauth import google_oauth, is_oauth_configured

# 프로젝트 루트 경로를 시스템 경로에 추가
project_root = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(project_root))

from utils.google_sheets import GoogleSheetsManager
from utils.data_analyzer import DataAnalyzer

# 환경 변수 로드
load_dotenv()

# Flask 앱 초기화
app = Flask(__name__)
app.config['SECRET_KEY'] = os.getenv('FLASK_SECRET_KEY', 'your-secret-key-here')
app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(hours=8)  # 8시간 세션 유지

# SocketIO 초기화 (실시간 업데이트용)
socketio = SocketIO(app, cors_allowed_origins="*")

# 로깅 설정
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# Google API 캐시 경고 억제
logging.getLogger('googleapiclient.discovery_cache').setLevel(logging.ERROR)

# 전역 변수 (캐싱 개선)
current_data = None
last_update = None
_data_cache = {}
_cache_expiry = 60   # 1분 캐시

# 감사 로그 시스템
AUDIT_LOG_DIR = os.path.join(os.path.dirname(__file__), 'audit_logs')
if not os.path.exists(AUDIT_LOG_DIR):
    os.makedirs(AUDIT_LOG_DIR)

def normalize_log_value(value, field_name):
    """감사 로그 값을 정규화 (표기법 통일)"""
    if not value or value == 'null' or value == 'undefined':
        return '-'
    
    # 금액 필드인 경우 화폐 기호 제거하고 숫자에 쉼표 추가
    money_fields = ['총액 1', '총액 2', '계약금', '중도금', '잔금', '미수금', '제품대', '도급비', '자재비', '기타비', '순익']
    if field_name in money_fields:
        # 화폐 기호와 쉼표 제거
        cleaned_value = str(value).replace('₩', '').replace('\\', '').replace(',', '').strip()
        try:
            # 숫자로 변환 가능한지 확인
            float_val = float(cleaned_value)
            if float_val.is_integer():
                # 정수인 경우 쉼표 추가
                return f"{int(float_val):,}"
            else:
                # 소수인 경우도 쉼표 추가
                return f"{float_val:,.2f}".rstrip('0').rstrip('.')
        except ValueError:
            return str(value)
    
    return str(value)

def log_user_action(action, details, project_code=None, field_name=None, old_value=None, new_value=None):
    """사용자 행동을 감사 로그에 기록"""
    try:
        user_info = session.get('user', {})
        
        # 값 정규화
        normalized_old_value = normalize_log_value(old_value, field_name)
        normalized_new_value = normalize_log_value(new_value, field_name)
        
        log_entry = {
            'timestamp': datetime.now().isoformat(),
            'user_name': user_info.get('name', 'Unknown'),
            'user_email': user_info.get('email', 'Unknown'),
            'user_role': user_info.get('permission_level', 'Unknown'),
            'action': action,
            'details': details,
            'project_code': project_code,
            'field_name': field_name,
            'old_value': normalized_old_value,
            'new_value': normalized_new_value,
            'ip_address': request.remote_addr if request else 'Unknown'
        }
        
        # 날짜별 로그 파일 생성
        log_date = datetime.now().strftime('%Y-%m-%d')
        log_file = os.path.join(AUDIT_LOG_DIR, f'audit_{log_date}.json')
        
        # 기존 로그 읽기
        logs = []
        if os.path.exists(log_file):
            try:
                with open(log_file, 'r', encoding='utf-8') as f:
                    logs = json.load(f)
            except:
                logs = []
        
        # 새 로그 추가
        logs.append(log_entry)
        
        # 로그 파일 저장
        with open(log_file, 'w', encoding='utf-8') as f:
            json.dump(logs, f, ensure_ascii=False, indent=2)
            
        logger.info(f"Audit log recorded: {action} by {user_info.get('name', 'Unknown')}")
        
    except Exception as e:
        logger.error(f"Failed to log user action: {e}")

def get_audit_logs(days=7):
    """감사 로그 조회 (최근 N일)"""
    try:
        all_logs = []
        for i in range(days):
            log_date = (datetime.now() - timedelta(days=i)).strftime('%Y-%m-%d')
            log_file = os.path.join(AUDIT_LOG_DIR, f'audit_{log_date}.json')
            
            if os.path.exists(log_file):
                try:
                    with open(log_file, 'r', encoding='utf-8') as f:
                        daily_logs = json.load(f)
                        all_logs.extend(daily_logs)
                except:
                    continue
        
        # 최신순으로 정렬
        all_logs.sort(key=lambda x: x['timestamp'], reverse=True)
        return all_logs
        
    except Exception as e:
        logger.error(f"Failed to get audit logs: {e}")
        return []

# 프로젝트 설정 로드
def _load_project_config():
    """project_config.json에서 설정 로드"""
    try:
        config_path = os.path.join(os.path.dirname(__file__), 'project_config.json')
        with open(config_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        logger.warning(f"프로젝트 설정 로드 실패: {e}")
        return {}

_project_config = _load_project_config()

def load_data():
    """구글 시트에서 데이터 로드"""
    global current_data, last_update
    
    try:
        sheet_id = os.getenv('GOOGLE_SHEET_ID')
        if not sheet_id:
            logger.error("GOOGLE_SHEET_ID가 설정되지 않았습니다.")
            return None
        
        # 구글 시트에서 데이터 가져오기
        manager = GoogleSheetsManager()
        df = manager.get_sheet_data(sheet_id)
        
        if df.empty:
            logger.warning("구글 시트에서 데이터를 가져올 수 없습니다.")
            return None
        
        current_data = df
        last_update = datetime.now()
        
        logger.info(f"데이터 로드 완료: {len(df)}행, 업데이트 시간: {last_update}")
        return df
        
    except Exception as e:
        logger.error(f"데이터 로드 오류: {str(e)}")
        
        # 로컬 엑셀 파일로 폴백
        try:
            import pandas as pd
            excel_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'data', '아이티 공사 현황 (2).xlsx')
            df = pd.read_excel(excel_path, sheet_name='공사 현황')
            current_data = df
            last_update = datetime.now()
            logger.info(f"로컬 파일에서 데이터 로드: {len(df)}행")
            return df
        except Exception as e2:
            logger.error(f"로컬 파일 로드도 실패: {str(e2)}")
            return None

# 프로젝트 코드 자동 생성 함수들
def _extract_number(code: str):
    """프로젝트 코드에서 숫자 부분 추출"""
    m = re.match(r'[A-Z](\d{4})-', str(code))
    return int(m.group(1)) if m else None

def _suffix_from_code(code: str):
    """프로젝트 코드에서 접미사 부분 추출"""
    m = re.match(r'[A-Z]\d{4}-([A-Z]+)$', str(code))
    return m.group(1) if m else None

def _build_company_prefix_map(df):
    """사업자-접두사 매핑 구축"""
    m = {}
    # 기존 데이터에서 학습
    if '프로젝트 코드' in df.columns and '사업자' in df.columns:
        for _, row in df.iterrows():
            code = str(row.get('프로젝트 코드',''))
            comp = str(row.get('사업자','')).strip()
            mm = re.match(r'([A-Z])\d{4}-', code)
            if comp and mm and comp not in m:
                m[comp] = mm.group(1)
    
    # 설정 파일에서 로드
    config_map = _project_config.get('company_prefix_map', {})
    for k, v in config_map.items():
        m.setdefault(k, v)
    
    return m

def _build_owner_suffix_map(df):
    """담당자-접미사 매핑 구축"""
    # 설정 파일에서 기본 매핑 로드
    m = {k: str(v).upper() for k, v in _project_config.get('owner_suffix_map', {}).items()}
    
    # 기존 데이터에서 학습
    if '프로젝트 코드' in df.columns and '담당자' in df.columns:
        grouped = defaultdict(list)
        for _, row in df.iterrows():
            name = str(row.get('담당자','')).strip()
            code = str(row.get('프로젝트 코드','')).strip()
            suf = _suffix_from_code(code)
            if name and suf:
                grouped[name].append(suf)
        
        for name, arr in grouped.items():
            if arr:
                common = Counter(arr).most_common(1)[0][0]
                m.setdefault(name, common)
    
    return m

def _next_running_number(df):
    """다음 순번 찾기"""
    nums = []
    if '프로젝트 코드' in df.columns:
        for c in df['프로젝트 코드'].astype(str):
            n = _extract_number(c)
            if n is not None:
                nums.append(n)
    return (max(nums) + 1) if nums else 1

def _safe_next_running_number_with_retry(company: str, owner: str, max_retries: int = 5):
    """재시도 로직이 있는 안전한 다음 순번 찾기 (동시성 대응)"""
    import threading
    import time
    
    # 전역 락 (메모리 기반, 단일 서버용)
    if not hasattr(_safe_next_running_number_with_retry, '_lock'):
        _safe_next_running_number_with_retry._lock = threading.RLock()
    
    for attempt in range(max_retries):
        with _safe_next_running_number_with_retry._lock:
            try:
                # 구글 시트에서 최신 데이터 다시 로드
                df = load_data()
                if df is None:
                    raise Exception("데이터를 불러올 수 없습니다")
                
                # 프로젝트 코드 생성
                code = _auto_project_code(df, company, owner)
                
                # 생성된 코드가 이미 존재하는지 확인
                if '프로젝트 코드' in df.columns:
                    existing_codes = df['프로젝트 코드'].astype(str).tolist()
                    if code in existing_codes:
                        logger.warning(f"프로젝트 코드 충돌 감지: {code} (시도 {attempt + 1}/{max_retries})")
                        if attempt < max_retries - 1:
                            time.sleep(0.1 * (attempt + 1))  # 지수백오프
                            continue
                        else:
                            raise Exception(f"프로젝트 코드 생성 실패: 최대 재시도 횟수 초과 ({code})")
                
                logger.info(f"프로젝트 코드 안전 생성 완료: {code} (시도 {attempt + 1}/{max_retries})")
                return code
                
            except Exception as e:
                if attempt == max_retries - 1:
                    raise e
                time.sleep(0.1 * (attempt + 1))
                continue
    
    raise Exception("프로젝트 코드 생성 실패: 예상치 못한 오류")

def _auto_project_code(df, company: str, owner: str) -> str:
    """자동 프로젝트 코드 생성"""
    comp_map = _build_company_prefix_map(df)
    own_map = _build_owner_suffix_map(df)
    
    prefix = comp_map.get(company.strip())
    suffix = own_map.get(owner.strip())
    
    if not prefix or not suffix:
        available_companies = list(comp_map.keys())
        available_owners = list(own_map.keys())
        error_msg = f'코드 생성 실패: 회사/담당자 매핑을 확인하세요.\n'
        error_msg += f'사용 가능한 회사: {", ".join(available_companies)}\n'
        error_msg += f'사용 가능한 담당자: {", ".join(available_owners)}'
        raise ValueError(error_msg)
    
    num = _next_running_number(df)
    return f"{prefix}{num:04d}-{suffix}"

@app.route('/')
def dashboard():
    """메인 페이지 - 프로젝트 관리로 리다이렉트"""
    from flask import redirect, url_for
    return redirect(url_for('project_list'))

@app.route('/login')
def login_page():
    """로그인 페이지 (Google OAuth 전용)"""
    if 'user' in session:
        return redirect(url_for('project_list'))
    
    # OAuth가 설정되지 않은 경우 오류 표시
    oauth_enabled = is_oauth_configured()
    if not oauth_enabled:
        logger.error("Google OAuth가 설정되지 않았습니다.")
    
    return render_template('login.html', oauth_enabled=oauth_enabled)


@app.route('/auth/google')
def google_login():
    """구글 OAuth 로그인 시작"""
    try:
        if not is_oauth_configured():
            return redirect(url_for('login_page', message='oauth_not_configured'))
        
        authorization_url, state = google_oauth.get_authorization_url(
            redirect_uri=url_for('google_callback', _external=True)
        )
        
        if not authorization_url:
            return redirect(url_for('login_page', message='oauth_error'))
        
        session['oauth_state'] = state
        return redirect(authorization_url)
        
    except Exception as e:
        logger.error(f"Google OAuth 시작 오류: {str(e)}")
        return redirect(url_for('login_page', message='oauth_error'))

@app.route('/auth/callback')
def google_callback():
    """구글 OAuth 콜백"""
    try:
        if not is_oauth_configured():
            return redirect(url_for('login_page', message='oauth_not_configured'))
        
        # 상태 검증
        state = request.args.get('state')
        if state != session.get('oauth_state'):
            logger.warning("OAuth state 불일치")
            return redirect(url_for('login_page', message='oauth_error'))
        
        # 에러 체크
        error = request.args.get('error')
        if error:
            logger.warning(f"OAuth 에러: {error}")
            return redirect(url_for('login_page', message='oauth_cancelled'))
        
        # 인증 코드 가져오기
        code = request.args.get('code')
        if not code:
            return redirect(url_for('login_page', message='oauth_error'))
        
        # 사용자 정보 가져오기
        user_info = google_oauth.get_user_info(
            code, state, 
            redirect_uri=url_for('google_callback', _external=True)
        )
        
        if not user_info:
            return redirect(url_for('login_page', message='domain_not_allowed'))
        
        # 사용자 인증 및 자동 등록
        user = user_manager.authenticate_google_user(user_info)
        if not user:
            return redirect(url_for('login_page', message='login_failed'))
        
        # 세션 생성
        session['user'] = user
        session['login_time'] = datetime.now().isoformat()
        session.permanent = True
        
        # OAuth state 정리
        session.pop('oauth_state', None)
        
        logger.info(f"Google OAuth 로그인 성공: {user['email']}")
        return redirect(url_for('project_list'))
        
    except Exception as e:
        logger.error(f"Google OAuth 콜백 오류: {str(e)}")
        return redirect(url_for('login_page', message='oauth_error'))

@app.route('/logout')
def logout():
    """로그아웃 처리"""
    session.clear()
    return redirect(url_for('login_page', message='logout_success'))

@app.route('/projects')
@login_required
def project_list():
    """프로젝트 목록 페이지"""
    user_role = get_user_role()
    return render_template('project_list.html', user_role=user_role)

@app.route('/project/new')
def project_form_new():
    """새 프로젝트 등록 페이지 (기존)"""
    return render_template('project_form.html')

@app.route('/project/new-auto')
def project_form_auto():
    """새 프로젝트 등록 페이지 (자동 코드 생성)"""
    return render_template('project_form_auto.html')

@app.route('/project/edit')
def project_form_edit():
    """프로젝트 수정 페이지"""
    return render_template('project_form.html')

@app.route('/data/<path:filename>')
def serve_data_files(filename):
    """data 폴더의 정적 파일 서빙"""
    data_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'data')
    return send_from_directory(data_dir, filename)

def convert_numpy_int64(obj):
    """numpy int64를 Python int로 변환"""
    import numpy as np
    if isinstance(obj, np.int64):
        return int(obj)
    elif isinstance(obj, dict):
        return {k: convert_numpy_int64(v) for k, v in obj.items()}
    elif isinstance(obj, list):
        return [convert_numpy_int64(v) for v in obj]
    return obj

@app.route('/api/summary')
def get_summary():
    """요약 통계 API"""
    try:
        df = current_data if current_data is not None else load_data()
        if df is None:
            return jsonify({'error': '데이터를 불러올 수 없습니다.'}), 500
        
        analyzer = DataAnalyzer(df)
        summary = analyzer.get_summary_stats()
        
        # numpy int64를 Python int로 변환
        summary = convert_numpy_int64(summary)
        
        # 추가 정보
        summary['last_update'] = last_update.isoformat() if last_update else None
        summary['total_records'] = int(len(df))  # 명시적으로 int로 변환
        
        return jsonify(summary)
        
    except Exception as e:
        logger.error(f"요약 통계 API 오류: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/monthly-sales')
def get_monthly_sales():
    """월별 매출 API"""
    try:
        year = request.args.get('year', datetime.now().year, type=int)
        
        df = current_data if current_data is not None else load_data()
        if df is None:
            return jsonify({'error': '데이터를 불러올 수 없습니다.'}), 500
        
        analyzer = DataAnalyzer(df)
        monthly_sales = analyzer.get_monthly_sales(year)
        
        # JSON 직렬화 가능한 형태로 변환
        result = monthly_sales.to_dict('records') if not monthly_sales.empty else []
        
        return jsonify(result)
        
    except Exception as e:
        logger.error(f"월별 매출 API 오류: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/regional-analysis')
def get_regional_analysis():
    """지역별 분석 API"""
    try:
        df = current_data if current_data is not None else load_data()
        if df is None:
            return jsonify({'error': '데이터를 불러올 수 없습니다.'}), 500
        
        analyzer = DataAnalyzer(df)
        regional = analyzer.get_regional_analysis()
        
        result = regional.to_dict('records') if not regional.empty else []
        return jsonify(result)
        
    except Exception as e:
        logger.error(f"지역별 분석 API 오류: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/outstanding-analysis')
def get_outstanding_analysis():
    """미수금 분석 API"""
    try:
        df = current_data if current_data is not None else load_data()
        if df is None:
            return jsonify({'error': '데이터를 불러올 수 없습니다.'}), 500
        
        analyzer = DataAnalyzer(df)
        outstanding = analyzer.get_outstanding_analysis()
        
        # DataFrame을 dict로 변환
        result = {}
        for key, value in outstanding.items():
            if hasattr(value, 'to_dict'):
                result[key] = value.to_dict('records')
            else:
                result[key] = value
        
        return jsonify(result)
        
    except Exception as e:
        logger.error(f"미수금 분석 API 오류: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/missing-data')
def get_missing_data():
    """누락 데이터 분석 API"""
    try:
        df = current_data if current_data is not None else load_data()
        if df is None:
            return jsonify({'error': '데이터를 불러올 수 없습니다.'}), 500
        
        analyzer = DataAnalyzer(df)
        missing = analyzer.check_missing_data()
        
        # DataFrame을 dict로 변환 (JSON 직렬화 가능하도록)
        import json
        result = {}
        for key, value in missing.items():
            if hasattr(value, 'to_dict'):
                result[key] = value.to_dict('records')
            elif isinstance(value, dict):
                # dict 내부의 numpy 타입들을 Python 기본 타입으로 변환
                converted_dict = {}
                for k, v in value.items():
                    if hasattr(v, 'item'):  # numpy 타입인 경우
                        converted_dict[k] = v.item()
                    elif hasattr(v, 'tolist'):  # numpy 배열인 경우
                        converted_dict[k] = v.tolist()
                    else:
                        converted_dict[k] = v
                result[key] = converted_dict
            else:
                # numpy 타입을 Python 기본 타입으로 변환
                if hasattr(value, 'item'):
                    result[key] = value.item()
                elif hasattr(value, 'tolist'):
                    result[key] = value.tolist()
                else:
                    result[key] = value
        
        return jsonify(result)
        
    except Exception as e:
        logger.error(f"누락 데이터 분석 API 오류: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/brand-analysis')
def get_brand_analysis():
    """브랜드별 분석 API"""
    try:
        df = current_data if current_data is not None else load_data()
        if df is None:
            return jsonify({'error': '데이터를 불러올 수 없습니다.'}), 500
        
        analyzer = DataAnalyzer(df)
        brands = analyzer.get_brand_analysis()
        
        result = brands.to_dict('records') if not brands.empty else []
        return jsonify(result)
        
    except Exception as e:
        logger.error(f"브랜드별 분석 API 오류: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/projects/auto', methods=['POST'])
def add_project_auto():
    """신규 프로젝트 자동 코드 생성 및 추가"""
    try:
        data = request.get_json()
        logger.info(f"add_project_auto에서 받은 데이터: {data}")
        
        company = str(data.get("사업자", "")).strip()
        owner = str(data.get("담당자", "")).strip()
        
        if not company or not owner:
            return jsonify({"ok": False, "error": "사업자/담당자는 필수입니다"}), 400

        # 현재 데이터 로드
        df = current_data if current_data is not None else load_data()
        if df is None:
            return jsonify({"ok": False, "error": "데이터를 불러올 수 없습니다"}), 500
        
        # 프로젝트 코드 처리 - 폼에서 전송된 코드가 있으면 우선 사용
        if "프로젝트 코드" in data and str(data["프로젝트 코드"]).strip():
            code = str(data["프로젝트 코드"]).strip()
            logger.info(f"폼에서 받은 프로젝트 코드 사용: {code}")
        else:
            # 안전한 자동 프로젝트 코드 생성 (동시성 대응)
            try:
                code = _safe_next_running_number_with_retry(company, owner)
                logger.info(f"자동 생성된 프로젝트 코드: {code}")
            except Exception as e:
                return jsonify({"ok": False, "error": str(e)}), 400
        
        data["프로젝트 코드"] = code
        
        # 필수 필드 검증
        required_fields = _project_config.get("required_fields", ["프로젝트 코드", "현장 주소"])
        missing_fields = [field for field in required_fields 
                         if field not in data or str(data.get(field, "")).strip() == ""]
        
        if missing_fields:
            return jsonify({"ok": False, "error": f"필수 필드 누락: {', '.join(missing_fields)}"}), 400

        # Google Sheets에 추가 (최종 중복 확인 포함)
        try:
            sheet_id = os.getenv('GOOGLE_SHEET_ID')
            if not sheet_id:
                return jsonify({"ok": False, "error": "GOOGLE_SHEET_ID가 설정되지 않았습니다"}), 500
            
            manager = GoogleSheetsManager()
            
            # 등록 직전 최종 중복 확인
            latest_df = load_data()
            if latest_df is not None and '프로젝트 코드' in latest_df.columns:
                existing_codes = latest_df['프로젝트 코드'].astype(str).tolist()
                if code in existing_codes:
                    logger.error(f"등록 직전 프로젝트 코드 중복 감지: {code}")
                    return jsonify({"ok": False, "error": f"프로젝트 코드가 중복됩니다: {code}. 다시 시도해주세요."}), 409
            
            values = convert_form_data_to_sheet_row(data, manager)
            manager.append_row(sheet_id, values)
            
            # 로컬 데이터 새로고침
            load_data()
            
            # 실시간 업데이트 알림
            socketio.emit('data_updated', {
                'message': f"새 프로젝트가 등록되었습니다: {code}",
                'timestamp': datetime.now().isoformat(),
                'action': 'create'
            })
            
            return jsonify({"ok": True, "project_code": code})
            
        except Exception as e:
            logger.error(f"Google Sheets 추가 오류: {str(e)}")
            return jsonify({"ok": False, "error": str(e)}), 500
        
    except Exception as e:
        logger.error(f"프로젝트 자동 생성 API 오류: {str(e)}")
        return jsonify({"ok": False, "error": str(e)}), 500

@app.route('/api/preview-project-code')
def preview_project_code():
    """프로젝트 코드 미리보기 생성"""
    try:
        company = request.args.get('company', '').strip()
        owner = request.args.get('owner', '').strip()
        
        if not company or not owner:
            return jsonify({"ok": False, "error": "사업자와 담당자가 필요합니다"})

        # 현재 데이터 로드
        df = current_data if current_data is not None else load_data()
        if df is None:
            return jsonify({"ok": False, "error": "데이터를 불러올 수 없습니다"})
        
        # 자동 프로젝트 코드 생성
        try:
            code = _auto_project_code(df, company, owner)
            return jsonify({"ok": True, "project_code": code})
        except Exception as e:
            return jsonify({"ok": False, "error": str(e)})
        
    except Exception as e:
        logger.error(f"프로젝트 코드 미리보기 오류: {e}")
        return jsonify({"ok": False, "error": "서버 오류가 발생했습니다"})

@app.route('/api/meta/options', methods=['GET'])
def get_meta_options():
    """드롭다운용 옵션 API (사업자, 담당자 목록)"""
    try:
        df = current_data if current_data is not None else load_data()
        if df is None:
            return jsonify({'error': '데이터를 불러올 수 없습니다.'}), 500
        
        # 사업자 목록 추출
        companies = []
        if "사업자" in df.columns:
            companies = sorted(set(x.strip() for x in df["사업자"].astype(str) 
                                 if x.strip() and x.strip() not in ("-", "없음", "N/A", "n/a")))
        
        # 담당자 목록 추출
        owners = []
        if "담당자" in df.columns:
            owners = sorted(set(x.strip() for x in df["담당자"].astype(str) 
                              if x.strip() and x.strip() not in ("-", "없음", "N/A", "n/a")))
        
        # 설정 파일에서도 추가
        config_companies = list(_project_config.get('company_prefix_map', {}).keys())
        config_owners = list(_project_config.get('owner_suffix_map', {}).keys())
        
        companies = sorted(set(companies + config_companies))
        owners = sorted(set(owners + config_owners))
        
        # 공사 구분, 기계 분류, 브랜드 추가 (2800-2803행 기준)
        work_categories = []
        machine_types = []
        brands = []
        
        try:
            # 2800-2803행 데이터 추출 (0-based index이므로 2799-2802)
            if len(df) >= 2803:
                sample_rows = df.iloc[2799:2803]
                
                if "공사 구분" in df.columns:
                    work_categories = sorted(set(x.strip() for x in sample_rows["공사 구분"].astype(str) 
                                               if x.strip() and x.strip() not in ("-", "없음", "N/A", "n/a")))
                
                if "기계 분류" in df.columns:  
                    machine_types = sorted(set(x.strip() for x in sample_rows["기계 분류"].astype(str)
                                             if x.strip() and x.strip() not in ("-", "없음", "N/A", "n/a")))
                
                if "브랜드" in df.columns:
                    brands = sorted(set(x.strip() for x in sample_rows["브랜드"].astype(str)
                                      if x.strip() and x.strip() not in ("-", "없음", "N/A", "n/a")))
        except Exception as e:
            logger.warning(f"샘플 데이터 추출 오류: {e}")
        
        return jsonify({
            "companies": companies,
            "owners": owners,
            "work_categories": work_categories,
            "machine_types": machine_types,
            "brands": brands
        })
        
    except Exception as e:
        logger.error(f"옵션 API 오류: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/refresh-data')
def refresh_data():
    """데이터 새로고침 API"""
    try:
        df = load_data()
        if df is None:
            return jsonify({'error': '데이터 새로고침 실패'}), 500
        
        # 실시간 업데이트 알림
        socketio.emit('data_updated', {
            'message': '데이터가 업데이트되었습니다.',
            'timestamp': last_update.isoformat() if last_update else None,
            'record_count': len(df)
        })
        
        return jsonify({
            'message': '데이터 새로고침 완료',
            'timestamp': last_update.isoformat() if last_update else None,
            'formatted_time': last_update.strftime('%Y-%m-%d %H:%M:%S') if last_update else None,
            'record_count': len(df)
        })
        
    except Exception as e:
        logger.error(f"데이터 새로고침 오류: {str(e)}")
        return jsonify({'error': str(e)}), 500

@socketio.on('connect')
def handle_connect():
    """클라이언트 연결 처리"""
    logger.info('클라이언트가 연결되었습니다.')
    emit('connected', {'message': '대시보드에 연결되었습니다.'})

@socketio.on('disconnect')
def handle_disconnect():
    """클라이언트 연결 해제 처리"""
    logger.info('클라이언트 연결이 해제되었습니다.')

@app.route('/api/projects/list')
def get_projects_list():
    """프로젝트 목록 API"""
    try:
        df = current_data if current_data is not None else load_data()
        if df is None:
            return jsonify({'error': '데이터를 불러올 수 없습니다.'}), 500
        
        # DataFrame을 dict 리스트로 변환 (NaN 값 처리)
        df = df.fillna('')  # NaN 값을 빈 문자열로 변환
        
        # 날짜 컬럼들을 문자열로 변환
        date_columns = ['공사 시작', '공사 종료', '수금 날짜', '공사 확정']
        for col in date_columns:
            if col in df.columns:
                df[col] = df[col].astype(str).replace('NaT', '').replace('nan', '')
        
        projects = df.to_dict('records')
        
        return jsonify(projects)
        
    except Exception as e:
        logger.error(f"프로젝트 목록 API 오류: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/next-project-code')
def get_next_project_code():
    """다음 프로젝트 코드 생성 API"""
    try:
        region_code = request.args.get('region', 'IT')
        
        sheet_id = os.getenv('GOOGLE_SHEET_ID')
        if not sheet_id:
            return jsonify({'error': 'GOOGLE_SHEET_ID가 설정되지 않았습니다.'}), 500
        
        manager = GoogleSheetsManager()
        project_code = manager.get_next_project_code(sheet_id, region_code)
        
        return jsonify({'project_code': project_code})
        
    except Exception as e:
        logger.error(f"프로젝트 코드 생성 API 오류: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/projects', methods=['POST'])
def create_project():
    """새 프로젝트 생성 API"""
    global current_data
    try:
        data = request.get_json()
        logger.info(f"받은 폼 데이터: {data}")
        
        sheet_id = os.getenv('GOOGLE_SHEET_ID')
        if not sheet_id:
            return jsonify({'error': 'GOOGLE_SHEET_ID가 설정되지 않았습니다.'}), 500
        
        manager = GoogleSheetsManager()
        
        # 데이터를 구글 시트 형식으로 변환
        values = convert_form_data_to_sheet_row(data, manager)
        logger.info(f"변환된 values 배열: {values[:10]}... (총 {len(values)}개)")
        
        # 구글 시트에 추가 시도
        result = manager.append_row(sheet_id, values)
        
        # 구글 시트 쓰기 실패 여부 확인 (로컬 전용 모드인지 체크)
        is_local_only = False
        if result and result.get('updates', {}).get('updatedRange', '').endswith('9999:AM9999'):
            is_local_only = True
            logger.info("로컬 전용 모드: 새 프로젝트를 로컬 DataFrame에 직접 추가")
            
            # 로컬 DataFrame에 새 데이터 직접 추가
            try:
                # 컬럼 매핑 가져오기
                column_mapping = manager.get_column_mapping()
                
                # values 배열을 DataFrame 행으로 변환
                new_row_data = {}
                for column_letter, column_name in column_mapping.items():
                    column_index = ord(column_letter) - ord('A') if len(column_letter) == 1 else \
                                  (ord(column_letter[0]) - ord('A') + 1) * 26 + (ord(column_letter[1]) - ord('A'))
                    
                    if column_index < len(values):
                        value = values[column_index]
                        
                        # 데이터 타입 변환 (GoogleSheetsManager의 _preprocess_data와 일치)
                        if column_name in ['공사 시작', '공사 종료', '수금 날짜', '공사 확정']:
                            if value and value != '':
                                try:
                                    new_row_data[column_name] = pd.to_datetime(value, errors='coerce')
                                except:
                                    new_row_data[column_name] = pd.NaType()
                            else:
                                new_row_data[column_name] = pd.NaType()
                        elif column_name in ['총액 1', '총액 2', '총액2', '계약금', '중도금', '잔금', 
                                            '미수금', '미수금W', '제품대', '도급비', '자재비', '기타비']:
                            if value and value != '':
                                try:
                                    # 쉼표, 원화기호 제거 후 숫자 변환
                                    clean_value = str(value).replace(',', '').replace('￦', '').replace('₩', '').replace('-', '').strip()
                                    new_row_data[column_name] = pd.to_numeric(clean_value, errors='coerce') if clean_value else pd.NA
                                except:
                                    new_row_data[column_name] = pd.NA
                            else:
                                new_row_data[column_name] = pd.NA
                        elif column_name in ['부가세', '수금 확인']:
                            if value == 'TRUE':
                                new_row_data[column_name] = True
                            elif value == 'FALSE':
                                new_row_data[column_name] = False
                            else:
                                new_row_data[column_name] = False
                        else:
                            new_row_data[column_name] = value if value != '' else None
                    else:
                        new_row_data[column_name] = None
                
                # 현재 데이터가 있으면 새 행 추가, 없으면 새 DataFrame 생성
                if current_data is not None and not current_data.empty:
                    # 새 행을 DataFrame에 추가
                    import pandas as pd
                    new_row_df = pd.DataFrame([new_row_data])
                    current_data = pd.concat([current_data, new_row_df], ignore_index=True)
                else:
                    # 현재 데이터가 없으면 구글 시트에서 다시 로드 시도
                    load_data()
                    if current_data is not None and not current_data.empty:
                        new_row_df = pd.DataFrame([new_row_data])
                        current_data = pd.concat([current_data, new_row_df], ignore_index=True)
                
                logger.info(f"로컬 DataFrame에 새 프로젝트 추가 완료: {data.get('projectCode', '')}")
                
            except Exception as local_error:
                logger.error(f"로컬 DataFrame 업데이트 오류: {str(local_error)}")
                # 실패해도 기존 로드 방식으로 폴백
                load_data()
        else:
            # 구글 시트 쓰기 성공 - 일반적인 데이터 새로고침
            load_data()
        
        # 실시간 업데이트 알림
        socketio.emit('data_updated', {
            'message': f"새 프로젝트가 등록되었습니다: {data.get('projectCode', '')}",
            'timestamp': datetime.now().isoformat(),
            'action': 'create',
            'local_only': is_local_only
        })
        
        return jsonify({
            'success': True, 
            'project_code': data.get('projectCode', ''),
            'local_only': is_local_only
        })
        
    except Exception as e:
        logger.error(f"프로젝트 생성 API 오류: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/projects/<project_code>', methods=['GET'])
def get_project(project_code):
    """프로젝트 상세 정보 API"""
    try:
        df = current_data if current_data is not None else load_data()
        if df is None:
            return jsonify({'error': '데이터를 불러올 수 없습니다.'}), 500
        
        # 프로젝트 코드로 찾기
        project_row = df[df['프로젝트 코드'] == project_code]
        
        if project_row.empty:
            return jsonify({'error': '프로젝트를 찾을 수 없습니다.'}), 404
        
        project = project_row.iloc[0].to_dict()
        
        return jsonify(project)
        
    except Exception as e:
        logger.error(f"프로젝트 조회 API 오류: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/projects/<project_code>', methods=['PUT'])
def update_project(project_code):
    """프로젝트 수정 API"""
    try:
        data = request.get_json()
        
        sheet_id = os.getenv('GOOGLE_SHEET_ID')
        if not sheet_id:
            return jsonify({'error': 'GOOGLE_SHEET_ID가 설정되지 않았습니다.'}), 500
        
        manager = GoogleSheetsManager()
        
        # 프로젝트가 있는 행 찾기 (직접 구현)
        search_result = manager.service.spreadsheets().values().get(
            spreadsheetId=sheet_id,
            range='공사 현황!A:A'
        ).execute()
        
        values = search_result.get('values', [])
        row_number = None
        
        for i, row in enumerate(values):
            if row and len(row) > 0 and row[0] == project_code:
                row_number = i + 1  # 1부터 시작
                break
        
        if not row_number:
            return jsonify({'error': '프로젝트를 찾을 수 없습니다.'}), 404
        
        # 인라인 편집 데이터인지 확인 (한국어 필드명 포함)
        korean_fields = ['현장 주소', '사업자', '현장 담당자', '도급 구분', '담당자 연락처', '시공자', '담당자 이메일', '견적서 및 계약서 폴더 경로']
        is_inline_data = any(field in data for field in korean_fields)
        
        if is_inline_data:
            # 인라인 편집 데이터 - 배치 업데이트 방식 사용
            updates = []
            field_column_mapping = {
                # 기본정보
                '사업자': 'B',
                '현장 담당자': 'N', 
                '도급 구분': 'L',
                '담당자 연락처': 'O',
                '시공자': 'M',
                '담당자 이메일': 'P',
                '현장 주소': 'E',
                # 공사정보
                '공사 구분': 'F',
                '기계 분류': 'G',
                '브랜드': 'H',
                '공사 시작': 'I',
                '공사 종료': 'J',
                '공사 내용': 'K',
                '공사 확정': 'AL',
                # 문서 정보
                '견적서 및 계약서 폴더 경로': 'AK'
            }
            
            for field_name, value in data.items():
                if field_name in field_column_mapping:
                    column = field_column_mapping[field_name]
                    updates.append({
                        'range': f'공사 현황!{column}{row_number}',
                        'values': [[value]]
                    })
            
            # 업데이트 전에 현재 값들을 조회해서 이전 값으로 기록
            old_values = {}
            try:
                for field_name in data.keys():
                    if field_name in field_column_mapping:
                        column = field_column_mapping[field_name]
                        current_value_result = manager.service.spreadsheets().values().get(
                            spreadsheetId=sheet_id,
                            range=f'공사 현황!{column}{row_number}'
                        ).execute()
                        current_values = current_value_result.get('values', [['']])
                        old_values[field_name] = current_values[0][0] if current_values and current_values[0] else ''
            except Exception as e:
                logger.warning(f"이전 값 조회 실패: {e}")
            
            if updates:
                batch_update_body = {
                    'valueInputOption': 'USER_ENTERED',
                    'data': updates
                }
                result = manager.service.spreadsheets().values().batchUpdate(
                    spreadsheetId=sheet_id,
                    body=batch_update_body
                ).execute()
                
                # 감사 로그 기록 (실제 이전 값 포함)
                try:
                    for field_name, new_value in data.items():
                        if field_name in field_column_mapping and field_name.strip() != '':
                            old_value = old_values.get(field_name, '')
                            log_user_action(
                                action='UPDATE_FIELD',
                                details=f'프로젝트 필드 수정: {field_name}',
                                project_code=project_code,
                                field_name=field_name,
                                old_value=old_value if old_value else '-',
                                new_value=str(new_value) if new_value else '-'
                            )
                except Exception as log_error:
                    logger.warning(f"감사 로그 기록 실패: {log_error}")
        else:
            # 기존 폼 데이터 - 전체 행 업데이트
            values = convert_form_data_to_sheet_row(data, manager)
            result = manager.update_row(sheet_id, row_number, values)
        
        # 로컬 데이터 새로고침
        load_data()
        
        # 실시간 업데이트 알림
        socketio.emit('data_updated', {
            'message': f"프로젝트가 수정되었습니다: {project_code}",
            'timestamp': datetime.now().isoformat(),
            'action': 'update'
        })
        
        return jsonify({'ok': True, 'success': True, 'project_code': project_code})
        
    except Exception as e:
        logger.error(f"프로젝트 수정 API 오류: {str(e)}")
        return jsonify({'ok': False, 'error': str(e)}), 500

# 삭제 기능 제거 (사용자 요청에 따라)
# @app.route('/api/projects/<project_code>', methods=['DELETE'])
# def delete_project(project_code):
#     """프로젝트 삭제 API (구글 시트에서는 빈 행으로 만들기)"""

def convert_form_data_to_sheet_row(form_data, manager):
    """폼 데이터를 구글 시트 행 형식으로 변환"""
    logger.info(f"convert_form_data_to_sheet_row 호출됨. 입력 데이터: {form_data}")
    
    column_mapping = manager.get_column_mapping()
    
    # 폼 필드명을 구글 시트 컬럼명으로 매핑 (영문명과 한글명 모두 지원)
    field_mapping = {
        # 영문 필드명 (기존 호환성)
        'projectCode': '프로젝트 코드',
        'company': '사업자',
        'region': '담당자',
        'client': '거래처',
        'address': '현장 주소',
        'workType': '공사 구분',
        'equipmentType': '기계 분류',
        'brand': '브랜드',
        'startDate': '공사 시작',
        'endDate': '공사 종료',
        'workDescription': '공사 내용',
        'contractType': '도급 구분',
        'constructor': '시공자',
        'siteManager': '현장 담당자',
        'managerPhone': '담당자 연락처',
        'managerEmail': '담당자 이메일',
        'amount1': '총액 1',
        'vatIncluded': '부가세',
        'amount2': '총액 2',
        'downPayment': '계약금',
        'middlePayment': '중도금',
        'finalPayment': '잔금',
        'outstanding': '미수금',
        'invoice': '계산서',
        'paymentDate': '수금 날짜',
        'paymentConfirmed': '수금 확인',
        'productCost': '제품대',
        'laborCost': '도급비',
        'materialCost': '자재비',
        'otherCost': '기타비',
        'netProfit': '순익',
        'marginRate': '마진율',
        'notes': '비고',
        'downPaymentPayer': '계약금 입금자명',
        'middlePaymentPayer': '중도금 입금자명',
        'finalPaymentPayer': '잔금 입금자명',
        
        # 한글 필드명 (project_list.html의 새 프로젝트 폼)
        '사업자': '사업자',
        '담당자': '담당자',
        '거래처': '거래처',
        '현장 주소': '현장 주소',
        '공사 구분': '공사 구분',
        '기계 분류': '기계 분류',
        '브랜드': '브랜드',
        '공사 시작': '공사 시작',
        '공사 종료': '공사 종료',
        '공사 내용': '공사 내용',
        '도급 구분': '도급 구분',
        '시공자': '시공자',
        '현장 담당자': '현장 담당자',
        '담당자 연락처': '담당자 연락처',
        '담당자 이메일': '담당자 이메일',
        '총액1': '총액 1',
        '총액 1': '총액 1', 
        '부가세 포함': '부가세',
        # '총액2': '총액 2',  # 제거: 총액2는 무조건 수식으로만 처리
        # '총액 2': '총액 2',  # 제거: 총액2는 무조건 수식으로만 처리
        '계약금': '계약금',
        '중도금': '중도금',
        '잔금': '잔금',
        '미수금': '미수금',
        '계산서': '계산서',
        '수금 날짜': '수금 날짜',
        '수금 확인': '수금 확인',
        '제품대': '제품대',
        '도급비': '도급비',
        '자재비': '자재비',
        '기타비': '기타비',
        '순익': '순익',
        '마진율': '마진율',
        '비고': '비고'
    }
    
    # 39개 컬럼에 맞춰 빈 리스트 생성
    values = [''] * 39
    
    # 각 컬럼에 해당하는 값 설정
    for column_letter, column_name in column_mapping.items():
        column_index = ord(column_letter) - ord('A') if len(column_letter) == 1 else \
                      (ord(column_letter[0]) - ord('A') + 1) * 26 + (ord(column_letter[1]) - ord('A'))
        
        # 프로젝트 코드가 없는 경우 자동 생성
        if column_name == '프로젝트 코드':
            if 'project_code' in form_data:
                values[column_index] = str(form_data['project_code'])
            elif '프로젝트 코드' in form_data and form_data['프로젝트 코드'].strip():
                values[column_index] = str(form_data['프로젝트 코드'])
                logger.info(f"폼에서 받은 프로젝트 코드: {form_data['프로젝트 코드']}")
            else:
                # 자동 생성: 담당자 정보를 기반으로 프로젝트 코드 생성
                region_code = 'IT'  # 기본값
                if '담당자' in form_data:
                    manager_name = form_data['담당자']
                    if manager_name in ['박정우', '정우']:
                        region_code = 'JW'
                    elif manager_name in ['박용구', '용구']:
                        region_code = 'YG' 
                    elif manager_name in ['황샛별', '샛별']:
                        region_code = 'SB'
                    elif manager_name in ['고광일', '광일']:
                        region_code = 'IT'
                
                # 프로젝트 코드 생성
                next_code = manager.get_next_project_code(os.getenv('GOOGLE_SHEET_ID'), region_code)
                values[column_index] = next_code
                logger.info(f"자동 생성된 프로젝트 코드: {next_code}")
            continue
            
        # 공사 확정 필드에 자동으로 등록 날짜 저장
        if column_name == '공사 확정':
            from datetime import datetime
            values[column_index] = datetime.now().strftime('%Y-%m-%d')
            continue
        
        # 폼 데이터에서 해당 값 찾기 - 다단계 매칭 전략
        form_field = None
        value = None
        
        # 1단계: 직접 매칭 (컬럼명이 폼 데이터에 그대로 있는 경우)
        if column_name in form_data:
            form_field = column_name
            value = form_data[column_name]
            logger.info(f"직접 매핑 성공: '{column_name}' -> 폼 데이터에서 직접 발견")
        
        # 2단계: 매핑 테이블을 통한 매칭
        if not form_field:
            for form_key, sheet_column in field_mapping.items():
                if sheet_column == column_name and form_key in form_data:
                    form_field = form_key
                    value = form_data[form_key]
                    logger.info(f"매핑 테이블 매칭 성공: '{form_key}' -> '{column_name}'")
                    break
        
        # 3단계: 특별한 한글 필드들을 위한 직접 매핑
        if not form_field:
            # 총액 관련 필드들
            if column_name == '총액 1':
                if '총액1' in form_data:
                    form_field = '총액1'
                    value = form_data['총액1']
                    logger.info(f"특별 매핑 성공: '총액1' -> '총액 1'")
                elif 'amount1' in form_data:
                    form_field = 'amount1'
                    value = form_data['amount1']
                    logger.info(f"특별 매핑 성공: 'amount1' -> '총액 1'")
# 총액 2 특별 처리 제거 - 강제 기본값에서만 처리함
            elif column_name == '부가세':
                if '부가세 포함' in form_data:
                    form_field = '부가세 포함'
                    value = form_data['부가세 포함']
                    logger.info(f"특별 매핑 성공: '부가세 포함' -> '부가세' (값: {value})")
                elif 'vatIncluded' in form_data:
                    form_field = 'vatIncluded'
                    value = form_data['vatIncluded']
                    logger.info(f"특별 매핑 성공: 'vatIncluded' -> '부가세' (값: {value})")
                else:
                    # 부가세 기본값 설정 (폼에서 체크박스가 기본적으로 체크된 상태)
                    form_field = 'default_vat'
                    value = True  # 기본값: 부가세 포함
                    logger.info(f"부가세 기본값 설정: True (체크됨)")
        
        # 매핑 실패 시 기본값 설정
        if not form_field:
            logger.info(f"매핑 실패: '{column_name}', 폼 데이터 키들: {list(form_data.keys())}")
            
            # 폼에서 보내지 않는 필드들에 대한 기본값 설정
            if column_name == '계약금':
                form_field = 'default_downPayment'
                value = '₩0'
                logger.info(f"기본값 설정: '계약금' -> '₩0'")
            elif column_name == '중도금':
                form_field = 'default_middlePayment'  
                value = '₩0'
                logger.info(f"기본값 설정: '중도금' -> '₩0'")
            elif column_name == '잔금':
                form_field = 'default_finalPayment'
                value = '₩0'
                logger.info(f"기본값 설정: '잔금' -> '₩0'")
            elif column_name == '미수금':
                form_field = 'default_outstanding'
                value = '=0-(S:S-(T:T+U:U+V:V))'
                logger.info(f"기본값 설정: '미수금' -> 수식 '{value}'")
            elif column_name == '계산서':
                form_field = 'default_invoice'
                value = '미발행'
                logger.info(f"기본값 설정: '계산서' -> '미발행'")
            elif column_name == '수금 확인':
                form_field = 'default_paymentConfirmed'
                value = False
                logger.info(f"기본값 설정: '수금 확인' -> False (체크박스 해제)")
            elif column_name == '순익':
                form_field = 'default_netProfit'
                value = '=$Q:Q-(($AA:AA)+($AB:AB)+($AC:AC)+($AD:AD))'
                logger.info(f"기본값 설정: '순익' -> 수식 '{value}'")
            elif column_name == '마진율':
                form_field = 'default_marginRate'
                value = '=(($AE:AE)/($Q:Q))'
                logger.info(f"기본값 설정: '마진율' -> 수식 '{value}'")
        
        if form_field and value is not None:
            logger.info(f"최종 매핑 발견: {form_field} -> {column_name} (인덱스 {column_index}): '{value}'")
            
            # 데이터 타입별 처리
            if isinstance(value, bool):
                # 체크박스 필드들은 TRUE/FALSE로 처리
                values[column_index] = 'TRUE' if value else 'FALSE'
            elif isinstance(value, (int, float)):
                # 금액 필드인 경우 포맷팅 적용
                if column_name in ['총액 1', '계약금', '중도금', '잔금', '제품대', '도급비', '자재비', '기타비', '순익']:
                    if value != 0:
                        # 숫자를 원화 형태로 포맷팅
                        formatted_value = f"₩{int(value):,}"
                        values[column_index] = formatted_value
                    else:
                        values[column_index] = ''
                else:
                    values[column_index] = str(value) if value != 0 else ''
            else:
                # 문자열 타입 처리
                # 수식인 경우 (=로 시작하는 경우) 그대로 전달
                if str(value).startswith('='):
                    values[column_index] = str(value)
                    logger.info(f"수식 전달: '{value}' -> '{values[column_index]}'")
                # 금액 필드 확인
                elif column_name in ['총액 1', '계약금', '중도금', '잔금', '제품대', '도급비', '자재비', '기타비', '순익']:
                    try:
                        # 문자열에서 숫자만 추출하여 포맷팅
                        numeric_value = float(str(value).replace(',', '').replace('₩', '').strip())
                        if numeric_value != 0:
                            formatted_value = f"₩{int(numeric_value):,}"
                            values[column_index] = formatted_value
                        else:
                            values[column_index] = ''
                    except (ValueError, TypeError):
                        values[column_index] = str(value) if value else ''
                else:
                    values[column_index] = str(value) if value else ''
                    
                # 디버깅: 문자열 변환 확인
                logger.info(f"문자열 변환: '{value}' -> '{values[column_index]}' (타입: {type(value)})")
        else:
            if form_field:
                logger.info(f"매핑 없음: {form_field} -> {column_name} (필드가 폼 데이터에 없음)")
    
    # 강제로 기본값 설정 - 폼에서 보내지 않는 필수 필드들
    for column_letter, column_name in column_mapping.items():
        column_index = ord(column_letter) - ord('A') if len(column_letter) == 1 else \
                      (ord(column_letter[0]) - ord('A') + 1) * 26 + (ord(column_letter[1]) - ord('A'))
        
        # 필수 필드들에 대한 강제 설정 (값이 있어도 덮어쓰는 필드들)
        if column_name == '총액 2':
            # 총액 2는 항상 수식으로 덮어쓰기 (기존 값 무시)
            values[column_index] = '=IF($R:R=TRUE,($Q:Q*1.1),$Q:Q)'
            logger.info(f"강제 수식 덮어쓰기: '총액 2' -> 수식")
        
        # 빈 값인 필드들에 대해 기본값 강제 설정
        elif not values[column_index] or str(values[column_index]).strip() == '':
            if column_name == '계약금':
                values[column_index] = '₩0'
                logger.info(f"강제 기본값 설정: '계약금' -> '₩0'")
            elif column_name == '중도금':
                values[column_index] = '₩0'
                logger.info(f"강제 기본값 설정: '중도금' -> '₩0'")
            elif column_name == '잔금':
                values[column_index] = '₩0'
                logger.info(f"강제 기본값 설정: '잔금' -> '₩0'")
            elif column_name == '미수금':
                values[column_index] = '=0-(S:S-(T:T+U:U+V:V))'
                logger.info(f"강제 수식 설정: '미수금' -> 수식")
            elif column_name == '계산서':
                values[column_index] = '미발행'
                logger.info(f"강제 기본값 설정: '계산서' -> '미발행'")
            elif column_name == '순익':
                values[column_index] = '=$Q:Q-(($AA:AA)+($AB:AB)+($AC:AC)+($AD:AD))'
                logger.info(f"강제 수식 설정: '순익' -> 수식")
            elif column_name == '마진율':
                values[column_index] = '=(($AE:AE)/($Q:Q))'
                logger.info(f"강제 수식 설정: '마진율' -> 수식")
            elif column_name == '수금 확인':
                values[column_index] = 'FALSE'
                logger.info(f"강제 기본값 설정: '수금 확인' -> 'FALSE'")

    # 최종 변환 결과 디버깅
    logger.info(f"최종 변환 결과 (처음 10개): {values[:10]}")
    non_empty_count = sum(1 for v in values if v and str(v).strip())
    logger.info(f"비어있지 않은 필드 수: {non_empty_count}/{len(values)}")
    
    return values

@app.route('/api/update-project-inline', methods=['POST'])
def update_project_inline():
    """프로젝트 인라인 편집 API - 구글 시트 직접 업데이트"""
    try:
        data = request.get_json()
        project_code = data.get('projectCode') or data.get('프로젝트 코드')
        
        if not project_code:
            return jsonify({'ok': False, 'error': '프로젝트 코드가 필요합니다.'}), 400
        
        sheet_id = os.getenv('GOOGLE_SHEET_ID')
        if not sheet_id:
            return jsonify({'ok': False, 'error': 'GOOGLE_SHEET_ID가 설정되지 않았습니다.'}), 500
        
        manager = GoogleSheetsManager()
        
        # 프로젝트가 있는 행 찾기
        row_number = manager.find_row_by_project_code(sheet_id, project_code)
        
        if not row_number:
            return jsonify({'ok': False, 'error': '프로젝트를 찾을 수 없습니다.'}), 404
        
        # 현재 행의 데이터를 가져오기 (전체 행 데이터 보존을 위해)
        current_row_range = f'공사 현황!A{row_number}:AM{row_number}'
        result = manager.service.spreadsheets().values().get(
            spreadsheetId=sheet_id,
            range=current_row_range
        ).execute()
        
        current_values = result.get('values', [[]])[0] if result.get('values') else []
        
        # 현재 값을 리스트로 확장 (39개 컬럼)
        while len(current_values) < 39:
            current_values.append('')
        
        # 컬럼 매핑 가져오기
        column_mapping = manager.get_column_mapping()
        
        # 업데이트할 필드만 변경
        for field_name, new_value in data.items():
            if field_name == '프로젝트 코드':
                continue  # 프로젝트 코드는 변경하지 않음
            
            # 필드명에 해당하는 컬럼 인덱스 찾기
            column_index = None
            for col_letter, col_name in column_mapping.items():
                if col_name == field_name:
                    # 컬럼 문자를 인덱스로 변환
                    if len(col_letter) == 1:
                        column_index = ord(col_letter) - ord('A')
                    else:
                        column_index = (ord(col_letter[0]) - ord('A') + 1) * 26 + (ord(col_letter[1]) - ord('A'))
                    break
            
            if column_index is not None and column_index < len(current_values):
                # 값 업데이트
                if new_value == '-' or new_value == '':
                    current_values[column_index] = ''
                else:
                    current_values[column_index] = str(new_value)
        
        # 구글 시트 업데이트
        update_result = manager.update_row(sheet_id, row_number, current_values)
        
        # 로컬 데이터 새로고침
        load_data()
        
        # 실시간 업데이트 알림
        socketio.emit('data_updated', {
            'message': f"프로젝트가 수정되었습니다: {project_code}",
            'timestamp': datetime.now().isoformat(),
            'action': 'inline_update',
            'project_code': project_code,
            'updated_fields': list(data.keys())
        })
        
        return jsonify({
            'ok': True,
            'message': '성공적으로 업데이트되었습니다.',
            'project_code': project_code
        })
        
    except Exception as e:
        logger.error(f"인라인 업데이트 오류: {str(e)}")
        return jsonify({'ok': False, 'error': str(e)}), 500

# 테스트용 간단한 엔드포인트 추가
@app.route('/api/test-inline', methods=['GET', 'POST'])
def test_inline_endpoint():
    """인라인 업데이트 테스트용 엔드포인트"""
    if request.method == 'GET':
        return jsonify({'ok': True, 'message': 'API 엔드포인트가 작동 중입니다.'})
    else:
        data = request.get_json()
        return jsonify({'ok': True, 'received_data': data})

@socketio.on('connect')
def handle_connect():
    """클라이언트 연결 처리"""
    logger.info('클라이언트가 연결되었습니다.')
    emit('connected', {'message': '대시보드에 연결되었습니다.'})

@socketio.on('disconnect')
def handle_disconnect():
    """클라이언트 연결 해제 처리"""
    logger.info('클라이언트 연결이 해제되었습니다.')

@socketio.on('request_update')
def handle_request_update():
    """실시간 업데이트 요청 처리"""
    try:
        df = load_data()
        if df is not None:
            emit('data_updated', {
                'message': '데이터가 업데이트되었습니다.',
                'timestamp': last_update.isoformat() if last_update else None,
                'record_count': len(df)
            })
    except Exception as e:
        emit('error', {'message': f'업데이트 오류: {str(e)}'})

@app.route('/api/debug/headers', methods=['GET'])
def debug_headers():
    """Google Sheets 헤더 확인용 디버깅 엔드포인트"""
    try:
        df = current_data if current_data is not None else load_data()
        if df is None:
            return jsonify({'error': '데이터를 불러올 수 없습니다.'}), 500
            
        headers = df.columns.tolist()
        
        # 샘플 데이터에서 날짜 문제 해결
        if not df.empty:
            sample_df = df.head(3).copy()
            # NaT 값을 None으로 변환
            for col in sample_df.columns:
                if sample_df[col].dtype == 'datetime64[ns]':
                    sample_df[col] = sample_df[col].dt.strftime('%Y-%m-%d').replace('NaT', None)
            sample_data = sample_df.to_dict('records')
        else:
            sample_data = []
        
        # 컬럼별 인덱스 정보
        column_mapping = {}
        for i, col in enumerate(headers):
            # A=0, B=1, C=2... -> A, B, C...
            column_letter = chr(ord('A') + i) if i < 26 else f"A{chr(ord('A') + i - 26)}"
            column_mapping[col] = {
                'index': i,
                'letter': column_letter
            }
        
        return jsonify({
            'headers': headers,
            'column_mapping': column_mapping,
            'sample_data': sample_data,
            'total_columns': len(headers)
        })
        
    except Exception as e:
        logger.error(f"디버깅 엔드포인트 오류: {str(e)}")
        return jsonify({'error': str(e)}), 500

def can_edit_field_server(user_role, card_type, field_name):
    """서버 사이드에서 필드 편집 권한 체크"""
    # 관리자는 모든 필드 편집 가능
    if user_role == 'Admin':
        return True
    
    # 편집자 권한 정의 (클라이언트와 동일하게 유지)
    editor_allowed_fields = {
        'basic': ['현장 담당자', '도급 구분', '담당자 연락처', '시공자', '담당자 이메일'],
        'construction': 'all',  # 공사정보는 모든 필드 수정 가능
        'financial': [],  # 금액정보 수정 불가
        'payment': [],  # 수금정보 수정 불가
        'profit': ['제품대', '도급비', '자재비', '기타비'],  # 특정 필드만 수정 가능
        'documents': ['견적서 및 계약서 폴더 경로']  # 문서 경로만 수정 가능
    }
    
    if user_role == 'Editor':
        allowed_fields = editor_allowed_fields.get(card_type, [])
        if allowed_fields == 'all':
            return True
        elif isinstance(allowed_fields, list):
            return field_name in allowed_fields
    
    # 뷰어는 편집 불가
    return False

def get_card_type_for_field(field_name):
    """필드명으로 카드 타입을 판단"""
    field_to_card_mapping = {
        # 기본정보
        '현장 담당자': 'basic', '도급 구분': 'basic', '담당자 연락처': 'basic', 
        '시공자': 'basic', '담당자 이메일': 'basic', '사업자': 'basic', '현장 주소': 'basic',
        
        # 공사정보
        '공사 구분': 'construction', '기계 분류': 'construction', '브랜드': 'construction',
        '공사 시작': 'construction', '공사 종료': 'construction', '공사 내용': 'construction',
        '공사 확정': 'construction',
        
        # 금액정보
        '총액 1': 'financial', '부가세': 'financial', '총액 2': 'financial',
        
        # 수금정보
        '계약금': 'payment', '중도금': 'payment', '잔금': 'payment', 
        '미수금': 'payment', '미수금W': 'payment', '수금 확인': 'payment', '수금 날짜': 'payment',
        
        # 손익정보
        '제품대': 'profit', '도급비': 'profit', '자재비': 'profit', '기타비': 'profit',
        
        # 문서정보
        '견적서 및 계약서 폴더 경로': 'documents'
    }
    
    return field_to_card_mapping.get(field_name, 'unknown')

@app.route('/api/inline-update', methods=['POST'])
@login_required
def inline_update_direct():
    """간단한 인라인 업데이트 API (직접 구현)"""
    try:
        data = request.get_json()
        logger.info(f"인라인 업데이트 요청: {data}")
        
        project_code = data.get('projectCode')
        if not project_code:
            return jsonify({'ok': False, 'error': '프로젝트 코드가 필요합니다.'}), 400
        
        # 사용자 권한 확인
        current_user_role = get_user_role()
        logger.info(f"사용자 권한 확인: {current_user_role}")
        
        # 권한별로 필드 수정 가능 여부 체크
        forbidden_fields = []
        for field_name, value in data.items():
            if field_name == 'projectCode':
                continue
                
            card_type = get_card_type_for_field(field_name)
            can_edit = can_edit_field_server(current_user_role, card_type, field_name)
            logger.info(f"필드 권한 체크: {field_name} ({card_type}) - {current_user_role} - {'허용' if can_edit else '거부'}")
            
            if not can_edit:
                forbidden_fields.append(field_name)
        
        if forbidden_fields:
            return jsonify({
                'ok': False, 
                'error': f'권한이 없습니다. 편집 불가 필드: {", ".join(forbidden_fields)}',
                'forbidden_fields': forbidden_fields
            }), 403
        
        sheet_id = os.getenv('GOOGLE_SHEET_ID')
        if not sheet_id:
            return jsonify({'ok': False, 'error': 'GOOGLE_SHEET_ID가 설정되지 않았습니다.'}), 500
        
        manager = GoogleSheetsManager()
        
        # 프로젝트 코드로 행 찾기
        logger.info(f"프로젝트 코드 {project_code}의 행 번호를 찾는 중...")
        
        # A열에서 프로젝트 코드 검색
        search_result = manager.service.spreadsheets().values().get(
            spreadsheetId=sheet_id,
            range='공사 현황!A:A'
        ).execute()
        
        values = search_result.get('values', [])
        row_number = None
        
        for i, row in enumerate(values):
            if row and len(row) > 0 and row[0] == project_code:
                row_number = i + 1  # 1부터 시작
                break
        
        if not row_number:
            logger.error(f"프로젝트 코드 {project_code}를 찾을 수 없습니다. 데이터 새로고침 후 재시도...")
            # 데이터를 새로 로드하고 재시도
            load_data()
            
            # 다시 검색 시도
            search_result = manager.service.spreadsheets().values().get(
                spreadsheetId=sheet_id,
                range='공사 현황!A:A'
            ).execute()
            
            values = search_result.get('values', [])
            for i, row in enumerate(values):
                if row and len(row) > 0 and row[0] == project_code:
                    row_number = i + 1
                    break
            
            if not row_number:
                return jsonify({'ok': False, 'error': f'프로젝트 코드 {project_code}를 찾을 수 없습니다.'}), 404
        
        logger.info(f"프로젝트 {project_code}을 {row_number}행에서 발견")
        
        # 업데이트할 셀들
        updates = []
        
        # 업데이트 전에 현재 값들을 조회해서 이전 값으로 기록
        old_values = {}
        
        # 필드별로 해당 열에 업데이트
        field_column_mapping = {
            # 기본정보
            '사업자': 'B',
            '현장 담당자': 'N', 
            '도급 구분': 'L',
            '담당자 연락처': 'O',
            '시공자': 'M',
            '담당자 이메일': 'P',
            '현장 주소': 'E',
            # 공사정보
            '공사 구분': 'F',
            '기계 분류': 'G',
            '브랜드': 'H',
            '공사 시작': 'I',
            '공사 종료': 'J',
            '공사 내용': 'K',
            '공사 확정': 'AL',
            # 금액정보
            '총액 1': 'Q',
            '부가세': 'R',
            '총액 2': 'S',
            '계약금': 'T',
            '중도금': 'U',
            '잔금': 'V',
            '미수금': 'W',
            '계산서': 'X',
            '수금 날짜': 'Y',
            '수금 확인': 'Z',
            '제품대': 'AA',
            '도급비': 'AB',
            '자재비': 'AC',
            '기타비': 'AD',
            '순익': 'AE',
            '마진율': 'AF',
            '비고': 'AG',
            '계약금 입금자명': 'AH',
            '중도금 입금자명': 'AI',
            '잔금 입금자명': 'AJ',
            '견적서 및 계약서 폴더 경로': 'AK'
        }
        
        for field_name, value in data.items():
            if field_name == 'projectCode':
                continue
            
            if field_name in field_column_mapping:
                column = field_column_mapping[field_name]
                range_name = f'공사 현황!{column}{row_number}'
                
                # 업데이트 전 이전 값 조회
                try:
                    current_value_result = manager.service.spreadsheets().values().get(
                        spreadsheetId=sheet_id,
                        range=range_name
                    ).execute()
                    current_values = current_value_result.get('values', [['']])
                    old_values[field_name] = current_values[0][0] if current_values and current_values[0] else ''
                except Exception as e:
                    logger.warning(f"이전 값 조회 실패 ({field_name}): {e}")
                    old_values[field_name] = ''
                
                # 숫자 필드인 경우 처리 개선 (콤마 뒤 0 손실 방지)
                processed_value = value
                if field_name in ['총액 1', '총액 2', '총액2', '계약금', '중도금', '잔금', '미수금', '제품대', '도급비', '자재비', '기타비', '순익']:
                    try:
                        # 원본 값 보존을 위해 조심스럽게 처리
                        original_value = str(value).strip()
                        logger.info(f"원본 입력값: {field_name} = '{original_value}'")
                        
                        if original_value and original_value != '-':
                            # 숫자가 아닌 문자들만 제거 (콤마는 보존)
                            clean_value = original_value.replace('₩', '').replace('원', '').strip()
                            
                            # 콤마가 포함된 숫자인지 확인
                            if ',' in clean_value:
                                # 콤마 형식 숫자 → 순수 숫자로 변환
                                try:
                                    # 콤마 제거 후 숫자로 파싱해서 유효성 확인
                                    numeric_value = float(clean_value.replace(',', ''))
                                    # 정수인 경우 .0 제거
                                    if numeric_value.is_integer():
                                        processed_value = str(int(numeric_value))
                                    else:
                                        processed_value = str(numeric_value)
                                    logger.info(f"콤마 형식 처리: '{clean_value}' → '{processed_value}'")
                                except ValueError:
                                    logger.warning(f"콤마 형식 파싱 실패: '{clean_value}'")
                                    processed_value = clean_value.replace(',', '')
                            else:
                                # 콤마가 없는 경우 그대로 사용
                                processed_value = clean_value
                        else:
                            processed_value = ''
                            
                    except Exception as e:
                        logger.warning(f"숫자 값 처리 실패 ({field_name}): {e}")
                        processed_value = value
                
                logger.info(f"업데이트 대상: {field_name} -> {range_name} = '{value}' -> '{processed_value}'")
                updates.append({
                    'range': range_name,
                    'values': [[processed_value]]
                })
        
        # 배치 업데이트 실행
        if updates:
            try:
                batch_update_body = {
                    'valueInputOption': 'USER_ENTERED',
                    'data': updates
                }
                
                logger.info(f"{len(updates)}개 셀 업데이트 시작...")
                
                batch_result = manager.service.spreadsheets().values().batchUpdate(
                    spreadsheetId=sheet_id,
                    body=batch_update_body
                ).execute()
                
                updated_cells = batch_result.get('totalUpdatedCells', 0)
                logger.info(f"업데이트 완료: {updated_cells}개 셀")
                
            except Exception as api_error:
                if "protected cell" in str(api_error):
                    logger.warning("보호된 셀 감지 - 단일 셀 업데이트로 재시도")
                    
                    # 단일 셀씩 개별 업데이트 시도
                    updated_cells = 0
                    failed_updates = []
                    
                    for update in updates:
                        try:
                            single_update_body = {
                                'valueInputOption': 'USER_ENTERED',
                                'data': [update]
                            }
                            
                            single_result = manager.service.spreadsheets().values().batchUpdate(
                                spreadsheetId=sheet_id,
                                body=single_update_body
                            ).execute()
                            
                            updated_cells += single_result.get('totalUpdatedCells', 0)
                            logger.info(f"개별 셀 업데이트 성공: {update['range']}")
                            
                        except Exception as single_error:
                            logger.error(f"개별 셀 업데이트 실패: {update['range']} - {str(single_error)}")
                            failed_updates.append({
                                'range': update['range'], 
                                'error': str(single_error)
                            })
                    
                    if updated_cells > 0:
                        logger.info(f"부분 업데이트 완료: {updated_cells}개 셀")
                        # 일부라도 성공했으면 성공으로 처리하되, 실패한 것들을 알림
                        message = f"일부 업데이트 완료: {updated_cells}개 셀"
                        if failed_updates:
                            message += f" (실패: {len(failed_updates)}개)"
                    else:
                        # 모든 업데이트 실패
                        return jsonify({
                            'ok': False, 
                            'error': f'모든 셀이 보호되어 있습니다. 서비스 계정 {manager.service_account_email}에게 편집 권한을 부여해주세요.',
                            'service_account': 'sheets-manager@smooth-unison-470801-p5.iam.gserviceaccount.com',
                            'failed_ranges': failed_updates
                        }), 400
                else:
                    raise api_error
        
        # 로컬 데이터 새로고침
        load_data()
        
        # 실시간 알림
        if socketio:
            socketio.emit('data_updated', {
                'message': f"프로젝트가 수정되었습니다: {project_code}",
                'timestamp': datetime.now().isoformat(),
                'action': 'inline_update',
                'project_code': project_code
            })
        
        # 업데이트 후 새로운 프로젝트 코드 확인 (수식으로 변경될 수 있음)
        try:
            updated_row_range = f'공사 현황!A{row_number}:A{row_number}'
            updated_result = manager.service.spreadsheets().values().get(
                spreadsheetId=sheet_id,
                range=updated_row_range
            ).execute()
            
            updated_values = updated_result.get('values', [[]])
            new_project_code = updated_values[0][0] if updated_values and updated_values[0] else project_code
            
            logger.info(f"업데이트 후 프로젝트 코드: {project_code} -> {new_project_code}")
            
        except Exception as e:
            logger.warning(f"새 프로젝트 코드 확인 실패: {e}")
            new_project_code = project_code
        
        # 감사 로그 기록 (실제 이전 값 포함)
        try:
            for field_name, new_value in data.items():
                # 유효한 필드만 로그 기록 (빈 문자열, None, 'projectCode' 제외)
                if (field_name and 
                    field_name != 'projectCode' and 
                    field_name.strip() != '' and 
                    field_name != 'undefined' and
                    field_name in field_column_mapping):
                    
                    old_value = old_values.get(field_name, '')
                    logger.info(f"필드 업데이트 성공 확인: {field_name} = {old_value} -> {new_value}")
                    log_user_action(
                        action='UPDATE_FIELD',
                        details=f'프로젝트 필드 수정: {field_name}',
                        project_code=project_code,
                        field_name=field_name,
                        old_value=old_value if old_value else '-',
                        new_value=str(new_value) if new_value else '-'
                    )
        except Exception as log_error:
            logger.warning(f"감사 로그 기록 실패: {log_error}")
        
        # 금액 관련 필드가 수정된 경우 나중에 미수금을 계산하기 위해 표시만 해둠
        amount_fields = ['총액 2', '총액2', '계약금', '중도금', '잔금']  # '총액2' (공백없음)도 포함
        updated_amount_fields = [field for field in data.keys() if field in amount_fields]
        calculated_fields = {}
        
        logger.info(f"수정된 데이터: {data}")
        logger.info(f"금액 관련 필드 체크: {updated_amount_fields}")
        
        # 각 수정된 필드의 원본 값과 타입 상세 로깅
        for field_name, value in data.items():
            if field_name != 'projectCode':
                logger.info(f"필드별 상세정보: {field_name} = '{value}' (타입: {type(value)}, 길이: {len(str(value))})")
        
        # 미수금 계산은 업데이트 후에 수행
        
        # 업데이트 후 실제 저장된 값 확인 (동기화 문제 해결)
        actual_values = {}
        if updated_amount_fields:
            try:
                # 방금 업데이트한 필드들의 실제 값을 Google Sheets에서 확인
                import time
                time.sleep(0.3)  # Google Sheets 업데이트 시간 대기 (단축)
                # 더 넓은 범위로 확장 (금액 관련 모든 필드 포함)
                verify_range = f'공사 현황!Q{row_number}:AM{row_number}'
                verify_result = manager.service.spreadsheets().values().get(
                    spreadsheetId=sheet_id,
                    range=verify_range,
                    valueRenderOption='FORMATTED_VALUE'
                ).execute()
                
                verify_values = verify_result.get('values', [[]])
                if verify_values and verify_values[0]:
                    verify_row = verify_values[0]
                    # Google Sheets 컬럼 매핑 (Q부터 AM까지)
                    all_field_mapping = [
                        '총액 1', '부가세', '총액 2', '계약금', '중도금', '잔금', '미수금', '계산서',
                        '수금 날짜', '수금 확인', '제품대', '도급비', '자재비', '기타비',
                        '비고', '계약금 입금자명', '중도금 입금자명', '잔금 입금자명', '견적서 및 계약서 폴더 경로',
                        '공사 확정', 'Airtable Record ID'
                    ]
                    
                    for i, field in enumerate(all_field_mapping):
                        if i < len(verify_row) and field in updated_amount_fields:
                            actual_values[field] = verify_row[i]
                            logger.info(f"실제 저장된 값 확인: {field} = {verify_row[i]}")
                            
            except Exception as verify_error:
                logger.warning(f"실제 값 확인 실패: {verify_error}")
        
        # 업데이트 완료 후 미수금 계산 (실제 저장된 값 기준)
        if updated_amount_fields:
            try:
                # 업데이트 완료 후 다시 Google Sheets에서 최신 값들을 가져와서 미수금 계산
                import time
                time.sleep(0.2)  # 잠시 대기
                amount_range = f'공사 현황!S{row_number}:V{row_number}'  # 총액2(S), 계약금(T), 중도금(U), 잔금(V)
                amount_result = manager.service.spreadsheets().values().get(
                    spreadsheetId=sheet_id,
                    range=amount_range,
                    valueRenderOption='UNFORMATTED_VALUE'  # 숫자 값으로 가져오기
                ).execute()
                
                amount_values = amount_result.get('values', [[]])
                if amount_values and amount_values[0]:
                    amount_row = amount_values[0]
                    
                    # 각 값을 안전하게 파싱
                    def safe_float(val):
                        if val is None or val == '':
                            return 0.0
                        try:
                            return float(str(val).replace(',', '').replace('₩', '').strip())
                        except (ValueError, TypeError):
                            return 0.0
                    
                    # Google Sheets에서 가져온 최신 값들 (업데이트 완료 후)
                    total_amount_2 = safe_float(amount_row[0] if len(amount_row) > 0 else 0)  # 총액2
                    contract_amount = safe_float(amount_row[1] if len(amount_row) > 1 else 0)  # 계약금
                    interim_amount = safe_float(amount_row[2] if len(amount_row) > 2 else 0)   # 중도금
                    final_amount = safe_float(amount_row[3] if len(amount_row) > 3 else 0)     # 잔금
                    
                    # 미수금 계산: 총액2 - (계약금 + 중도금 + 잔금)
                    outstanding_amount = total_amount_2 - (contract_amount + interim_amount + final_amount)
                    
                    # 계산된 미수금을 반환할 필드에 추가 (항상 포함, 0이어도)
                    calculated_fields['미수금'] = int(outstanding_amount) if outstanding_amount.is_integer() else outstanding_amount
                    
                    logger.info(f"업데이트 후 미수금 계산: 총액2({total_amount_2}) - 계약금({contract_amount}) - 중도금({interim_amount}) - 잔금({final_amount}) = {outstanding_amount}")
                
            except Exception as calc_error:
                logger.warning(f"업데이트 후 미수금 계산 실패: {calc_error}")
        
        return jsonify({
            'ok': True,
            'message': '성공적으로 업데이트되었습니다.',
            'project_code': project_code,
            'new_project_code': new_project_code,
            'project_code_changed': new_project_code != project_code,
            'updated_cells': updated_cells if updates else 0,
            'auto_calculated': len(updated_amount_fields) > 0,
            'calculated_fields': calculated_fields,  # 계산된 필드 값들 반환
            'actual_values': actual_values  # 실제 저장된 값들 반환
        })
        
    except Exception as e:
        logger.error(f"인라인 업데이트 오류: {str(e)}", exc_info=True)
        return jsonify({'ok': False, 'error': str(e)}), 500

# 감사 로그 API
@app.route('/api/audit-logs', methods=['GET'])
@login_required
def get_audit_logs_api():
    """감사 로그 조회 API (페이지네이션 지원)"""
    try:
        # 쿼리 파라미터 처리
        days = int(request.args.get('days', 7))  # 기본 7일
        page = int(request.args.get('page', 1))  # 페이지 번호 (1부터 시작)
        per_page = int(request.args.get('per_page', 50))  # 페이지당 항목 수 (기본 50개)
        
        # 전체 로그 조회
        all_logs = get_audit_logs(days)
        
        # 관리자가 아닌 경우 본인 로그만 조회
        user_email = session['user']['email']
        user_role = session['user'].get('permission_level', 'viewer')
        
        if user_role != 'admin':
            all_logs = [log for log in all_logs if log.get('user_email') == user_email]
        
        # 최신순 정렬 (timestamp 기준 내림차순)
        all_logs.sort(key=lambda x: x.get('timestamp', ''), reverse=True)
        
        # 페이지네이션 계산
        total_count = len(all_logs)
        total_pages = (total_count + per_page - 1) // per_page  # 올림 계산
        start_index = (page - 1) * per_page
        end_index = start_index + per_page
        
        # 해당 페이지의 로그만 추출
        paginated_logs = all_logs[start_index:end_index]
        
        return jsonify({
            'success': True, 
            'logs': paginated_logs,
            'pagination': {
                'current_page': page,
                'per_page': per_page,
                'total_count': total_count,
                'total_pages': total_pages,
                'has_prev': page > 1,
                'has_next': page < total_pages,
                'prev_page': page - 1 if page > 1 else None,
                'next_page': page + 1 if page < total_pages else None
            }
        })
        
    except Exception as e:
        logger.error(f"감사 로그 조회 오류: {str(e)}")
        return jsonify({'success': False, 'message': '로그 조회에 실패했습니다.'}), 500

# 파비콘 처리 (404 오류 방지)
@app.route('/favicon.ico')
def favicon():
    return '', 204  # No Content 응답

if __name__ == '__main__':
    # 인라인 업데이트 라우트 등록 (중복 제거를 위해 주석 처리)
    # register_inline_update_routes(app, socketio, load_data)
    
    # 등록된 라우트 확인 (디버깅용)
    logger.info("등록된 라우트:")
    for rule in app.url_map.iter_rules():
        logger.info(f"  {rule.endpoint}: {rule.rule} {rule.methods}")
    
    # 초기 데이터 로드
    logger.info("초기 데이터 로드 중...")
    load_data()
    
    # 서버 시작
    port = int(os.getenv('PORT', 5000))
    debug = os.getenv('DEBUG', 'True').lower() == 'true'
    
    # 사용자 관리 API 엔드포인트
    @app.route('/api/users', methods=['GET'])
    @admin_required
    def get_users():
        """사용자 목록 조회 (관리자만)"""
        try:
            users = user_manager.get_all_users()
            return jsonify({'success': True, 'users': users})
        except Exception as e:
            logger.error(f"사용자 목록 조회 오류: {str(e)}")
            return jsonify({'success': False, 'message': '사용자 목록을 불러올 수 없습니다.'}), 500

    @app.route('/api/users/permission', methods=['POST'])
    @admin_required
    def update_user_permission():
        """사용자 권한 업데이트 (관리자만)"""
        try:
            data = request.get_json()
            email = data.get('email')
            permission = data.get('permission')
            
            if not email or not permission:
                return jsonify({'success': False, 'message': '이메일과 권한이 필요합니다.'}), 400
            
            success, message = user_manager.update_user_permission(email, permission)
            return jsonify({'success': success, 'message': message})
            
        except Exception as e:
            logger.error(f"권한 업데이트 오류: {str(e)}")
            return jsonify({'success': False, 'message': '권한 업데이트에 실패했습니다.'}), 500

    @app.route('/api/users/status', methods=['POST'])
    @admin_required
    def toggle_user_status():
        """사용자 상태 변경 (관리자만)"""
        try:
            data = request.get_json()
            email = data.get('email')
            is_active = data.get('is_active')
            
            if email is None or is_active is None:
                return jsonify({'success': False, 'message': '이메일과 상태가 필요합니다.'}), 400
            
            success, message = user_manager.toggle_user_status(email, is_active)
            return jsonify({'success': success, 'message': message})
            
        except Exception as e:
            logger.error(f"사용자 상태 변경 오류: {str(e)}")
            return jsonify({'success': False, 'message': '사용자 상태 변경에 실패했습니다.'}), 500

    @app.route('/api/users', methods=['POST'])
    @admin_required
    def create_user():
        """새 사용자 생성 (관리자만)"""
        try:
            data = request.get_json()
            name = data.get('name', '').strip()
            email = data.get('email', '').strip().lower()
            password = data.get('password', '')
            permission = data.get('permission', 'viewer')
            
            if not name or not email or not password:
                return jsonify({'success': False, 'message': '이름, 이메일, 비밀번호가 필요합니다.'}), 400
            
            success, message = user_manager.create_user(name, email, password, permission)
            return jsonify({'success': success, 'message': message})
            
        except Exception as e:
            logger.error(f"사용자 생성 오류: {str(e)}")
            return jsonify({'success': False, 'message': '사용자 생성에 실패했습니다.'}), 500

    @app.route('/api/users/<email>', methods=['DELETE'])
    @admin_required
    def delete_user(email):
        """사용자 삭제 (관리자만)"""
        try:
            # 본인 계정 삭제 방지
            if session['user']['email'] == email:
                return jsonify({'success': False, 'message': '본인 계정은 삭제할 수 없습니다.'}), 400
            
            success, message = user_manager.delete_user(email)
            return jsonify({'success': success, 'message': message})
            
        except Exception as e:
            logger.error(f"사용자 삭제 오류: {str(e)}")
            return jsonify({'success': False, 'message': '사용자 삭제에 실패했습니다.'}), 500

    logger.info(f"대시보드 서버 시작: http://localhost:{port}")
    socketio.run(app, debug=debug, host='0.0.0.0', port=port)