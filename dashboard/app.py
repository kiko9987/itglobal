from flask import Flask, render_template, jsonify, request, send_from_directory
from flask_socketio import SocketIO, emit
import os
import sys
import logging
from datetime import datetime, timedelta
import json
import re
from collections import Counter, defaultdict
from dotenv import load_dotenv

# 프로젝트 루트 경로를 시스템 경로에 추가
project_root = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(project_root))

from dashboard.utils.google_sheets import GoogleSheetsManager
from dashboard.utils.data_analyzer import DataAnalyzer

# 환경 변수 로드
load_dotenv()

# Flask 앱 초기화
app = Flask(__name__)
app.config['SECRET_KEY'] = os.getenv('FLASK_SECRET_KEY', 'your-secret-key-here')

# SocketIO 초기화 (실시간 업데이트용)
socketio = SocketIO(app, cors_allowed_origins="*")

# 로깅 설정
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# 전역 변수 (캐싱 개선)
current_data = None
last_update = None
_data_cache = {}
_cache_expiry = 60   # 1분 캐시

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
    """메인 대시보드 페이지"""
    return render_template('dashboard.html')

@app.route('/projects')
def project_list():
    """프로젝트 목록 페이지"""
    return render_template('project_list.html')

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
        
        company = str(data.get("사업자", "")).strip()
        owner = str(data.get("담당자", "")).strip()
        
        if not company or not owner:
            return jsonify({"ok": False, "error": "사업자/담당자는 필수입니다"}), 400

        # 현재 데이터 로드
        df = current_data if current_data is not None else load_data()
        if df is None:
            return jsonify({"ok": False, "error": "데이터를 불러올 수 없습니다"}), 500
        
        # 안전한 자동 프로젝트 코드 생성 (동시성 대응)
        try:
            code = _safe_next_running_number_with_retry(company, owner)
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
    try:
        data = request.get_json()
        
        sheet_id = os.getenv('GOOGLE_SHEET_ID')
        if not sheet_id:
            return jsonify({'error': 'GOOGLE_SHEET_ID가 설정되지 않았습니다.'}), 500
        
        manager = GoogleSheetsManager()
        
        # 데이터를 구글 시트 형식으로 변환
        values = convert_form_data_to_sheet_row(data, manager)
        
        # 구글 시트에 추가
        result = manager.append_row(sheet_id, values)
        
        # 로컬 데이터 새로고침
        load_data()
        
        # 실시간 업데이트 알림
        socketio.emit('data_updated', {
            'message': f"새 프로젝트가 등록되었습니다: {data.get('projectCode', '')}",
            'timestamp': datetime.now().isoformat(),
            'action': 'create'
        })
        
        return jsonify({'success': True, 'project_code': data.get('projectCode', '')})
        
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
        
        # 프로젝트가 있는 행 찾기
        row_number = manager.find_row_by_project_code(sheet_id, project_code)
        
        if not row_number:
            return jsonify({'error': '프로젝트를 찾을 수 없습니다.'}), 404
        
        # 데이터를 구글 시트 형식으로 변환
        values = convert_form_data_to_sheet_row(data, manager)
        
        # 구글 시트 업데이트
        result = manager.update_row(sheet_id, row_number, values)
        
        # 로컬 데이터 새로고침
        load_data()
        
        # 실시간 업데이트 알림
        socketio.emit('data_updated', {
            'message': f"프로젝트가 수정되었습니다: {project_code}",
            'timestamp': datetime.now().isoformat(),
            'action': 'update'
        })
        
        return jsonify({'success': True, 'project_code': project_code})
        
    except Exception as e:
        logger.error(f"프로젝트 수정 API 오류: {str(e)}")
        return jsonify({'error': str(e)}), 500

# 삭제 기능 제거 (사용자 요청에 따라)
# @app.route('/api/projects/<project_code>', methods=['DELETE'])
# def delete_project(project_code):
#     """프로젝트 삭제 API (구글 시트에서는 빈 행으로 만들기)"""

def convert_form_data_to_sheet_row(form_data, manager):
    """폼 데이터를 구글 시트 행 형식으로 변환"""
    column_mapping = manager.get_column_mapping()
    
    # 폼 필드명을 구글 시트 컬럼명으로 매핑
    field_mapping = {
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
        'finalPaymentPayer': '잔금 입금자명'
    }
    
    # 39개 컬럼에 맞춰 빈 리스트 생성
    values = [''] * 39
    
    # 각 컬럼에 해당하는 값 설정
    for column_letter, column_name in column_mapping.items():
        column_index = ord(column_letter) - ord('A') if len(column_letter) == 1 else \
                      (ord(column_letter[0]) - ord('A') + 1) * 26 + (ord(column_letter[1]) - ord('A'))
        
        # 폼 데이터에서 해당 값 찾기
        form_field = None
        for form_key, sheet_column in field_mapping.items():
            if sheet_column == column_name:
                form_field = form_key
                break
        
        if form_field and form_field in form_data:
            value = form_data[form_field]
            
            # 데이터 타입별 처리
            if isinstance(value, bool):
                values[column_index] = 'TRUE' if value else 'FALSE'
            elif isinstance(value, (int, float)):
                values[column_index] = str(value) if value != 0 else ''
            else:
                values[column_index] = str(value) if value else ''
    
    return values

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

if __name__ == '__main__':
    # 초기 데이터 로드
    logger.info("초기 데이터 로드 중...")
    load_data()
    
    # 서버 시작
    port = int(os.getenv('PORT', 5000))
    debug = os.getenv('DEBUG', 'True').lower() == 'true'
    
    logger.info(f"대시보드 서버 시작: http://localhost:{port}")
    socketio.run(app, debug=debug, host='0.0.0.0', port=port)