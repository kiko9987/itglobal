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
            
            if updates:
                batch_update_body = {
                    'valueInputOption': 'USER_ENTERED',
                    'data': updates
                }
                result = manager.service.spreadsheets().values().batchUpdate(
                    spreadsheetId=sheet_id,
                    body=batch_update_body
                ).execute()
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

@app.route('/api/inline-update', methods=['POST'])
def inline_update_direct():
    """간단한 인라인 업데이트 API (직접 구현)"""
    try:
        data = request.get_json()
        logger.info(f"인라인 업데이트 요청: {data}")
        
        project_code = data.get('projectCode')
        if not project_code:
            return jsonify({'ok': False, 'error': '프로젝트 코드가 필요합니다.'}), 400
        
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
            '공사 확정': 'AL'
        }
        
        for field_name, value in data.items():
            if field_name == 'projectCode':
                continue
            
            if field_name in field_column_mapping:
                column = field_column_mapping[field_name]
                range_name = f'공사 현황!{column}{row_number}'
                logger.info(f"업데이트 대상: {field_name} -> {range_name} = {value}")
                updates.append({
                    'range': range_name,
                    'values': [[value]]
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
        
        return jsonify({
            'ok': True,
            'message': '성공적으로 업데이트되었습니다.',
            'project_code': project_code,
            'new_project_code': new_project_code,
            'project_code_changed': new_project_code != project_code,
            'updated_cells': updated_cells if updates else 0
        })
        
    except Exception as e:
        logger.error(f"인라인 업데이트 오류: {str(e)}", exc_info=True)
        return jsonify({'ok': False, 'error': str(e)}), 500

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
    port = int(os.getenv('PORT', 8000))
    debug = os.getenv('DEBUG', 'True').lower() == 'true'
    
    # 파비콘 처리 (404 오류 방지)
    @app.route('/favicon.ico')
    def favicon():
        return '', 204  # No Content 응답
    
    logger.info(f"대시보드 서버 시작: http://localhost:{port}")
    socketio.run(app, debug=debug, host='0.0.0.0', port=port)