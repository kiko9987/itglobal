"""
프로젝트 전역 상수 정의
메모 기능 및 기타 공통 상수를 중앙 집중 관리
"""

# 메모 기능 관련 상수
# =====================

# 메모 가능한 필드 목록
MEMOABLE_FIELDS = ['계약금', '중도금', '잔금']

# 필드명 → 구글 시트 컬럼 매핑
PAYMENT_FIELD_TO_COLUMN = {
    '계약금': 'T',
    '중도금': 'U',
    '잔금': 'V'
}

# 컬럼 → 필드명 매핑 (역방향)
COLUMN_TO_PAYMENT_FIELD = {v: k for k, v in PAYMENT_FIELD_TO_COLUMN.items()}

# 메모 필드명 (데이터프레임 컬럼)
MEMO_FIELD_SUFFIX = '_메모'
PAYMENT_MEMO_FIELDS = {
    '계약금_메모': 'T',
    '중도금_메모': 'U',
    '잔금_메모': 'V'
}

# 메모 길이 제한
MAX_MEMO_LENGTH = 2000

# Google Sheets 관련 상수
# =====================

# 셀 노트 캐시 키 템플릿
CELL_NOTES_CACHE_KEY_TEMPLATE = "cell_notes_{sheet_id}"

# 프로젝트 데이터 캐시 키
PROJECT_DATA_CACHE_KEY = "current_sheet_data"

# 캐시 무효화 전략 가이드
# =====================
#
# 메모 저장 시 (save_field_memo):
#   - 셀 노트 캐시만 무효화 (CELL_NOTES_CACHE_KEY_TEMPLATE)
#   - 프로젝트 데이터는 변경되지 않으므로 PROJECT_DATA_CACHE_KEY 무효화 불필요
#   - 최적화: 전체 데이터 재로드 방지, 메모리 부하 감소
#
# 프로젝트 데이터 변경 시 (update_project, update_project_inline, add_project_auto):
#   - 프로젝트 데이터 캐시 무효화 (PROJECT_DATA_CACHE_KEY)
#   - load_data(force_refresh=True) 호출 시 셀 노트도 자동으로 재로드됨
#   - 셀 노트 캐시 명시적 무효화 불필요 (load_data 내부에서 처리)

# 검증 관련 상수
# =====================

# 프로젝트 코드 패턴
PROJECT_CODE_PATTERN = r'^[A-Z]\d{4}-[A-Z]+$'

# 날짜 형식
DATE_FORMAT = '%Y-%m-%d'

# 통화 형식
CURRENCY_SYMBOLS = r'[₩원$€¥£＄￦]'
CURRENCY_NUMBER_PATTERN = r'-?[\d,.]+'

# 에러 메시지
# =====================

ERROR_MESSAGES = {
    'memo': {
        'project_not_found': '프로젝트를 찾을 수 없습니다',
        'invalid_field': '지원하지 않는 필드명입니다',
        'save_failed': '메모 저장에 실패했습니다',
        'validation_failed': '입력 데이터 검증 실패',
    },
    'project': {
        'not_found': '프로젝트를 찾을 수 없습니다',
        'cancelled': '취소된 프로젝트입니다',
        'no_permission': '편집 권한이 없습니다',
    }
}

# 성공 메시지
# =====================

SUCCESS_MESSAGES = {
    'memo': {
        'saved': '메모가 저장되었습니다',
        'deleted': '메모가 삭제되었습니다',
    }
}
