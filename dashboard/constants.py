"""
프로젝트 전역 상수 정의
메모 기능 및 기타 공통 상수를 중앙 집중 관리
"""

# 메모 기능 관련 상수
# =====================

# 금액(수금) 필드 — 금액 0 이하 시 메모 자동삭제 등 '금액 전용' 로직에만 사용.
#   (MEMOABLE_FIELDS 를 이 용도로 재사용하면 계산서 텍스트가 금액 파싱→0 으로 읽혀 메모가 지워짐)
PAYMENT_AMOUNT_FIELDS = ['계약금', '중도금', '잔금']

# 메모 가능한 필드 목록 (수금 3단계 + 계산서 요약(Y열))
MEMOABLE_FIELDS = ['계약금', '중도금', '잔금', '계산서']

def payment_field_column(field: str):
    """메모 대상 필드(MEMOABLE_FIELDS) → 공사 현황 열 글자. 대상이 아니면 None.

    예전엔 {'계약금': 'U', ...} 로 박아 둬서 열을 끼워 넣으면 입금 메모·금액이 엉뚱한 칸에
    쓰일 상황이었다 → 헤더 지도에서 찾는다 (2026-10-22).
    """
    if field not in MEMOABLE_FIELDS:
        return None
    from dashboard.utils.google_sheets import main_field_to_letter
    return main_field_to_letter().get(field)


# 메모 필드명 (데이터프레임 컬럼)
MEMO_FIELD_SUFFIX = '_메모'

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
# 프로젝트 데이터 변경 시 (update_project, add_project_auto):
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
