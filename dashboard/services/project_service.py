import json
import logging
import os
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

import pandas as pd
import re
import time

from dashboard.utils.google_sheets import GoogleSheetsManager
from dashboard.utils.smart_cache_manager import CacheStrategy, smart_get, smart_invalidate, smart_set
from dashboard.utils.error_handler import handle_error, ErrorCategory
from dashboard.utils.background_prefetch import get_background_prefetch, start_background_prefetch, register_sheets_prefetch

logger = logging.getLogger(__name__)

_sheets_manager: Optional[GoogleSheetsManager] = None
_load_lock = None  # lazy initialization

PROJECT_CONFIG: Dict[str, Any] = {}

def _load_project_config() -> Dict[str, Any]:
    try:
        config_path = Path(__file__).resolve().parent.parent / 'project_config.json'
        with config_path.open('r', encoding='utf-8') as handle:
            return json.load(handle)
    except Exception as exc:  # pragma: no cover
        logger.warning("Failed to load project configuration: %s", exc)
        return {}

PROJECT_CONFIG = _load_project_config()


def get_project_config() -> Dict[str, Any]:
    return PROJECT_CONFIG


def get_sheets_manager() -> GoogleSheetsManager:
    global _sheets_manager
    if _sheets_manager is None:
        _sheets_manager = GoogleSheetsManager()
    return _sheets_manager


def get_last_update() -> Optional[datetime]:
    """캐시된 데이터의 마지막 업데이트 시간 반환

    Returns:
        캐시된 데이터의 생성 시간 (datetime), 캐시 미스 시 None
    """
    from dashboard.utils.smart_cache_manager import smart_get_timestamp

    timestamp = smart_get_timestamp("current_sheet_data")
    if timestamp is not None:
        return datetime.fromtimestamp(timestamp)
    return None


def _fetch_fresh_data() -> Optional[pd.DataFrame]:
    """백그라운드 프리패치용 데이터 가져오기 함수"""
    try:
        sheet_id = os.getenv('GOOGLE_SHEET_ID')
        if not sheet_id:
            logger.error("GOOGLE_SHEET_ID is not configured")
            return None

        manager = get_sheets_manager()
        sheet_range = PROJECT_CONFIG.get('sheet_range', '공사 현황!A:AN')
        logger.debug("[PREFETCH] Google Sheets 데이터 가져오기 시작")
        df = manager.get_sheet_data(sheet_id, sheet_range)

        if df.empty:
            logger.warning("[PREFETCH] Google Sheet returned no rows")
            return None

        # 셀 노트(댓글) 가져오기 (계약금, 중도금, 잔금)
        try:
            sheet_name = PROJECT_CONFIG.get('sheet_name', '공사 현황의 사본')

            # 셀 노트 캐시 확인 (메모리 부하 최적화: TEMPORARY 전략 사용)
            notes_cache_key = f"cell_notes_{sheet_id}"
            notes_by_row = smart_get(notes_cache_key, CacheStrategy.TEMPORARY)

            if notes_by_row is None:
                logger.debug("[PREFETCH] 셀 노트 가져오기")
                notes_by_row = manager.get_cell_notes(sheet_id, sheet_name, ['T', 'U', 'V'])
                smart_set(notes_cache_key, notes_by_row, CacheStrategy.TEMPORARY)
            else:
                logger.debug(f"[PREFETCH] 캐시된 셀 노트 사용: {len(notes_by_row)}개 행")

            # DataFrame에 메모 컬럼 추가
            df['계약금_메모'] = None
            df['중도금_메모'] = None
            df['잔금_메모'] = None

            # 행 번호 기준으로 메모 매핑
            for row_num, notes in notes_by_row.items():
                # Redis 캐시에서 row_num이 문자열로 올 수 있으므로 정수로 변환
                try:
                    df_index = int(row_num) - 2  # 헤더 제외, 0-based 인덱스
                except (ValueError, TypeError):
                    continue  # 숫자로 변환할 수 없으면 건너뛰기

                if 0 <= df_index < len(df):
                    if 'T' in notes:
                        df.at[df_index, '계약금_메모'] = notes['T']
                    if 'U' in notes:
                        df.at[df_index, '중도금_메모'] = notes['U']
                    if 'V' in notes:
                        df.at[df_index, '잔금_메모'] = notes['V']

            logger.debug(f"[PREFETCH] 셀 노트 {len(notes_by_row)}개 행 로드 완료")

        except Exception as note_error:
            logger.warning(f"[PREFETCH] 셀 노트 가져오기 실패: {note_error}")
            # 노트 로드 실패해도 데이터는 정상 반환
            df['계약금_메모'] = None
            df['중도금_메모'] = None
            df['잔금_메모'] = None

        # ✨ 백엔드에서 계산 필드 계산 (미수금, 순익, 마진율)
        df = _apply_backend_calculations(df)

        logger.info(f"[PREFETCH] {len(df)} rows 가져오기 완료")
        return df

    except Exception as exc:
        logger.error(f"[PREFETCH] 데이터 가져오기 실패: {exc}")
        return None


def init_background_prefetch():
    """백그라운드 프리패치 시스템 초기화"""
    try:
        # 프리패치 작업 등록
        register_sheets_prefetch()
        # 백그라운드 프리패치 시작
        start_background_prefetch()
        logger.info("[PREFETCH] 백그라운드 프리패치 시스템 초기화 완료")
    except Exception as e:
        logger.error("[PREFETCH] 초기화 실패: %s", e)


def _safe_parse_amount(value: Any) -> float:
    """금액 값을 안전하게 float로 변환

    Args:
        value: 변환할 값 (str, int, float, None 등)

    Returns:
        float 값 (변환 실패 시 0.0)
    """
    if pd.isna(value) or value is None or value == '':
        return 0.0

    if isinstance(value, (int, float)):
        return float(value)

    if isinstance(value, str):
        # 쉼표, 공백, 원화 기호 제거
        cleaned = re.sub(r'[,\s원₩]', '', value)
        try:
            return float(cleaned)
        except (ValueError, TypeError):
            return 0.0

    return 0.0


def _parse_vat_flag(vat_value: Any) -> bool:
    """부가세 플래그를 boolean으로 변환

    Args:
        vat_value: '포함', '미포함', 'true', 'false', True, False 등

    Returns:
        True: 부가세 포함, False: 부가세 미포함
    """
    if pd.isna(vat_value) or vat_value is None or vat_value == '':
        return False  # 기본값: 미포함

    if isinstance(vat_value, bool):
        return vat_value

    if isinstance(vat_value, str):
        vat_str = vat_value.strip().lower()
        # '포함', 'true', 'yes', '1' 등은 True
        return vat_str in ['포함', 'true', 'yes', '1', 'vat 포함']

    return False


def _calculate_total2(total1: float, vat_flag: bool) -> float:
    """총액2 계산: 부가세 플래그에 따라 총액1에 10% 부가세 적용 여부 결정

    Google Sheets 수식과 동일한 절사 + 끝자리 1/9원 보정 방식 사용:
    =IF(R:R=TRUE, Q:Q+FLOOR(Q:Q*0.1,1) + IF(MOD(Q:Q+FLOOR(Q:Q*0.1,1), 10)=1, -1, IF(MOD(Q:Q+FLOOR(Q:Q*0.1,1), 10)=9, 1, 0)), Q:Q)

    세법 근거: 부가가치세법 시행령 제61조 (세액 계산 시 원 단위 미만 절사)
    FLOOR/ROUNDDOWN 사용 (0.5 이상도 내림)
    Python round()는 은행가 반올림 → 사용 금지

    Args:
        total1: 총액1 (공급가액)
        vat_flag: 부가세 포함 여부 (True = 포함, False = 미포함)

    Returns:
        총액2 (부가세 포함 금액)
    """
    if vat_flag:
        # 부가세 포함: 총액2 = 총액1 + FLOOR(총액1 × 0.1, 1) + 끝자리 1/9원 보정
        # 절사 (FLOOR/ROUNDDOWN) - 실무 표준
        from decimal import Decimal, ROUND_DOWN

        total1_decimal = Decimal(str(total1))
        vat_amount = (total1_decimal * Decimal('0.1')).quantize(
            Decimal('1'),
            rounding=ROUND_DOWN  # 절사
        )
        base_total = int(total1_decimal + vat_amount)

        # 끝자리 1/9원 보정 (Google Sheets 수식과 일치)
        last_digit = base_total % 10
        if last_digit == 1:
            adjustment = -1
        elif last_digit == 9:
            adjustment = 1
        else:
            adjustment = 0

        return float(base_total + adjustment)
    else:
        # 부가세 미포함: 총액2 = 총액1
        return total1


def _calculate_outstanding_amount(row: pd.Series) -> float:
    """미수금 계산: 총액2 - (계약금 + 중도금 + 잔금)

    총액2는 Google Sheets 수식에 의존하지 않고 직접 계산:
    - 부가세 포함: 총액1 + ROUND(총액1 × 0.1)
    - 부가세 미포함: 총액1

    Args:
        row: DataFrame의 한 행

    Returns:
        미수금 금액
    """
    # 총액2 계산 (부가세 플래그 기반)
    total1 = _safe_parse_amount(row.get('총액 1', 0))
    vat_flag = _parse_vat_flag(row.get('부가세'))
    total2 = _calculate_total2(total1, vat_flag)

    contract = _safe_parse_amount(row.get('계약금', 0))
    mid = _safe_parse_amount(row.get('중도금', 0))
    final = _safe_parse_amount(row.get('잔금', 0))

    outstanding = total2 - (contract + mid + final)
    return outstanding


def _calculate_net_profit(row: pd.Series) -> float:
    """순익 계산: 총액1 - (제품대 + 도급비 + 자재비 + 기타비)

    Args:
        row: DataFrame의 한 행

    Returns:
        순익 금액
    """
    total1 = _safe_parse_amount(row.get('총액 1', 0))
    product_cost = _safe_parse_amount(row.get('제품대', 0))
    contract_cost = _safe_parse_amount(row.get('도급비', 0))
    material_cost = _safe_parse_amount(row.get('자재비', 0))
    other_cost = _safe_parse_amount(row.get('기타비', 0))

    total_costs = product_cost + contract_cost + material_cost + other_cost
    net_profit = total1 - total_costs

    return net_profit


def _calculate_margin_rate(row: pd.Series, net_profit: float) -> float:
    """마진율 계산: (순익 / 총액1) * 100

    Args:
        row: DataFrame의 한 행
        net_profit: 순익 (미리 계산된 값)

    Returns:
        마진율 (%) - -1000% ~ 1000% 범위로 제한
    """
    total1 = _safe_parse_amount(row.get('총액 1', 0))

    # 총액1이 0이면 마진율 0
    if total1 == 0:
        return 0.0

    # 총액1이 음수면 마진율 0 (비정상 케이스)
    if total1 < 0:
        logger.warning(f"[MARGIN] 총액1이 음수: {total1} (프로젝트: {row.get('프로젝트 코드', 'N/A')})")
        return 0.0

    margin_rate = (net_profit / total1) * 100

    # NaN/Infinity 검증
    if not float('inf') > margin_rate > float('-inf'):
        logger.warning(f"[MARGIN] 마진율 계산 오류: {margin_rate} (프로젝트: {row.get('프로젝트 코드', 'N/A')})")
        return 0.0

    # -1000% ~ 1000% 범위로 제한
    if margin_rate > 1000:
        logger.warning(f"[MARGIN] 마진율 1000% 초과: {margin_rate:.1f}% → 1000% (프로젝트: {row.get('프로젝트 코드', 'N/A')})")
        return 1000.0
    elif margin_rate < -1000:
        logger.warning(f"[MARGIN] 마진율 -1000% 미만: {margin_rate:.1f}% → -1000% (프로젝트: {row.get('프로젝트 코드', 'N/A')})")
        return -1000.0

    # 소수점 1자리로 반올림 (UI에서 과도한 소수점 노출 방지)
    return round(margin_rate, 1)


def _apply_backend_calculations(df: pd.DataFrame) -> pd.DataFrame:
    """백엔드에서 계산 필드(총액2, 미수금, 순익, 마진율) 계산

    Google Sheets 수식에 의존하지 않고 백엔드에서 직접 계산하여
    레이스 컨디션 및 수식 재계산 타이밍 이슈 해결

    Args:
        df: 원본 DataFrame

    Returns:
        계산 필드가 추가된 DataFrame
    """
    logger.info("[BACKEND_CALC] 백엔드 계산 시작")

    # 총액2 계산 (부가세 플래그 기반)
    def calculate_total2_for_row(row):
        total1 = _safe_parse_amount(row.get('총액 1', 0))
        vat_flag = _parse_vat_flag(row.get('부가세'))
        return _calculate_total2(total1, vat_flag)

    df['총액 2'] = df.apply(calculate_total2_for_row, axis=1)

    # 미수금 계산 (총액2 사용)
    df['미수금'] = df.apply(_calculate_outstanding_amount, axis=1)

    # 순익 계산
    df['순익'] = df.apply(_calculate_net_profit, axis=1)

    # 마진율 계산 (순익 값을 재사용)
    df['마진율'] = df.apply(lambda row: _calculate_margin_rate(row, row['순익']), axis=1)

    logger.info(f"[BACKEND_CALC] 완료: {len(df)}행 계산됨")

    return df


@handle_error(category=ErrorCategory.HIGH, reraise=False, fallback_value=None)
def load_data(force_refresh: bool = False, skip_cache: bool = False) -> Optional[pd.DataFrame]:
    import threading
    import time
    global _load_lock

    # Lazy lock initialization (thread-safe)
    if _load_lock is None:
        _load_lock = threading.RLock()

    cache_key = "current_sheet_data"

    if not force_refresh:
        cached = smart_get(cache_key, CacheStrategy.CRITICAL_DATA)
        if cached is not None:
            logger.debug("Using cached project data")
            return cached

    # Lock으로 보호: 동시에 여러 스레드가 Google Sheets API를 호출하지 않도록
    with _load_lock:
        # Double-check: lock을 얻은 후 다시 캐시 확인 (다른 스레드가 이미 로드했을 수 있음)
        if not force_refresh:
            cached = smart_get(cache_key, CacheStrategy.CRITICAL_DATA)
            if cached is not None:
                logger.info("[LOAD_DATA] 다른 스레드가 이미 로드함 - 캐시 사용")
                return cached

        # 🔒 레이스 컨디션 방지: 데이터 수집 시작 시각 기록
        fetch_start = time.time()
        logger.info(f"[LOAD_DATA][PID:{os.getpid()}] Google Sheets 로드 시작 (Lock 획득, fetch_start: {fetch_start})")

        try:
            sheet_id = os.getenv('GOOGLE_SHEET_ID')
            if not sheet_id:
                logger.error("GOOGLE_SHEET_ID is not configured")
                return None

            manager = get_sheets_manager()
            sheet_range = PROJECT_CONFIG.get('sheet_range', '공사 현황!A:AN')
            logger.info(f"Loading project data from range: {sheet_range}")
            df = manager.get_sheet_data(sheet_id, sheet_range)

            if df.empty:
                logger.warning("Google Sheet returned no rows")
                return None

            # 셀 노트(댓글) 가져오기 (계약금, 중도금, 잔금)
            try:
                sheet_name = PROJECT_CONFIG.get('sheet_name', '공사 현황의 사본')

                # 셀 노트 캐시 확인 (메모리 부하 최적화: TEMPORARY 전략 사용)
                notes_cache_key = f"cell_notes_{sheet_id}"
                notes_by_row = smart_get(notes_cache_key, CacheStrategy.TEMPORARY)

                if notes_by_row is None or force_refresh:
                    logger.info(f"[LOAD_NOTES] 셀 노트 가져오기 시작: {sheet_name}")
                    notes_by_row = manager.get_cell_notes(sheet_id, sheet_name, ['T', 'U', 'V'])

                    # 셀 노트 캐시에 저장 (TEMPORARY 전략)
                    smart_set(notes_cache_key, notes_by_row, CacheStrategy.TEMPORARY)
                    logger.info(f"[LOAD_NOTES] 셀 노트 캐시 저장: {len(notes_by_row)}개 행")
                else:
                    logger.info(f"[LOAD_NOTES] 캐시된 셀 노트 사용: {len(notes_by_row)}개 행")

                # DataFrame에 메모 컬럼 추가
                df['계약금_메모'] = None
                df['중도금_메모'] = None
                df['잔금_메모'] = None

                # 행 번호 기준으로 메모 매핑 (구글 시트 행 번호 - 2 = DataFrame 인덱스)
                # 구글 시트: 1행(헤더), 2행(데이터0), 3행(데이터1), ...
                mapped_count = 0
                for row_num, notes in notes_by_row.items():
                    # Redis 캐시에서 row_num이 문자열로 올 수 있으므로 정수로 변환
                    try:
                        df_index = int(row_num) - 2  # 헤더 제외, 0-based 인덱스
                    except (ValueError, TypeError):
                        continue  # 숫자로 변환할 수 없으면 건너뛰기

                    # 디버깅: 첫 3개 매핑 로그
                    if mapped_count < 3:
                        logger.info(f"[MEMO_MAP] 시트 행 {row_num} → DF 인덱스 {df_index}, notes: {notes}")

                    if 0 <= df_index < len(df):
                        if 'T' in notes:
                            df.at[df_index, '계약금_메모'] = notes['T']
                            if mapped_count < 3:
                                logger.info(f"[MEMO_MAP] 계약금_메모 매핑 완료: '{notes['T']}'")
                        if 'U' in notes:
                            df.at[df_index, '중도금_메모'] = notes['U']
                        if 'V' in notes:
                            df.at[df_index, '잔금_메모'] = notes['V']
                        mapped_count += 1
                    else:
                        if mapped_count < 3:
                            logger.warning(f"[MEMO_MAP] 인덱스 범위 벗어남: df_index={df_index}, len(df)={len(df)}")

                memo_count = sum(1 for row in notes_by_row.values() for _ in row)
                logger.info(f"[LOAD_NOTES] 완료: {len(notes_by_row)}개 행, {memo_count}개 메모")

            except Exception as note_error:
                logger.warning(f"[LOAD_NOTES] 셀 노트 가져오기 실패 (데이터는 정상 로드됨): {note_error}")
                # 노트 로드 실패해도 데이터는 정상 반환
                df['계약금_메모'] = None
                df['중도금_메모'] = None
                df['잔금_메모'] = None

            # ✨ 백엔드에서 계산 필드 계산 (미수금, 순익, 마진율)
            # Google Sheets 수식 재계산 타이밍에 의존하지 않음
            df = _apply_backend_calculations(df)

            # 정상 완료 시 캐시 저장 및 반환
            # 🔒 레이스 컨디션 방지: fetch_start 전달하여 무효화 마커와 비교
            # skip_cache=True면 캐시에 쓰지 않음 (재시도 검증용)
            if not skip_cache:
                smart_set(cache_key, df, CacheStrategy.CRITICAL_DATA, fetched_at=fetch_start)
                logger.info(f"[LOAD_DATA][PID:{os.getpid()}] 완료: {len(df)}행 로드됨 (캐시 저장, fetch_start: {fetch_start})")
            else:
                logger.info(f"[LOAD_DATA][PID:{os.getpid()}] 완료: {len(df)}행 로드됨 (캐시 스킵 - 검증용)")
            return df

        except Exception as exc:
            error_msg = str(exc)
            logger.error(f"[LOAD_DATA] Google Sheets 로드 실패: {error_msg}", exc_info=True)

            # 에러 유형별 상세 메시지
            if 'credentials' in error_msg.lower() or 'auth' in error_msg.lower():
                logger.error("[LOAD_DATA] 인증 오류: credentials.json 파일을 확인하세요")
            elif 'quota' in error_msg.lower() or 'rate' in error_msg.lower():
                logger.error("[LOAD_DATA] API 할당량 초과: 잠시 후 다시 시도하세요")
            elif 'network' in error_msg.lower() or 'connection' in error_msg.lower():
                logger.error("[LOAD_DATA] 네트워크 오류: 인터넷 연결을 확인하세요")

            # 로컬 백업 파일 시도
            try:
                logger.warning("[LOAD_DATA] 로컬 백업 파일로 전환 시도...")
                data_dir = Path(__file__).resolve().parent.parent / 'data'
                fallback_path = data_dir / '프로젝트공사 현황 (2).xlsx'
                df = pd.read_excel(fallback_path, sheet_name='공사 현황')
                logger.warning(f"[LOAD_DATA] 백업 파일 로드 성공: {len(df)}행 (최신 데이터가 아닐 수 있음)")
                return df
            except Exception as fallback_exc:
                logger.error(f"[LOAD_DATA] 백업 파일 로드 실패: {fallback_exc}")
                # 백업 파일도 실패하면 사용자에게 명확한 에러 제공
                logger.critical(
                    "[LOAD_DATA] 데이터 로드 완전 실패! "
                    "구글 시트와 백업 파일 모두 접근 불가. "
                    "1) 인터넷 연결 확인 2) credentials.json 확인 3) API 할당량 확인"
                )
                return None


def get_project_records(force_refresh: bool = False) -> List[Dict[str, Any]]:
    """
    프로젝트 레코드 조회 (dict 리스트 반환)

    Note:
        날짜 필드(공사 시작, 공사 종료, 수금 날짜, 공사 확정)는
        pandas Timestamp 객체를 문자열('YYYY-MM-DD')로 변환하여 반환
    """
    df = load_data(force_refresh=force_refresh)
    if df is None:
        return []
    if isinstance(df, pd.DataFrame):
        normalized = df.where(pd.notna(df), None)
        projects = normalized.to_dict('records')

        # 날짜 필드 후처리: Timestamp/NaT 객체를 문자열로 변환
        # (calendar_sync_scheduler에서 문자열 비교 시 TypeError 방지)
        date_columns = ['공사 시작', '공사 종료', '수금 날짜', '공사 확정']
        for project in projects:
            for date_col in date_columns:
                if date_col in project:
                    value = project[date_col]
                    # Timestamp 객체인 경우 문자열로 변환
                    if pd.notna(value) and hasattr(value, 'strftime'):
                        try:
                            project[date_col] = value.strftime('%Y-%m-%d')
                        except:
                            project[date_col] = ''
                    # NaT, None, 빈 문자열은 모두 빈 문자열로 통일
                    elif pd.isna(value) or value == '':
                        project[date_col] = ''

        return projects
    return df


def invalidate_project_cache(project_code: Optional[str] = None, trigger_refresh: bool = True) -> None:
    """프로젝트 데이터 캐시 무효화 (레이스 컨디션 방지 + 백그라운드 리프레시)

    Args:
        project_code: 특정 프로젝트 코드 (선택사항)
        trigger_refresh: True이면 백그라운드에서 즉시 새 데이터 가져오기

    Note:
        메타데이터 캐시(metadata_options)는 의도적으로 무효화하지 않음.
        사업자/담당자/거래처 목록은 거의 변경되지 않으므로 10분 TTL로 독립 관리됨.

        🔒 레이스 컨디션 방지:
        1. 무효화 마커 설정 → 오래된 데이터가 캐시에 쓰이는 것을 차단
        2. 백그라운드 리프레시 트리거 → 즉시 최신 데이터 가져오기
    """
    from dashboard.utils.smart_cache_manager import smart_set_invalidation_marker
    import threading

    # 1️⃣ 무효화 마커 설정 (레이스 컨디션 방지)
    smart_set_invalidation_marker("current_sheet_data")
    logger.info("[CACHE] 무효화 마커 설정: current_sheet_data")

    # 2️⃣ 캐시 삭제
    smart_invalidate("current_sheet_data")
    smart_invalidate("projects_list")
    if project_code:
        smart_invalidate(f"project_{project_code}")

    logger.info(f"[CACHE] 프로젝트 캐시 무효화 완료 (project_code: {project_code}, trigger_refresh: {trigger_refresh})")

    # 3️⃣ 백그라운드에서 즉시 새 데이터 가져오기 (비동기)
    if trigger_refresh:
        def background_refresh():
            """백그라운드에서 최신 데이터 가져오기 (간소화)

            백엔드 계산으로 전환하여 Google Sheets 수식 재계산 대기 불필요
            - 미수금, 순익, 마진율: 백엔드에서 직접 계산 (load_data 내부)
            - 재시도 로직 제거: 계산 필드가 즉시 생성됨
            - 락 점유 시간 최소화: 단일 API 호출만 수행
            """
            try:
                logger.info("[CACHE] 백그라운드 리프레시 시작")

                # 단순히 force_refresh=True로 최신 데이터 가져오기
                # load_data 내부에서 백엔드 계산이 자동으로 실행됨
                fresh_data = load_data(force_refresh=True)

                if fresh_data is not None:
                    logger.info(f"[CACHE] 백그라운드 리프레시 완료: {len(fresh_data)} rows")
                else:
                    logger.warning("[CACHE] 백그라운드 리프레시 실패: 데이터 없음")

            except Exception as e:
                logger.error(f"[CACHE] 백그라운드 리프레시 오류: {e}")

        refresh_thread = threading.Thread(target=background_refresh, daemon=True)
        refresh_thread.start()
        logger.debug("[CACHE] 백그라운드 리프레시 스레드 시작")


def invalidate_metadata_cache() -> None:
    """메타데이터 캐시 무효화 (수동 호출 전용)

    사업자/담당자/거래처가 실제로 변경되었을 때만 호출해야 함.
    일반적인 프로젝트 생성/수정/삭제에서는 호출하지 않음.
    """
    smart_invalidate("metadata_options")
    logger.info("[CACHE] 메타데이터 캐시 무효화됨")


def can_user_edit_project(project: Dict[str, Any], user_email: str, user_role: str) -> bool:
    try:
        if project.get('수금 관련 특이사항') == '공사취소':
            return False
        # 대소문자 구분 없이 역할 비교
        user_role_lower = user_role.lower()
        if user_role_lower == 'admin':
            return True
        owner_email = project.get('담당자 이메일', '')
        if owner_email == user_email:
            return True
        if user_role_lower in ['editor', 'user']:
            return True
        return False
    except Exception as exc:
        logger.error("Permission check failed: %s", exc)
        return False


def check_overdue_status(project: Dict[str, Any]) -> bool:
    try:
        confirm_date = project.get('공사 확정')
        deposit_date = project.get('계약금 입금일')
        if confirm_date and not deposit_date:
            confirm_dt = pd.to_datetime(confirm_date, errors='coerce')
            if pd.notna(confirm_dt):
                days_passed = (datetime.now() - confirm_dt.to_pydatetime()).days
                return days_passed > PROJECT_CONFIG.get('overdue_confirm_days', 2)
        return False
    except Exception as exc:
        logger.error("Overdue status check failed: %s", exc)
        return False


def _extract_number(code: str) -> Optional[int]:
    match = re.match(r'[A-Z](\d{4})-', str(code))
    return int(match.group(1)) if match else None


def _suffix_from_code(code: str) -> Optional[str]:
    match = re.match(r'[A-Z]\d{4}-([A-Z]+)$', str(code))
    return match.group(1) if match else None


def _build_company_prefix_map(df: pd.DataFrame) -> Dict[str, str]:
    mapping: Dict[str, str] = {}
    if '프로젝트 코드' in df.columns and ('프로젝트 담당자' in df.columns or '회사' in df.columns):
        for _, row in df.iterrows():
            code = str(row.get('프로젝트 코드', ''))
            company = str(row.get('프로젝트 담당자', row.get('회사', ''))).strip()
            match = re.match(r'([A-Z])\d{4}-', code)
            if company and match and company not in mapping:
                mapping[company] = match.group(1)
    for key, value in PROJECT_CONFIG.get('company_prefix_map', {}).items():
        mapping.setdefault(key, value)
    return mapping


def _build_owner_suffix_map(df: pd.DataFrame) -> Dict[str, str]:
    mapping = {k: str(v).upper() for k, v in PROJECT_CONFIG.get('owner_suffix_map', {}).items()}
    if '프로젝트 코드' in df.columns and ('프로젝트 담당자' in df.columns or '담당자 이메일' in df.columns):
        grouped: Dict[str, List[str]] = defaultdict(list)
        for _, row in df.iterrows():
            name = str(row.get('프로젝트 담당자', row.get('담당자 이메일', ''))).strip()
            code = str(row.get('프로젝트 코드', '')).strip()
            suffix = _suffix_from_code(code)
            if name and suffix:
                grouped[name].append(suffix)
        for name, suffixes in grouped.items():
            if suffixes:
                mapping.setdefault(name, Counter(suffixes).most_common(1)[0][0])
    return mapping


def _next_running_number(df: pd.DataFrame) -> int:
    """
    다음 프로젝트 번호 반환 (DB 시퀀스 기반, 전체 통합 번호)

    Args:
        df: 프로젝트 데이터프레임 (기존 데이터 max값 확인용)

    Returns:
        int: 다음 사용 가능한 번호

    Note:
        - 기존 규칙 유지: 모든 회사가 같은 번호 시퀀스 공유
        - ITG0001, ABC0002, ITG0003, XYZ0004 (전체 통합 번호)
        - DB 시퀀스를 사용하여 다중 프로세스 환경에서 안전
        - 시퀀스 최초 생성 시 기존 시트의 max값으로 초기화 (중복 방지)
    """
    from dashboard.utils.user_database import get_user_database

    user_db = get_user_database()

    # 전체 통합 번호 시퀀스 (기존 규칙 유지)
    pattern = "GLOBAL"

    # 기존 데이터에서 최대 번호 추출 (시퀀스 초기화용)
    max_number = 0
    if '프로젝트 코드' in df.columns:
        numbers: List[int] = []
        for code in df['프로젝트 코드'].astype(str):
            number = _extract_number(code)
            if number is not None:
                numbers.append(number)
        max_number = max(numbers) if numbers else 0

    # DB 시퀀스에서 다음 번호 가져오기 (ATOMIC)
    try:
        # 매번 스프레드시트 max와 비교하여 시퀀스 자동 조정
        next_number = user_db.get_next_project_number(pattern, current_max_from_sheet=max_number)
        return next_number
    except Exception as e:
        # DB 실패 시 폴백: 기존 방식 사용
        import logging
        logger = logging.getLogger(__name__)
        logger.warning(f"DB 시퀀스 실패, 폴백 사용: {e}")
        return max_number + 1


def _safe_next_running_number_with_retry(company: str, owner: str, max_retries: int = 5) -> str:
    """
    프로젝트 코드 생성 (DB 시퀀스 기반)

    Args:
        company: 회사명
        owner: 담당자명
        max_retries: 재시도 횟수 (사용 안 함, 하위 호환성)

    Returns:
        str: 생성된 프로젝트 코드 (예: "ITG0001-KW")

    Note:
        - threading.RLock 제거 (다중 프로세스 환경에서 작동 안 함)
        - DB 시퀀스로 변경하여 Gunicorn workers=4 환경에서도 안전
        - 재시도 로직 제거 (DB 시퀀스가 ATOMIC하게 처리)
    """
    df = load_data()
    if df is None:
        raise RuntimeError('Failed to load project data from source')

    # DB 시퀀스 기반으로 코드 생성 (중복 불가능)
    code = _auto_project_code(df, company, owner)
    return code


def _auto_project_code(df: pd.DataFrame, company: str, owner: str) -> str:
    comp_map = _build_company_prefix_map(df)
    own_map = _build_owner_suffix_map(df)

    prefix = comp_map.get(company.strip())
    suffix = own_map.get(owner.strip())

    if not prefix or not suffix:
        available_companies = ', '.join(comp_map.keys()) or 'unregistered'
        available_owners = ', '.join(own_map.keys()) or 'unregistered'
        message = (
            'Failed to generate project code. Verify company/owner mappings.\n'
            f'Available companies: {available_companies}\n'
            f'Available owners: {available_owners}'
        )
        raise ValueError(message)

    # 전체 통합 번호 사용 (기존 규칙 유지)
    number = _next_running_number(df)
    return f"{prefix}{number:04d}-{suffix}"
