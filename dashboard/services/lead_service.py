"""
고객 리드 관리 서비스
별도 Google Sheets '온라인 문의 현황' (탭명: '고객 리드 관리')와 연동
15열 구조: A 리드 No / B 상담 시간 / C 플랫폼 / D 상태 / E 방문 예정일 /
           F 고객 연락처 / G 이메일 / H 고객명 / I 방문 주소 / J 상담 내용 /
           K 키워드 / L 온라인 상담자 / M 영업 담당자 / N 마지막 연락일 / O 피드백
"""

import logging
import os
from datetime import datetime
from typing import Any, Dict, List, Optional

import pandas as pd

from dashboard.utils.google_sheets import GoogleSheetsManager
from dashboard.utils.smart_cache_manager import CacheStrategy, smart_get, smart_invalidate, smart_set
from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)



# ─────────────────────────────────────────────────────────────
# 시트 구조 상수 (헤더 행: 1행, 데이터: 2행~)
# ─────────────────────────────────────────────────────────────
LEAD_COLUMN_MAPPING: Dict[str, str] = {
    '리드 No':        'A',
    '상담 시간':       'B',
    '플랫폼':          'C',
    '상태':           'D',
    '방문 예정일':     'E',
    '고객 연락처':     'F',
    '이메일':         'G',
    '고객명':         'H',
    '방문 주소':       'I',
    '문의 내용':       'J',  # 인입 시 원본 (옛 '상담 내용')
    '상담 내용':       'K',  # 매니저 상담 결과 입력 (옛 '피드백')
    '키워드':         'L',
    '온라인 상담자':   'M',
    '영업 담당자':     'N',
    '마지막 연락일':   'O',
    '폴더 ID':        'P',  # 방문 사진 업로드 시 자동 write (2026-07 신규)
}

# 헤더 순서대로의 컬럼 리스트 (새 행 작성 시 사용)
LEAD_COLUMN_ORDER: List[str] = list(LEAD_COLUMN_MAPPING.keys())

# 컬럼 alias - 기존 코드 호환을 위해 일부 옛 이름도 받아들임
LEAD_FIELD_ALIASES: Dict[str, str] = {
    '거래처':   '플랫폼',
    '담당자':   '영업 담당자',
    '연락처':   '고객 연락처',
    '비고':     '상담 내용',   # 옛 '비고' → 옛 '피드백' → 새 '상담 내용'
    '피드백':   '상담 내용',   # 옛 '피드백' → 새 '상담 내용' (매니저 입력)
}


def _normalize_field_name(name: str) -> str:
    """옛 필드명을 새 컬럼명으로 정규화."""
    return LEAD_FIELD_ALIASES.get(name, name)


def get_sheets_manager() -> GoogleSheetsManager:
    """GoogleSheetsManager 반환 (클래스 자체가 스레드별 인스턴스 관리)."""
    return GoogleSheetsManager()


def _get_sheet_config() -> Optional[Dict[str, str]]:
    """환경변수에서 시트 ID/탭명 조회. 없으면 None."""
    sheet_id = os.getenv('ONLINE_LEADS_SHEET_ID')
    sheet_name = os.getenv('ONLINE_LEADS_SHEET_NAME', '고객 리드 관리')
    if not sheet_id:
        logger.error("[LEADS] ONLINE_LEADS_SHEET_ID 환경변수가 설정되지 않았습니다")
        return None
    return {'sheet_id': sheet_id, 'sheet_name': sheet_name}


def load_leads_data(force_refresh: bool = False) -> Optional[pd.DataFrame]:
    """
    온라인 문의 현황 시트 데이터 로드 (15열).

    Args:
        force_refresh: 캐시 무시하고 강제로 새로 로드

    Returns:
        리드 데이터 DataFrame, 실패 시 None
    """
    try:
        cache_key = "online_leads_sheet_data"

        if not force_refresh:
            cached_data = smart_get(cache_key, CacheStrategy.TEMPORARY)
            if cached_data is not None:
                logger.debug(f"[LEADS] 캐시된 리드 데이터 사용: {len(cached_data)}행")
                return cached_data

        cfg = _get_sheet_config()
        if cfg is None:
            return None

        manager = get_sheets_manager()
        sheet_range = f"{cfg['sheet_name']}!A:P"  # 16열 (A~P, 2026-07 폴더 ID 추가)

        logger.debug(f"[LEADS] Google Sheets 데이터 가져오기: {sheet_range}")
        df = manager.get_sheet_data(cfg['sheet_id'], sheet_range)

        if df.empty:
            logger.warning("[LEADS] Google Sheet에서 빈 데이터 반환됨")
            return None

        # 시트 헤더가 'No.'로 되어 있으면 '리드 No'로 정규화
        # (시트 운영자가 헤더를 다르게 적었어도 코드는 '리드 No'로 일관 처리)
        rename_map = {}
        if 'No.' in df.columns and '리드 No' not in df.columns:
            rename_map['No.'] = '리드 No'
        if rename_map:
            df = df.rename(columns=rename_map)

        logger.info(f"[LEADS] Google Sheets 로드 성공: {len(df)}행 / 컬럼 {list(df.columns)}")

        # 캐시 저장
        smart_set(cache_key, df, CacheStrategy.TEMPORARY)

        return df

    except Exception as exc:
        logger.error(f"[LEADS] Google Sheets 로드 실패: {exc}", exc_info=True)
        return None


def get_lead_records() -> List[Dict[str, Any]]:
    """
    리드 목록 조회

    Returns:
        리드 딕셔너리 리스트 (15개 컬럼 키)
    """
    try:
        df = load_leads_data()

        if df is None or df.empty:
            logger.warning("[LEADS] 리드 데이터가 없습니다")
            return []

        # DataFrame을 딕셔너리 리스트로 변환
        records = df.to_dict('records')

        # NaN 값을 빈 문자열로 변환
        for record in records:
            for key, value in list(record.items()):
                if pd.isna(value):
                    record[key] = ''

        logger.debug(f"[LEADS] 리드 레코드 반환: {len(records)}건")
        return records

    except Exception as exc:
        logger.error(f"[LEADS] 리드 레코드 조회 실패: {exc}", exc_info=True)
        return []


def get_lead_by_no(lead_no: str) -> Optional[Dict[str, Any]]:
    """
    리드 번호로 단일 리드 조회

    Args:
        lead_no: 리드 번호 (예: L-00001)

    Returns:
        리드 딕셔너리, 없으면 None
    """
    try:
        records = get_lead_records()

        for record in records:
            if str(record.get('리드 No', '')).strip() == str(lead_no).strip():
                return record

        logger.warning(f"[LEADS] 리드 번호 {lead_no}를 찾을 수 없습니다")
        return None

    except Exception as exc:
        logger.error(f"[LEADS] 리드 조회 실패: {exc}", exc_info=True)
        return None


def _generate_next_lead_no(df: Optional[pd.DataFrame]) -> str:
    """
    다음 리드 번호 생성.
    시트의 기존 포맷(L-00001 / L0001 둘 다)을 감지해 큰 쪽으로 +1.
    """
    max_number = 0
    if df is not None and not df.empty and '리드 No' in df.columns:
        for lead_no in df['리드 No'].dropna().tolist():
            if not isinstance(lead_no, str):
                continue
            s = lead_no.strip()
            if not s.upper().startswith('L'):
                continue
            digits = s[1:].lstrip('-')  # "L-00001" → "00001", "L0001" → "0001"
            try:
                num = int(digits)
                max_number = max(max_number, num)
            except ValueError:
                continue

    # 시트의 패턴이 L-00001 (하이픈 + 5자리)이므로 그 포맷을 따름
    return f"L-{max_number + 1:05d}"


def create_lead(lead_data: Dict[str, Any]) -> Dict[str, Any]:
    """
    새 리드 생성

    Args:
        lead_data: 리드 정보 딕셔너리 (15개 컬럼 키 또는 옛 키 alias 허용)
            - 고객명, 고객 연락처 등 필수
            - 상태 미지정 시 '유선 상담' 기본값

    Returns:
        {'success': bool, 'lead_no': str, 'message': str}
    """
    try:
        # alias 정규화 (옛 키가 들어오면 새 키로 변환)
        normalized = {_normalize_field_name(k): v for k, v in lead_data.items()}

        # 필수 필드 검증
        required_fields = ['고객명', '고객 연락처']
        for field in required_fields:
            if not normalized.get(field):
                return {
                    'success': False,
                    'message': f'필수 필드가 누락되었습니다: {field}'
                }

        cfg = _get_sheet_config()
        if cfg is None:
            return {'success': False, 'message': 'ONLINE_LEADS_SHEET_ID가 설정되지 않았습니다'}

        manager = get_sheets_manager()

        # 다음 리드 번호 생성
        df = load_leads_data(force_refresh=True)
        new_lead_no = _generate_next_lead_no(df)

        # 상담 시간 기본값 (현재 시각)
        if not normalized.get('상담 시간'):
            normalized['상담 시간'] = datetime.now().strftime('%Y.%m.%d. %H:%M')

        # 상태 기본값
        if not normalized.get('상태'):
            normalized['상태'] = '유선 상담'

        # 리드 No는 자동 생성값으로 덮어쓰기
        normalized['리드 No'] = new_lead_no

        # 헤더 순서에 맞춰 행 데이터 구성
        new_row = [normalized.get(col, '') for col in LEAD_COLUMN_ORDER]

        manager.append_row(cfg['sheet_id'], cfg['sheet_name'], new_row)

        logger.info(f"[LEADS] 새 리드 생성 완료: {new_lead_no}")

        # 캐시 무효화
        invalidate_leads_cache()

        return {
            'success': True,
            'lead_no': new_lead_no,
            'message': f'리드 {new_lead_no}가 생성되었습니다'
        }

    except Exception as exc:
        logger.error(f"[LEADS] 리드 생성 실패: {exc}", exc_info=True)
        return {'success': False, 'message': f'리드 생성 중 오류 발생: {str(exc)}'}


def update_lead(lead_no: str, update_data: Dict[str, Any]) -> Dict[str, Any]:
    """
    리드 정보 업데이트

    Args:
        lead_no: 리드 번호
        update_data: 업데이트할 필드 딕셔너리 (옛 키 alias 허용)

    Returns:
        {'success': bool, 'message': str}
    """
    try:
        df = load_leads_data(force_refresh=True)

        if df is None or df.empty:
            return {'success': False, 'message': '리드 데이터를 찾을 수 없습니다'}

        # 리드 행 찾기 (문자열 strip 비교)
        lead_row_index = None
        for idx, row in df.iterrows():
            if str(row.get('리드 No', '')).strip() == str(lead_no).strip():
                lead_row_index = idx
                break

        if lead_row_index is None:
            return {'success': False, 'message': f'리드 {lead_no}를 찾을 수 없습니다'}

        # Google Sheets 행 번호 (헤더 1행 + 0-based → +2)
        sheet_row_number = lead_row_index + 2

        cfg = _get_sheet_config()
        if cfg is None:
            return {'success': False, 'message': 'ONLINE_LEADS_SHEET_ID가 설정되지 않았습니다'}

        manager = get_sheets_manager()

        # 업데이트할 셀 준비 (alias 정규화)
        updates: List[tuple] = []
        for field_name, value in update_data.items():
            canonical = _normalize_field_name(field_name)
            col = LEAD_COLUMN_MAPPING.get(canonical)
            if not col:
                logger.warning(f"[LEADS] 알 수 없는 필드 무시: {field_name}")
                continue
            # 리드 No는 식별자이므로 업데이트 금지
            if canonical == '리드 No':
                continue
            cell_range = f"{col}{sheet_row_number}"
            updates.append((cell_range, value))

        if not updates:
            return {'success': False, 'message': '업데이트할 필드가 없습니다'}

        # 1회 batchUpdate API 호출로 모든 셀 갱신 (GoogleSheetsManager.update_cell 부재)
        # 일시 에러(SSL/timeout/connection/NoneType .read 등)는 최대 2회 자동 재시도
        import time as _t
        batch_body = {
            'valueInputOption': 'USER_ENTERED',
            'data': [
                {
                    'range': f"'{cfg['sheet_name']}'!{cell_range}",
                    'values': [[value]],
                }
                for cell_range, value in updates
            ],
        }
        last_exc = None
        for attempt in range(3):
            try:
                manager.service.spreadsheets().values().batchUpdate(
                    spreadsheetId=cfg['sheet_id'], body=batch_body,
                ).execute()
                last_exc = None
                break
            except Exception as exc:
                last_exc = exc
                err_l = str(exc).lower()
                is_transient = (
                    'ssl' in err_l or 'wrong_version' in err_l
                    or 'timeout' in err_l or 'connection' in err_l
                    or 'nonetype' in err_l or 'read' in err_l
                )
                if not is_transient or attempt == 2:
                    raise
                logger.warning(
                    f"[LEADS] 리드 {lead_no} 일시 에러, "
                    f"{attempt+1}/2 재시도: {exc}"
                )
                _t.sleep(2)
        if last_exc:
            raise last_exc

        logger.info(f"[LEADS] 리드 {lead_no} 업데이트 완료: {len(updates)}개 필드")

        invalidate_leads_cache()

        return {
            'success': True,
            'message': f'리드 {lead_no}가 업데이트되었습니다'
        }

    except Exception as exc:
        logger.error(f"[LEADS] 리드 업데이트 실패: {exc}", exc_info=True)
        return {'success': False, 'message': f'리드 업데이트 중 오류 발생: {str(exc)}'}


def update_lead_status(lead_no: str, new_status: str) -> Dict[str, Any]:
    """
    리드 상태만 업데이트 (빠른 업데이트)
    """
    return update_lead(lead_no, {'상태': new_status})


def invalidate_leads_cache():
    """리드 데이터 캐시 무효화"""
    try:
        smart_invalidate("online_leads_sheet_data")
        logger.debug("[LEADS] 캐시 무효화 완료")
    except Exception as exc:
        logger.warning(f"[LEADS] 캐시 무효화 실패: {exc}")


def search_leads_by_customer(customer_name: str, phone: str = None) -> List[Dict[str, Any]]:
    """
    고객명/연락처로 리드 검색 (프로젝트 등록 시 자동완성용)

    Args:
        customer_name: 고객명 (부분 일치)
        phone: 고객 연락처 (선택, 부분 일치)

    Returns:
        일치하는 리드 리스트
    """
    try:
        records = get_lead_records()

        results = []
        needle = customer_name.lower()
        for record in records:
            name_match = needle in str(record.get('고객명', '')).lower()

            phone_match = True
            if phone:
                phone_match = phone in str(record.get('고객 연락처', ''))

            if name_match and phone_match:
                results.append(record)

        logger.debug(f"[LEADS] 고객 검색 결과: {len(results)}건 (검색어: {customer_name})")
        return results

    except Exception as exc:
        logger.error(f"[LEADS] 고객 검색 실패: {exc}", exc_info=True)
        return []
