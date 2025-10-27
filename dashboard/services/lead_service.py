"""
고객 리드 관리 서비스
Google Sheets '고객 리드 관리' 시트와 연동
"""

import logging
import os
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

import pandas as pd

from dashboard.utils.google_sheets import GoogleSheetsManager
from dashboard.utils.smart_cache_manager import CacheStrategy, smart_get, smart_invalidate, smart_set
from dashboard.utils.error_handler import handle_error, ErrorCategory
from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)

_sheets_manager: Optional[GoogleSheetsManager] = None


def get_sheets_manager() -> GoogleSheetsManager:
    """Google Sheets Manager 싱글톤 인스턴스 반환"""
    global _sheets_manager
    if _sheets_manager is None:
        _sheets_manager = GoogleSheetsManager()
    return _sheets_manager


def load_leads_data(force_refresh: bool = False) -> Optional[pd.DataFrame]:
    """
    고객 리드 관리 시트 데이터 로드

    Args:
        force_refresh: 캐시 무시하고 강제로 새로 로드

    Returns:
        리드 데이터 DataFrame, 실패 시 None
    """
    try:
        # 캐시 확인
        cache_key = "leads_sheet_data"

        if not force_refresh:
            cached_data = smart_get(cache_key, CacheStrategy.TEMPORARY)
            if cached_data is not None:
                logger.debug(f"[LEADS] 캐시된 리드 데이터 사용: {len(cached_data)}행")
                return cached_data

        # Google Sheets에서 데이터 가져오기
        sheet_id = os.getenv('GOOGLE_SHEET_ID')
        if not sheet_id:
            logger.error("[LEADS] GOOGLE_SHEET_ID가 설정되지 않았습니다")
            return None

        manager = get_sheets_manager()
        sheet_range = "고객 리드 관리!A:K"  # 리드 No ~ 비고

        logger.debug(f"[LEADS] Google Sheets 데이터 가져오기: {sheet_range}")
        df = manager.get_sheet_data(sheet_id, sheet_range)

        if df.empty:
            logger.warning("[LEADS] Google Sheet에서 빈 데이터 반환됨")
            return None

        logger.info(f"[LEADS] Google Sheets 로드 성공: {len(df)}행")

        # 캐시 저장 (TEMPORARY 전략 사용)
        smart_set(cache_key, df, CacheStrategy.TEMPORARY)

        return df

    except Exception as exc:
        logger.error(f"[LEADS] Google Sheets 로드 실패: {exc}", exc_info=True)
        return None


def get_lead_records() -> List[Dict[str, Any]]:
    """
    리드 목록 조회

    Returns:
        리드 딕셔너리 리스트
    """
    try:
        df = load_leads_data()

        if df is None or df.empty:
            logger.warning("[LEADS] 리드 데이터가 없습니다")
            return []

        # DataFrame을 딕셔너리 리스트로 변환
        records = df.to_dict('records')

        # NaN 값을 None 또는 빈 문자열로 변환
        for record in records:
            for key, value in record.items():
                if pd.isna(value):
                    record[key] = '' if isinstance(value, str) else None

        logger.debug(f"[LEADS] 리드 레코드 반환: {len(records)}건")
        return records

    except Exception as exc:
        logger.error(f"[LEADS] 리드 레코드 조회 실패: {exc}", exc_info=True)
        return []


def get_lead_by_no(lead_no: str) -> Optional[Dict[str, Any]]:
    """
    리드 번호로 단일 리드 조회

    Args:
        lead_no: 리드 번호 (예: L0001)

    Returns:
        리드 딕셔너리, 없으면 None
    """
    try:
        records = get_lead_records()

        for record in records:
            if record.get('리드 No') == lead_no:
                return record

        logger.warning(f"[LEADS] 리드 번호 {lead_no}를 찾을 수 없습니다")
        return None

    except Exception as exc:
        logger.error(f"[LEADS] 리드 조회 실패: {exc}", exc_info=True)
        return None


def create_lead(lead_data: Dict[str, Any]) -> Dict[str, Any]:
    """
    새 리드 생성

    Args:
        lead_data: 리드 정보 딕셔너리
            - 거래처 (필수)
            - 상태 (기본값: "유선 상담")
            - 방문 예정일
            - 담당자
            - 연락처 (필수)
            - 고객명 (필수)
            - 방문 주소
            - 상담 내용
            - 마지막 연락일
            - 비고

    Returns:
        {'success': bool, 'lead_no': str, 'message': str}
    """
    try:
        # 필수 필드 검증
        required_fields = ['거래처', '연락처', '고객명']
        for field in required_fields:
            if not lead_data.get(field):
                return {
                    'success': False,
                    'message': f'필수 필드가 누락되었습니다: {field}'
                }

        # Google Sheets Manager 가져오기
        sheet_id = os.getenv('GOOGLE_SHEET_ID')
        if not sheet_id:
            return {'success': False, 'message': 'GOOGLE_SHEET_ID가 설정되지 않았습니다'}

        manager = get_sheets_manager()
        sheet_name = "고객 리드 관리"

        # 다음 리드 번호 생성 (L0001, L0002, ...)
        df = load_leads_data(force_refresh=True)

        if df is None or df.empty:
            next_number = 1
        else:
            # 기존 리드 번호에서 최대값 찾기
            lead_numbers = df['리드 No'].dropna().tolist()
            max_number = 0

            for lead_no in lead_numbers:
                if isinstance(lead_no, str) and lead_no.startswith('L'):
                    try:
                        num = int(lead_no[1:])  # L0001 → 1
                        max_number = max(max_number, num)
                    except ValueError:
                        continue

            next_number = max_number + 1

        new_lead_no = f"L{next_number:04d}"

        # 새 행 데이터 준비
        new_row = [
            new_lead_no,                                    # A: 리드 No
            lead_data.get('거래처', ''),                    # B: 거래처
            lead_data.get('상태', '유선 상담'),              # C: 상태
            lead_data.get('방문 예정일', ''),                # D: 방문 예정일
            lead_data.get('담당자', ''),                    # E: 담당자
            lead_data.get('연락처', ''),                    # F: 연락처
            lead_data.get('고객명', ''),                    # G: 고객명
            lead_data.get('방문 주소', ''),                  # H: 방문 주소
            lead_data.get('상담 내용', ''),                  # I: 상담 내용
            lead_data.get('마지막 연락일', ''),              # J: 마지막 연락일
            lead_data.get('비고', '')                       # K: 비고
        ]

        # Google Sheets에 추가
        manager.append_row(sheet_id, sheet_name, new_row)

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
        update_data: 업데이트할 필드 딕셔너리

    Returns:
        {'success': bool, 'message': str}
    """
    try:
        # 리드 존재 확인
        df = load_leads_data(force_refresh=True)

        if df is None or df.empty:
            return {'success': False, 'message': '리드 데이터를 찾을 수 없습니다'}

        # 리드 행 찾기
        lead_row_index = None
        for idx, row in df.iterrows():
            if row['리드 No'] == lead_no:
                lead_row_index = idx
                break

        if lead_row_index is None:
            return {'success': False, 'message': f'리드 {lead_no}를 찾을 수 없습니다'}

        # Google Sheets에서 행 번호 계산 (헤더 +1, 0-based index +1)
        sheet_row_number = lead_row_index + 2

        # Google Sheets Manager
        sheet_id = os.getenv('GOOGLE_SHEET_ID')
        manager = get_sheets_manager()
        sheet_name = "고객 리드 관리"

        # 컬럼 매핑 (업데이트 가능한 필드만)
        column_mapping = {
            '거래처': 'B',
            '상태': 'C',
            '방문 예정일': 'D',
            '담당자': 'E',
            '연락처': 'F',
            '고객명': 'G',
            '방문 주소': 'H',
            '상담 내용': 'I',
            '마지막 연락일': 'J',
            '비고': 'K'
        }

        # 업데이트할 셀 준비
        updates = []
        for field_name, value in update_data.items():
            if field_name in column_mapping:
                col = column_mapping[field_name]
                cell_range = f"{col}{sheet_row_number}"
                updates.append((cell_range, value))

        if not updates:
            return {'success': False, 'message': '업데이트할 필드가 없습니다'}

        # Google Sheets 업데이트
        for cell_range, value in updates:
            manager.update_cell(sheet_id, sheet_name, cell_range, value)

        logger.info(f"[LEADS] 리드 {lead_no} 업데이트 완료: {len(updates)}개 필드")

        # 캐시 무효화
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

    Args:
        lead_no: 리드 번호
        new_status: 새 상태 (유선 상담, 방문 예약, 견적 제출, 방문 취소, 공사 확정, 공사 드랍)

    Returns:
        {'success': bool, 'message': str}
    """
    return update_lead(lead_no, {'상태': new_status})


def invalidate_leads_cache():
    """리드 데이터 캐시 무효화"""
    try:
        smart_invalidate("leads_sheet_data")
        logger.debug("[LEADS] 캐시 무효화 완료")
    except Exception as exc:
        logger.warning(f"[LEADS] 캐시 무효화 실패: {exc}")


def search_leads_by_customer(customer_name: str, phone: str = None) -> List[Dict[str, Any]]:
    """
    고객명/연락처로 리드 검색 (프로젝트 등록 시 자동완성용)

    Args:
        customer_name: 고객명 (부분 일치)
        phone: 연락처 (선택)

    Returns:
        일치하는 리드 리스트
    """
    try:
        records = get_lead_records()

        results = []
        for record in records:
            # 고객명 매칭 (대소문자 무시, 부분 일치)
            name_match = customer_name.lower() in str(record.get('고객명', '')).lower()

            # 연락처 매칭 (선택)
            phone_match = True
            if phone:
                phone_match = phone in str(record.get('연락처', ''))

            if name_match and phone_match:
                results.append(record)

        logger.debug(f"[LEADS] 고객 검색 결과: {len(results)}건 (검색어: {customer_name})")
        return results

    except Exception as exc:
        logger.error(f"[LEADS] 고객 검색 실패: {exc}", exc_info=True)
        return []
