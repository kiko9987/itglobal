import json
import logging
import os
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

import pandas as pd
import re
from collections import Counter, defaultdict

from ..utils.google_sheets import GoogleSheetsManager
from ..utils.smart_cache_manager import CacheStrategy, smart_get, smart_invalidate, smart_set
from ..utils.error_handler import handle_error, ErrorCategory

logger = logging.getLogger(__name__)

_current_data: Optional[pd.DataFrame] = None
_last_update: Optional[datetime] = None
_sheets_manager: Optional[GoogleSheetsManager] = None


def _load_project_config() -> Dict[str, Any]:
    try:
        config_path = Path(__file__).resolve().parent.parent / 'project_config.json'
        with config_path.open('r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as exc:  # pragma: no cover
        logger.warning("프로젝트 설정 로드 실패: %s", exc)
        return {}


PROJECT_CONFIG: Dict[str, Any] = _load_project_config()


def get_project_config() -> Dict[str, Any]:
    return PROJECT_CONFIG


def get_sheets_manager() -> GoogleSheetsManager:
    global _sheets_manager
    if _sheets_manager is None:
        _sheets_manager = GoogleSheetsManager()
    return _sheets_manager


def get_current_data() -> Optional[pd.DataFrame]:
    return _current_data


def get_last_update() -> Optional[datetime]:
    return _last_update


def set_current_data(df: Optional[pd.DataFrame], update_timestamp: bool = True) -> None:
    global _current_data, _last_update
    _current_data = df
    if update_timestamp:
        _last_update = datetime.now()


def clear_current_data() -> None:
    set_current_data(None, update_timestamp=False)


@handle_error(category=ErrorCategory.HIGH, reraise=False, fallback_value=None)
def load_data(force_refresh: bool = False) -> Optional[pd.DataFrame]:
    cache_key = "current_sheet_data"

    if not force_refresh:
        cached = smart_get(cache_key, CacheStrategy.CRITICAL_DATA)
        if cached is not None:
            set_current_data(cached, update_timestamp=False)
            logger.debug("스마트 캐시 데이터를 사용합니다")
            return cached

    try:
        sheet_id = os.getenv('GOOGLE_SHEET_ID')
        if not sheet_id:
            logger.error("GOOGLE_SHEET_ID가 설정되지 않았습니다.")
            return None

        manager = get_sheets_manager()
        sheet_range = PROJECT_CONFIG.get('sheet_range', '공사 현황!A:AM')
        logger.info("데이터 로드 범위: %s", sheet_range)
        df = manager.get_sheet_data(sheet_id, sheet_range)

        if df.empty:
            logger.warning("구글 시트에서 데이터를 가져올 수 없습니다.")
            return None

        set_current_data(df)
        smart_set(cache_key, df, CacheStrategy.CRITICAL_DATA)
        logger.info("데이터 로드 완료: %s행", len(df))
        return df

    except Exception as exc:
        logger.error("데이터 로드 오류: %s", exc)

        try:
            data_dir = Path(__file__).resolve().parent.parent / 'data'
            fallback_path = data_dir / '프로젝트공사 현황 (2).xlsx'
            df = pd.read_excel(fallback_path, sheet_name='공사 현황')
            set_current_data(df)
            logger.info("로컬 파일에서 데이터를 로드했습니다: %s행", len(df))
            return df
        except Exception as fallback_exc:  # pragma: no cover
            logger.error("로컬 파일 로드도 실패: %s", fallback_exc)
            return None




def can_user_edit_project(project, user_email, user_role):
    """사용자가 프로젝트를 편집할 수 있는지 확인"""
    try:
        if project.get('수금 관련 특이사항') == '공사취소':
            return False
        if user_role == 'admin':
            return True
        project_owner_email = project.get('담당자 이메일', '')
        if project_owner_email == user_email:
            return True
        if user_role in ['editor', 'user']:
            return True
        return False
    except Exception as exc:
        logger.error("편집 권한 확인 오류: %s", exc)
        return False


def check_overdue_status(project):
    """프로젝트 연체 상태 확인"""
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
        logger.error("연체 상태 확인 오류: %s", exc)
        return False


def _extract_number(code: str) -> Optional[int]:
    match = re.match(r'[A-Z](\d{4})-', str(code))
    return int(match.group(1)) if match else None


def _suffix_from_code(code: str) -> Optional[str]:
    match = re.match(r'[A-Z]\d{4}-([A-Z]+)$', str(code))
    return match.group(1) if match else None


def _build_company_prefix_map(df: pd.DataFrame) -> Dict[str, str]:
    mapping: Dict[str, str] = {}
    if '프로젝트 코드' in df.columns and '프로젝트 담당자' in df.columns:
        for _, row in df.iterrows():
            code = str(row.get('프로젝트 코드', ''))
            company = str(row.get('프로젝트 담당자', '')).strip()
            match = re.match(r'([A-Z])\d{4}-', code)
            if company and match and company not in mapping:
                mapping[company] = match.group(1)
    for key, value in PROJECT_CONFIG.get('company_prefix_map', {}).items():
        mapping.setdefault(key, value)
    return mapping


def _build_owner_suffix_map(df: pd.DataFrame) -> Dict[str, str]:
    mapping = {k: str(v).upper() for k, v in PROJECT_CONFIG.get('owner_suffix_map', {}).items()}
    if '프로젝트 코드' in df.columns and '프로젝트 담당자' in df.columns:
        grouped: Dict[str, List[str]] = defaultdict(list)
        for _, row in df.iterrows():
            name = str(row.get('프로젝트 담당자', '')).strip()
            code = str(row.get('프로젝트 코드', '')).strip()
            suffix = _suffix_from_code(code)
            if name and suffix:
                grouped[name].append(suffix)
        for name, suffixes in grouped.items():
            if suffixes:
                mapping.setdefault(name, Counter(suffixes).most_common(1)[0][0])
    return mapping


def _next_running_number(df: pd.DataFrame) -> int:
    numbers: List[int] = []
    if '프로젝트 코드' in df.columns:
        for code in df['프로젝트 코드'].astype(str):
            number = _extract_number(code)
            if number is not None:
                numbers.append(number)
    return (max(numbers) + 1) if numbers else 1



def _safe_next_running_number_with_retry(company: str, owner: str, max_retries: int = 5) -> str:
    import threading
    import time

    if not hasattr(_safe_next_running_number_with_retry, '_lock'):
        _safe_next_running_number_with_retry._lock = threading.RLock()

    for attempt in range(max_retries):
        with _safe_next_running_number_with_retry._lock:
            df = load_data()
            if df is None:
                raise RuntimeError('Failed to load project data from source')

            code = _auto_project_code(df, company, owner)
            if '프로젝트 코드' in df.columns:
                existing = df['프로젝트 코드'].astype(str).tolist()
                if code in existing:
                    if attempt < max_retries - 1:
                        time.sleep(0.1 * (attempt + 1))
                        continue
                    raise RuntimeError(f'Project code generation failed after {max_retries} attempts ({code})')
            return code

    raise RuntimeError('Project code generation failed: unknown error')



def _auto_project_code(df: pd.DataFrame, company: str, owner: str) -> str:
    comp_map = _build_company_prefix_map(df)
    own_map = _build_owner_suffix_map(df)

    prefix = comp_map.get(company.strip())
    suffix = own_map.get(owner.strip())

    if not prefix or not suffix:
        available_companies = ', '.join(comp_map.keys()) or 'unregistered'
        available_owners = ', '.join(own_map.keys()) or 'unregistered'
        message = ("Failed to generate project code. Verify company/owner mappings.\n"
                   f"Available companies: {available_companies}\n"
                   f"Available owners: {available_owners}")
        raise ValueError(message)

    number = _next_running_number(df)
    return f"{prefix}{number:04d}-{suffix}"

def get_project_records(force_refresh: bool = False) -> List[Dict[str, Any]]:
    df = load_data(force_refresh=force_refresh)
    if df is None:
        return []
    if isinstance(df, pd.DataFrame):
        normalized = df.where(pd.notna(df), None)
        return normalized.to_dict('records')
    return df


def invalidate_project_cache(project_code: Optional[str] = None) -> None:
    smart_invalidate("current_sheet_data")
    smart_invalidate("projects_list")
    if project_code:
        smart_invalidate(f"project_{project_code}")


def can_user_edit_project(project, user_email, user_role):
    """사용자가 프로젝트를 편집할 수 있는지 확인"""
    try:
        if project.get('수금 관련 특이사항') == '공사취소':
            return False
        if user_role == 'admin':
            return True
        project_owner_email = project.get('담당자 이메일', '')
        if project_owner_email == user_email:
            return True
        if user_role in ['editor', 'user']:
            return True
        return False
    except Exception as exc:
        logger.error("편집 권한 확인 오류: %s", exc)
        return False


def check_overdue_status(project):
    """프로젝트 연체 상태 확인"""
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
        logger.error("연체 상태 확인 오류: %s", exc)
        return False


def extract_numeric_suffix(code: str) -> Optional[int]:
    match = re.match(r'[A-Z](\d{4})-', str(code))
    return int(match.group(1)) if match else None


def extract_owner_suffix(code: str) -> Optional[str]:
    match = re.match(r'[A-Z]\d{4}-([A-Z]+)$', str(code))
    return match.group(1) if match else None


def build_company_prefix_map(df: pd.DataFrame) -> Dict[str, str]:
    mapping: Dict[str, str] = {}
    if '프로젝트 코드' in df.columns and '프로젝트 담당자' in df.columns:
        for _, row in df.iterrows():
            code = str(row.get('프로젝트 코드', ''))
            company = str(row.get('프로젝트 담당자', '')).strip()
            match = re.match(r'([A-Z])\d{4}-', code)
            if company and match and company not in mapping:
                mapping[company] = match.group(1)
    for key, value in PROJECT_CONFIG.get('company_prefix_map', {}).items():
        mapping.setdefault(key, value)
    return mapping


def build_owner_suffix_map(df: pd.DataFrame) -> Dict[str, str]:
    mapping = {k: str(v).upper() for k, v in PROJECT_CONFIG.get('owner_suffix_map', {}).items()}
    if '프로젝트 코드' in df.columns and '프로젝트 담당자' in df.columns:
        grouped: Dict[str, List[str]] = defaultdict(list)
        for _, row in df.iterrows():
            name = str(row.get('프로젝트 담당자', '')).strip()
            code = str(row.get('프로젝트 코드', '')).strip()
            suffix = extract_owner_suffix(code)
            if name and suffix:
                grouped[name].append(suffix)
        for name, suffixes in grouped.items():
            if suffixes:
                mapping.setdefault(name, Counter(suffixes).most_common(1)[0][0])
    return mapping


def next_running_number(df: pd.DataFrame) -> int:
    numbers: List[int] = []
    if '프로젝트 코드' in df.columns:
        for code in df['프로젝트 코드'].astype(str):
            number = extract_numeric_suffix(code)
            if number is not None:
                numbers.append(number)
    return (max(numbers) + 1) if numbers else 1
