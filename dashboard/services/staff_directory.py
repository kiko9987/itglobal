# -*- coding: utf-8 -*-
"""직원 명단 한곳 — 공사 코드 접미사·비영업 인원·현재 영업 담당자 (2026-10-28 3차 정리).

예전엔 같은 명단이 파이썬 2곳·JS 3곳·제외 목록 5곳에 따로 박혀 있어, 입·퇴사 때마다 빠지는 곳이 생겼다
(10-26 김지훈 입사 때 PM 펼침 담당자 목록에 안 나오고, 10-28 이미영 합류 때 슬랙 A/S 목록에만 남는 식).

정본:
  - 접미사: project_config.json owner_suffix_map (+ 관리자 화면 /admin/users 에서 고친 Redis 값이 우선)
  - 비영업(경영지원·회사 이름·예전 비영업): project_config.json non_sales_owners
  - 퇴사·비활성: 사용자 DB(users) is_resigned / is_active
화면은 GET /api/staff-directory 로 받아 쓴다(utils/staffDirectory.js).
"""
from typing import Dict, List

from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)


def _config() -> dict:
    from dashboard.services.project_service import PROJECT_CONFIG
    return PROJECT_CONFIG or {}


def suffix_map() -> Dict[str, str]:
    """담당자 이름 → 코드 접미사(대문자, '-' 없음). Redis(관리자 편집) > 설정."""
    try:
        import pandas as pd
        from dashboard.services.project_service import _build_owner_suffix_map
        return {k: v for k, v in _build_owner_suffix_map(pd.DataFrame()).items() if k and v}
    except Exception as exc:
        logger.warning(f'[STAFF] 접미사 표 조회 실패 — 설정만 사용: {exc}')
        return {str(k): str(v).upper() for k, v in (_config().get('owner_suffix_map') or {}).items() if k and v}


def suffix_to_owner() -> Dict[str, str]:
    """접미사 → 담당자 이름 (같은 접미사가 둘이면 설정 순서상 앞사람)."""
    out: Dict[str, str] = {}
    for name, sfx in suffix_map().items():
        out.setdefault(sfx.upper(), name)
    return out


def company_prefix_map() -> Dict[str, str]:
    """내부 사업자 → 코드 앞 글자 (글로벌 G · 글로벌그룹 R · 플렌트 P)."""
    return {str(k): str(v) for k, v in (_config().get('company_prefix_map') or {}).items() if k and v}


def non_sales_owners() -> List[str]:
    """영업 담당자 목록·매출 통계에서 빼는 이름 (경영지원·회사 이름 등)."""
    return [str(x) for x in (_config().get('non_sales_owners') or []) if str(x).strip()]


def _users() -> list:
    try:
        from dashboard.utils.user_database import get_user_database
        return get_user_database().get_all_users() or []
    except Exception as exc:
        logger.warning(f'[STAFF] 사용자 DB 조회 실패: {exc}')
        return []


def inactive_names() -> List[str]:
    """퇴사·비활성 사용자 이름."""
    return sorted({(u.get('name') or '').strip() for u in _users()
                   if (u.get('name') or '').strip() and (u.get('is_resigned') or not u.get('is_active', True))})


def active_sales_managers() -> List[str]:
    """지금 공사를 맡길 수 있는 영업 담당자 — 접미사가 있고, 비영업 아니고, 퇴사·비활성 아닌 사람.

    사용자 DB 에 없는 접미사 주인(예: '아이티' 같은 회사 이름)은 비영업 목록으로 거른다."""
    smap = suffix_map()
    gone = set(inactive_names())
    skip = set(non_sales_owners())
    db_names = {(u.get('name') or '').strip() for u in _users()}
    names = [n for n in smap if n not in gone and n not in skip and (not db_names or n in db_names)]
    return sorted(names)


def directory() -> dict:
    """화면용 한 번에 (GET /api/staff-directory)."""
    return {
        'suffix_map': suffix_map(),
        'company_prefix_map': company_prefix_map(),
        'non_sales': non_sales_owners(),
        'inactive': inactive_names(),
        'active_sales': active_sales_managers(),
    }
