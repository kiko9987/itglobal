# -*- coding: utf-8 -*-
"""공사 취소·재개 → 연결 리드 상태 맞추기 (2026-11-16 상태 머신 점검 운영 확인).

공사를 등록하거나 PM 에서 리드 번호를 붙이면 리드가 '공사 확정'이 되지만, 공사를 취소해도 리드는
'공사 확정'으로 남아 리드 화면·통계에 계약으로 잡혔다(11-16 시트 확인 7건).

- 취소: 리드가 '공사 확정'이고, 같은 리드에 붙은 다른 공사가 모두 취소됐으면 → '공사 취소'
- 재개: 리드가 '공사 취소'면 → '공사 확정'
- 그 밖의 리드 상태(방문 완료·문의 드랍 등)는 사람이 정한 것이라 건드리지 않는다.

취소·재개 공용 함수(`apply_project_cancel_to_slack` / `apply_project_uncancel_to_slack`)에서 부르므로
슬랙 버튼·PM 화면·시트 직접 정리(리컨사일러) 세 경로가 모두 탄다. 시트 읽기·쓰기가 있어 백그라운드 스레드로 돈다.
"""
import logging
import os
import re
import threading
from typing import List, Optional

logger = logging.getLogger(__name__)

STATUS_CONFIRMED = '공사 확정'
STATUS_CANCELLED = '공사 취소'
_LEAD_TOKEN = re.compile(r'\b((?:L|ETC)-[0-9A-Za-z-]+)\b')


def lead_nos_of(value) -> List[str]:
    """공사현황 'Lead No' 칸 값 → 리드 번호 목록 (빈칸·'-'·nan 은 없음)."""
    s = str(value if value is not None else '').strip()
    if not s or s in ('-', 'nan', 'None'):
        return []
    seen = []
    for m in _LEAD_TOKEN.finditer(s):
        if m.group(1) not in seen:
            seen.append(m.group(1))
    return seen


def _projects_df():
    from dashboard.services.project_service import load_data
    return load_data()


def _lead_nos_for(code: str, project: Optional[dict]) -> List[str]:
    nos = lead_nos_of((project or {}).get('Lead No'))
    if nos:
        return nos
    try:
        df = _projects_df()
        if df is not None and not df.empty and 'Lead No' in df.columns:
            hit = df[df['프로젝트 코드'].astype(str).str.strip() == code]
            if not hit.empty:
                return lead_nos_of(hit.iloc[0].get('Lead No'))
    except Exception as exc:
        logger.warning(f'[LEAD_LINK] {code} 리드 번호 조회 실패: {exc}')
    return []


def _other_live_projects(lead_no: str, code: str) -> Optional[List[str]]:
    """같은 리드에 붙은 다른 공사 중 취소 안 된 것. 공사 목록을 못 읽으면 None(=판단 못 함)."""
    from dashboard.services.project_state import spec_cancelled
    try:
        df = _projects_df()
    except Exception as exc:
        logger.warning(f'[LEAD_LINK] 공사 목록 읽기 실패 ({lead_no}): {exc}')
        return None
    if df is None or df.empty or 'Lead No' not in df.columns:
        return None
    live = []
    for _, r in df.iterrows():
        c = str(r.get('프로젝트 코드') or '').strip()
        if not c or c == code or lead_no not in lead_nos_of(r.get('Lead No')):
            continue
        if not spec_cancelled(r.get('수금 관련 특이사항')):
            live.append(c)
    return live


def _lead_status(lead_no: str) -> Optional[str]:
    from dashboard.services.lead_service import get_lead_by_no
    lead = get_lead_by_no(lead_no)
    if not lead:
        return None
    return str(lead.get('상태') or '').strip()


def _set_status(lead_no: str, status: str) -> bool:
    from dashboard.services.lead_service import update_lead_status
    res = update_lead_status(lead_no, status)
    return not (isinstance(res, dict) and res.get('success') is False)


def sync_lead_status(code: str, project: Optional[dict], cancelled: bool) -> List[str]:
    """공사 취소(cancelled=True)·재개(False) 뒤 연결 리드 상태를 맞춘다. 바꾼 리드 번호 목록을 돌려준다."""
    code = (code or '').strip()
    changed = []
    for lead_no in _lead_nos_for(code, project):
        try:
            cur = _lead_status(lead_no)
            if cur is None:
                logger.info(f'[LEAD_LINK] {code} → {lead_no} 리드를 못 찾아 그대로 둠')
                continue
            if cancelled:
                if cur != STATUS_CONFIRMED:
                    continue
                others = _other_live_projects(lead_no, code)
                if others is None:
                    logger.warning(f'[LEAD_LINK] {code} → {lead_no} 다른 공사 확인 못 해 그대로 둠')
                    continue
                if others:
                    logger.info(f'[LEAD_LINK] {code} 취소 — {lead_no} 에 진행 중 공사 {others} 있어 공사 확정 유지')
                    continue
                new = STATUS_CANCELLED
            else:
                if cur != STATUS_CANCELLED:
                    continue
                new = STATUS_CONFIRMED
            if _set_status(lead_no, new):
                changed.append(lead_no)
                logger.info(f"[LEAD_LINK] {code} {'취소' if cancelled else '재개'} → {lead_no} '{cur}' → '{new}'")
            else:
                logger.error(f"[LEAD_LINK] {code} → {lead_no} '{new}' 저장 실패 (리드는 '{cur}' 그대로)")
        except Exception as exc:
            logger.error(f'[LEAD_LINK] {code} → {lead_no} 상태 맞추기 실패: {exc}')
    return changed


def sync_lead_status_async(code: str, project: Optional[dict], cancelled: bool) -> None:
    # 다른 테스트가 취소·재개 함수를 부를 때 진짜 시트를 읽고 쓰지 않게 (이 모듈 테스트는 sync_lead_status 를 직접 부름)
    if os.environ.get('PYTEST_CURRENT_TEST') and not os.environ.get('LEAD_LINK_TEST'):
        return
    if os.getenv('LEAD_LINK_ENABLED', '1').strip().lower() in ('0', 'false', 'off'):
        return
    threading.Thread(target=sync_lead_status, args=(code, dict(project or {}), cancelled),
                     daemon=True, name=f'lead-link-{code}').start()
