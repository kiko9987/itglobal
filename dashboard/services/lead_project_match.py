"""새 프로젝트 수동 입력 시 '이미 방문한 리드' 대조 (2026-11-06).

방문한 현장인데 '수동 입력하기'로 리드 번호 없이 등록되는 일을 막기 위해, 등록 창에서
연락처·주소를 넣는 순간 최근 방문 리드를 찾아 알려 준다 (프론트 ModernProjectModal).

대상 리드 — '방문 현장 불러오기'(leads.api_search_leads_for_project)로 고를 수 있는 것과 같은 범위:
  - L- 번호, 아직 프로젝트에 연결 안 됨, 상태 '공사 확정' 아님
  - 실제 방문한 것: 방문 완료 · 견적 제출 · 방문 예약(예정일이 오늘 이전)
  - 최근 180일 (방문 예정일, 없으면 상담 시간)

맞추는 기준:
  - 주소: 도로명/동 + 번지가 같으면 (lead_sync._addr_core — 층·호·상호 차이 무시)
  - 연락처: 끝 8자리. 단 거래처 건(입력한 유입 구분 또는 리드 플랫폼이 거래처)은 안 맞춤 —
    거래처 담당자 번호 하나에 방문이 수십 건씩 붙어 있어 엉뚱한 리드가 걸림
"""
import re
from datetime import date, timedelta
from typing import Iterable, List, Optional, Set

from .lead_sync import _addr_core

RECENT_DAYS = 180
MAX_RESULTS = 3
VISITED_STATUSES = ('방문 완료', '견적 제출', '방문 예약')
PARTNER = '거래처'

_DATE_RE = re.compile(r'(\d{4})[-./](\d{1,2})[-./](\d{1,2})')


def _parse_date(v) -> Optional[date]:
    m = _DATE_RE.search(str(v or ''))
    if not m:
        return None
    try:
        return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    except ValueError:
        return None


def _phone_key(v) -> str:
    d = re.sub(r'\D', '', str(v or ''))
    return d[-8:] if len(d) >= 9 else ''


def match_visited_leads(leads: Iterable[dict], registered: Set[str], phone: str = '',
                        address: str = '', inflow: str = '', today: Optional[date] = None,
                        limit: int = MAX_RESULTS) -> List[dict]:
    """입력한 연락처·주소와 맞는 최근 방문 리드 (최근 방문일 순, 최대 limit 건).

    반환 항목: 리드 원본 dict + '_match'('연락처'·'주소'·'연락처·주소') + '_visit_date'(YYYY-MM-DD).
    """
    today = today or date.today()
    since = today - timedelta(days=RECENT_DAYS)
    want_phone = '' if (inflow or '').strip() == PARTNER else _phone_key(phone)
    want_addr = _addr_core(address or '')
    if not want_phone and not want_addr:
        return []

    out = []
    for lead in leads:
        lead_no = str(lead.get('리드 No') or '').strip()
        if not lead_no.startswith('L-') or lead_no in registered:
            continue
        status = str(lead.get('상태') or '').strip()
        if status not in VISITED_STATUSES:
            continue
        visit = _parse_date(lead.get('방문 예정일'))
        if status == '방문 예약' and not (visit and visit <= today):
            continue
        when = visit or _parse_date(lead.get('상담 시간'))
        if not when or when < since:
            continue

        how = []
        if want_phone and str(lead.get('플랫폼') or '').strip() != PARTNER \
                and _phone_key(lead.get('고객 연락처')) == want_phone:
            how.append('연락처')
        if want_addr and _addr_core(str(lead.get('방문 주소') or '')) == want_addr:
            how.append('주소')
        if not how:
            continue
        out.append({**lead, '_match': '·'.join(how), '_visit_date': when.isoformat()})

    # 연락처·주소 둘 다 맞는 것 먼저, 그다음 최근 방문 순
    out.sort(key=lambda x: (len(x['_match'].split('·')), x['_visit_date']), reverse=True)
    return out[:limit]
