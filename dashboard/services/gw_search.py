"""그룹웨어 빠른 찾기 — 프로젝트·리드를 한 번에 (당근 비즈니스의 `/` 메뉴 찾기처럼, 2026-11-10).

- 프로젝트: 공사현황 캐시(load_data — 캐시가 있으면 시트를 다시 읽지 않음)
- 리드: 리드 시트 캐시(get_lead_records), 플랫폼 '기타'(ETC- 사후관리 등)는 리드 화면처럼 뺌
- 찾는 말: 코드·리드 번호·거래처/고객 이름·현장 주소·발주처 담당자·전화번호(숫자 4자리 이상이면 숫자만 비교)
- 검색어는 로그에 남기지 않는다(전화번호일 수 있음).
- 결과를 누르면 기존 화면으로 간다: /projects?search=코드&fresh=1 · /leads?search=리드번호
"""
from __future__ import annotations

import re
from urllib.parse import quote

from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)

LIMIT = 6
_MAX_Q = 50

# (칸 이름, 점수) — 이름·주소가 맞으면 담당자 이름만 맞은 것보다 위로
_PROJECT_TEXT = (('거래처명', 50), ('사업자명', 50), ('현장 주소', 40), ('발주처 담당자', 40), ('담당자', 10))
_PROJECT_PHONE = ('발주처 연락처',)
_LEAD_TEXT = (('고객명', 50), ('방문 주소', 40), ('영업 담당자', 10), ('온라인 상담자', 10))
_LEAD_PHONE = ('고객 연락처',)


def _s(v) -> str:
    if v is None:
        return ''
    try:
        import pandas as pd
        if pd.isna(v):
            return ''
    except (TypeError, ValueError):
        pass
    if hasattr(v, 'strftime'):
        try:
            return v.strftime('%Y-%m-%d')
        except Exception:
            return ''
    return str(v).strip()


def _compact(s: str) -> str:
    return re.sub(r'\s+', '', s).lower()


def _digits(s: str) -> str:
    return re.sub(r'\D', '', s)


def _num(code: str) -> int:
    m = re.search(r'\d+', code or '')
    return int(m.group()) if m else 0


class _Query:
    def __init__(self, q: str):
        q = (q or '').strip()[:_MAX_Q]
        self.raw = q
        self.low = q.lower()
        self.compact = _compact(q)
        d = _digits(q)
        # 숫자·하이픈·띄어쓰기·괄호만으로 된 4자리 이상 → 전화번호로도 비교.
        #   글자가 섞이면(G4266·L-0451 같은 코드) 전화번호로 보지 않음 — 코드 숫자가 남의 전화번호에 걸리던 것(11-10 운영 확인)
        self.digits = d if len(d) >= 4 and re.fullmatch(r'[\d\s\-+().]+', q) else ''

    @property
    def usable(self) -> bool:
        return len(self.compact) >= 2 or bool(self.digits)

    def code_score(self, code: str) -> int:
        c = code.lower()
        if not c or not self.compact:
            return 0
        if c == self.compact:
            return 100
        if c.startswith(self.compact):
            return 80
        if self.compact in c:
            return 60
        return 0

    def text_score(self, value: str, weight: int) -> int:
        if not value:
            return 0
        if self.low in value.lower() or (self.compact and self.compact in _compact(value)):
            return weight
        return 0

    def phone_score(self, value: str) -> int:
        return 45 if self.digits and value and self.digits in _digits(value) else 0


def _rank(rows, limit):
    rows.sort(key=lambda r: (-r[0], -r[1]))
    return [r[2] for r in rows[:limit]], len(rows)


def search_projects(query: _Query, limit: int = LIMIT) -> tuple[list[dict], int]:
    from dashboard.services.project_service import load_data
    df = load_data()
    if df is None or len(df) == 0:
        return [], 0
    cols = [c for c in {'프로젝트 코드', '공사 시작', '수금 관련 특이사항',
                        *(c for c, _ in _PROJECT_TEXT), *_PROJECT_PHONE} if c in df.columns]
    hits = []
    for rec in df[cols].to_dict('records'):
        code = _s(rec.get('프로젝트 코드'))
        if not code:
            continue
        score = query.code_score(code)
        for col, w in _PROJECT_TEXT:
            score = max(score, query.text_score(_s(rec.get(col)), w))
        for col in _PROJECT_PHONE:
            score = max(score, query.phone_score(_s(rec.get(col))))
        if not score:
            continue
        title = _s(rec.get('거래처명')) or _s(rec.get('사업자명')) or _s(rec.get('발주처 담당자'))
        meta = ' · '.join(x for x in (_s(rec.get('담당자')), _s(rec.get('공사 시작'))[:10]) if x)
        hits.append((score, _num(code), {
            'kind': 'project', 'code': code, 'title': title, 'sub': _s(rec.get('현장 주소')), 'meta': meta,
            'cancelled': bool(re.search(r'공사\s*취소', _s(rec.get('수금 관련 특이사항')))),
            'href': f'/projects?search={quote(code)}&fresh=1',
        }))
    return _rank(hits, limit)


def search_leads(query: _Query, limit: int = LIMIT) -> tuple[list[dict], int]:
    from dashboard.services.lead_service import get_lead_records
    hits = []
    for rec in get_lead_records() or []:
        if _s(rec.get('플랫폼')) == '기타':
            continue
        no = _s(rec.get('리드 No'))
        if not no:
            continue
        score = query.code_score(no)
        for col, w in _LEAD_TEXT:
            score = max(score, query.text_score(_s(rec.get(col)), w))
        for col in _LEAD_PHONE:
            score = max(score, query.phone_score(_s(rec.get(col))))
        if not score:
            continue
        meta = ' · '.join(x for x in (_s(rec.get('플랫폼')), _s(rec.get('상태')), _s(rec.get('상담 시간'))[:10]) if x)
        hits.append((score, _num(no), {
            'kind': 'lead', 'code': no, 'title': _s(rec.get('고객명')), 'sub': _s(rec.get('방문 주소')), 'meta': meta,
            'href': f'/leads?search={quote(no)}',
        }))
    return _rank(hits, limit)


def search(q: str, limit: int = LIMIT) -> dict:
    query = _Query(q)
    out = {'q': query.raw, 'projects': [], 'leads': [], 'project_total': 0, 'lead_total': 0}
    if not query.usable:
        return out
    for key, fn in (('projects', search_projects), ('leads', search_leads)):
        try:
            items, total = fn(query, limit)
            out[key] = items
            out[key[:-1] + '_total'] = total
        except Exception as e:  # 한쪽이 실패해도 다른 쪽은 보여 줌
            logger.warning(f'[GW검색] {key} 찾기 실패: {type(e).__name__}')
    return out
