# -*- coding: utf-8 -*-
"""거래처명(공사 현황 '거래처명' 열) 규칙 — PM·슬랙 공용 (2026-10-22 사용자 확정).

- 유입 구분: 이 공사가 들어온 경로. 공사별로 그대로 둔다(광고 성과 분석 기준).
- 거래처명: 이 공사를 맡긴 상대(곳). 거래처(인테리어 등)거나 두 번 이상 공사한 곳이면 그 이름(법인 표기를
  뺀 부르는 이름), 아니면 유입 묶음(온라인·전화·소개·중고판매·기타). 같은 곳의 두 번째 공사가 생기면
  첫 공사의 유입 묶음도 그 이름으로 바꾼다(유입 구분은 그대로).
- 사업자명: 실제 사업자등록증 상호 — 계산서·등록증·입금 매칭의 기준. 거래처명으로 바꾸지 말 것.

'같은 곳' = 짧은 사업자명이 같은 공사. 온라인 고객은 등록 때 사업자명이 비어 있다가 등록증 OCR·편집으로
나중에 채워지므로, 등록 때와 사업자명이 바뀔 때 두 번 묶는다(apply_place_on_biz_change).
"""
from __future__ import annotations

import logging
import os
import re
from collections import Counter
from typing import Iterable, List, Optional, Tuple

logger = logging.getLogger(__name__)

# 거래처명 칸에 들어가는 '이름이 아닌' 분류값 — 이름처럼 묶거나 비교하면 안 됨
GROUP_LABELS = ('온라인', '전화', '소개', '중고판매', '기타')
_ONLINE_INFLOWS = {'온라인', '당근', '홈페이지', '숨고', '큐플레이스', '카카오톡', '채널톡', '메일'}
_NOT_NAMES = set(GROUP_LABELS) | _ONLINE_INFLOWS | {'거래처', '-', '', 'nan'}   # nan = 판다스 빈 칸
_LEGAL_MARK_RE = re.compile(r'\[\s*폐업\s*\]|주식회사|\(주\)|㈜|유한회사|\(유\)')


def inflow_group(inflow: str) -> str:
    """유입 구분 → 거래처가 아닐 때 거래처명에 넣는 묶음 (온라인 계열·빈칸은 '온라인')."""
    s = str(inflow or '').strip()
    if s in GROUP_LABELS:
        return s
    return '온라인' if (s in _ONLINE_INFLOWS or not s) else '기타'


def short_name(name: str) -> str:
    """부르는 이름 — '주식회사 메이드바이디자인' → '메이드바이디자인' (법인 표기·[폐업] 빼고 공백 정리)."""
    return re.sub(r'\s+', ' ', _LEGAL_MARK_RE.sub(' ', str(name or ''))).strip()


def _norm(name: str) -> str:
    try:
        from dashboard.services.business_license_handler import _norm_biz
        return _norm_biz(name)
    except Exception:
        return re.sub(r'\s+', '', str(name or '')).lower()


def is_partner_name(value) -> bool:
    """실제 거래처 이름인가 (분류값·'-'·자리표시가 아니면 True)."""
    s = short_name(value)
    return bool(s) and s not in _NOT_NAMES and bool(_norm(s))


def place_key(name) -> str:
    """같은 곳 판단 키 — 짧은 이름을 정규화(공백·대소문자 무시). 이름이 아니면 ''."""
    return _norm(short_name(name)) if is_partner_name(name) else ''


def display_name(rec: dict) -> str:
    """화면에서 공사를 부르는 이름 — 거래처명이 이름이면 그것, 아니면 사업자명(없으면 '-')."""
    p = str((rec or {}).get('거래처명') or '').strip()
    if is_partner_name(p):
        return p
    b = str((rec or {}).get('사업자명') or '').strip()
    return b if b and b != '-' else '-'


def _code(r: dict) -> str:
    return str(r.get('프로젝트 코드') or '').strip()


def same_place(biz: str, records: Iterable[dict], exclude_code: str = '') -> List[dict]:
    """사업자명이 같은 곳(짧은 이름 기준)의 기존 공사들."""
    k = place_key(biz)
    if not k:
        return []
    return [r for r in records or [] if _code(r) != exclude_code and place_key(r.get('사업자명')) == k]


def place_name(biz: str, prev: List[dict]) -> str:
    """같은 곳 공사들이 쓰는 그곳 이름.

    ① 같은 곳 공사의 거래처명 중 사업자명과 같은 곳을 가리키는 이름(띄어쓰기만 다른 것 포함)
    ② 직접 맡긴 공사(유입이 거래처가 아닌 것)에 붙은 이름 — 사람이 부르는 이름으로 고쳐 둔 경우
    ③ 짧은 사업자명. 거래처 유입 공사의 거래처명은 인테리어일 수 있어(고객이 돈 낸 경우) ②에서 뺀다.
    """
    k = place_key(biz)
    names = Counter(str(r.get('거래처명') or '').strip() for r in prev if is_partner_name(r.get('거래처명')))
    for n, _ in names.most_common():
        if place_key(n) == k:
            return n
    direct = Counter(str(r.get('거래처명') or '').strip() for r in prev
                     if str(r.get('유입 구분') or '').strip() != '거래처' and is_partner_name(r.get('거래처명')))
    if direct:
        return direct.most_common(1)[0][0]
    return short_name(biz)


def _labelled(recs: Iterable[dict]) -> List[str]:
    """거래처명이 아직 유입 묶음인 공사 코드('-' = 확인 필요는 사람이 볼 몫이라 건드리지 않음)."""
    return [_code(r) for r in recs
            if not is_partner_name(r.get('거래처명')) and str(r.get('거래처명') or '').strip() != '-']


def resolve_new_partner(inflow: str, biz: str, given: str, records: Iterable[dict],
                        code: str = '') -> Tuple[str, List[str], str]:
    """새 공사의 거래처명 + 소급할 같은 곳 공사 코드 + 그곳 이름.

    새 공사: ① 입력한 거래처명이 이름이면 그것 ② 같은 곳 공사가 있으면 그곳 이름 ③ 유입 '거래처'면
    짧은 사업자명(이름이 아니면 '-' = 확인 필요) ④ 그 밖은 유입 묶음.
    소급: 같은 곳 공사가 있으면 그중 거래처명이 유입 묶음인 것 → 그곳 이름(새 공사가 인테리어 경유여도
    그곳 자체는 2건 이상이 됐으므로). Returns (거래처명, 소급 코드, 그곳 이름).
    """
    prev = same_place(biz, list(records or []), exclude_code=code)
    pname = place_name(biz, prev) if prev else ''
    if is_partner_name(given):
        name = str(given).strip()
    elif prev:
        name = pname
    elif str(inflow or '').strip() == '거래처':
        name = short_name(biz) if is_partner_name(biz) else '-'
    else:
        name = inflow_group(inflow)
    return name, (_labelled(prev) if prev else []), pname


def promote_previous(codes: List[str], name: str, why: str = '', editor: str = '') -> int:
    """공사들의 거래처명(유입 묶음) → 거래처 이름. 공사확정 카드에 '시스템 자동기록' 댓글.

    PM 편집과 같은 길: 시트 행을 새로 읽어 지금 값이 유입 묶음일 때만, 쓰기 큐로 거래처명 + _version(+1)
    두 칸만 쓴다(버전을 올려야 오래 열린 PM 화면 저장이 덮어쓰지 못함). 캐시·감사 로그·카드 댓글.
    실패한 공사는 건너뜀. Returns 바꾼 수.
    """
    if not codes or not is_partner_name(name):
        return 0
    from dashboard.services.lead_service import get_sheets_manager
    from dashboard.services.project_service import get_project_records, update_project_in_cache
    from dashboard.services.sheet_write_queue import enqueue
    sid = os.getenv('GOOGLE_SHEET_ID', '').strip()
    sn = os.getenv('GOOGLE_SHEET_NAME', '').strip()
    mgr = get_sheets_manager()
    f2l, f2i = mgr.get_field_to_letter(), mgr.get_field_to_index()
    recs = {_code(r): r for r in (get_project_records() or [])}
    why = why or '같은 곳 공사 2건 이상'
    done = 0
    for code in dict.fromkeys(c for c in codes if c):
        rec = recs.get(code) or {}
        try:
            row = mgr.find_row_by_project_code(sid, code, f"{sn}!A:A")
            if not row:
                continue
            vals = mgr.get_row_values(sid, sn, row)
            cell = lambda f: str(vals[f2i[f]]).strip() if f2i[f] < len(vals) else ''
            if cell('프로젝트 코드') != code:
                continue
            old = cell('거래처명')
            if is_partner_name(old) or old == '-':
                continue                  # 그 사이 이름이 생겼거나 '확인 필요' 표시 → 그대로
            # 버전: 시트 값과 캐시 값(큐에 아직 안 쓰인 PM 저장) 중 큰 쪽 + 1
            vers = []
            for v in (cell('_version'), rec.get('_version')):
                try:
                    vers.append(int(float(str(v or '0').strip() or '0')))
                except ValueError:
                    pass
            ver = str(max(vers or [0]) + 1)
            change = {'field_name': '거래처명', 'old_value': old, 'new_value': name}
            enqueue('project_update_sheet', {
                'sheet_id': sid, 'sheet_name': sn, 'row_number': row,
                'cells': [{'col': f2l['거래처명'], 'value': name}, {'col': f2l['_version'], 'value': ver}],
                'field_changes': [change], 'project_code': code,
            }, meta={'user_email': editor or 'system'})
            update_project_in_cache(code, {'거래처명': name, '_version': ver})
            done += 1
        except Exception as exc:
            logger.warning(f'[PARTNER] 거래처명 소급 예외 {code}: {exc}')
            continue
        try:
            from dashboard.utils.redis_client import get_redis_client
            get_redis_client().redis.setex(f'project_sys_edit:{code}', 900, '시스템 자동기록')
        except Exception:
            pass
        try:
            from dashboard.utils.user_database import get_audit_repository
            get_audit_repository().log_action(
                editor or 'system', 'project_update', details=f'거래처명 자동 소급: {why}',
                project_code=code, field_name='거래처명', old_value=old, new_value=name)
        except Exception:
            pass
        try:
            from dashboard.services.project_slack_notifier import notify_project_field_changes
            notify_project_field_changes(
                code, [{'field_name': '거래처명', 'old_value': old, 'new_value': name}],
                latest_data=dict(rec, 거래처명=name), editor=f'시스템 자동기록 · {why}')
        except Exception:
            pass
    if done:
        logger.info(f'[PARTNER] 거래처명 소급 {done}건 → {name} ({why})')
    return done


def apply_place_on_biz_change(code: str, biz: str, editor: str = '') -> int:
    """사업자명이 새로 채워지거나 바뀐 공사 — 같은 곳 공사가 있으면 이 공사와 그 공사들의 유입 묶음을
    그곳 이름으로 묶는다(등록 때 사업자명이 비어 있던 온라인 고객이 등록증으로 채워지는 경우 등).
    Returns 바꾼 수."""
    if not code or not is_partner_name(biz):
        return 0
    from dashboard.services.project_service import get_project_records
    records = get_project_records() or []
    prev = same_place(biz, records, exclude_code=code)
    if not prev:
        return 0
    me = [r for r in records if _code(r) == code] or [{'프로젝트 코드': code, '거래처명': ''}]
    targets = _labelled(me + prev)
    if not targets:
        return 0
    return promote_previous(targets, place_name(biz, prev), why=f'{code} 사업자명 확인 — 같은 곳 공사 묶음',
                            editor=editor)


def known_partner_names(records: Iterable[dict], limit: Optional[int] = None) -> List[str]:
    """추천 목록 — 실제 거래처 이름을 공사 수 많은 순으로 (분류값 제외)."""
    c = Counter(str(r.get('거래처명') or '').strip() for r in records or [] if is_partner_name(r.get('거래처명')))
    names = [n for n, _ in c.most_common()]
    return names[:limit] if limit else names
