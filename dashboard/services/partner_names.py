# -*- coding: utf-8 -*-
"""거래처명(공사 현황 '거래처명' 열) 규칙 — PM·슬랙 공용 (2026-10-22 사용자 확정, 10-23 보완).

- 유입 구분: 이 공사가 들어온 경로. 공사별로 그대로 둔다(광고 성과 분석 기준). 진짜 거래처(인테리어 등)
  여부는 이 칸('거래처')으로 본다.
- 거래처명: 이 공사를 맡긴 곳의 부르는 이름(법인 표기 뺀). 거래처 유입이면 일을 준 업체, 그 밖은 고객 자신
  (짧은 사업자명, 10-23 사용자 요청 — 시트에서 '온라인'만 보여 누군지 몰랐음). 사업자명이 없을 때만 유입
  묶음(온라인·전화·소개·중고판매·기타). 같은 곳 공사가 있으면 그곳 이름을 이어 쓰고, 앞 공사의 묶음도 소급.
- 사업자명: 실제 사업자등록증 상호 — 계산서·등록증·입금 매칭의 기준. 거래처명으로 바꾸지 말 것.

'같은 곳' = 짧은 사업자명이 같은 공사. 온라인 고객은 등록 때 사업자명이 비어 있다가 등록증 OCR·편집으로
나중에 채워지므로, 등록 때와 사업자명이 바뀔 때 두 번 정한다(apply_place_on_biz_change).
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
# 부르는 이름에서 빼는 것 (10-23 넓힘 — '(재)성곡미술문화재단', '카페 숨(suum)'이 그대로 남았음)
_LEGAL_MARK_RE = re.compile(
    r'\[\s*폐업\s*\]|주식회사|유한책임회사|유한회사|합자회사|합명회사|재단법인|사단법인|의료법인|학교법인'
    r'|사회복지법인|농업회사법인|\(\s*(?:주|유|재|사|의|학|복|합)\s*\)|㈜|㈔|㈐')
_PAREN_NO_HANGUL_RE = re.compile(r'\(([^()가-힣]*)\)')        # 한글 없는 괄호: (PK) (suum) (AB STUDI0) (完)
_OPEN_PAREN_TAIL_RE = re.compile(r'\([^()가-힣]*$')           # 닫히지 않은 꼬리: '(JH ENG Co.,Ltd.'
# 한글 이름 뒤 영문·한자 꼬리: '완디자인 完 Design Interior' → '완디자인'. 꼬리에 영문 2자 이상이나 한자가
#   있어야 하고 로마 숫자('골드로즈 II')·숫자만은 남김
_FOREIGN_TAIL_RE = re.compile(r'^(.*[가-힣].*?)\s+([^가-힣]+)$')
_ROMAN_RE = re.compile(r'^[IVX]+$')
# 앞이 업종 이름뿐이면 영문 꼬리가 실제 상호 — '건축사사무소 VINZIP', '세무회계컨설팅 SONG&PARTNERS'
_GENERIC_HEAD_RE = re.compile(r'(사무소|사무실|컨설팅|법무법인|회계법인|특허법인|세무법인|건축사)$')
# 건물 관리 단체 꼬리 — '관악산 휴먼시아아파트 입주자대표회의' → '관악산 휴먼시아아파트' (10-23 사용자 요청)
_ORG_SUFFIX_RE = re.compile(r'\s*(입주자\s*대표\s*회의?|관리\s*사무소|관리단|관리\s*위원회|운영\s*위원회)$')


def inflow_group(inflow: str) -> str:
    """유입 구분 → 거래처가 아닐 때 거래처명에 넣는 묶음 (온라인 계열·빈칸은 '온라인')."""
    s = str(inflow or '').strip()
    if s in GROUP_LABELS:
        return s
    return '온라인' if (s in _ONLINE_INFLOWS or not s) else '기타'


def short_name(name: str) -> str:
    """부르는 이름 — '주식회사 메이드바이디자인' → '메이드바이디자인'.

    빼는 것: 법인 표기((주)·(재)·재단법인 등·[폐업]), 한글 없는 괄호('(PK)', '(suum)'), 닫히지 않은 영문 괄호
    꼬리, 한글 이름 뒤 영문·한자 꼬리('完 Design Interior'), 건물 관리 단체 꼬리('입주자대표회의'·'관리사무소'·
    '관리단'·'운영위원회'). 남기는 것: 지점('신도림점')·한글 괄호('(송파점)')
    — 다른 지점이 한 곳으로 묶이지 않게. 다 빼서 비면 법인 표기만 뺀 이름. JS shortPartnerName 과 같게.
    """
    base = re.sub(r'\s+', ' ', _LEGAL_MARK_RE.sub(' ', str(name or ''))).strip()
    s = _OPEN_PAREN_TAIL_RE.sub('', _PAREN_NO_HANGUL_RE.sub('', base))   # '피케이(PK)대치' → '피케이대치'
    s = re.sub(r'\s+', ' ', s).strip()
    m = _FOREIGN_TAIL_RE.match(s)
    if m:
        tail = m.group(2).strip()
        if not _ROMAN_RE.match(tail) and not _GENERIC_HEAD_RE.search(m.group(1).strip()) \
                and (len(re.findall(r'[A-Za-z]', tail)) >= 2 or re.search(r'[一-鿿]', tail)):
            s = m.group(1)
    t = _ORG_SUFFIX_RE.sub('', s).strip()
    if len(re.findall(r'[가-힣]', t)) >= 2:      # '관리단'만 있으면 그대로
        s = t
    s = s.strip(' .,-_/·')
    return s if re.search(r'[가-힣A-Za-z0-9]', s) else base


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


def card_partner(value) -> str:
    """슬랙 카드 '🤝 거래처명' 줄에 쓸 값 — 이름이면 그대로, '-'(이름 미입력)는 '-', 분류값은 ''(줄 생략).

    '거래처명 : 전화'가 '전화'라는 업체로 읽혀서 분류값은 안 보임 — 유입 구분 줄에 이미 나옴(2026-10-28).
    """
    s = str(value or '').strip()
    if is_partner_name(s):
        return s
    return '-' if s == '-' else ''


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

    새 공사: ① 입력한 거래처명이 이름이면 그것 ② 같은 곳 공사가 있으면 그곳 이름 ③ 사업자명이 있으면
    짧은 사업자명 ④ 없으면 거래처 유입은 '-'(확인 필요), 그 밖은 유입 묶음.
    소급: 같은 곳 공사가 있으면 그중 거래처명이 유입 묶음인 것 → 그곳 이름. Returns (거래처명, 소급 코드, 그곳 이름).
    """
    prev = same_place(biz, list(records or []), exclude_code=code)
    pname = place_name(biz, prev) if prev else ''
    if is_partner_name(given):
        name = str(given).strip()
    elif prev:
        name = pname
    elif is_partner_name(biz):
        name = short_name(biz)
    elif str(inflow or '').strip() == '거래처':
        name = '-'
    else:
        name = inflow_group(inflow)
    return name, (_labelled(prev) if prev else []), pname


def quiet_hours(now=None) -> bool:
    """업무시간(평일 08~17시) 밖이면 True — 카드 댓글(슬랙 알림) 없이 조용히 기록 (10-23 새벽 알림 계기)."""
    from datetime import datetime
    now = now or datetime.now()
    return now.weekday() >= 5 or not (8 <= now.hour < 17)


def promote_previous(codes: List[str], name: str, why: str = '', editor: str = '',
                     replace: Optional[dict] = None, quiet: Optional[bool] = None) -> int:
    """공사들의 거래처명(유입 묶음) → 거래처 이름. 공사확정 카드에 '시스템 자동기록' 댓글.

    PM 편집과 같은 길: 시트 행을 새로 읽어 지금 값이 유입 묶음일 때만, 쓰기 큐로 거래처명 + _version(+1)
    두 칸만 쓴다(버전을 올려야 오래 열린 PM 화면 저장이 덮어쓰지 못함). 캐시·감사 로그·카드 댓글.
    replace={코드: 옛 이름} — 그 공사는 지금 값이 그 옛 이름(같은 곳)이어도 바꿈(사업자명을 고쳐 자동 이름이
    따라가는 경우). quiet(기본=업무시간 밖) — 댓글 없이 PM 편집 마커만(리컨사일러가 기준값만 맞춤).
    실패한 공사는 건너뜀. Returns 바꾼 수.
    """
    if not codes or not is_partner_name(name):
        return 0
    quiet = quiet_hours() if quiet is None else quiet
    replace = replace or {}
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
            replaceable = code in replace and is_partner_name(old) and place_key(old) == place_key(replace[code])
            if (is_partner_name(old) and not replaceable) or old == '-' or place_key(old) == place_key(name):
                continue                  # 그 사이 이름이 생겼거나 '확인 필요' 표시·이미 같은 이름 → 그대로
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
            if quiet:   # 리컨사일러가 댓글 없이 기준값만 맞춤 (project_reconcile PM 편집 마커)
                get_redis_client().redis.setex(f'project_pm_edit:{code}', 900, '1')
            else:
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
        if quiet:
            # 댓글은 없이 원본 카드만 새 이름으로 — 안 그러면 카드에 '거래처명 : 전화'가 그대로 남음
            #   (2026-10-28 R4227·R4209 등 6장). 캐시가 비워져 있을 수 있어(등록증 OCR) 이름을 직접 넣어 그림
            try:
                from dashboard.services.project_slack_notifier import refresh_project_card_license
                refresh_project_card_license(code, latest_data=dict(rec, 거래처명=name) if rec else None)
            except Exception as exc:
                logger.debug(f'[PARTNER] 카드 갱신 실패 {code}: {exc}')
            continue
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


def apply_place_on_biz_change(code: str, biz: str, editor: str = '', old_biz: str = '') -> int:
    """사업자명이 새로 채워지거나 바뀐 공사의 거래처명 맞추기 (등록 때 사업자명이 비어 있던 온라인 고객이
    등록증으로 채워지는 경우 등). 이름 = 같은 곳 공사가 있으면 그곳 이름, 없으면 짧은 사업자명.
    - 이 공사: 거래처명이 유입 묶음이면 바꿈. 거래처 유입이 아니고 거래처명이 옛 사업자명에서 나온 이름이면
      (old_biz) 새 이름으로 따라감 — 거래처 유입의 거래처명(일 준 업체)은 사업자명이 바뀌어도 그대로.
      이 공사 쪽은 사업자명 변경 알림이 따로 나가므로 조용히(댓글 없이).
    - 같은 곳 다른 공사: 유입 묶음으로 남은 것을 그곳 이름으로 소급.
    Returns 바꾼 수."""
    if not code or not is_partner_name(biz):
        return 0
    # 나눠 내는 공사의 사업자명('최영희 / 기종원')은 부담자 이름 모음이라 '곳' 이름이 아님 (2026-10-28)
    from dashboard.services.payer_split import is_split, is_split_biz_value
    if is_split_biz_value(biz):
        return 0
    from dashboard.services.project_service import get_project_records
    records = get_project_records() or []
    if is_split(next((r for r in records if _code(r) == code), None) or {}):
        return 0
    prev = same_place(biz, records, exclude_code=code)
    name = place_name(biz, prev) if prev else short_name(biz)
    me = next((r for r in records if _code(r) == code), None) or {'프로젝트 코드': code, '거래처명': ''}
    cur = str(me.get('거래처명') or '').strip()
    replace = {}
    if str(me.get('유입 구분') or '').strip() != '거래처' and is_partner_name(old_biz) \
            and is_partner_name(cur) and place_key(cur) == place_key(old_biz):
        replace[code] = cur
    done = 0
    if _labelled([me]) or replace:
        done += promote_previous([code], name, why=f'{code} 사업자명 확인', editor=editor,
                                 replace=replace, quiet=True)
    others = _labelled(prev)
    if others:
        done += promote_previous(others, name, why=f'{code} 사업자명 확인 — 같은 곳 공사 묶음', editor=editor)
    return done


def known_partner_names(records: Iterable[dict], limit: Optional[int] = None) -> List[str]:
    """추천 목록 — 실제 거래처 이름을 공사 수 많은 순으로 (분류값 제외)."""
    c = Counter(str(r.get('거래처명') or '').strip() for r in records or [] if is_partner_name(r.get('거래처명')))
    names = [n for n, _ in c.most_common()]
    return names[:limit] if limit else names
