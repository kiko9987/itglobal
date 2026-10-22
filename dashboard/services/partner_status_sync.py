"""거래처 탭 사업자등록 상태 갱신 — 국세청 상태조회 → J(상태)/K(최종확인일) 기록.

거래처 탭이 '마지막 세금계산서 발행 시점 스냅샷'이라 상대방 폐업/신설을 못 따라가는
문제의 안전망. 폐업 번호로 세금계산서 발행하는 사고 방지 (2026-07-28 도입).

- 소스: 메인 시트(GOOGLE_SHEET_ID) '거래처' 탭 A열 등록번호
- 조회: dashboard.services.nts_status.check_business_status (국세청 odcloud)
- 기록: J열 = 상태(계속사업자/휴업자/폐업자/조회안됨, 폐업은 폐업일 병기), K열 = 최종확인일
"""
from __future__ import annotations

import logging
import os
import re
from datetime import date
from typing import List, Optional, Tuple

from dashboard.services.nts_status import check_business_status, normalize_bno

logger = logging.getLogger(__name__)

_TAB = '거래처'
_COL_STATUS = 'J'
_COL_CHECKED = 'K'
_WRITE_CHUNK = 500  # batchUpdate 1회 range 수


def _norm_name(s: str) -> str:
    """상호명 정규화 — 공백 제거 + 전각괄호 반각화 + 소문자. exact match 용.

    (주)/주식회사 등 법인표기는 **제거 안 함** — 별개 사업자 오매칭 방지 (정밀 우선).
    """
    s = str(s or '').replace('（', '(').replace('）', ')')
    return re.sub(r'\s+', '', s).strip().lower()


def _is_flagged_status(s: str) -> bool:
    s = str(s or '')
    return s.startswith('폐업') or s == '휴업자'


def lookup_partner_status_by_name(biz_name: str) -> Optional[dict]:
    """상호명으로 거래처 탭 상태 조회 (정규화 exact match).

    계산서 요청 시 거래처 폐업/휴업 경고용. 계산서 flow 엔 사업자번호가 없어
    상호명 매칭이 유일한 링크 — 정밀(정규화 exact) 매칭으로 오경보 최소화.

    Returns None (매칭 없음/전부 정상/조회 실패) 또는:
      {'matches': [{'bno','status'}], 'flagged': [...], 'label': '폐업'|'휴업',
       'ambiguous': bool}  # ambiguous = 매칭 여러 건이고 일부만 flagged (상호 중복)
    """
    key = _norm_name(biz_name)
    if not key or key == '-':
        return None
    try:
        m, sid = _sheet()
        vals = m.service.spreadsheets().values().get(
            spreadsheetId=sid, range=f'{_TAB}!A1:J',
        ).execute().get('values', [])
    except Exception as exc:
        logger.warning(f'[거래처상태] 상호 조회 실패 ({biz_name}): {exc}')
        return None

    matches = []
    for r in vals:
        name = r[1] if len(r) > 1 else ''
        if _norm_name(name) != key:
            continue
        matches.append({
            'bno': (r[0] if len(r) > 0 else '').strip(),
            'status': (r[9] if len(r) > 9 else '').strip(),
        })
    if not matches:
        return None
    flagged = [x for x in matches if _is_flagged_status(x['status'])]
    if not flagged:
        return None
    label = '폐업' if any(x['status'].startswith('폐업') for x in flagged) else '휴업'
    return {
        'matches': matches, 'flagged': flagged, 'label': label,
        'ambiguous': len(matches) > len(flagged),
    }


def _sheet():
    from dashboard.utils.google_sheets import GoogleSheetsManager
    sid = os.getenv('GOOGLE_SHEET_ID', '').strip()
    if not sid:
        raise RuntimeError('GOOGLE_SHEET_ID 미설정')
    return GoogleSheetsManager(), sid


def _fmt_end(end: str) -> str:
    e = (end or '').strip()
    return f'{e[:4]}-{e[4:6]}-{e[6:8]}' if len(e) == 8 and e.isdigit() else e


def refresh_partner_status(dry_run: bool = True,
                            today_str: Optional[str] = None,
                            only_blank: bool = False) -> dict:
    """거래처 탭 등록번호 상태 조회 후 J/K 기록.

    dry_run=True: API 조회·집계만, 시트 미기록. (폐업 건수 확인용)
    dry_run=False: J(상태)/K(최종확인일) 실제 기록 + 헤더 세팅.
    only_blank=True: **J가 빈 행만** 대상 (신규 추가 거래처 채우기용, 데일리 증분).
        빈 행 없으면 조회·쓰기 없이 즉시 반환.

    Returns: {total_rows, queried, resolved, summary{상태:건수}, closed[(row,bno,end)], ...}
    """
    m, sid = _sheet()
    today = today_str or date.today().strftime('%Y-%m-%d')

    # 등록번호(A) + 현재 상태(J) 로드 (행 번호 보존). only_blank 시 J로 필터.
    a_vals = m.service.spreadsheets().values().get(
        spreadsheetId=sid, range=f'{_TAB}!A1:J',
    ).execute().get('values', [])
    rows: List[Tuple[int, str]] = []  # (sheet_row_1base, bno_norm)
    for i, r in enumerate(a_vals):
        nb = normalize_bno(r[0] if r else '')
        if not nb:
            continue
        if only_blank:
            cur_j = (r[9] if len(r) > 9 else '').strip()
            if cur_j:  # 이미 상태 기록됨 → 증분 대상 아님
                continue
        rows.append((i + 1, nb))

    if only_blank and not rows:
        logger.info('[거래처상태] 증분: 채울 빈 행 없음 — skip')
        return {'total_rows': 0, 'queried': 0, 'resolved': 0,
                'summary': {}, 'closed': [], 'dry_run': dry_run,
                'today': today, 'only_blank': True, 'no_blank': True}

    bnos = list({nb for _, nb in rows})
    status_map = check_business_status(bnos)

    summary = {'계속사업자': 0, '휴업자': 0, '폐업자': 0, '조회안됨': 0}
    closed: List[Tuple[int, str, str]] = []
    updates: List[Tuple[int, str, str]] = []  # (row, status_text, checked)
    for row, nb in rows:
        d = status_map.get(nb)
        stt = (d or {}).get('b_stt', '').strip() if d else ''
        if stt == '폐업자':
            end_fmt = _fmt_end((d or {}).get('end_dt', ''))
            status_text = f'폐업자 (폐업 {end_fmt})' if end_fmt else '폐업자'
            closed.append((row, nb, end_fmt))
            summary['폐업자'] += 1
        elif stt in ('계속사업자', '휴업자'):
            status_text = stt
            summary[stt] += 1
        else:
            status_text = '조회안됨'
            summary['조회안됨'] += 1
        updates.append((row, status_text, today))

    result = {
        'total_rows': len(rows), 'queried': len(bnos),
        'resolved': len(status_map), 'summary': summary,
        'closed': closed, 'dry_run': dry_run, 'today': today,
    }
    if dry_run:
        return result

    # 안전장치 — 조회할 번호가 있는데 결과 0건이면 (키 미설정/API 전체 실패)
    # 시트를 '조회안됨'으로 덮어쓰지 않고 기존 J/K 보존 (주기 갱신 사고 방지).
    if rows and not status_map:
        logger.warning('[거래처상태] 국세청 조회 0건 (키 미설정/API 실패) — 시트 쓰기 skip (기존 J/K 보존)')
        result['skipped_write'] = True
        return result

    # 실제 기록 — 헤더 + 행별 J:K (비데이터 행 보호 위해 행별 range)
    data = [{'range': f'{_TAB}!J1:K1', 'values': [['상태', '최종확인일']]}]
    for row, stt, chk in updates:
        data.append({'range': f'{_TAB}!J{row}:K{row}', 'values': [[stt, chk]]})

    written = 0
    for i in range(0, len(data), _WRITE_CHUNK):
        chunk = data[i:i + _WRITE_CHUNK]
        m.service.spreadsheets().values().batchUpdate(
            spreadsheetId=sid,
            body={'valueInputOption': 'USER_ENTERED', 'data': chunk},
        ).execute()
        written += len(chunk)
        logger.info(f'[거래처상태] 기록 진행 {written}/{len(data)}')
    result['written_ranges'] = written
    logger.info(
        f"[거래처상태] 완료 — 총 {len(rows)}행, 폐업 {summary['폐업자']}, "
        f"휴업 {summary['휴업자']}, 조회안됨 {summary['조회안됨']}"
    )
    return result


# ─────────────────────────────────────────────────────────────
# 거래처 탭 Redis 캐시 — ①상호→이메일(계산서 pre-fill) ②번호→상호(OCR 역조회)
# ─────────────────────────────────────────────────────────────
_EMAIL_CACHE_KEY = 'partner_email_map'      # Redis hash: norm_biz -> email
_NAME_BY_BNO_KEY = 'partner_name_by_bno'    # Redis hash: bno(10자리) -> 상호
_NAME_SET_KEY = 'partner_name_set'          # Redis set: norm_biz (거래처 탭 존재 = 발행 가능)
_INFO_MAP_KEY = 'partner_info_map'          # Redis hash: norm_biz -> {bno,name,ceo,addr,biz_type,biz_item,email}


def rebuild_partner_caches() -> dict:
    """거래처 탭 캐시 재구성 (거래처 탭 A:G 1회 로드). 10분마다 + 매일 07:00 국세청 증분 뒤.

    거래처 탭에 줄을 넣으면 10분 안에 계산서 창 이메일·등록증 대신 발행 가능 판정에 반영
    (2026-10-22 — 원래 매일 07:00 국세청 증분 잡에 묶여 다음 날 아침에야 반영됐음).

    ① 상호→이메일 (계산서 모달 pre-fill) — 모호(같은 상호 다른 이메일)·빈값 제외, 정밀 우선.
    ② 번호→상호 (사업자등록증 OCR 역조회) — 사업자번호는 유일키라 모호성 없음.
       숫자 OCR(90%)로 번호만 뽑아 정답 상호를 얻는 용도.

    Returns: {'email_cached', 'ambiguous', 'name_by_bno'}
    """
    m, sid = _sheet()
    vals = m.service.spreadsheets().values().get(
        spreadsheetId=sid, range=f'{_TAB}!A1:G',
    ).execute().get('values', [])
    import json as _json
    name_emails: dict = {}   # norm_biz -> set(email)
    name_by_bno: dict = {}   # bno -> 상호
    name_set: set = set()    # norm_biz 전체 (거래처 탭에 존재 = 계산서 발행 가능 판정용)
    info_map: dict = {}      # norm_biz -> {bno,name,ceo,addr,biz_type,biz_item,email} (등록증 대체 댓글용)
    for r in vals:
        bno_raw = (r[0] if len(r) > 0 else '').strip()
        bno = normalize_bno(bno_raw)
        name = (r[1] if len(r) > 1 else '').strip()
        email = (r[6] if len(r) > 6 else '').strip()
        key = _norm_name(name)
        if key:
            name_set.add(key)
            info_map[key] = {  # 마지막 행 우선 (동일 상호 시)
                'bno': bno_raw, 'name': name,
                'ceo': (r[2] if len(r) > 2 else '').strip(),
                'addr': (r[3] if len(r) > 3 else '').strip(),
                'biz_type': (r[4] if len(r) > 4 else '').strip(),
                'biz_item': (r[5] if len(r) > 5 else '').strip(),
                'email': email,
            }
        if key and email and '@' in email:
            name_emails.setdefault(key, set()).add(email)
        if bno and name:
            name_by_bno[bno] = name  # 번호 유일 → 마지막 값 (거래처 탭 정렬상 무해)
    email_cache = {k: next(iter(v)) for k, v in name_emails.items() if len(v) == 1}
    ambiguous = sum(1 for v in name_emails.values() if len(v) > 1)
    added = set()
    try:
        from dashboard.utils.redis_client import get_redis_client
        rc = get_redis_client().redis
        old_names = set(rc.smembers(_NAME_SET_KEY) or [])
        # 지우기+쓰기를 한 트랜잭션으로 — 10분마다 다시 만들 때 조회가 빈 캐시를 만나
        #   '거래처 아님(발행 불가)'으로 잘못 판정되지 않게 (2026-10-22 매일 07:00 → 10분).
        pipe = rc.pipeline(transaction=True)
        pipe.delete(_EMAIL_CACHE_KEY)
        if email_cache:
            pipe.hset(_EMAIL_CACHE_KEY, mapping=email_cache)
        pipe.delete(_NAME_BY_BNO_KEY)
        if name_by_bno:
            pipe.hset(_NAME_BY_BNO_KEY, mapping=name_by_bno)
        pipe.delete(_NAME_SET_KEY)
        if name_set:
            pipe.sadd(_NAME_SET_KEY, *name_set)
        pipe.delete(_INFO_MAP_KEY)
        if info_map:
            pipe.hset(_INFO_MAP_KEY, mapping={k: _json.dumps(v, ensure_ascii=False) for k, v in info_map.items()})
        pipe.execute()
        if old_names:  # 처음 만들 때(빈 캐시)는 전부 '새 상호'라 제외
            added = name_set - old_names
    except Exception as exc:
        logger.warning(f'[거래처캐시] 저장 실패: {exc}')
    if added:
        _clear_license_negatives(added)
    (logger.info if added else logger.debug)(
        f'[거래처캐시] 재구성 — 이메일 {len(email_cache)}건(모호 {ambiguous} 제외), '
        f'번호→상호 {len(name_by_bno)}건, 상호세트 {len(name_set)}건, 정보맵 {len(info_map)}건'
        + (f', 새 상호 {len(added)}건' if added else '')
    )
    return {'email_cached': len(email_cache), 'ambiguous': ambiguous,
            'name_by_bno': len(name_by_bno), 'name_set': len(name_set), 'info_map': len(info_map),
            'added': len(added)}


def _clear_license_negatives(names: set) -> int:
    """거래처 탭에 새로 들어온 상호의 공사 → 등록증 판정 캐시 비움.

    거래처 탭에 있으면 등록증 파일 없이도 계산서 발행 가능('partner')인데, 그 전에 내려진
    '등록증 없음' 판정(1시간 캐시)이 남아 있으면 줄을 넣은 뒤에도 요청이 막힌다 (2026-10-22).
    """
    n = 0
    try:
        from dashboard.services.project_service import get_project_records
        from dashboard.services.business_license_handler import invalidate_license_state
        for r in (get_project_records() or []):
            if _norm_name(r.get('사업자명') or '') in names:
                invalidate_license_state(str(r.get('프로젝트 코드') or '').strip())
                n += 1
    except Exception as exc:
        logger.warning(f'[거래처캐시] 등록증 판정 캐시 비우기 실패: {exc}')
    if n:
        logger.info(f'[거래처캐시] 새 거래처 {len(names)}건 → 공사 {n}건 등록증 판정 다시 계산')
    return n


def get_partner_info(biz_name: str) -> Optional[dict]:
    """상호 → 거래처 탭 전체정보 {bno,name,ceo,addr,biz_type,biz_item,email} (없으면 None).
    등록증 파일 없는 계산서 요청 시 '거래처 탭 정보로 대신' 댓글용 (2026-09-02)."""
    key = _norm_name(biz_name)
    if not key or key == '-':
        return None
    try:
        import json as _json
        from dashboard.utils.redis_client import get_redis_client
        v = get_redis_client().redis.hget(_INFO_MAP_KEY, key)
        if v:
            return _json.loads(v.decode() if isinstance(v, bytes) else v)
    except Exception:
        pass
    return None


_SEARCH_CACHE = {'at': 0.0, 'rows': []}


def search_partners(query: str, limit: int = 8) -> list:
    """거래처 탭 상호·사업자번호·대표자 부분 검색 — PM [부담자 설정] 자동완성 (2026-10-19).

    Redis 정보맵(partner_info_map) 전체를 5분 캐시해 메모리에서 찾는다(3,500건 내외).
    Returns [{name, bno, ceo, email}] — 상호가 검색어로 시작하는 것 먼저.
    """
    import json as _json
    import time as _time
    q = _norm_name(query)
    q_digits = re.sub(r'\D', '', str(query or ''))
    if len(q) < 2 and len(q_digits) < 3:
        return []
    if _time.time() - _SEARCH_CACHE['at'] > 300 or not _SEARCH_CACHE['rows']:
        try:
            from dashboard.utils.redis_client import get_redis_client
            raw = get_redis_client().redis.hgetall(_INFO_MAP_KEY) or {}
            rows = []
            for v in raw.values():
                try:
                    d = _json.loads(v.decode() if isinstance(v, bytes) else v)
                except Exception:
                    continue
                if d.get('name'):
                    rows.append(d)
            _SEARCH_CACHE.update(at=_time.time(), rows=rows)
        except Exception:
            return []
    hits = []
    for d in _SEARCH_CACHE['rows']:
        nm = _norm_name(d.get('name'))
        bno = re.sub(r'\D', '', str(d.get('bno') or ''))
        ceo = _norm_name(d.get('ceo'))
        if (q and (q in nm or q in ceo)) or (len(q_digits) >= 3 and q_digits in bno):
            hits.append((0 if nm.startswith(q) else 1, d))
    hits.sort(key=lambda t: (t[0], len(str(t[1].get('name') or ''))))
    return [{'name': d.get('name') or '', 'bno': d.get('bno') or '', 'ceo': d.get('ceo') or '',
             'email': d.get('email') or ''} for _, d in hits[:limit]]


def is_partner_known(biz_name: str) -> bool:
    """상호가 거래처 탭에 존재하면 True — '계산서 발행 가능'(홈택스 발행 이력 있는 거래처) 판정.

    등록증 파일이 없어도 거래처 탭에 정보가 있으면 발행 가능(사용자 통찰, 2026-09-02).
    정규화 exact 매칭. 캐시 미구성/오류 시 False (안전: 미지정으로 취급).
    """
    key = _norm_name(biz_name)
    if not key or key == '-':
        return False
    try:
        from dashboard.utils.redis_client import get_redis_client
        return bool(get_redis_client().redis.sismember(_NAME_SET_KEY, key))
    except Exception:
        return False


def get_partner_name_by_bno(bno: str) -> Optional[str]:
    """사업자번호 → 캐시된 거래처 상호 (O(1) Redis). 없으면 None.

    사업자등록증 OCR 역조회용 — 숫자 번호로 정답 상호를 얻어 한글 상호 오인식 회피.
    """
    b = normalize_bno(bno)
    if not b:
        return None
    try:
        from dashboard.utils.redis_client import get_redis_client
        v = get_redis_client().redis.hget(_NAME_BY_BNO_KEY, b)
    except Exception:
        return None
    if v is None:
        return None
    return v.decode() if isinstance(v, bytes) else v


def get_cached_partner_email(biz_name: str) -> Optional[str]:
    """상호명 → 캐시된 거래처 이메일 (O(1) Redis). 없으면 None. trigger 안전."""
    key = _norm_name(biz_name)
    if not key or key == '-':
        return None
    try:
        from dashboard.utils.redis_client import get_redis_client
        v = get_redis_client().redis.hget(_EMAIL_CACHE_KEY, key)
    except Exception:
        return None
    if v is None:
        return None
    return v.decode() if isinstance(v, bytes) else v
