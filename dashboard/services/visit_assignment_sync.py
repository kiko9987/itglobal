"""방문 일정 조정 캔버스 (JW 전용) 파서 → 시트 담당자 업데이트 (2026-07-15).

JW 가 매일 방문 일정을 담당자 배정 형태로 붙여넣기 하면 봇이 파싱해서
시트 '영업 담당자' 컬럼을 업데이트하고 방문 캔버스 A 를 rebuild.

파싱 규칙:
  - 라인이 방문 라인 (연락처 있음) → 앞 이니셜 배정 or 섹션 헤더 상속
  - 라인이 짧고 이니셜/이름/별명만 → 섹션 헤더 저장
  - 채널 태그 `(당)`, `(카톡)` 등은 배정 판정에서 제외
  - 별명 매핑: 대표님 → YG, 정우 → JW

두 슬래시 명령으로 조작:
  /일정확인 — dry-run (변경 예정 리스트만 미리보기)
  /일정확정 — 실제 시트 update + 방문 캔버스 A rebuild
"""
from __future__ import annotations

import html as _html_lib
import logging
import os
import re
import threading
from typing import Dict, List, Optional, Tuple

import requests

logger = logging.getLogger(__name__)

_ASSIGN_LOCK = threading.Lock()

# 이름 → 이니셜 별명
_ALIAS_MAP = {
    '대표님': 'YG',
    '정우': 'JW',
}

# 채널 태그 (배정 판정에서 제외)
_CHANNEL_TAGS = {'당', '당근', '카톡', '홈', '홈페이지', '전화', '숨고', '큐플레이스', '메일', '온라인'}

# 연락처 정규식
_PHONE_RE = re.compile(r'0\d{1,2}-?\d{3,4}-?\d{4}')

# 이니셜 조합 (대문자 워드 + / , 로 조합)
_INITIAL_TOKEN_RE = re.compile(r'\b([A-Z]{1,4})\b')


def _fetch_canvas_html(token: str, canvas_id: str) -> Optional[str]:
    """캔버스 HTML content fetch."""
    try:
        info = requests.get(
            'https://slack.com/api/files.info',
            params={'file': canvas_id},
            headers={'Authorization': f'Bearer {token}'},
            timeout=10,
        ).json()
    except Exception as exc:
        logger.error(f'[ASSIGN] files.info 실패: {exc}')
        return None
    if not info.get('ok'):
        logger.warning(f'[ASSIGN] files.info 응답 오류: {info.get("error")}')
        return None
    url = info.get('file', {}).get('url_private', '')
    if not url:
        return None
    try:
        resp = requests.get(url, headers={'Authorization': f'Bearer {token}'}, timeout=10)
    except Exception as exc:
        logger.error(f'[ASSIGN] canvas 다운로드 실패: {exc}')
        return None
    if resp.status_code != 200:
        return None
    return resp.text


def _html_to_lines(html: str) -> List[str]:
    """canvas HTML → 개행 분리된 라인 리스트."""
    text = re.sub(r'<br\s*/?>', '\n', html, flags=re.IGNORECASE)
    text = re.sub(r'</(p|div|li|h\d)>', '\n', text, flags=re.IGNORECASE)
    text = re.sub(r'<[^>]+>', '', text)
    text = _html_lib.unescape(text)
    return [ln.strip() for ln in text.split('\n') if ln.strip()]


def _load_initial_maps() -> Tuple[Dict[str, str], Dict[str, str]]:
    """users.db 로부터 두 매핑 반환.

    Returns:
        (initial_to_name, name_to_initial)
        이니셜 → 이름 (예: YG → 박용구), 이름 → 이니셜 (박용구 → YG)
    """
    initial_to_name: Dict[str, str] = {}
    name_to_initial: Dict[str, str] = {}
    try:
        from dashboard.utils.user_database import UserDatabase
        db = UserDatabase()
        for u in db.get_all_users():
            name = (u.get('name') or '').strip()
            email = (u.get('email') or '').strip()
            if not name or not email:
                continue
            local = email.split('@')[0].strip().upper()
            if local:
                initial_to_name[local] = name
                name_to_initial[name] = local
    except Exception as exc:
        logger.warning(f'[ASSIGN] 이니셜 매핑 로드 실패: {exc}')
    return initial_to_name, name_to_initial


def _is_section_header(line: str, initial_to_name: Dict[str, str]) -> Optional[str]:
    """라인이 섹션 헤더면 배정 대상 이니셜 반환, 아니면 None.

    섹션 헤더 조건:
      - 라인이 짧음 (~30자)
      - 연락처·주소 keyword 없음
      - 이니셜 or 별명 매치
    """
    if len(line) > 40:
        return None
    if _PHONE_RE.search(line):
        return None
    stripped = line.strip()
    # 별명 매핑 (대표님, 정우)
    for alias, ini in _ALIAS_MAP.items():
        if stripped == alias or stripped.startswith(alias):
            return ini
    # "MW", "JW", "SH" 등 이니셜 단독
    m = re.match(r'^([A-Z]{1,4})(?:\s*\([^)]*\))?$', stripped)
    if m and m.group(1) in initial_to_name:
        return m.group(1)
    # 이니셜 조합 (YG+JK, MJ+MW, JW+JSH 등) — 2026-07-20 관측: JW 가 여러 매니저를
    # 한 섹션 아래 묶는 케이스. 반환은 '+' 로 join 된 문자열, 사용부에서 split.
    m = re.match(r'^([A-Z]{1,4}(?:\s*\+\s*[A-Z]{1,4})+)$', stripped)
    if m:
        parts = [p.strip() for p in re.split(r'\+', m.group(1)) if p.strip()]
        if parts and all(p in initial_to_name for p in parts):
            return '+'.join(parts)
    # "온라인 (SD현장지원)" 같이 헤더 카테고리
    if stripped.startswith('온라인'):
        return None  # 온라인은 카테고리 헤더, 배정 대상 아님
    return None


def _extract_lead_initials(line: str, initial_to_name: Dict[str, str]) -> Optional[List[str]]:
    """방문 라인 앞에서 배정 이니셜 조합 추출.

    반환: 이니셜 리스트 (예: ['JW', 'MS', 'JK']) 또는 None (매치 없음)
    """
    # 연락처 나오기 전 부분에서 매치
    phone_match = _PHONE_RE.search(line)
    prefix = line[:phone_match.start()] if phone_match else line

    # 괄호 앞 이니셜 조합 (JW+MS+JK, SJ+JK, TH, 대표님+SD, 대표님 + TH * SD 등)
    # 첫 괄호 or 첫 슬래시 or 첫 숫자 이전 부분
    # 2026-07-19: `*` 도 조합 구분자로 취급 (JW 관행)
    m = re.match(r'^([가-힣A-Z+,·/*\s]+?)(?:\s*\(|\s*/|\s*\d)', prefix)
    if not m:
        # 첫 워드만 시도
        m = re.match(r'^([가-힣A-Z+,·/*\s]+)', prefix)
    if not m:
        return None
    candidate = m.group(1).strip()
    if not candidate:
        return None

    # 별명 치환 (대표님 → YG, 정우 → JW)
    for alias, ini in _ALIAS_MAP.items():
        candidate = candidate.replace(alias, ini)

    # 조합 분리 (+, ,, ·, /, *)
    tokens = [t.strip().upper() for t in re.split(r'[+,·/*]', candidate) if t.strip()]
    # 이니셜 매핑 있는 것만
    valid = [t for t in tokens if t in initial_to_name]
    if not valid:
        return None
    return valid


def _extract_bracket_initial(line: str, initial_to_name: Dict[str, str]) -> Optional[str]:
    """라인 안 괄호 이니셜 (원 담당) 추출. 채널 태그는 무시.

    예: '(YM) ...' → 'YM'
        '(당) ...' → None (채널 태그)
        'SJ (YG) ...' → 'YG'
    """
    for m in re.finditer(r'\(([^)]+)\)', line):
        inner = m.group(1).strip()
        if inner in _CHANNEL_TAGS:
            continue
        # 별명 치환
        for alias, ini in _ALIAS_MAP.items():
            if inner == alias:
                inner = ini
                break
        if inner.upper() in initial_to_name:
            return inner.upper()
    return None


def _normalize_phone(phone: str) -> str:
    """전화번호 → 숫자만."""
    return re.sub(r'\D', '', phone)


def parse_assignment_canvas(html: str) -> List[Dict]:
    """캔버스 HTML → 배정 리스트.

    Returns:
      [{'phone': '010-...', 'phone_digits': '010...', 'assign': ['JW','MS'],
        'original': 'YG'|None, 'raw': line}, ...]
    """
    initial_to_name, _ = _load_initial_maps()
    lines = _html_to_lines(html)
    results: List[Dict] = []
    current_section: Optional[str] = None

    for line in lines:
        # 방문 라인 (연락처 있음) 우선 판정
        phone_m = _PHONE_RE.search(line)
        if phone_m:
            phone = phone_m.group(0)
            # 하이픈 정규화 010-XXXX-XXXX
            phone_digits = _normalize_phone(phone)
            if len(phone_digits) == 11:
                phone_normalized = f'{phone_digits[:3]}-{phone_digits[3:7]}-{phone_digits[7:]}'
            else:
                phone_normalized = phone
            # 배정 이니셜 결정
            assign = _extract_lead_initials(line, initial_to_name)
            if not assign and current_section:
                # 섹션 헤더가 조합('YG+JK') 이면 split 해서 다중 배정.
                assign = [p for p in re.split(r'\+', current_section) if p]
            original = _extract_bracket_initial(line, initial_to_name)
            results.append({
                'phone': phone_normalized,
                'phone_digits': phone_digits,
                'assign': assign or [],
                'original': original,
                'raw': line,
            })
            continue

        # 방문 라인 아니면 섹션 헤더 판정
        section = _is_section_header(line, initial_to_name)
        if section:
            current_section = section

    return results


def _match_leads_by_phone(parsed: List[Dict]) -> Dict[str, Dict]:
    """전화번호 → lead 매핑 (시트에서 조회)."""
    from dashboard.services.lead_service import load_leads_data
    df = load_leads_data(force_refresh=True)
    if df is None or df.empty:
        return {}
    phone_to_lead: Dict[str, Dict] = {}
    for _, row in df.iterrows():
        raw = str(row.get('고객 연락처') or '').strip()
        digits = _normalize_phone(raw)
        if not digits:
            continue
        # 최신 lead 하나만 (시트 순 마지막)
        phone_to_lead[digits] = row.to_dict()
    return phone_to_lead


def _resolve_lead_for_assignment(a: Dict, phone_map: Dict[str, Dict],
                                    addr_candidates: List[Dict]) -> Optional[Dict]:
    """assignment 하나에 대해 lead 매칭. phone 우선, 실패 시 주소 substring fallback.

    2026-07-20: dry_run·commit·DM 발송 모두 이 helper 로 통일. 인투익스·산들해 등
    phone '-' lead 는 주소 substring 매칭으로 배정 반영.
    """
    if a.get('phone_digits'):
        lead = phone_map.get(a['phone_digits'])
        if lead:
            return lead
    if a.get('address'):
        cand_addr = str(a['address']).strip()
        if cand_addr and len(cand_addr) >= 8:
            for _l in addr_candidates:
                _sa = str(_l.get('방문 주소','') or '').strip()
                if not _sa or _sa == '-':
                    continue
                if cand_addr in _sa or _sa in cand_addr:
                    return _l
    return None


def _load_addr_candidates() -> List[Dict]:
    """주소 fallback 매칭 대상 sheet lead (phone 없는 것만 — 오탐 최소)."""
    from dashboard.services.lead_service import load_leads_data
    df = load_leads_data(force_refresh=False)
    if df is None or df.empty:
        return []
    out = []
    for _, row in df.iterrows():
        raw = str(row.get('고객 연락처') or '').strip()
        if raw in ('', '-'):
            out.append(row.to_dict())
    return out


def dry_run() -> Dict:
    """캔버스 파싱 결과 표 반환 (변경 X).

    2026-07-19 확장: online_duty·off_duty·target_date 도 응답에 포함.
    """
    from datetime import date as _date
    token = (
        os.getenv('SLACK_VISIT_BOT_TOKEN', '').strip()
        or os.getenv('SLACK_BOT_TOKEN', '').strip()
    )
    canvas_id = os.getenv('SLACK_VISIT_ASSIGNMENT_CANVAS_ID', '').strip()
    if not token or not canvas_id:
        return {'ok': False, 'reason': 'env 미설정', 'rows': []}
    html = _fetch_canvas_html(token, canvas_id)
    if html is None:
        return {'ok': False, 'reason': '캔버스 fetch 실패', 'rows': []}
    parsed_full = parse_assignment_canvas_full(html)
    parsed = parsed_full['assignments']
    online_duty = parsed_full['online_duty']
    off_duty = parsed_full['off_duty']

    initial_to_name, _ = _load_initial_maps()
    phone_map = _match_leads_by_phone(parsed)
    addr_candidates = _load_addr_candidates()
    rows: List[Dict] = []
    today = _date.today()
    future_starts = []
    for p in parsed:
        lead = _resolve_lead_for_assignment(p, phone_map, addr_candidates)
        current_assign = str(lead.get('영업 담당자') or '').strip() if lead else ''
        new_names = ','.join(initial_to_name.get(i, i) for i in p['assign'])
        changed = bool(lead) and (current_assign != new_names)
        start = _parse_visit_date_start(lead.get('방문 예정일')) if lead else None
        if start is not None and start > today:
            future_starts.append(start)
        rows.append({
            'phone': p['phone'],
            'assign_initials': '+'.join(p['assign']) or '-',
            'assign_names': new_names or '-',
            'matched': bool(lead),
            'lead_no': str(lead.get('리드 No', '')).strip() if lead else '-',
            'current': current_assign or '-',
            'changed': changed,
            'original': p.get('original') or '-',
        })
    target_date = min(future_starts).isoformat() if future_starts else None
    return {
        'ok': True,
        'rows': rows,
        'total': len(rows),
        'online_duty': online_duty,
        'off_duty': off_duty,
        'target_date': target_date,
    }


def commit() -> Dict:
    """실제 시트 업데이트 + Slack List 담당자 + 담당자/온라인 당번 DM + 방문 캔버스 A rebuild.

    2026-07-19 확장:
      - Slack List 담당자 컬럼 update
      - 방문 담당자별 DM (v9 양식)
      - 온라인 당번 DM (v13 양식, 방문 참고 + 휴무 포함)
      - Redis dm_sent:{lead_no} flag (방문일 변경 시 알림 대상 마킹)
    """
    from dashboard.services.lead_service import update_lead
    from dashboard.services.visit_canvas_sync import rebuild_canvas_async

    with _ASSIGN_LOCK:
        # 캔버스 확장 파싱 (assignments + online_duty + off_duty)
        token = (
            os.getenv('SLACK_VISIT_BOT_TOKEN', '').strip()
            or os.getenv('SLACK_BOT_TOKEN', '').strip()
        )
        canvas_id = os.getenv('SLACK_VISIT_ASSIGNMENT_CANVAS_ID', '').strip()
        if not token or not canvas_id:
            return {'ok': False, 'reason': 'env 미설정'}
        html = _fetch_canvas_html(token, canvas_id)
        if html is None:
            return {'ok': False, 'reason': '캔버스 fetch 실패'}
        parsed_full = parse_assignment_canvas_full(html)

        assignments = parsed_full['assignments']
        online_duty = parsed_full['online_duty']
        off_duty = parsed_full['off_duty']

        initial_to_name, _ = _load_initial_maps()
        phone_map = _match_leads_by_phone(assignments)
        addr_candidates = _load_addr_candidates()

        # 1. 시트 update
        updated: List[str] = []
        failed: List[Tuple[str, str]] = []
        for p in assignments:
            lead = _resolve_lead_for_assignment(p, phone_map, addr_candidates)
            if not lead:
                continue
            if not p['assign']:
                continue
            new_names = ','.join(initial_to_name.get(i, i) for i in p['assign'])
            if not new_names or new_names == '-':
                continue
            current = str(lead.get('영업 담당자') or '').strip()
            lno = str(lead.get('리드 No') or '').strip()
            if not lno:
                continue
            if current == new_names:
                continue
            try:
                update_lead(lno, {'영업 담당자': new_names})
                updated.append(lno)
            except Exception as exc:
                failed.append((lno, str(exc)))
                logger.error(f'[ASSIGN] {lno} 시트 update 실패: {exc}')

        # 2. Slack List 담당자 update (주소 fallback 위해 sheet_leads 넘김)
        try:
            from dashboard.services.visit_canvas_sync import _fetch_visit_leads
            _sheet_leads = _fetch_visit_leads()
        except Exception:
            _sheet_leads = None
        list_ok, list_fail = _update_slack_list_managers(
            assignments, phone_map, sheet_leads=_sheet_leads,
        )

        # 3. DM 발송 (target_date = min(방문일 시작일) > today)
        dm_result = _send_dms_for_next_visit(
            assignments, phone_map, initial_to_name, online_duty, off_duty,
            addr_candidates=addr_candidates,
        )

        # 4. 방문 캔버스 A rebuild
        rebuild_canvas_async()

        return {
            'ok': True,
            'total_rows': len(assignments),
            'updated': updated,
            'updated_count': len(updated),
            'failed': failed,
            'failed_count': len(failed),
            'list_updated': list_ok,
            'list_failed': list_fail,
            'dm': dm_result,
            'online_duty': online_duty,
            'off_duty': off_duty,
        }


# ---------------------------------------------------------------------------
# 2026-07-19 확장: 온라인 당번·휴무 파싱 + List update + DM 발송
# ---------------------------------------------------------------------------

# List 담당자 컬럼 옵션은 스키마 실시간 조회 (2026-07-19 자동화).
# 10분 캐시.
_LIST_COL_LEAD = 'Col087VA2RG3G'
_LIST_COL_MGR = 'Col087WMMT84W'

_LIST_OPT_CACHE: Dict[str, object] = {'map': {}, 'ts': 0.0}


def _load_list_manager_option_map(list_id: str) -> Dict[str, str]:
    """List 담당자 컬럼 스키마 → {이니셜: option value}.

    Slack List 스키마의 담당자 컬럼 (select) choices 를 users.db 이름 매핑으로
    이니셜 (email local part 대문자) 로 변환.
    """
    if not list_id:
        return {}
    client = _get_visit_client()
    if not client:
        return {}
    try:
        r = client.files_info(file=list_id)
        schema = r['file'].get('list_metadata', {}).get('schema', [])
    except Exception as exc:
        logger.warning(f'[ASSIGN/LIST] 스키마 조회 실패: {exc}')
        return {}

    _, name_to_initial = _load_initial_maps()
    result: Dict[str, str] = {}
    for col in schema:
        if col.get('id') != _LIST_COL_MGR:
            continue
        for ch in col.get('options', {}).get('choices', []):
            label = (ch.get('label') or '').strip()
            value = (ch.get('value') or '').strip()
            if not label or not value:
                continue
            ini = name_to_initial.get(label)
            if ini:
                result[ini] = value
    return result


def _get_list_manager_option_map(list_id: str) -> Dict[str, str]:
    """캐시 10분."""
    import time as _time
    now = _time.time()
    ts = float(_LIST_OPT_CACHE.get('ts') or 0)
    cached = _LIST_OPT_CACHE.get('map') or {}
    if now - ts < 600 and cached:
        return cached  # type: ignore
    fresh = _load_list_manager_option_map(list_id)
    if fresh:
        _LIST_OPT_CACHE['map'] = fresh
        _LIST_OPT_CACHE['ts'] = now
    return fresh or cached  # type: ignore

_SEP = '--------------------------------------------'
_BLANK = '⠀'
_EMOJIS = [':one:', ':two:', ':three:', ':four:', ':five:',
           ':six:', ':seven:', ':eight:', ':nine:', ':keycap_ten:']


def parse_assignment_canvas_full(html: str) -> Dict:
    """캔버스 파싱 확장 — 배정 + 온라인 당번 + 휴무.

    Returns:
      {'assignments': [...], 'online_duty': ['JK'], 'off_duty': ['JSH', 'YM', 'SJ']}
    """
    initial_to_name, _ = _load_initial_maps()
    lines = _html_to_lines(html)
    assignments: List[Dict] = []
    online_duty: List[str] = []
    off_duty: List[str] = []
    current_section: Optional[str] = None  # 'YG' etc or '_ONLINE_' or '_OFF_'

    for line in lines:
        # 휴무인원 ( XX + YY + ZZ ) 라인
        if '휴무인원' in line:
            m = re.search(r'\(([^)]+)\)', line)
            if m:
                off_duty = [x for x in re.findall(r'[A-Z]{2,4}', m.group(1))
                            if x in initial_to_name]
            continue

        phone_m = _PHONE_RE.search(line)
        if not phone_m:
            stripped = line.strip()
            if stripped == '온라인':
                current_section = '_ONLINE_'
                continue
            if stripped == '휴무':
                current_section = '_OFF_'
                continue
            # 온라인 섹션 안 이니셜 = 상담 당번
            if current_section == '_ONLINE_' and stripped in initial_to_name:
                if stripped not in online_duty:
                    online_duty.append(stripped)
                continue
            # 개인 담당자 섹션 헤더
            section = _is_section_header(line, initial_to_name)
            if section:
                current_section = section
                continue
            # phone 없는 방문 라인 (인투익스 등 연락처 '-' 케이스, 2026-07-20)
            # '/' 로 여러 필드 구분된 형태에 주소가 있으면 assignment 로 추가 → 주소 fallback 매칭용
            if current_section and current_section not in ('_ONLINE_', '_OFF_') and line.count('/') >= 2:
                parts = [p.strip() for p in line.split('/')]
                # 형태: "(MW) 7월 21일 / - / 주소 / 내용"  →  parts[2] = 주소
                address = parts[2] if len(parts) >= 3 else ''
                assign = _extract_lead_initials(line, initial_to_name)
                if not assign:
                    assign = [p for p in re.split(r'\+', current_section) if p]
                if address and address != '-':
                    assignments.append({
                        'phone': '',
                        'phone_digits': '',
                        'assign': assign or [],
                        'original': _extract_bracket_initial(line, initial_to_name),
                        'raw': line,
                        'address': address,
                    })
            continue

        # phone 있음 = 방문 라인
        if current_section in (None, '_ONLINE_', '_OFF_'):
            continue

        phone = phone_m.group(0)
        phone_digits = _normalize_phone(phone)
        if len(phone_digits) == 11:
            phone_normalized = f'{phone_digits[:3]}-{phone_digits[3:7]}-{phone_digits[7:]}'
        else:
            phone_normalized = phone
        assign = _extract_lead_initials(line, initial_to_name)
        if not assign and current_section:
            # 섹션 헤더가 조합('YG+JK') 이면 split 해서 다중 배정.
            assign = [p for p in re.split(r'\+', current_section) if p]
        original = _extract_bracket_initial(line, initial_to_name)
        assignments.append({
            'phone': phone_normalized,
            'phone_digits': phone_digits,
            'assign': assign or [],
            'original': original,
            'raw': line,
        })

    return {
        'assignments': assignments,
        'online_duty': online_duty,
        'off_duty': off_duty,
    }


def _parse_visit_date_start(vd) -> Optional['date']:
    """방문 예정일 → 시작 date."""
    from datetime import date as _date
    s = str(vd or '').strip().lstrip("'")
    if not s:
        return None
    m = re.match(r'^(\d{4})[-./](\d{1,2})[-./](\d{1,2})', s)
    if not m:
        return None
    try:
        return _date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    except ValueError:
        return None


def _get_visit_client():
    """방문봇 client 반환."""
    from dashboard.blueprints.slack_bot import _init_visit_slack_app, _visit_slack_app
    _init_visit_slack_app()
    return _visit_slack_app.client if _visit_slack_app else None


def _get_visit_card_permalink(client, lead_no: str) -> str:
    """Redis visit_notice_msg 매핑 → 방문 카드 permalink."""
    try:
        from dashboard.utils.redis_client import get_redis_client
        rc = get_redis_client().redis
        v = rc.get(f'visit_notice_msg:{lead_no}')
        if not v:
            return ''
        val = v.decode() if isinstance(v, bytes) else v
        ch, ts = val.split('|', 1)
        return client.chat_getPermalink(channel=ch, message_ts=ts).get('permalink', '')
    except Exception:
        return ''


def _email_from_initial(initial: str, initial_to_name: Dict[str, str]) -> str:
    """이니셜 → 이메일. users.db 매핑에서 역산."""
    from dashboard.utils.user_database import UserDatabase
    try:
        db = UserDatabase()
        for u in db.get_all_users():
            email = (u.get('email') or '').strip()
            if not email:
                continue
            local = email.split('@')[0].strip().upper()
            if local == initial.upper():
                return email
    except Exception:
        pass
    return ''


def _update_slack_list_managers(assignments: List[Dict],
                                  phone_map: Dict[str, Dict],
                                  sheet_leads: Optional[List[Dict]] = None) -> Tuple[int, int]:
    """Slack List 담당자 컬럼 update. (성공, 실패) 반환.

    2026-07-20: 주소 fallback 추가. phone 없는 캔버스2 라인은 assignment 의 address 로
    시트 lead 주소를 substring 매칭. (인투익스·관악구자활센터·산들해 등 phone '-' 대응)
    """
    import json as _json
    import urllib.request

    list_id = os.getenv('SLACK_VISIT_LIST_ID', '').strip()
    token = os.getenv('SLACK_VISIT_BOT_TOKEN', '').strip()
    if not list_id or not token:
        logger.warning('[ASSIGN/LIST] SLACK_VISIT_LIST_ID or token 미설정 — skip')
        return 0, 0

    client = _get_visit_client()
    if not client:
        return 0, 0

    # 옵션 매핑 실시간 조회 (10분 캐시)
    opt_map = _get_list_manager_option_map(list_id)
    if not opt_map:
        logger.warning('[ASSIGN/LIST] 옵션 매핑 조회 실패 — skip')
        return 0, 0

    # List 항목 조회 → lead_no → row_id 매핑
    try:
        r = client.api_call('slackLists.items.list', http_verb='GET',
                            params={'list_id': list_id})
        items = r.data.get('items', []) if hasattr(r, 'data') else r.get('items', [])
    except Exception as exc:
        logger.warning(f'[ASSIGN/LIST] items 조회 실패: {exc}')
        return 0, 0

    lead_to_row: Dict[str, str] = {}
    for it in items:
        for f in it.get('fields', []):
            if f.get('column_id') == _LIST_COL_LEAD:
                lead_to_row[f.get('text', '')] = it['id']
                break

    # 주소 substring 매칭용 sheet lead 후보 (phone 없는 lead 만 대상 — 오탐 최소)
    _addr_candidates: List[Dict] = []
    if sheet_leads:
        for _l in sheet_leads:
            _phone = str(_l.get('고객 연락처','') or '').strip()
            if _phone in ('', '-'):
                _addr_candidates.append(_l)

    # 2026-07-20: 같은 lead 가 여러 assignment 로 나올 때 (동행 방문 — TH 섹션·MS 섹션
    # 양쪽 나온 경우) 순차 update 로 뒤엣것이 앞엣것을 덮어쓰는 이슈. lead_no 로 병합해
    # 이니셜 union → 한 번에 update.
    merged: Dict[str, Dict] = {}  # lead_no → {'row_id': ..., 'inis': set()}
    for p in assignments:
        lead = None
        if p.get('phone_digits'):
            lead = phone_map.get(p['phone_digits'])
        if not lead and p.get('address'):
            cand_addr = p['address'].strip()
            if cand_addr and len(cand_addr) >= 8:
                for _l in _addr_candidates:
                    _sheet_addr = str(_l.get('방문 주소','') or '').strip()
                    if not _sheet_addr or _sheet_addr == '-':
                        continue
                    if cand_addr in _sheet_addr or _sheet_addr in cand_addr:
                        lead = _l
                        break
        if not lead:
            continue
        lno = str(lead.get('리드 No') or '').strip()
        row_id = lead_to_row.get(lno)
        if not row_id:
            continue
        entry = merged.setdefault(lno, {'row_id': row_id, 'inis': set()})
        for i in (p.get('assign') or []):
            entry['inis'].add(i)

    ok = 0
    fail = 0
    for lno, entry in merged.items():
        missing = [i for i in entry['inis'] if i not in opt_map]
        if missing:
            logger.warning(
                f'[ASSIGN/LIST] {lno}: 옵션 매핑 누락 {missing} — 해당만 skip'
            )
        opts = [opt_map[i] for i in entry['inis'] if i in opt_map]
        if not opts:
            continue
        body = {
            'list_id': list_id,
            'cells': [{'row_id': entry['row_id'], 'column_id': _LIST_COL_MGR, 'select': opts}],
        }
        try:
            req = urllib.request.Request(
                'https://slack.com/api/slackLists.items.update',
                data=_json.dumps(body).encode('utf-8'),
                headers={'Content-Type': 'application/json; charset=utf-8',
                         'Authorization': f'Bearer {token}'},
                method='POST',
            )
            with urllib.request.urlopen(req, timeout=10) as resp:
                result = _json.loads(resp.read().decode())
            if result.get('ok'):
                ok += 1
            else:
                fail += 1
                logger.warning(f'[ASSIGN/LIST] {lno} update 실패: {result.get("error")}')
        except Exception as exc:
            fail += 1
            logger.warning(f'[ASSIGN/LIST] {lno} 예외: {exc}')
    return ok, fail


def _send_dms_for_next_visit(assignments: List[Dict],
                              phone_map: Dict[str, Dict],
                              initial_to_name: Dict[str, str],
                              online_duty: List[str],
                              off_duty: List[str],
                              addr_candidates: Optional[List[Dict]] = None) -> Dict:
    """min(방문일 시작일) > today 인 리드에 담당자/온라인 당번 DM.

    2026-07-19 확장: dm_sent 를 JSON 으로 관리해 재실행 시 신규/유지/제거 분류.
      - 신규 매니저 → v9 DM
      - 유지 매니저 → skip (중복 방지)
      - 제거 매니저 → v20 배정 해제 알림

    Returns:
        {'target_date': 'YYYY-MM-DD' or None,
         'visit_mgr_sent': N, 'visit_mgr_failed': N, 'deassign_sent': N,
         'online_duty_sent': N, 'lead_nos_flagged': [...]}
    """
    import json as _json
    from datetime import date as _date
    result = {'target_date': None, 'visit_mgr_sent': 0, 'visit_mgr_failed': 0,
              'deassign_sent': 0, 'online_duty_sent': 0, 'lead_nos_flagged': []}

    today = _date.today()
    _addr_cands = addr_candidates or _load_addr_candidates()
    # 각 assignment → (lead, start_date)
    enriched = []
    for p in assignments:
        lead = _resolve_lead_for_assignment(p, phone_map, _addr_cands)
        if not lead:
            continue
        vd = lead.get('방문 예정일', '')
        start = _parse_visit_date_start(vd)
        if start is None or start <= today:
            continue
        enriched.append((p, lead, start))

    if not enriched:
        logger.info('[ASSIGN/DM] 오늘 이후 방문 없음 — skip')
        return result

    target_date = min(x[2] for x in enriched)
    result['target_date'] = target_date.isoformat()

    filtered = [(p, lead) for p, lead, s in enriched if s == target_date]

    # 담당자별 그룹핑 + lead → 담당자 매핑
    from collections import defaultdict
    lead_to_mgrs: Dict[str, List[str]] = {}
    for p, lead in filtered:
        lno = str(lead.get('리드 No') or '').strip()
        lead_to_mgrs[lno] = p['assign']

    # 각 lead 이전 dm_sent 조회 (JSON) — 신규/유지/제거 분류용
    try:
        from dashboard.utils.redis_client import get_redis_client
        rc_pre = get_redis_client().redis
    except Exception:
        rc_pre = None

    prev_mgrs_by_lead: Dict[str, set] = {}
    for lno in lead_to_mgrs:
        if rc_pre is None:
            prev_mgrs_by_lead[lno] = set()
            continue
        try:
            raw = rc_pre.get(f'dm_sent:{lno}')
            if not raw:
                prev_mgrs_by_lead[lno] = set()
                continue
            val = raw.decode() if isinstance(raw, bytes) else raw
            # JSON format {"date":..., "mgrs":[...]} 또는 legacy 문자열
            if val.startswith('{'):
                data = _json.loads(val)
                prev_mgrs_by_lead[lno] = set(data.get('mgrs', []))
            else:
                # legacy — target_date 만 있고 mgrs 정보 없음, 첫 확장 실행 시엔
                # 이미 발송된 것으로 간주 (모두 유지) 하려 해도 정보 없으므로 빈 set
                prev_mgrs_by_lead[lno] = set()
        except Exception as exc:
            logger.debug(f'[ASSIGN/DM] dm_sent 파싱 실패 ({lno}): {exc}')
            prev_mgrs_by_lead[lno] = set()

    # 신규 배정 (첫 발송 대상) 과 제거 (배정 해제) 분류
    by_mgr_new: Dict[str, List[Tuple[str, Dict]]] = defaultdict(list)
    deassigned_by_mgr: Dict[str, List[Tuple[str, Dict]]] = defaultdict(list)
    for p, lead in filtered:
        lno = str(lead.get('리드 No') or '').strip()
        current_set = set(p['assign'])
        prev_set = prev_mgrs_by_lead.get(lno, set())
        for ini in current_set - prev_set:
            by_mgr_new[ini].append((lno, lead))
        for ini in prev_set - current_set:
            deassigned_by_mgr[ini].append((lno, lead))

    # 방문봇 client
    client = _get_visit_client()
    if not client:
        logger.warning('[ASSIGN/DM] 방문봇 client 초기화 실패')
        return result

    lead_nos_flagged = set()

    # 배정 해제 알림 (v20)
    for ini, dea_leads in deassigned_by_mgr.items():
        mgr_name = initial_to_name.get(ini, ini)
        email = _email_from_initial(ini, initial_to_name)
        if not email:
            logger.warning(
                f'[ASSIGN/DM] 해제 알림 {mgr_name}({ini}): 이메일 매핑 실패 — skip'
            )
            continue
        lines = [_BLANK]
        lines.append(
            f'>:no_entry: *{mgr_name}님*, {target_date.isoformat()} '
            f'방문 배정에서 제외되었습니다 ({len(dea_leads)}건).'
        )
        lines.append(f'>{_SEP}')
        for i, (lno, lead) in enumerate(dea_leads):
            e = _EMOJIS[i] if i < len(_EMOJIS) else f'*{i+1}.*'
            lines.append(f'>{e} *{lno} · {lead.get("고객명") or "-"}*')
        lines.append(
            '>:information_source: 다른 매니저에게 재배정되었습니다.'
        )
        lines.append(f'>{_SEP}')
        lines.append(_BLANK)
        try:
            u = client.users_lookupByEmail(email=email)
            uid = u['user']['id']
            r = client.chat_postMessage(
                channel=uid, text='\n'.join(lines),
                unfurl_links=False, unfurl_media=False,
            )
            if r.get('ok'):
                result['deassign_sent'] += 1
                logger.info(
                    f'[ASSIGN/DM] 해제 알림 {mgr_name}({ini}) → {len(dea_leads)}건'
                )
        except Exception as exc:
            logger.warning(
                f'[ASSIGN/DM] 해제 알림 {mgr_name}({ini}) 발송 예외: {exc}'
            )

    # 담당자별 v9 DM — 신규 매니저만 발송
    for ini, mgr_leads in by_mgr_new.items():
        mgr_name = initial_to_name.get(ini, ini)
        email = _email_from_initial(ini, initial_to_name)
        if not email:
            logger.warning(
                f'[ASSIGN/DM] {mgr_name}({ini}): 이메일 매핑 실패 — skip'
            )
            result['visit_mgr_failed'] += 1
            continue
        # 동행 계산
        companions_per_lead = {}
        for lno, _ in mgr_leads:
            others = [x for x in lead_to_mgrs.get(lno, []) if x != ini]
            companions_per_lead[lno] = tuple(others)
        uniq = set(companions_per_lead.values())
        common = list(uniq)[0] if len(uniq) == 1 and uniq != {tuple()} else None

        lines = [_BLANK]
        lines.append(
            f'>:wave: *{mgr_name}님*, {target_date.isoformat()} '
            f'배정된 방문 일정 {len(mgr_leads)}건 입니다.'
        )
        if common:
            lines.append(f'>:busts_in_silhouette: 동행 : {"+".join(common)}')
        lines.append(f'>{_SEP}')
        for i, (lno, lead) in enumerate(mgr_leads):
            e = _EMOJIS[i] if i < len(_EMOJIS) else f'*{i+1}.*'
            lines.append(f'>{e} *{lno} · {lead.get("고객명") or "-"}*')
            if not common and companions_per_lead.get(lno):
                lines.append(
                    f'>   :busts_in_silhouette: 동행 {"+".join(companions_per_lead[lno])}'
                )
            lines.append(f'>   :iphone: {lead.get("고객 연락처") or "-"}')
            lines.append(f'>   :round_pushpin: {lead.get("방문 주소") or "-"}')
            note = (lead.get('상담 내용') or lead.get('문의 내용') or '').strip()
            if note:
                lines.append(f'>   :speech_balloon: {note[:200]}')
            pl = _get_visit_card_permalink(client, lno)
            if pl:
                lines.append(f'>   :link: <{pl}|방문 카드>')
            lines.append(f'>{_SEP}')
            lead_nos_flagged.add(lno)
        lines.append(_BLANK)

        try:
            u = client.users_lookupByEmail(email=email)
            uid = u['user']['id']
            r = client.chat_postMessage(
                channel=uid, text='\n'.join(lines),
                unfurl_links=False, unfurl_media=False,
            )
            if r.get('ok'):
                result['visit_mgr_sent'] += 1
                logger.info(
                    f'[ASSIGN/DM] {mgr_name}({ini}) → {len(mgr_leads)}건 발송 OK'
                )
            else:
                result['visit_mgr_failed'] += 1
                logger.warning(
                    f'[ASSIGN/DM] {mgr_name}({ini}) 응답 not ok: {r.get("error")}'
                )
        except Exception as exc:
            result['visit_mgr_failed'] += 1
            logger.warning(f'[ASSIGN/DM] {mgr_name}({ini}) 발송 예외: {exc}')

    # 온라인 당번 v13 DM
    if online_duty:
        # 방문 담당자 참고 (조합별 카운트)
        combo_counts: Dict[Tuple[str, ...], int] = {}
        for lno, mgrs in lead_to_mgrs.items():
            key = tuple(mgrs)
            combo_counts[key] = combo_counts.get(key, 0) + 1
        combo_lines = [f'   • {"·".join(k)} ({v}건)'
                       for k, v in combo_counts.items()]

        for duty_ini in online_duty:
            duty_name = initial_to_name.get(duty_ini, duty_ini)
            email = _email_from_initial(duty_ini, initial_to_name)
            if not email:
                logger.warning(
                    f'[ASSIGN/DM] 온라인 당번 {duty_name}({duty_ini}): '
                    f'이메일 매핑 실패 — skip'
                )
                continue
            # 재실행 중복 방지 (2026-07-19)
            duty_flag = f'duty_sent:{target_date.isoformat()}:{duty_ini}'
            try:
                if rc_pre is not None and rc_pre.get(duty_flag):
                    logger.info(
                        f'[ASSIGN/DM] 온라인 당번 {duty_name}({duty_ini}) '
                        f'이미 발송 — skip'
                    )
                    continue
            except Exception:
                pass
            lines = [_BLANK]
            lines.append(
                f'>:wave: *{duty_name}님*, {target_date.isoformat()} 온라인 상담 당번입니다.'
            )
            lines.append(f'>{_SEP}')
            lines.append('>:headphones: 사무실에서 문의 응대 부탁드립니다.')
            lines.append('>')
            lines.append(f'>:car: *{target_date.isoformat()} 방문 담당자 참고*')
            for cl in combo_lines:
                lines.append(f'>{cl}')
            if off_duty:
                lines.append('>')
                lines.append(f'>:palm_tree: *휴무* : {"·".join(off_duty)}')
            lines.append(f'>{_SEP}')
            lines.append(_BLANK)
            try:
                u = client.users_lookupByEmail(email=email)
                uid = u['user']['id']
                r = client.chat_postMessage(
                    channel=uid, text='\n'.join(lines),
                    unfurl_links=False, unfurl_media=False,
                )
                if r.get('ok'):
                    result['online_duty_sent'] += 1
                    logger.info(
                        f'[ASSIGN/DM] 온라인 당번 {duty_name}({duty_ini}) 발송 OK'
                    )
                    # 중복 방지 flag 세팅 (3일 TTL)
                    try:
                        if rc_pre is not None:
                            rc_pre.set(duty_flag, '1', ex=86400 * 3)
                    except Exception:
                        pass
                else:
                    logger.warning(
                        f'[ASSIGN/DM] 당번 {duty_name}({duty_ini}) 응답 not ok: '
                        f'{r.get("error")}'
                    )
            except Exception as exc:
                logger.warning(
                    f'[ASSIGN/DM] 당번 {duty_name}({duty_ini}) 발송 예외: {exc}'
                )

    # Redis dm_sent JSON 저장 — 모든 filtered lead 에 대해 최종 mgrs 기록.
    # {"date": "YYYY-MM-DD", "mgrs": ["YG","TH",...]}
    # 재실행 시 diff 판정용. 신규/유지/제거 분류.
    all_flagged = set(lead_to_mgrs.keys())
    try:
        from dashboard.utils.redis_client import get_redis_client
        rc = get_redis_client().redis
        ttl_days = (target_date - today).days + 7
        ttl = max(ttl_days * 86400, 86400)
        for lno in all_flagged:
            payload = _json.dumps({
                'date': target_date.isoformat(),
                'mgrs': list(lead_to_mgrs.get(lno, [])),
            })
            try:
                rc.set(f'dm_sent:{lno}', payload, ex=ttl)
            except Exception:
                pass
        result['lead_nos_flagged'] = sorted(all_flagged)
    except Exception as exc:
        logger.warning(f'[ASSIGN/DM] Redis flag 세팅 실패: {exc}')

    return result


def send_visit_cancel_notification(lead_no: str, canceller_initial: str,
                                     reason: str) -> bool:
    """방문 취소 시 dm_sent JSON 의 mgrs 조회 → 취소자 제외한 매니저에게 v21 알림.

    Args:
        lead_no: 리드 No
        canceller_initial: 취소자 이니셜
        reason: 취소 사유 (모달 입력)
    """
    import json as _json
    try:
        from dashboard.utils.redis_client import get_redis_client
        rc = get_redis_client().redis
        raw = rc.get(f'dm_sent:{lead_no}')
        if not raw:
            logger.info(f'[ASSIGN/CANCEL] {lead_no}: dm_sent 없음 — skip')
            return False
        val = raw.decode() if isinstance(raw, bytes) else raw
        try:
            data = _json.loads(val) if val.startswith('{') else {'mgrs': []}
        except Exception:
            data = {'mgrs': []}
        mgrs = list(data.get('mgrs', []))
    except Exception as exc:
        logger.warning(f'[ASSIGN/CANCEL] dm_sent 조회 실패 ({lead_no}): {exc}')
        return False

    # 취소자 제외
    targets = [i for i in mgrs if i != canceller_initial]
    if not targets:
        logger.info(f'[ASSIGN/CANCEL] {lead_no}: 알림 대상 없음 (혼자 배정)')
        return False

    # 시트 조회
    try:
        from dashboard.services.lead_service import load_leads_data
        df = load_leads_data(force_refresh=True)
        row = df[df['리드 No'] == lead_no]
        if row.empty:
            return False
        lead = row.iloc[0].to_dict()
    except Exception:
        return False

    lead_name = lead.get('고객명') or '-'
    visit_date = str(lead.get('방문 예정일') or '').strip().lstrip("'") or '-'
    initial_to_name, _ = _load_initial_maps()

    client = _get_visit_client()
    if not client:
        return False

    lines = [_BLANK]
    lines.append(f'>:x: *방문 취소 알림*')
    lines.append(f'>{_SEP}')
    lines.append(f'>*{lead_no} · {lead_name}*')
    lines.append(f'>{visit_date} 방문이 취소되었습니다.')
    if reason:
        lines.append(f'>:memo: {reason[:200]}')
    lines.append(f'>:information_source: 취소자 : {canceller_initial}')
    lines.append(f'>{_SEP}')
    lines.append(_BLANK)
    text = '\n'.join(lines)

    sent = 0
    for ini in targets:
        mgr_name = initial_to_name.get(ini, ini)
        email = _email_from_initial(ini, initial_to_name)
        if not email:
            logger.warning(f'[ASSIGN/CANCEL] {mgr_name}({ini}): 이메일 매핑 실패')
            continue
        try:
            u = client.users_lookupByEmail(email=email)
            uid = u['user']['id']
            r = client.chat_postMessage(
                channel=uid, text=text,
                unfurl_links=False, unfurl_media=False,
            )
            if r.get('ok'):
                sent += 1
        except Exception as exc:
            logger.warning(f'[ASSIGN/CANCEL] {mgr_name}({ini}) 발송 실패: {exc}')

    logger.info(f'[ASSIGN/CANCEL] {lead_no}: {sent}/{len(targets)}명 알림 발송')
    return sent > 0


def send_visit_change_notification(lead_no: str, old_visit_date: str,
                                     new_visit_date: str, note: str = '') -> bool:
    """방문일 변경 시 담당자에게 v19 양식 DM. dm_sent flag 있는 lead 만 대상.

    Args:
        lead_no: 리드 No
        old_visit_date, new_visit_date: 변경 전/후 방문일 문자열
        note: 상담 내용 (매니저 변경 사유)
    """
    try:
        from dashboard.utils.redis_client import get_redis_client
        rc = get_redis_client().redis
        flag = rc.get(f'dm_sent:{lead_no}')
        if not flag:
            logger.info(f'[ASSIGN/CHANGE] {lead_no}: dm_sent flag 없음 — skip')
            return False
    except Exception:
        return False

    # 시트에서 lead 조회 → 이름·담당자
    try:
        from dashboard.services.lead_service import load_leads_data
        df = load_leads_data(force_refresh=True)
        row = df[df['리드 No'] == lead_no]
        if row.empty:
            return False
        lead = row.iloc[0].to_dict()
    except Exception:
        return False

    lead_name = lead.get('고객명') or '-'
    mgr_names_raw = str(lead.get('영업 담당자') or '').strip()
    if not mgr_names_raw:
        return False

    initial_to_name, name_to_initial = _load_initial_maps()

    client = _get_visit_client()
    if not client:
        return False

    lines = [_BLANK]
    lines.append(f'>:arrows_counterclockwise: *방문 일정 변경*')
    lines.append(f'>{_SEP}')
    lines.append(f'>*{lead_no} · {lead_name}*')
    lines.append(f'>~{old_visit_date}~ → *{new_visit_date}*')
    if note:
        lines.append(f'>:memo: {note[:200]}')
    lines.append('>:information_source: 변경된 날짜 전날 다시 안내 드릴 예정입니다')
    lines.append(f'>{_SEP}')
    lines.append(_BLANK)
    text = '\n'.join(lines)

    sent = 0
    for name in [n.strip() for n in mgr_names_raw.split(',') if n.strip()]:
        ini = name_to_initial.get(name)
        if not ini:
            continue
        email = _email_from_initial(ini, initial_to_name)
        if not email:
            continue
        try:
            u = client.users_lookupByEmail(email=email)
            uid = u['user']['id']
            r = client.chat_postMessage(
                channel=uid, text=text,
                unfurl_links=False, unfurl_media=False,
            )
            if r.get('ok'):
                sent += 1
        except Exception as exc:
            logger.warning(f'[ASSIGN/CHANGE] {name} 발송 실패: {exc}')

    logger.info(f'[ASSIGN/CHANGE] {lead_no}: {sent}명 변경 알림 발송')
    return sent > 0
