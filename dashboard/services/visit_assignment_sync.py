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

    # 괄호 앞 이니셜 조합 (JW+MS+JK, SJ+JK, TH, 대표님+SD 등)
    # 첫 괄호 or 첫 슬래시 or 첫 숫자 이전 부분
    m = re.match(r'^([가-힣A-Z+,·/\s]+?)(?:\s*\(|\s*/|\s*\d)', prefix)
    if not m:
        # 첫 워드만 시도
        m = re.match(r'^([가-힣A-Z+,·/\s]+)', prefix)
    if not m:
        return None
    candidate = m.group(1).strip()
    if not candidate:
        return None

    # 별명 치환 (대표님 → YG, 정우 → JW)
    for alias, ini in _ALIAS_MAP.items():
        candidate = candidate.replace(alias, ini)

    # 조합 분리 (+, ,, ·, /)
    tokens = [t.strip().upper() for t in re.split(r'[+,·/]', candidate) if t.strip()]
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
                assign = [current_section]
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


def dry_run() -> Dict:
    """캔버스 파싱 결과 표 반환 (변경 X)."""
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
    parsed = parse_assignment_canvas(html)
    initial_to_name, _ = _load_initial_maps()
    phone_map = _match_leads_by_phone(parsed)
    rows: List[Dict] = []
    for p in parsed:
        lead = phone_map.get(p['phone_digits'])
        current_assign = str(lead.get('영업 담당자') or '').strip() if lead else ''
        new_names = ','.join(initial_to_name.get(i, i) for i in p['assign'])
        changed = bool(lead) and (current_assign != new_names)
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
    return {'ok': True, 'rows': rows, 'total': len(rows)}


def commit() -> Dict:
    """실제 시트 업데이트 + 방문 캔버스 A rebuild."""
    from dashboard.services.lead_service import update_lead
    from dashboard.services.visit_canvas_sync import rebuild_canvas_async

    with _ASSIGN_LOCK:
        dr = dry_run()
        if not dr.get('ok'):
            return dr
        updated: List[str] = []
        failed: List[Tuple[str, str]] = []
        for row in dr['rows']:
            if not row['matched'] or not row['changed']:
                continue
            if row['assign_names'] == '-':
                continue
            try:
                update_lead(row['lead_no'], {'영업 담당자': row['assign_names']})
                updated.append(row['lead_no'])
            except Exception as exc:
                failed.append((row['lead_no'], str(exc)))
                logger.error(f'[ASSIGN] {row["lead_no"]} 업데이트 실패: {exc}')
        # 방문 캔버스 A 재빌드 (백그라운드)
        rebuild_canvas_async()
        return {
            'ok': True,
            'total_rows': dr['total'],
            'updated': updated,
            'updated_count': len(updated),
            'failed': failed,
            'failed_count': len(failed),
        }
