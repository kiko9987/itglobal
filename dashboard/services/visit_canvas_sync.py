"""방문 일정 캔버스 자동 sync (2026-07-15 도입).

봇이 슬랙 캔버스 (SLACK_VISIT_CANVAS_ID) 전체를 관리.
방문 예약 lead 를 3개 카테고리로 분류·정렬해 markdown 으로 렌더 후
canvases.edit API 로 전체 replace.

카테고리 매핑:
  - 기타 = 리드 No 가 'ETC-' prefix (기타 방문 pseudo lead)
  - 거래처 = 플랫폼 '거래처' or '소개'
  - 온라인 방문 = 플랫폼 '당근', '홈페이지', '카카오톡', '전화'

트리거:
  - 방문 예약 등록 후 rebuild_canvas() 호출 (_post_to_slack_list 훅)
  - 방문 완료 처리 시 상태 변경 → rebuild 재호출 → 자동 제외
"""
from __future__ import annotations

import logging
import os
import re
import threading
from datetime import date, datetime
from typing import Dict, List, Optional

import requests

logger = logging.getLogger(__name__)

_CANVAS_LOCK = threading.Lock()


def _get_initial_map() -> Dict[str, str]:
    """users.db 에서 {이름: 이니셜} 매핑.

    이니셜 = 이메일 로컬 파트 대문자. 매번 신선한 데이터 (매니저 추가/제거 즉시 반영).
    """
    mapping: Dict[str, str] = {}
    try:
        from dashboard.utils.user_database import UserDatabase
        db = UserDatabase()
        for u in db.get_all_users():
            name = (u.get('name') or '').strip()
            email = (u.get('email') or '').strip()
            if not name or not email:
                continue
            local = email.split('@')[0].strip()
            if not local:
                continue
            mapping[name] = local.upper()
    except Exception as exc:
        logger.warning(f'[VISIT_CANVAS] 이니셜 매핑 로드 실패: {exc}')
    return mapping


def _initial_from_name(name: str, mapping: Dict[str, str]) -> str:
    """이름 → 이니셜. 콤마 조합 (박용구,권태훈) 도 처리."""
    if not name or name.strip() in ('', '-'):
        return '-'
    parts = [p.strip() for p in re.split(r'[,·+/]', name) if p.strip()]
    initials = [mapping.get(p, p) for p in parts]
    return ','.join(initials) or '-'


def _categorize(lead: Dict) -> Optional[str]:
    lead_no = str(lead.get('리드 No') or '').strip()
    platform = str(lead.get('플랫폼') or '').strip()
    if lead_no.startswith('ETC-') or platform == '기타':
        return '기타'
    if platform in ('거래처', '소개'):
        return '거래처'
    if platform in ('당근', '홈페이지', '카카오톡', '전화'):
        return '온라인 방문'
    return None


def _fmt_visit_date(raw) -> str:
    """방문 예정일 → 'MM월 DD일' 또는 'MM월 DD~DD일' 표시."""
    s = str(raw or '').strip().lstrip("'")
    if not s or s == '-':
        return '-'
    # YYYY-MM-DD~DD 또는 YYYY-MM-DD ~ YYYY-MM-DD
    m_range = re.match(
        r'^(\d{4})-(\d{1,2})-(\d{1,2})\s*~\s*(?:(\d{4})-(\d{1,2})-)?(\d{1,2})',
        s,
    )
    if m_range:
        mm1, dd1 = int(m_range.group(2)), int(m_range.group(3))
        dd2 = int(m_range.group(6))
        return f'{mm1}월 {dd1}~{dd2}일'
    m = re.match(r'^(\d{4})-(\d{1,2})-(\d{1,2})', s)
    if m:
        return f'{int(m.group(2))}월 {int(m.group(3))}일'
    # 이미 사람 읽는 양식이면 그대로
    return s[:20]


def _visit_date_sort_key(raw) -> str:
    """방문 예정일 정렬 키 — ISO 앞부분 (YYYY-MM-DD)."""
    s = str(raw or '').strip().lstrip("'")
    m = re.match(r'^(\d{4}-\d{1,2}-\d{1,2})', s)
    return m.group(1) if m else s


def _render_item(lead: Dict, initial_map: Dict[str, str]) -> str:
    """캔버스 각 항목 텍스트 렌더.

    양식: (이니셜) MM월 DD일 / 연락처 / 주소 상호 / 내용
    """
    ini = _initial_from_name(str(lead.get('영업 담당자') or ''), initial_map)
    vd = _fmt_visit_date(lead.get('방문 예정일'))
    phone = str(lead.get('고객 연락처') or '').strip() or '-'
    address = str(lead.get('방문 주소') or '').strip() or '-'
    biz = str(lead.get('고객명') or '').strip()
    inquiry = str(lead.get('상담 내용') or lead.get('문의 내용') or '').strip()
    # 개행 flatten
    address = re.sub(r'\s*\n\s*', ' ', address)
    inquiry = re.sub(r'\s*\n\s*', ' ', inquiry)
    # 주소 뒤에 상호 붙이기 (기존 매니저 양식)
    addr_biz = f'{address} {biz}'.strip() if biz else address
    # 내용 길이 제한 (200자)
    if len(inquiry) > 200:
        inquiry = inquiry[:200] + '...'
    return f'({ini}) {vd} / {phone} / {addr_biz} / {inquiry}'


def _fetch_visit_leads() -> List[Dict]:
    """상태 = 방문 예약 + 방문 예정일 오늘 이후 lead 만 fetch.

    지난 방문 예정일은 완료 처리 안됐어도 목록에서 자동 제외 (매니저 카톡 원본 관행).
    """
    from dashboard.services.lead_service import load_leads_data
    df = load_leads_data(force_refresh=True)
    if df is None or df.empty:
        return []
    today = date.today()
    leads: List[Dict] = []
    for _, row in df.iterrows():
        status = str(row.get('상태') or '').strip()
        if status != '방문 예약':
            continue
        # 방문 예정일 파싱 — 오늘 이전이면 제외 (범위면 종료일 기준)
        vd_raw = str(row.get('방문 예정일') or '').strip().lstrip("'")
        vd_date: Optional[date] = None
        # `-`, `.`, `/` 구분자 혼용 허용 (매니저 표기 다양 — 2026-01.29~30 등)
        m_range = re.match(
            r'^(\d{4})[-./](\d{1,2})[-./](\d{1,2})\s*~\s*'
            r'(?:(\d{4})[-./](\d{1,2})[-./])?(\d{1,2})',
            vd_raw,
        )
        m_single = re.match(r'^(\d{4})[-./](\d{1,2})[-./](\d{1,2})', vd_raw)
        try:
            if m_range:
                y = int(m_range.group(4) or m_range.group(1))
                mo = int(m_range.group(5) or m_range.group(2))
                d = int(m_range.group(6))
                vd_date = date(y, mo, d)
            elif m_single:
                vd_date = date(
                    int(m_single.group(1)),
                    int(m_single.group(2)),
                    int(m_single.group(3)),
                )
        except ValueError:
            vd_date = None
        if not vd_date:
            continue  # 방문 예정일 미확정·파싱 실패 → 캔버스 제외 (매니저 관행)
        if vd_date < today:
            continue  # 지난 방문 → 완료 처리 안됐어도 목록 제외
        leads.append(row.to_dict())
    return leads


def build_canvas_markdown() -> str:
    """캔버스 전체 markdown 조립."""
    initial_map = _get_initial_map()
    leads = _fetch_visit_leads()

    buckets: Dict[str, List[Dict]] = {'기타': [], '거래처': [], '온라인 방문': []}
    for lead in leads:
        cat = _categorize(lead)
        if not cat:
            continue
        buckets[cat].append(lead)

    # 각 카테고리 방문일 오름차순
    for cat in buckets:
        buckets[cat].sort(key=lambda l: _visit_date_sort_key(l.get('방문 예정일')))

    lines: List[str] = []
    lines.append('# 방문 일정')
    lines.append('')
    lines.append('> 봇이 자동 관리 — 시트의 방문 예약 lead 를 실시간 반영합니다.')
    lines.append('> 방문 완료 처리되면 자동으로 목록에서 제거됩니다.')
    lines.append('')
    lines.append('**작성 요령**: `(이니셜) 날짜 / 연락처 / 주소 상호 / 내용`')
    lines.append('')

    category_titles = {
        '기타': '기타 (AS, 계약서작성 등)',
        '거래처': '거래처',
        '온라인 방문': '온라인 방문',
    }
    for cat in ('기타', '거래처', '온라인 방문'):
        title = category_titles[cat]
        items = buckets.get(cat, [])
        lines.append(f'## {title}')
        if not items:
            lines.append('_없음_')
        else:
            # 방문 시작 날짜 별 그룹핑 (이미 시간순 정렬 상태) — 날짜 바뀌면 빈 줄 삽입
            prev_key: Optional[str] = None
            for lead in items:
                cur_key = _visit_date_sort_key(lead.get('방문 예정일'))
                if prev_key is not None and cur_key != prev_key:
                    lines.append('')  # 날짜 구분 여백
                lines.append(f'- {_render_item(lead, initial_map)}')
                prev_key = cur_key
        lines.append('')

    now = datetime.now().strftime('%Y-%m-%d %H:%M')
    lines.append(f'_마지막 갱신 : {now}_')
    return '\n'.join(lines)


def rebuild_canvas() -> Dict:
    """캔버스 전체 replace. 결과 dict 반환."""
    result = {'ok': False, 'reason': ''}
    canvas_id = os.getenv('SLACK_VISIT_CANVAS_ID', '').strip()
    if not canvas_id:
        result['reason'] = 'SLACK_VISIT_CANVAS_ID 미설정'
        logger.debug(f'[VISIT_CANVAS] {result["reason"]}')
        return result
    # 방문 일정 봇 (canvases:write scope 필요) 우선, 없으면 기본 봇 fallback
    token = (
        os.getenv('SLACK_VISIT_BOT_TOKEN', '').strip()
        or os.getenv('SLACK_BOT_TOKEN', '').strip()
    )
    if not token:
        result['reason'] = 'SLACK_VISIT_BOT_TOKEN / SLACK_BOT_TOKEN 미설정'
        logger.warning(f'[VISIT_CANVAS] {result["reason"]}')
        return result

    with _CANVAS_LOCK:
        try:
            markdown = build_canvas_markdown()
        except Exception as exc:
            result['reason'] = f'markdown build 실패: {exc}'
            logger.error(f'[VISIT_CANVAS] {result["reason"]}', exc_info=True)
            return result

        payload = {
            'canvas_id': canvas_id,
            'changes': [{
                'operation': 'replace',
                'document_content': {
                    'type': 'markdown',
                    'markdown': markdown,
                },
            }],
        }
        try:
            resp = requests.post(
                'https://slack.com/api/canvases.edit',
                headers={
                    'Authorization': f'Bearer {token}',
                    'Content-Type': 'application/json; charset=utf-8',
                },
                json=payload,
                timeout=15,
            )
            data = resp.json()
        except Exception as exc:
            result['reason'] = f'API 호출 실패: {exc}'
            logger.error(f'[VISIT_CANVAS] {result["reason"]}', exc_info=True)
            return result

        if not data.get('ok'):
            result['reason'] = f'API 응답 오류: {data.get("error")}'
            logger.warning(f'[VISIT_CANVAS] {result["reason"]}')
            return result

        result['ok'] = True
        logger.info(f'[VISIT_CANVAS] 캔버스 갱신 완료 ({len(markdown)}자)')
        return result


def rebuild_canvas_async() -> None:
    """훅에서 부담 없이 호출할 background 실행."""
    def _bg():
        try:
            rebuild_canvas()
        except Exception as exc:
            logger.error(f'[VISIT_CANVAS] async rebuild 예외: {exc}', exc_info=True)
    threading.Thread(target=_bg, daemon=True).start()
