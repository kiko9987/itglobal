"""부재중/미완료 리마인드 — 매일 아침 9시.

어제 인입된 lead 중 다음 조건 대상을 온라인 문의 채널에 요약 카드로 발송:
  A. 상태 = '상담 대기' & 온라인 상담자 미배정 (매니저가 아예 놓친 것)
  B. 상태 = '부재중' & 영업 담당자 없음  (콜백했으나 미연결, 재연락 필요)

카드 구성 (v6):
  ⠀
  :bell: *어제 미처리 문의 (N건) — 오늘 다시 연락 부탁드립니다*
  ─── SEP ───
  :speech_balloon: *미완료 (X건)*        # A 케이스
  • `lead_no` [플랫폼] 고객명 · 어제 HH:MM  |  <확인하기>
  ...
  :phone: *매니저 (INI) 부재중 (Y건)*    # B 케이스 (매니저별 그룹)
  • `lead_no` [플랫폼] 고객명 · 연락처  |  <확인하기>
  ...
  ─── SEP ───
  :information_source: 부재중 2번 이상일 시 문의 드랍 처리 해주세요.
  ⠀
"""
from __future__ import annotations

import logging
import os
import re
from collections import defaultdict
from datetime import date, timedelta
from typing import Dict, List, Optional, Tuple

logger = logging.getLogger(__name__)

_ONLINE_CHANNEL_DEFAULT = 'C0BB9SRMEA1'
_SEP = '--------------------------------------------'
_BLANK = '⠀'


def _hhmm(t: str) -> str:
    m = re.search(r'(\d{2}):(\d{2})', t or '')
    return f'{m.group(1)}:{m.group(2)}' if m else ''


def _get_online_client():
    """온라인 채널 접근용 client — SLACK_BOT_TOKEN (online_bot)."""
    from slack_sdk import WebClient
    tok = os.environ.get('SLACK_BOT_TOKEN', '').strip()
    if not tok:
        return None
    return WebClient(token=tok)


def _find_card_permalink(client, channel: str, lead_no: str,
                         history_cache: Optional[List] = None) -> str:
    """lead_no 의 슬랙 카드 permalink 조회.

    (1) Redis lead_card_msg:{lno} → 있으면 ts 로 permalink API
    (2) 없으면 채널 history 500건 뒤로 lead_no substring 매칭
    """
    try:
        from dashboard.utils.redis_client import get_redis_client
        rc = get_redis_client().redis
        v = rc.get(f'lead_card_msg:{lead_no}')
    except Exception:
        v = None
    ch, ts = None, None
    if v and '|' in (v if isinstance(v, str) else v.decode()):
        raw = v if isinstance(v, str) else v.decode()
        ch, ts = raw.split('|', 1)
    else:
        if history_cache is None:
            return ''
        for m in history_cache:
            text = m.get('text', '') + str(m.get('blocks', ''))
            if lead_no in text and m.get('thread_ts', m['ts']) == m['ts']:
                ch, ts = channel, m['ts']
                break
    if not ch or not ts:
        return ''
    try:
        r = client.chat_getPermalink(channel=ch, message_ts=ts)
        return r.get('permalink', '') or ''
    except Exception as exc:
        logger.debug(f'[ABSENT] permalink 실패 {lead_no}: {exc}')
        return ''


def collect_absent_leads(target_date: Optional[date] = None) -> Tuple[List[Dict], Dict[str, List[Dict]]]:
    """부재중 리마인드 대상 수집.

    Args:
        target_date: 인입 기준일 (기본: 어제)

    Returns:
        (unassigned, retry_by_manager)
            unassigned: 상담 대기 & 온라인 상담자 미배정
            retry_by_manager: {매니저이름: [lead, ...]}  상태='부재중' & 영업 담당자 없음
    """
    from dashboard.services.lead_service import get_lead_records
    target_date = target_date or (date.today() - timedelta(days=1))
    ymd_dot = target_date.strftime('%Y.%m.%d')
    leads = get_lead_records()
    yday = [l for l in leads if str(l.get('상담 시간', '')).startswith(ymd_dot)]

    unassigned: List[Dict] = []
    retry: Dict[str, List[Dict]] = defaultdict(list)
    for l in yday:
        status = str(l.get('상태', '')).strip()
        consultant = str(l.get('온라인 상담자', '')).strip()
        sales = str(l.get('영업 담당자', '')).strip()
        if status == '상담 대기' and not consultant:
            unassigned.append(l)
        elif status == '부재중' and not sales:
            retry[consultant or '(미배정)'].append(l)
    return unassigned, dict(retry)


def build_remind_text(unassigned: List[Dict], retry: Dict[str, List[Dict]],
                       client=None, channel: str = _ONLINE_CHANNEL_DEFAULT) -> Tuple[str, int]:
    """리마인드 카드 텍스트 조립.

    Returns: (text, total_count)
    """
    from dashboard.services.visit_assignment_sync import _load_initial_maps
    _initial_to_name, name_to_initial = _load_initial_maps()

    # permalink 조회 최적화 — 채널 history 1번만 fetch
    history_cache = None
    if client is not None:
        try:
            h = client.conversations_history(channel=channel, limit=500)
            history_cache = h.get('messages') or []
        except Exception as exc:
            logger.debug(f'[ABSENT] history fetch 실패: {exc}')

    total = len(unassigned) + sum(len(v) for v in retry.values())

    def _line(l: Dict, mode: str) -> str:
        lno = str(l.get('리드 No', '')).strip()
        pl = _find_card_permalink(client, channel, lno, history_cache) if client else ''
        link = f'  |  <{pl}|확인하기>' if pl else ''
        name = str(l.get('고객명', ''))[:20]
        plat = str(l.get('플랫폼', ''))
        if mode == 'unassigned':
            t = _hhmm(str(l.get('상담 시간', '')))
            return f'• `{lno}` [{plat}] {name} · 어제 {t}{link}'
        return f'• `{lno}` [{plat}] {name} · {l.get("고객 연락처", "")}{link}'

    lines = [_BLANK]
    lines.append(f':bell: *어제 미처리 문의 ({total}건) — 오늘 다시 연락 부탁드립니다*')
    lines.append(_SEP)
    if unassigned:
        lines.append(f':speech_balloon: *미완료 ({len(unassigned)}건)*')
        for l in unassigned:
            lines.append(_line(l, 'unassigned'))
        lines.append('')
    for mgr_name, items in retry.items():
        ini = name_to_initial.get(mgr_name, '?')
        lines.append(f':phone: *{mgr_name} ({ini}) 부재중 ({len(items)}건)*')
        for l in items:
            lines.append(_line(l, 'retry'))
        lines.append('')
    while lines and lines[-1] == '':
        lines.pop()
    lines.append(_SEP)
    lines.append(':information_source: 부재중 2번 이상일 시 문의 드랍 처리 해주세요.')
    lines.append(_BLANK)
    return '\n'.join(lines), total


def send_daily_remind() -> Dict:
    """어제 미처리 문의 리마인드 카드 발송 — 매일 아침 9시 스케줄러 진입점.

    Returns:
        {'ok': bool, 'total': int, 'ts': str or '', 'reason': str or None}
    """
    channel = os.getenv('SLACK_ONLINE_CHANNEL', _ONLINE_CHANNEL_DEFAULT).strip() or _ONLINE_CHANNEL_DEFAULT
    client = _get_online_client()
    if not client:
        return {'ok': False, 'total': 0, 'ts': '', 'reason': 'SLACK_BOT_TOKEN 미설정'}

    unassigned, retry = collect_absent_leads()
    total = len(unassigned) + sum(len(v) for v in retry.values())
    if total == 0:
        logger.info('[ABSENT] 어제 미처리 문의 0건 — 카드 발송 skip')
        return {'ok': True, 'total': 0, 'ts': '', 'reason': None}

    text, _ = build_remind_text(unassigned, retry, client=client, channel=channel)
    try:
        r = client.chat_postMessage(
            channel=channel, text=text,
            unfurl_links=False, unfurl_media=False,
        )
        if r.get('ok'):
            logger.info(f'[ABSENT] 리마인드 카드 발송 완료 (total={total}, ts={r.get("ts")})')
            return {'ok': True, 'total': total, 'ts': r.get('ts', ''), 'reason': None}
        return {'ok': False, 'total': total, 'ts': '', 'reason': r.get('error', 'unknown')}
    except Exception as exc:
        logger.error(f'[ABSENT] 리마인드 발송 예외: {exc}', exc_info=True)
        return {'ok': False, 'total': total, 'ts': '', 'reason': str(exc)}
