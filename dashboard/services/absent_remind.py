"""부재중/미완료/견적요청 리마인드 — 매일 아침 9시.

lead 중 다음 조건 대상을 온라인 문의 채널에 요약 카드로 발송:
  A. 상태 = '상담 대기' & 온라인 상담자 미배정 (매니저가 아예 놓친 것) — date_range 인입분
  B. 상태 = '부재중' & 영업 담당자 없음  (콜백했으나 미연결, 재연락 필요) — date_range 인입분
  C. 상태 = '견적 요청' (견적 미제출) — **날짜 무관 전체 스캔**, 제출·드랍될 때까지 매일 (2026-07-27)

카드 구성 (v7):
  ⠀
  :bell: *어제 미처리 문의 (N건) — 오늘 다시 연락 부탁드립니다*
  ─── SEP ───
  :speech_balloon: *미완료 (X건)*        # A 케이스
  • `lead_no` [플랫폼] 고객명 · 어제 HH:MM  |  <확인하기>
  ...
  :phone: *매니저 (INI) 부재중 (Y건)*    # B 케이스 (매니저별 그룹)
  • `lead_no` [플랫폼] 고객명 · 연락처  |  <확인하기>
  ...
  :receipt: *매니저 (INI) 견적 요청 (미제출) (Z건)*   # C 케이스 (매니저별 그룹, 접수일 병기)
  • `lead_no` [플랫폼] 고객명 · 07.25(토) HH:MM  |  <확인하기>
  ...
  ─── SEP ───
  :information_source: 부재중 2번 이상일 시 문의 드랍 처리 해주세요.
  :information_source: 견적 요청은 견적 제출·방문 예약·드랍 처리 전까지 매일 표시됩니다.
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
_WEEKDAY_KR = ['월', '화', '수', '목', '금', '토', '일']


def _hhmm(t: str) -> str:
    m = re.search(r'(\d{2}):(\d{2})', t or '')
    return f'{m.group(1)}:{m.group(2)}' if m else ''


def _md_weekday(d: date) -> str:
    """`07.26(일)` 형식 — 헤더·라인 date 병기용."""
    return f'{d.month:02d}.{d.day:02d}({_WEEKDAY_KR[d.weekday()]})'


def _disp_ini(ini: str) -> str:
    """이니셜 표시 정규화 — 이니셜 맵이 'KIKO' 로 주는 케이스를 'KiKO' 로 (사용자 예외 표기)."""
    return 'KiKO' if str(ini).upper() == 'KIKO' else ini


def _lead_date(l: Dict) -> Optional[date]:
    """lead 의 상담 시간 → date 파싱. `2026.07.26. HH:MM` 포맷."""
    s = str(l.get('상담 시간') or '')
    m = re.match(r'(\d{4})\.(\d{1,2})\.(\d{1,2})', s)
    if not m:
        return None
    try:
        return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    except ValueError:
        return None


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


def _is_business_day(d: date) -> bool:
    """평일 & 한국 공휴일 아님 → 영업일."""
    if d.weekday() >= 5:  # 토(5), 일(6)
        return False
    try:
        import holidays
        return d not in holidays.KR()
    except Exception:
        # holidays 미설치 or 오류 → 주말만 skip
        return True


def _previous_business_day(d: date) -> date:
    """d 이전 첫 영업일. (오늘이 화요일 → 월요일, 월요일 → 지난 금요일)"""
    prev = d - timedelta(days=1)
    while not _is_business_day(prev):
        prev = prev - timedelta(days=1)
    return prev


def collect_absent_leads(target_date: Optional[date] = None,
                          date_range: Optional[List[date]] = None) -> Tuple[List[Dict], Dict[str, List[Dict]], Dict[str, List[Dict]]]:
    """부재중 리마인드 대상 수집.

    Args:
        target_date: 단일 date (하위 호환)
        date_range: 복수 date 리스트 — 우선 사용. 각 date 인입 lead 다 포함.
                    (주말·공휴일 다음 영업일 리마인드에서 여러 날 잡기 위함)

    Returns:
        (unassigned, retry_by_manager, quote_pending_by_manager)
            unassigned: 상담 대기 & 온라인 상담자 미배정 (date_range 인입분)
            retry_by_manager: {매니저이름: [lead, ...]}  상태='부재중' & 영업 담당자 없음 (date_range 인입분)
            quote_pending_by_manager: {매니저이름: [lead, ...]}  상태='견적 요청' — 견적 제출 전까지
                **날짜 무관 전체 스캔** (며칠 걸릴 수 있어 제출·드랍될 때까지 매일 리마인드)
    """
    from dashboard.services.lead_service import get_lead_records
    if date_range is None:
        date_range = [target_date or (date.today() - timedelta(days=1))]
    ymd_dots = {d.strftime('%Y.%m.%d') for d in date_range}
    leads = get_lead_records()
    yday = [
        l for l in leads
        if any(str(l.get('상담 시간', '')).startswith(p) for p in ymd_dots)
    ]

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

    # 견적 요청 = 견적 제출 전까지 매일 리마인드 (날짜 무관 — 전체 lead 스캔).
    #   상태가 아직 '견적 요청' 이면 미제출. 제출/방문예약/드랍 시 상태가 바뀌어 자동 이탈.
    #   등록자(온라인 상담자)별 그룹 — 부재중과 동일 accountability.
    quote_pending: Dict[str, List[Dict]] = defaultdict(list)
    for l in leads:
        if str(l.get('상태', '')).strip().replace(' ', '') == '견적요청':
            consultant = str(l.get('온라인 상담자', '')).strip()
            quote_pending[consultant or '(미배정)'].append(l)

    return unassigned, dict(retry), dict(quote_pending)


def build_remind_text(unassigned: List[Dict], retry: Dict[str, List[Dict]],
                       client=None, channel: str = _ONLINE_CHANNEL_DEFAULT,
                       date_range: Optional[List[date]] = None,
                       quote_pending: Optional[Dict[str, List[Dict]]] = None) -> Tuple[str, int]:
    """리마인드 카드 텍스트 조립.

    Args:
        date_range: 리마인드 대상 date 리스트. 크기 별 헤더·라인 표기 분기.
                    - 1일: `어제 미처리 문의` + 라인 `어제 HH:MM`
                    - 2일 이상: `{start} ~ {end} 미처리 문의` + 라인 `MM.DD(요일) HH:MM`
        quote_pending: {매니저: [lead,...]} 상태='견적 요청' 미제출 (날짜 무관). 별도 섹션.

    Returns: (text, total_count)
    """
    from dashboard.services.visit_assignment_sync import _load_initial_maps
    _initial_to_name, name_to_initial = _load_initial_maps()
    quote_pending = quote_pending or {}

    # permalink 조회 최적화 — 채널 history 1번만 fetch
    history_cache = None
    if client is not None:
        try:
            h = client.conversations_history(channel=channel, limit=500)
            history_cache = h.get('messages') or []
        except Exception as exc:
            logger.debug(f'[ABSENT] history fetch 실패: {exc}')

    total = (len(unassigned) + sum(len(v) for v in retry.values())
             + sum(len(v) for v in quote_pending.values()))
    _range = sorted(date_range or [])
    _multi_day = len(_range) >= 2

    def _line(l: Dict, mode: str) -> str:
        lno = str(l.get('리드 No', '')).strip()
        pl = _find_card_permalink(client, channel, lno, history_cache) if client else ''
        link = f'  |  <{pl}|확인하기>' if pl else ''
        name = str(l.get('고객명', ''))[:20]
        plat = str(l.get('플랫폼', ''))
        if mode == 'unassigned':
            t = _hhmm(str(l.get('상담 시간', '')))
            # range 2일 이상이면 date 병기 (07.26(일) 12:39), 1일이면 `어제 HH:MM`
            if _multi_day:
                _d = _lead_date(l)
                _date_tag = _md_weekday(_d) if _d else '어제'
                return f'• `{lno}` [{plat}] {name} · {_date_tag} {t}{link}'
            return f'• `{lno}` [{plat}] {name} · 어제 {t}{link}'
        if mode == 'quote':
            # 날짜 무관 스캔 → 접수일 병기 (얼마나 대기 중인지 파악)
            t = _hhmm(str(l.get('상담 시간', '')))
            _d = _lead_date(l)
            _when = (f'{_md_weekday(_d)} {t}' if _d else t).strip() or '-'
            return f'• `{lno}` [{plat}] {name} · {_when}{link}'
        return f'• `{lno}` [{plat}] {name} · {l.get("고객 연락처", "")}{link}'

    # 헤더 문구 — range 크기별 분기
    if _multi_day:
        _hdr_range = f'{_md_weekday(_range[0])} ~ {_md_weekday(_range[-1])}'
        _hdr = f':bell: *{_hdr_range} 미처리 문의 ({total}건) — 오늘 다시 연락 부탁드립니다*'
    else:
        _hdr = f':bell: *어제 미처리 문의 ({total}건) — 오늘 다시 연락 부탁드립니다*'

    lines = [_BLANK]
    lines.append(_hdr)
    lines.append(_SEP)
    if unassigned:
        lines.append(f':speech_balloon: *미완료 ({len(unassigned)}건)*')
        for l in unassigned:
            lines.append(_line(l, 'unassigned'))
        lines.append('')
    for mgr_name, items in retry.items():
        ini = _disp_ini(name_to_initial.get(mgr_name, '?'))
        lines.append(f':phone: *{mgr_name} ({ini}) 부재중 ({len(items)}건)*')
        for l in items:
            lines.append(_line(l, 'retry'))
        lines.append('')
    # 견적 요청 (미제출) — 매니저별 그룹, 부재중 섹션 뒤에
    for mgr_name, items in quote_pending.items():
        ini = _disp_ini(name_to_initial.get(mgr_name, '?'))
        _mgr_label = f'{mgr_name} ({ini})' if mgr_name != '(미배정)' else '(미배정)'
        lines.append(f':receipt: *{_mgr_label} 견적 요청 (미제출) ({len(items)}건)*')
        for l in items:
            lines.append(_line(l, 'quote'))
        lines.append('')
    while lines and lines[-1] == '':
        lines.pop()
    lines.append(_SEP)
    lines.append(':information_source: 부재중 2번 이상일 시 문의 드랍 처리 해주세요.')
    if quote_pending:
        lines.append(':information_source: 견적 요청은 견적 제출·방문 예약·드랍 처리 전까지 매일 표시됩니다.')
    lines.append(_BLANK)
    return '\n'.join(lines), total


def send_daily_remind() -> Dict:
    """어제 미처리 문의 리마인드 카드 발송 — 매일 아침 9시 스케줄러 진입점.

    2026-07-26 주말·공휴일 skip + 직전 영업일 이후 range 수집:
      - 오늘이 주말·공휴일 → 발송 skip (매니저 대응 불가)
      - 오늘이 영업일 → 직전 영업일 다음날부터 어제까지 모든 date 대상
        (예: 월요일 아침 → 금·토·일 3일치 미처리 잡음)

    Returns:
        {'ok': bool, 'total': int, 'ts': str or '', 'reason': str or None}
    """
    today = date.today()
    if not _is_business_day(today):
        logger.info(f'[ABSENT] 비영업일 ({today.strftime("%Y-%m-%d %a")}) — 리마인드 skip')
        return {'ok': True, 'total': 0, 'ts': '', 'reason': 'non_business_day'}

    # 직전 영업일 계산 → 그 다음날부터 어제까지가 리마인드 대상 date range.
    #   예: today=화 → prev=월, range=[월] (하루)
    #       today=월 → prev=금, range=[토, 일] (주말 인입)
    #       today=목(수요일이 공휴일) → prev=화, range=[수] (공휴일 인입)
    prev_bday = _previous_business_day(today)
    date_range: List[date] = []
    d = prev_bday + timedelta(days=1)
    while d < today:
        date_range.append(d)
        d = d + timedelta(days=1)
    # date_range 가 비어있으면 (연속 영업일 화·수 등) 어제 하나만
    if not date_range:
        date_range = [today - timedelta(days=1)]

    channel = os.getenv('SLACK_ONLINE_CHANNEL', _ONLINE_CHANNEL_DEFAULT).strip() or _ONLINE_CHANNEL_DEFAULT
    client = _get_online_client()
    if not client:
        return {'ok': False, 'total': 0, 'ts': '', 'reason': 'SLACK_BOT_TOKEN 미설정'}

    unassigned, retry, quote_pending = collect_absent_leads(date_range=date_range)
    total = (len(unassigned) + sum(len(v) for v in retry.values())
             + sum(len(v) for v in quote_pending.values()))
    if total == 0:
        logger.info('[ABSENT] 미처리 문의 0건 — 카드 발송 skip')
        return {'ok': True, 'total': 0, 'ts': '', 'reason': None}

    text, _ = build_remind_text(unassigned, retry, client=client, channel=channel,
                                  date_range=date_range, quote_pending=quote_pending)
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
