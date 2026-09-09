"""부재중/미완료/견적요청 리마인드 — 매일 아침 9시.

lead 중 다음 조건 대상을 온라인 문의 채널에 요약 카드로 발송:
  A. 상태 = '상담 대기' & 온라인 상담자 미배정 (매니저가 아예 놓친 것) — 최근 N영업일 (배정 전까지)
  B. 상태 = '부재중' & 영업 담당자 없음  (콜백했으나 미연결, 재연락 필요) — 최근 N영업일 (재시도/드랍 전까지)
  C. 상태 = '견적 요청' (견적 미제출) — **날짜 무관 전체 스캔**, 제출·드랍될 때까지 매일 (2026-07-27)
  A·B 는 _LOOKBACK_BDAYS(2) 영업일 하한(cutoff)까지 지속 (2026-09-09, 하루만 뜨고 사라지던 누락 해소).

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


# 미완료·부재중 지속 리마인드 범위 (최근 N영업일). 처리(배정·재시도 성공·드랍) 전까지
# 매일 표시하되, 오래된 미처리 백로그(예: 5~7월 부재중 47건, 1~5월 상담대기 등)가 홍수처럼
# 딸려오지 않게 하한을 둔다.
# 2026-09-09: 미완료·부재중이 직전영업일 range(=1영업일)만 봐서 다음날 처리 안 하면 사라지던
#   누락 해소 (부재중 L-03938 백상현 09-07 계기). 원래 1영업일 → 2영업일로 확장
#   (사용자 결정: '2일 정도면 됨'. 미완료도 동일 적용).
_LOOKBACK_BDAYS = 2


def _nth_previous_business_day(d: date, n: int) -> date:
    """d 로부터 n 영업일 전 date (n>=1). 지속 리마인드의 하한(cutoff)."""
    cur = d
    for _ in range(max(1, n)):
        cur = _previous_business_day(cur)
    return cur


def _reminder_ordinal(lead_date: date, today: date) -> int:
    """접수일 이후 오늘까지 '몇 번째 영업일 알림'인가. 1=첫 알림, 2+=재알림.
    (lead_date, today] 사이 영업일 수. 접수 다음 영업일=1, 그 다음=2 …
    주말·공휴일은 알림이 안 나가므로 자동 제외 → 월요일이 금요일 건의 '1일차'로 정확히 계산됨.
    """
    n = 0
    d = today
    while d > lead_date:
        if _is_business_day(d):
            n += 1
        d = d - timedelta(days=1)
    return n


def collect_absent_leads(target_date: Optional[date] = None,
                          date_range: Optional[List[date]] = None) -> Tuple[List[Dict], Dict[str, List[Dict]], Dict[str, List[Dict]]]:
    """부재중 리마인드 대상 수집.

    Args:
        target_date: 단일 date (하위 호환, 미사용 — 수집은 cutoff 기반)
        date_range: 하위 호환용 (미사용). 미완료·부재중은 최근 _LOOKBACK_BDAYS 영업일
                    cutoff 로 수집하고, 견적요청은 날짜 무관 스캔한다.

    Returns:
        (unassigned, retry_by_manager, quote_pending_by_manager)
            unassigned: 상담 대기 & 온라인 상담자 미배정 (최근 N영업일)
            retry_by_manager: {매니저이름: [lead, ...]}  상태='부재중' & 영업 담당자 없음 (최근 N영업일)
            quote_pending_by_manager: {매니저이름: [lead, ...]}  상태='견적 요청' — 견적 제출 전까지
                **날짜 무관 전체 스캔** (며칠 걸릴 수 있어 제출·드랍될 때까지 매일 리마인드)
    """
    from dashboard.services.lead_service import get_lead_records
    if date_range is None:
        date_range = [target_date or (date.today() - timedelta(days=1))]
    leads = get_lead_records()

    # 미완료·부재중 공통 하한(cutoff) — 최근 _LOOKBACK_BDAYS 영업일. 이 날짜 이후 인입분만.
    #   date_range(직전영업일)만 보던 옛 방식은 다음날 처리 안 하면 사라져 누락됐음.
    cutoff = _nth_previous_business_day(date.today(), _LOOKBACK_BDAYS)

    def _consultant(l):
        c = str(l.get('온라인 상담자', '')).strip()
        return '' if c == '-' else c  # '-' 플레이스홀더 = 미배정 (빈값과 동일)

    def _sales(l):
        s = str(l.get('영업 담당자', '')).strip()
        return '' if s == '-' else s

    def _recent(l):
        ld = _lead_date(l)
        return ld is not None and ld >= cutoff

    # A. 미완료 (상태='상담 대기' & 온라인 상담자 미배정) — 배정 전까지 매일 (최근 N영업일).
    #   2026-09-09: '-' 플레이스홀더 정규화(큐플레이스 김시현 누락 사고) + date_range→cutoff 전환
    #   (부재중과 동일하게, 배정 안 하면 다음날 사라지던 문제 해소. 사용자 결정 '2일 기준').
    unassigned: List[Dict] = []
    for l in leads:
        if str(l.get('상태', '')).strip() == '상담 대기' and not _consultant(l) and _recent(l):
            unassigned.append(l)

    # B. 부재중 (상태='부재중' & 영업 담당자 미배정) — 재시도(성공)·드랍 전까지 매일 (최근 N영업일).
    #   2026-09-09: date_range(직전영업일)만 보면 다음날 재시도 안 하면 사라져 누락됨
    #   (L-03938 백상현 09-07 건). 견적요청처럼 지속 스캔하되, 오래된 백로그(5~7월 47건)
    #   홍수 방지 위해 cutoff(최근 _LOOKBACK_BDAYS 영업일) 제한.
    retry: Dict[str, List[Dict]] = defaultdict(list)
    for l in leads:
        if str(l.get('상태', '')).strip() == '부재중' and not _sales(l) and _recent(l):
            retry[_consultant(l) or '(미배정)'].append(l)

    # C. 견적 요청 = 견적 제출 전까지 매일 리마인드 (날짜 무관 — 전체 lead 스캔).
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
        date_range: 하위 호환용 (미사용). 헤더는 '최근 미처리 문의', 각 라인은 접수일
                    (MM.DD(요일) HH:MM) 병기 — 미완료·부재중이 최근 N영업일 창이라.
        quote_pending: {매니저: [lead,...]} 상태='견적 요청' 미제출 (날짜 무관). 별도 섹션.

    Returns: (text, total_count)
    """
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

    def _line(l: Dict, mode: str) -> str:
        lno = str(l.get('리드 No', '')).strip()
        pl = _find_card_permalink(client, channel, lno, history_cache) if client else ''
        link = f'  |  <{pl}|확인하기>' if pl else ''
        name = str(l.get('고객명', ''))[:20]
        plat = str(l.get('플랫폼', ''))
        _d = _lead_date(l)
        # 재알림 여부 — 1일차 알림에도 처리 안 돼 2영업일째 뜨는 건은 강조 (2026-09-09 사용자 요청).
        _realert = _d is not None and _reminder_ordinal(_d, date.today()) >= 2
        if mode in ('unassigned', 'quote'):
            # 미완료·견적요청: 접수일 병기. 최근 N영업일 창이라 며칠 전 건일 수 있어
            #   '어제' 고정 대신 실제 접수일(MM.DD(요일) HH:MM)로 표기 (2026-09-09).
            t = _hhmm(str(l.get('상담 시간', '')))
            _when = (f'{_md_weekday(_d)} {t}' if _d else t).strip() or '-'
            # 미완료 재알림: 어제 알림에도 미배정 → 강조 (견적요청은 원래 지속이라 제외).
            _esc = ' · :warning: *접수 후 아직 미배정*' if (mode == 'unassigned' and _realert) else ''
            return f'• `{lno}` [{plat}] {name} · {_when}{_esc}{link}'
        # 부재중 — 연락처 표기. 재알림이면 접수일 + 재시도 없음 강조.
        _esc = f' · :warning: *{_md_weekday(_d)} 부재중 접수 후 재시도 없음*' if _realert else ''
        return f'• `{lno}` [{plat}] {name} · {l.get("고객 연락처", "")}{_esc}{link}'

    # 헤더 — 미완료·부재중이 최근 N영업일 창이라 '어제' 고정 대신 '최근' (각 라인에 접수일 표기).
    _hdr = f':bell: *최근 미처리 문의 ({total}건) — 오늘 다시 연락 부탁드립니다*'

    lines = [_BLANK]
    lines.append(_hdr)
    lines.append(_SEP)
    if unassigned:
        lines.append(f':speech_balloon: *미완료 ({len(unassigned)}건)*')
        for l in unassigned:
            lines.append(_line(l, 'unassigned'))
        lines.append('')
    # 부재중 — 담당자 무관 카테고리 단위 (그냥 알림이라 이름/이니셜 그룹 제거, 2026-08-06).
    #   책임 매니저는 카드 permalink(확인하기)에서 확인 가능 → 헤더 노이즈만 제거.
    _retry_all = [l for items in retry.values() for l in items]
    if _retry_all:
        lines.append(f':phone: *부재중 ({len(_retry_all)}건)*')
        for l in _retry_all:
            lines.append(_line(l, 'retry'))
        lines.append('')
    # 견적 요청 (미제출) — 카테고리 단위 (부재중 섹션 뒤), 날짜 무관 스캔.
    _quote_all = [l for items in quote_pending.values() for l in items]
    if _quote_all:
        lines.append(f':receipt: *견적 요청 (미제출) ({len(_quote_all)}건)*')
        for l in _quote_all:
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
