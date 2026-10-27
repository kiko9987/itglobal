# -*- coding: utf-8 -*-
"""방문 일정 ↔ 캔버스1·캔버스2 누락 점검 (2026-10-27 사장님 요청).

① 오후 4시(월~금, 공휴일 제외): #방문_일정에 진행 중인 카드([방문 완료] 버튼이 살아 있는
   카드)가 있는데 캔버스1에 없는 건 → 사장님 DM (있을 때만)
② /일정확정 직후: ①과 같은 점검 + 캔버스1에 있는데 일정확정 대상일 방문인데도 캔버스2에
   안 옮긴 건 + 캔버스2에서 리드를 못 찾은 줄 → 실행한 사람 화면 + 사장님 DM (있을 때만)

배경 — 사람이 매번 대조해서 찾던 사고들:
  L-04341: 방문일 '10-24~20' 오입력(끝<시작) → '지난 방문'으로 캔버스1 제외
  L-04364·L-04368: 일정확정 뒤에 등록 → 캔버스2 누락, 담당 DM 없음
  ETC-587764: 캔버스2 번호 공백 '010- 4327-1111' → 매칭 실패 → 담당 DM 누락
  L-04289: 카드는 진행 중인데 시트는 '방문 완료'
"""
import html as _html
import json
import logging
import os
import re
import time
import urllib.parse
import urllib.request
from datetime import date, datetime
from typing import Dict, List, Optional

logger = logging.getLogger(__name__)

_ACTIVE_STATUS = ('방문 예약', '공사 확정')
# 막 등록된 카드는 캔버스1 재생성(10분 주기) 전일 수 있어 '캔버스1에 없음' 판정에서 제외
_FRESH_CARD_SEC = 15 * 60
_ADDR_KEY_RE = re.compile(r'[가-힣0-9]+(?:로|길)\s*\d+(?:-\d+)?')
_CANVAS_DATE_RE = re.compile(r'\d{1,2}\s*월\s*\d{1,2}')


def _digits(v) -> str:
    return re.sub(r'\D', '', str(v or ''))


# ── 판단 (순수 함수) ─────────────────────────────────────────────────────────

def lead_in_canvas_lines(lead: Dict, lines: List[str]) -> bool:
    """캔버스 줄들에 이 리드가 있는지 — 전화 뒤 8자리, 없으면 주소(도로명+번지)로 판단.

    거래처/기타 줄은 상호를 표시하지 않아 이름으로는 못 찾음(ETC-587764 오해 계기).
    """
    ph = _digits(lead.get('고객 연락처'))
    if len(ph) >= 9 and any(ph[-8:] in _digits(x) for x in lines):
        return True
    m = _ADDR_KEY_RE.search(str(lead.get('방문 주소') or ''))
    return bool(m and any(m.group(0) in x for x in lines))


def classify_card(lead: Optional[Dict], in_canvas1: bool, today: date) -> Optional[str]:
    """진행 중 방문 카드 하나 → 캔버스1에 없을 사유(사람이 읽는 문장). 문제 없으면 None."""
    from dashboard.services.visit_assignment_sync import (
        _parse_visit_date_end, _parse_visit_date_start,
    )
    if lead is None:
        return '시트에 리드가 없음'
    status = str(lead.get('상태') or '').strip()
    if status not in _ACTIVE_STATUS:
        return f"카드는 진행 중인데 시트 상태는 '{status or '빈칸'}'"
    vd = str(lead.get('방문 예정일') or '').strip().lstrip("'")
    start, end = _parse_visit_date_start(vd), _parse_visit_date_end(vd)
    if not start or not end:
        return f"방문일을 읽을 수 없음('{vd or '빈칸'}')"
    if end < start:
        return f"방문일 끝 날짜가 시작보다 앞('{vd}')"
    if end < today:
        return f"방문일({vd})이 지났는데 완료·취소 처리 안 됨"
    if not in_canvas1:
        return f"캔버스1에 없음(방문일 {vd})"
    return None


def canvas2_missing(c1_leads: List[Dict], matched: set, target: date) -> List[Dict]:
    """캔버스1(방문 예약) 중 방문 기간에 target 일이 들어가는데 캔버스2에 매칭되지 않은 건."""
    from dashboard.services.visit_assignment_sync import (
        _parse_visit_date_end, _parse_visit_date_start,
    )
    out = []
    for lead in c1_leads:
        ln = str(lead.get('리드 No') or '').strip()
        if not ln or ln in matched:
            continue
        vd = str(lead.get('방문 예정일') or '').strip().lstrip("'")
        start = _parse_visit_date_start(vd)
        end = _parse_visit_date_end(vd) or start
        if start and end and start <= target <= end:
            out.append(lead)
    return out


# ── 슬랙·시트 조회 ───────────────────────────────────────────────────────────

def _visit_token() -> str:
    return (os.getenv('SLACK_VISIT_BOT_TOKEN', '').strip()
            or os.getenv('SLACK_BOT_TOKEN', '').strip())


def _slack_get(method: str, params: Dict) -> Dict:
    req = urllib.request.Request(
        f'https://slack.com/api/{method}?{urllib.parse.urlencode(params)}',
        headers={'Authorization': f'Bearer {_visit_token()}'})
    return json.loads(urllib.request.urlopen(req, timeout=20).read().decode('utf-8'))


def active_visit_cards(days: int = 60) -> List[Dict]:
    """#방문_일정 최근 N일 봇 카드 중 [방문 완료] 버튼이 살아 있는(=진행 중) 카드. 리드당 최신 1장."""
    channel = os.getenv('SLACK_VISIT_CHANNEL', '').strip()
    oldest = f'{time.time() - days * 86400:.0f}'
    out, seen, cursor = [], set(), None
    while True:
        params = {'channel': channel, 'oldest': oldest, 'limit': 200}
        if cursor:
            params['cursor'] = cursor
        r = _slack_get('conversations.history', params)
        if not r.get('ok'):
            raise RuntimeError(f"#방문_일정 조회 실패: {r.get('error')}")
        for m in r.get('messages', []):          # 최신순
            if not m.get('bot_id'):
                continue
            for b in m.get('blocks') or []:
                if b.get('type') != 'actions':
                    continue
                for el in b.get('elements') or []:
                    ln = str(el.get('value') or '').strip()
                    if el.get('action_id') == 'visit_complete' and ln and ln not in seen:
                        seen.add(ln)
                        out.append({'lead_no': ln, 'ts': m.get('ts', '')})
        cursor = (r.get('response_metadata') or {}).get('next_cursor')
        if not cursor:
            break
    return out


def _canvas1_lines() -> List[str]:
    from dashboard.services.visit_assignment_sync import _fetch_canvas_html
    raw = _fetch_canvas_html(_visit_token(), os.getenv('SLACK_VISIT_CANVAS_ID', '').strip())
    if raw is None:
        raise RuntimeError('캔버스1 읽기 실패')
    text = _html.unescape(re.sub(r'<[^>]+>', '\n', raw))
    return [x.strip() for x in text.split('\n') if _CANVAS_DATE_RE.search(x)]


def check_cards_vs_canvas1(today: Optional[date] = None) -> List[Dict]:
    """진행 중 방문 카드 중 캔버스1에 없는 건(사유 포함)."""
    from dashboard.services.lead_service import load_leads_data
    today = today or date.today()
    cards = active_visit_cards()
    lines = _canvas1_lines()
    df = load_leads_data(force_refresh=True)
    if df is None or df.empty:
        raise RuntimeError('리드 시트 읽기 실패')
    sheet = {str(r.get('리드 No') or '').strip(): r for r in df.to_dict('records')}
    now = time.time()
    out = []
    for c in cards:
        lead = sheet.get(c['lead_no'])
        in_c1 = lead is not None and lead_in_canvas_lines(lead, lines)
        reason = classify_card(lead, in_c1, today)
        if not reason:
            continue
        if reason.startswith('캔버스1에 없음') and now - float(c['ts'] or 0) < _FRESH_CARD_SEC:
            continue
        out.append({'lead_no': c['lead_no'],
                    'name': str((lead or {}).get('고객명') or '').strip(),
                    'reason': reason})
    return sorted(out, key=lambda x: x['lead_no'])


def audit_after_commit(commit_result: Optional[Dict]) -> Dict:
    """일정확정 직후 점검 — 카드↔캔버스1, 캔버스1↔캔버스2(대상일), 캔버스2 매칭 실패 줄."""
    from dashboard.services import visit_assignment_sync as vas
    from dashboard.services.visit_canvas_sync import _fetch_visit_leads
    out = {'cards': [], 'c2_missing': [], 'c2_unmatched': [], 'target': None}
    out['cards'] = check_cards_vs_canvas1()
    dr = vas.dry_run()
    if not dr.get('ok'):
        raise RuntimeError(f"캔버스2 읽기 실패: {dr.get('reason')}")
    rows = dr.get('rows', [])
    out['c2_unmatched'] = [r for r in rows if not r.get('matched')]
    target_s = (((commit_result or {}).get('dm') or {}).get('target_date')
                or dr.get('target_date'))
    if target_s:
        out['target'] = target_s
        matched = {r.get('lead_no') for r in rows if r.get('matched')}
        for lead in canvas2_missing(_fetch_visit_leads(), matched, date.fromisoformat(target_s)):
            mgr = str(lead.get('영업 담당자') or '').strip()
            out['c2_missing'].append({
                'lead_no': str(lead.get('리드 No') or '').strip(),
                'name': str(lead.get('고객명') or '').strip(),
                'addr': str(lead.get('방문 주소') or '').strip(),
                'visit': str(lead.get('방문 예정일') or '').strip().lstrip("'"),
                'mgr': '' if mgr in ('', '-') else mgr,
            })
    return out


# ── 메시지 ───────────────────────────────────────────────────────────────────

def _card_line(x: Dict) -> str:
    return f"• `{x['lead_no']}` {x.get('name') or '-'} — {x['reason']}"


def format_daily(items: List[Dict], now: Optional[datetime] = None) -> str:
    now = now or datetime.now()
    lines = [f':mag: *방문 일정 점검* ({now.month}월 {now.day}일 {now:%H:%M})',
             f'#방문_일정에 진행 중인 카드는 있는데 캔버스1에 없는 건 {len(items)}건이에요.']
    lines += [_card_line(x) for x in items]
    return '\n'.join(lines)


def format_commit_audit(au: Dict) -> str:
    """일정확정 결과 화면·DM 에 붙일 '확인 필요' 문단. 확인할 게 없으면 빈 문자열."""
    if not (au.get('cards') or au.get('c2_missing') or au.get('c2_unmatched')):
        return ''
    lines = [':warning: *확인 필요*']
    if au.get('cards'):
        lines.append(f"*캔버스1에 없는 방문 카드 {len(au['cards'])}건*")
        lines += [_card_line(x) for x in au['cards']]
    if au.get('c2_missing'):
        t = date.fromisoformat(au['target'])
        lines.append(f"*캔버스1에는 있는데 캔버스2에 안 옮긴 {t.month}월 {t.day}일 방문 "
                     f"{len(au['c2_missing'])}건* — 담당 DM이 나가지 않았어요")
        for x in au['c2_missing']:
            who = f"담당 {x['mgr']}(이전 배정)" if x['mgr'] else '담당 미정'
            lines.append(f"• `{x['lead_no']}` {x['name'] or '-'} · {x['addr'][:30]} "
                         f"· 방문일 {x['visit']} · {who}")
    if au.get('c2_unmatched'):
        lines.append(f"*캔버스2에서 리드를 못 찾은 줄 {len(au['c2_unmatched'])}건* "
                     "— 전화번호·주소가 시트와 같은지 확인해 주세요")
        for r in au['c2_unmatched']:
            raw = str(r.get('raw') or '').replace(':man-raising-hand:', '').strip()
            lines.append(f"• {raw[:100] or r.get('phone') or '-'}")
    return '\n'.join(lines)


def send_admin_dm(text: str) -> bool:
    """사장님(관리자) DM — 방문 일정 알림 봇으로 발송."""
    channel = (os.getenv('VISIT_AUDIT_ALERT_CHANNEL', '').strip()
               or os.getenv('SLACK_ADMIN_CHANNEL', '').strip())
    if not channel or not _visit_token():
        logger.warning('[VISIT_AUDIT] 발송 채널/토큰 미설정 — skip')
        return False
    try:
        body = json.dumps({'channel': channel, 'text': text,
                           'unfurl_links': False, 'unfurl_media': False}).encode('utf-8')
        req = urllib.request.Request(
            'https://slack.com/api/chat.postMessage', data=body,
            headers={'Authorization': f'Bearer {_visit_token()}',
                     'Content-Type': 'application/json; charset=utf-8'})
        r = json.loads(urllib.request.urlopen(req, timeout=10).read().decode('utf-8'))
        if not r.get('ok'):
            logger.warning(f"[VISIT_AUDIT] DM 실패: {r.get('error')}")
        return bool(r.get('ok'))
    except Exception as exc:
        logger.warning(f'[VISIT_AUDIT] DM 예외: {exc}')
        return False


def run_daily_canvas1_audit(force: bool = False) -> Dict:
    """오후 4시 점검 — 월~금(공휴일 제외), 문제 있을 때만 사장님 DM."""
    from dashboard.services.absent_remind import _is_business_day
    today = date.today()
    if not force and not _is_business_day(today):
        return {'skipped': '영업일 아님'}
    items = check_cards_vs_canvas1(today)
    if not items:
        logger.info('[VISIT_AUDIT] 오후 점검 — 캔버스1 누락 없음')
        return {'items': 0}
    sent = send_admin_dm(format_daily(items))
    logger.info(f'[VISIT_AUDIT] 오후 점검 — {len(items)}건, DM={sent}')
    return {'items': len(items), 'sent': sent}
