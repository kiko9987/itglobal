"""실제 운전자 받기 (2026-11-22 부장님) — 누가 운전할지는 출발할 때 동행자끼리 정하므로 배정 때 고르지 않고,
그날 저녁 그 차에 탄 사람들에게 슬랙 DM 으로 한 번 묻는다.

- 확정한 날(vehicle_assign_day.confirmed_at)만 묻는다 — 부장님이 시험 삼아 저장만 한 날에 직원 DM 이 가지 않게.
- 탄 사람이 한 명이면 묻지 않고 그 사람으로 적는다(by='auto').
- 버튼: 탄 사람 이름마다 한 개 + [나눠서 운전](오전·오후 운전자를 따로). 한 명이 답하면 같은 차 다른 사람 DM 도
  '○○ 님이 답했어요'로 바뀐다. 나중에 다시 누르면 고쳐진다(마지막 답이 맞음).
- 운행일지(vehicle_log)는 답이 있으면 배정보다 이 값을 먼저 쓴다(13:00 기준 오전·오후).
- 회사로 돌아오지 않고 바로 퇴근하는 날이 많아(11월 운행일지: 카니발·스파크) '복귀 버튼' 대신 DM 으로 했다.
"""
from __future__ import annotations

import json
import urllib.request
from contextlib import contextmanager
from datetime import date, datetime
from typing import Dict, List, Optional, Tuple

from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)

_DDL = [
    '''CREATE TABLE IF NOT EXISTS vehicle_driver (
        day TEXT NOT NULL,
        plate TEXT NOT NULL,
        am TEXT,
        pm TEXT,
        answered_by TEXT,
        answered_at TEXT,
        PRIMARY KEY (day, plate)
    )''',
    '''CREATE TABLE IF NOT EXISTS vehicle_driver_msg (
        day TEXT NOT NULL,
        plate TEXT NOT NULL,
        code TEXT NOT NULL,
        channel TEXT,
        ts TEXT,
        PRIMARY KEY (day, plate, code)
    )''',
]
_table_ready = False
A_ONE, A_SPLIT, A_AM, A_PM = 'vdrv_one', 'vdrv_split', 'vdrv_am', 'vdrv_pm'


def _db():
    from dashboard.utils.user_database import get_user_database
    return get_user_database()


@contextmanager
def _conn():
    global _table_ready
    with _db()._get_connection() as conn:
        if not _table_ready:
            for ddl in _DDL:
                conn.execute(ddl)
            conn.commit()
            _table_ready = True
        yield conn
        conn.commit()


def _now() -> str:
    return datetime.now().strftime('%Y-%m-%d %H:%M:%S')


# ── 슬랙 ────────────────────────────────────────────────────────────────
def _slack(method: str, payload: Dict) -> Dict:
    from dashboard.services import vehicle_assign as va
    tok = va._token()
    if not tok:
        return {'ok': False, 'error': 'no_token'}
    try:
        req = urllib.request.Request(f'https://slack.com/api/{method}', data=json.dumps(payload).encode('utf-8'),
                                     headers={'Authorization': f'Bearer {tok}',
                                              'Content-Type': 'application/json; charset=utf-8'})
        r = json.loads(urllib.request.urlopen(req, timeout=10).read().decode('utf-8'))
        if not r.get('ok'):
            logger.warning(f"[운전자] {method} 실패: {r.get('error')}")
        return r
    except Exception as e:
        logger.warning(f'[운전자] {method} 예외: {e}')
        return {'ok': False, 'error': str(e)}


def _btn(text: str, action: str, value: Dict, primary: bool = False) -> Dict:
    b = {'type': 'button', 'text': {'type': 'plain_text', 'text': text[:75]}, 'action_id': action,
         'value': json.dumps(value, ensure_ascii=False)}
    if primary:
        b['style'] = 'primary'
    return b


# ── 읽기 ────────────────────────────────────────────────────────────────
def answer(day: str, plate: str) -> Optional[Dict]:
    with _conn() as conn:
        r = conn.execute('SELECT * FROM vehicle_driver WHERE day = ? AND plate = ?', (day, plate)).fetchone()
    return dict(r) if r else None


def actual_drivers(start: str, end: str) -> Dict[Tuple[str, str], Dict]:
    """운행일지용 — (날짜, 번호판) → {'am': 이니셜, 'pm': 이니셜} (답이 있는 차만, 반쪽만 답했으면 그쪽만)."""
    with _conn() as conn:
        rows = conn.execute('SELECT day, plate, am, pm FROM vehicle_driver WHERE day BETWEEN ? AND ?', (start, end)).fetchall()
    return {(r['day'], r['plate']): {'am': r['am'] or '', 'pm': r['pm'] or ''} for r in rows if (r['am'] or r['pm'])}


def _cars_to_ask(day: str) -> List[Tuple[str, List[str]]]:
    """확정한 날의 차마다 탄 사람(오전 먼저). 확정 안 한 날은 []."""
    from dashboard.services import vehicle_assign as va
    info = va._day_info(day) or {}
    if not info.get('confirmed_at'):
        return []
    return [(plate, va.riders(r)) for plate, r in va._day_rows(day).items() if va.riders(r)]


# ── 묻는 글 ──────────────────────────────────────────────────────────────
def _blocks(day: str, plate: str, riders: List[str], split_open: bool = False) -> List[Dict]:
    from dashboard.services import vehicle_assign as va
    who = va._who_map()
    cars = {c['plate']: c for c in va.fleet(True)}
    car = va._car_text(plate, cars)
    a = answer(day, plate) or {}
    blocks = [{'type': 'section', 'text': {'type': 'mrkdwn',
                                           'text': f":car: {va.day_label(day)} {car} 는 누가 운전했나요?\n"
                                                   '눌러 주시면 운행일지에 들어가요. 같은 차 한 분만 답하면 돼요.'}}]
    if not split_open:
        btns = [_btn(va.display(c, who), A_ONE, {'day': day, 'plate': plate, 'code': c}) for c in riders]
        btns.append(_btn('나눠서 운전(오전·오후)', A_SPLIT, {'day': day, 'plate': plate}))
        blocks.append({'type': 'actions', 'elements': btns[:25]})
        return blocks
    for half, action, label in (('am', A_AM, '오전(13시 전)'), ('pm', A_PM, '오후(13시부터)')):
        blocks.append({'type': 'context', 'elements': [{'type': 'mrkdwn', 'text': f'*{label}* 운전'}]})
        blocks.append({'type': 'actions', 'elements': [
            _btn(va.display(c, who), action, {'day': day, 'plate': plate, 'code': c}, primary=(a.get(half) == c))
            for c in riders][:25]})
    return blocks


def _done_text(day: str, plate: str, by: str) -> str:
    from dashboard.services import vehicle_assign as va
    who = va._who_map()
    cars = {c['plate']: c for c in va.fleet(True)}
    a = answer(day, plate) or {}
    am, pm = a.get('am') or '', a.get('pm') or ''
    if am and am == pm:
        drove = va.display(am, who)
    else:
        drove = f"오전 {va.display(am, who) if am else '-'} · 오후 {va.display(pm, who) if pm else '-'}"
    return (f":white_check_mark: {va.day_label(day)} {va._car_text(plate, cars)} 운전: {drove}"
            f" (답: {va.display(by, who)}) — 틀렸으면 다시 눌러 고칠 수 있어요.")


def _refresh_messages(day: str, plate: str, by: str) -> None:
    """같은 차에 탄 모두의 DM 을 '답했어요'로 바꾼다(버튼은 남겨 고칠 수 있게)."""
    riders = dict(_cars_to_ask_all(day)).get(plate, [])
    text = _done_text(day, plate, by)
    blocks = [{'type': 'section', 'text': {'type': 'mrkdwn', 'text': text}}] + _blocks(day, plate, riders)[1:]
    with _conn() as conn:
        msgs = conn.execute('SELECT channel, ts FROM vehicle_driver_msg WHERE day = ? AND plate = ?', (day, plate)).fetchall()
    for m in msgs:
        if m['channel'] and m['ts']:
            _slack('chat.update', {'channel': m['channel'], 'ts': m['ts'], 'text': text, 'blocks': blocks})


def _cars_to_ask_all(day: str) -> List[Tuple[str, List[str]]]:
    """확정 여부와 상관없이 그날 차마다 탄 사람(버튼 누른 뒤 다시 그릴 때)."""
    from dashboard.services import vehicle_assign as va
    return [(plate, va.riders(r)) for plate, r in va._day_rows(day).items() if va.riders(r)]


# ── 보내기·받기 ──────────────────────────────────────────────────────────
def send_asks(day: Optional[str] = None) -> Dict:
    """그날 저녁 — 확정한 날, 탄 사람이 둘 이상이고 아직 답이 없는 차마다 탄 사람 모두에게 DM.
    한 명뿐인 차는 묻지 않고 그 사람으로 적는다. 이미 보낸 사람에게는 다시 안 보낸다."""
    from dashboard.services import vehicle_assign as va
    day = day or date.today().isoformat()
    who = va._who_map()
    sent = auto = 0
    for plate, riders in _cars_to_ask(day):
        if answer(day, plate):
            continue
        if len(riders) == 1:
            _save(day, plate, riders[0], riders[0], 'auto')
            auto += 1
            continue
        blocks = _blocks(day, plate, riders)
        text = f"{va.day_label(day)} 차량 운전자 확인"
        for c in riders:
            p = who.get(c)
            if not p or not p.get('slack_id'):
                continue
            with _conn() as conn:
                if conn.execute('SELECT 1 FROM vehicle_driver_msg WHERE day = ? AND plate = ? AND code = ?',
                                (day, plate, c)).fetchone():
                    continue
            r = _slack('chat.postMessage', {'channel': p['slack_id'], 'text': text, 'blocks': blocks,
                                            'unfurl_links': False, 'unfurl_media': False})
            if r.get('ok'):
                sent += 1
                with _conn() as conn:
                    conn.execute('INSERT OR REPLACE INTO vehicle_driver_msg (day, plate, code, channel, ts) VALUES (?,?,?,?,?)',
                                 (day, plate, c, r.get('channel') or p['slack_id'], r.get('ts') or ''))
    if sent or auto:
        logger.info(f'[운전자] {day} 운전자 묻기 DM {sent}건, 혼자 탄 차 {auto}대 자동')
    return {'day': day, 'sent': sent, 'auto': auto}


def _save(day: str, plate: str, am: Optional[str], pm: Optional[str], by: str) -> None:
    with _conn() as conn:
        conn.execute('INSERT INTO vehicle_driver (day, plate, am, pm, answered_by, answered_at) VALUES (?,?,?,?,?,?) '
                     'ON CONFLICT(day, plate) DO UPDATE SET am = COALESCE(excluded.am, vehicle_driver.am), '
                     'pm = COALESCE(excluded.pm, vehicle_driver.pm), answered_by = excluded.answered_by, '
                     'answered_at = excluded.answered_at', (day, plate, am, pm, by, _now()))


def _code_of_slack(slack_id: str) -> str:
    from dashboard.services import vehicle_assign as va
    return next((p['code'] for p in va.staff() if p.get('slack_id') == slack_id), '')


def handle_action(action_id: str, value: str, slack_user: str, channel: str = '', ts: str = '') -> Optional[str]:
    """버튼 처리. 돌려주는 값 = 누른 사람에게 보일 짧은 안내(없으면 None)."""
    try:
        v = json.loads(value or '{}')
    except ValueError:
        return None
    day, plate, code = v.get('day') or '', v.get('plate') or '', v.get('code') or ''
    by = _code_of_slack(slack_user) or slack_user
    riders = dict(_cars_to_ask_all(day)).get(plate, [])
    if not riders:
        return '이 날 차량 배정이 바뀌어 더 받을 수 없어요.'
    if action_id == A_SPLIT:          # 누른 사람 DM 만 오전·오후 고르기로 바꾼다
        if channel and ts:
            _slack('chat.update', {'channel': channel, 'ts': ts, 'text': '오전·오후 운전자를 골라 주세요',
                                   'blocks': _blocks(day, plate, riders, split_open=True)})
        return None
    if code not in riders:
        return None
    if action_id == A_ONE:
        _save(day, plate, code, code, by)
    elif action_id in (A_AM, A_PM):
        _save(day, plate, code if action_id == A_AM else None, code if action_id == A_PM else None, by)
        a = answer(day, plate) or {}
        if not (a.get('am') and a.get('pm')):       # 반쪽만 골랐으면 누른 사람 DM 만 선택 표시를 갱신
            if channel and ts:
                _slack('chat.update', {'channel': channel, 'ts': ts, 'text': '오전·오후 운전자를 골라 주세요',
                                       'blocks': _blocks(day, plate, riders, split_open=True)})
            return None
    else:
        return None
    _refresh_messages(day, plate, by)
    logger.info(f'[운전자] {day} {plate} 운전 답 by {by}: {answer(day, plate)}')
    return None
