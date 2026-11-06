# -*- coding: utf-8 -*-
"""시공팀 현장 출퇴근 보고·월말 조출/연장 정산 (2026-11-06) — 휴가 봇에 붙인다. 계산은 services/attendance.py.

흐름: 세콤 태그가 없는 날 → 다음 날 아침 본인 DM [🕒 출퇴근 시각 보고] → 창(날짜·출근·퇴근·메모) →
경영지원 DM [승인/반려] → 승인되면 정산에 들어간다(예전 카톡 보고 → SB 수기 입력 대신).
매달 1일 → 경영지원 DM 으로 지난달 정산표. ATTENDANCE_ENABLED 전에는 시공팀에게 묻지 않고 KiKO 에게만 보낸다.
"""
import json
import logging
import threading
from datetime import date, timedelta
from typing import Optional

from dashboard.services import attendance as at

logger = logging.getLogger(__name__)
_LOG = '[SLACK/근태]'


def _lv():
    from dashboard.blueprints import slack_leave as sl
    return sl


def _plain(t: str) -> dict:
    return {'type': 'plain_text', 'text': t, 'emoji': True}


def _md(t: str) -> dict:
    return {'type': 'section', 'text': {'type': 'mrkdwn', 'text': t[:2900]}}


def _kday(day: str) -> str:
    from dashboard.services.leave import kdate
    return kdate(date.fromisoformat(day), True)


# ── 보고 창 ───────────────────────────────────────────────────────────
def report_view(name: str, day: str, in_hm: str = '', out_hm: str = '') -> dict:
    def _time(block_id, label, initial):
        el = {'type': 'timepicker', 'action_id': block_id, 'placeholder': _plain('시각 고르기')}
        if initial:
            el['initial_time'] = initial
        return {'type': 'input', 'block_id': block_id, 'label': _plain(label), 'element': el, 'optional': True}

    return {
        'type': 'modal', 'callback_id': at.CB_REPORT, 'private_metadata': json.dumps({'name': name}),
        'title': _plain('출퇴근 시각 보고'), 'submit': _plain('보내기'), 'close': _plain('닫기'),
        'blocks': [
            _md(f'*{name}* 님 — 세콤에 찍히지 않은 출근·퇴근 시각을 알려 주세요. 경영지원이 확인하면 정산에 들어가요.'),
            {'type': 'input', 'block_id': 'att_day', 'label': _plain('날짜'),
             'element': {'type': 'datepicker', 'action_id': 'att_day', 'initial_date': day}},
            _time('att_in', '출근 시각 (현장 출근)', in_hm),
            _time('att_out', '퇴근 시각 (현장 퇴근)', out_hm),
            {'type': 'input', 'block_id': 'att_note', 'label': _plain('메모 (현장 이름 등)'), 'optional': True,
             'element': {'type': 'plain_text_input', 'action_id': 'att_note', 'max_length': 100}},
        ],
    }


def read_report(view: dict) -> dict:
    v = view['state']['values']
    return {'day': (v.get('att_day', {}).get('att_day') or {}).get('selected_date') or '',
            'in': (v.get('att_in', {}).get('att_in') or {}).get('selected_time') or '',
            'out': (v.get('att_out', {}).get('att_out') or {}).get('selected_time') or '',
            'note': ((v.get('att_note', {}).get('att_note') or {}).get('value') or '').strip()}


def _report_line(r: dict) -> str:
    parts = []
    if r.get('in_hm'):
        parts.append(f"출근 {r['in_hm']}")
    if r.get('out_hm'):
        parts.append(f"퇴근 {r['out_hm']}")
    return f"{_kday(r['day'])} " + ' · '.join(parts) + (f" ({r['note']})" if r.get('note') else '')


def manager_blocks(r: dict, decided: Optional[str] = None) -> list:
    rec = at.day_record(r['name'], r['day'])
    secom = f"세콤 기록: 출근 {at._hm(rec['in']) if rec['in_src'] == 'secom' else '-'} · " \
            f"퇴근 {at._hm(rec['out']) if rec['out_src'] == 'secom' else '-'}"
    blocks = [_md(f"🕒 *현장 출퇴근 보고 — {r['name']}*\n• {_report_line(r)}\n• {secom}\n"
                  f"_세콤 기록이 있는 쪽은 세콤 시각으로 계산해요._")]
    if decided:
        blocks.append({'type': 'context', 'elements': [{'type': 'mrkdwn', 'text': decided}]})
    else:
        blocks.append({'type': 'actions', 'block_id': f"att_{r['id']}", 'elements': [
            {'type': 'button', 'style': 'primary', 'text': _plain('✅ 승인'), 'action_id': at.A_APPROVE,
             'value': str(r['id'])},
            {'type': 'button', 'text': _plain('반려'), 'action_id': at.A_REJECT, 'value': str(r['id'])},
        ]})
    return blocks


def ask_blocks(name: str, items: list) -> list:
    lines = [f"🕒 *{name}* 님, 세콤에 출퇴근이 다 찍히지 않은 날이 있어요. 현장 출퇴근 시각을 알려 주세요."]
    for it in items:
        have = ' · '.join(x for x in (f"출근 {it['in']}" if it['in'] else '', f"퇴근 {it['out']}" if it['out'] else '') if x)
        lines.append(f"• {_kday(it['day'])} — {it['why']}" + (f" ({have})" if have else ''))
    lines.append('_쉬는 날이었으면 그냥 두셔도 돼요._')
    btns = [{'type': 'button', 'text': _plain(f"🕒 {int(it['day'][5:7])}/{int(it['day'][8:])} 보고"),
             'action_id': f"{at.A_REPORT}_{k}",
             'value': json.dumps({'name': name, 'day': it['day'], 'in': it['in'][:5] if it['in'] else '',
                                  'out': it['out'][:5] if it['out'] and '다음' not in it['out'] else ''})}
            for k, it in enumerate(items[:5])]
    return [_md('\n'.join(lines)), {'type': 'actions', 'elements': btns}]


# ── 보내기 ─────────────────────────────────────────────────────────────
def _dm(client, uid: str, text: str, blocks: list):
    ch = client.conversations_open(users=uid)['channel']['id']
    r = client.chat_postMessage(channel=ch, text=text, blocks=blocks)
    return ch, r['ts']


def submit_report(client, uid: str, name: str, vals: dict) -> dict:
    r = at.add_report(name, vals['day'], vals['in'], vals['out'], vals['note'], uid, _lv()._eap()._name(uid))
    refs = []
    for m in at.manager_ids():
        try:
            ch, ts = _dm(client, m, f"현장 출퇴근 보고 — {name} {r['day']}", manager_blocks(r))
            refs.append({'ch': ch, 'ts': ts})
        except Exception as exc:
            logger.warning(f'{_LOG} 경영지원 DM 실패 ({m}): {exc}')
    at.set_report_refs(r['id'], json.dumps(refs))
    logger.info(f"{_LOG} 보고 #{r['id']} {name} {r['day']} 출근 {r['in_hm']} 퇴근 {r['out_hm']}")
    return r


def decide(client, rid: int, ok: bool, uid: str) -> Optional[str]:
    actor = _lv()._eap()._name(uid)
    r = at.decide_report(rid, ok, actor)
    if not r:
        return '이미 처리된 보고예요.'
    note = f"{'✅ 승인' if ok else '❌ 반려'} — {actor}"
    for ref in json.loads(r.get('dm_refs') or '[]'):
        try:
            client.chat_update(channel=ref['ch'], ts=ref['ts'], text=f"현장 출퇴근 보고 — {r['name']}",
                               blocks=manager_blocks(r, note))
        except Exception as exc:
            logger.warning(f'{_LOG} 보고 카드 고치기 실패 (#{rid}): {exc}')
    if r.get('reporter_id'):
        _lv()._notify(client, r['reporter_id'],
                      f"{'✅' if ok else '❌'} 출퇴근 보고가 {'승인' if ok else '반려'}됐어요: {_report_line(r)}"
                      + ('' if ok else ' — 경영지원에 확인해 주세요.'))
    logger.info(f'{_LOG} 보고 #{rid} {"승인" if ok else "반려"} by {actor}')
    return None


def run_ask_missing(today: Optional[date] = None) -> int:
    """매일 아침 — 지난 3일 중 세콤이 덜 찍힌 날을 본인에게 묻는다(같은 날은 한 번). 켜기 전엔 KiKO 에게 미리보기."""
    client = _lv()._client()
    if not client:
        return 0
    at.ingest_dir()
    from dashboard.services import leave as lv
    sent = 0
    for name in at.staff():
        items = [it for it in at.missing_days(name, today) if at.mark_asked(name, it['day'])]
        if not items:
            continue
        p = lv.staff_by_name(name) or {}
        to = p.get('slack_id') if at.enabled() else at.manager_ids()[0]
        if not to:
            continue
        try:
            _dm(client, to, f'{name} 님 출퇴근 시각 보고 요청', ask_blocks(name, items))
            sent += 1
        except Exception as exc:
            logger.warning(f'{_LOG} 누락 묻기 DM 실패 ({name}): {exc}')
    return sent


def run_monthly(ym: str = '') -> bool:
    """매달 1일 — 지난달 정산표를 경영지원 DM 으로. 켜기 전엔 KiKO 에게만."""
    client = _lv()._client()
    if not client:
        return False
    at.ingest_dir(days=40)
    if not ym:
        first = date.today().replace(day=1)
        ym = (first - timedelta(days=1)).strftime('%Y-%m')
    s = at.summary(ym)
    text = at.summary_text(ym, s)
    blocks = [_md(text), _md(at.detail_text(ym, s))]
    ok = False
    for m in at.manager_ids():
        try:
            _dm(client, m, f'{ym} 시공팀 조출·연장 정산', blocks)
            ok = True
        except Exception as exc:
            logger.warning(f'{_LOG} 월말 정산 DM 실패 ({m}): {exc}')
    return ok


# ── 핸들러 ────────────────────────────────────────────────────────────
def register_attendance_handlers(app) -> None:
    import re
    se = _lv()._eap()

    @app.action(re.compile(rf'^{at.A_REPORT}(_\d+)?$'))
    def handle_att_open(ack, body, client):
        ack()
        try:
            v = json.loads(body['actions'][0].get('value') or '{}')
        except ValueError:
            v = {}
        uid = body['user']['id']
        from dashboard.services import leave as lv
        me = lv.staff_by_slack(uid) or {}
        name = v.get('name') or me.get('name') or ''
        if name and me.get('name') and name != me['name'] and uid not in at.manager_ids():
            se._ephemeral(client, body, '본인 출퇴근만 보고할 수 있어요.')
            return
        try:
            client.views_open(trigger_id=body['trigger_id'],
                              view=report_view(name, v.get('day') or date.today().isoformat(),
                                               v.get('in', ''), v.get('out', '')))
        except Exception as exc:
            logger.error(f'{_LOG} 보고 창 열기 실패: {exc}', exc_info=True)

    @app.view(at.CB_REPORT)
    def handle_att_submit(ack, body, view):
        vals = read_report(view)
        errs = at.validate_report(vals['in'], vals['out'])
        if errs:
            ack(response_action='errors', errors=errs)
            return
        ack()
        uid = body['user']['id']
        name = json.loads(view.get('private_metadata') or '{}').get('name', '')

        def _bg():
            client = _lv()._client()
            try:
                r = submit_report(client, uid, name, vals)
                _lv()._notify(client, uid, f'🕒 출퇴근 보고를 보냈어요: {_report_line(r)} — 경영지원이 확인해요.')
            except Exception as exc:
                logger.error(f'{_LOG} 보고 처리 실패: {exc}', exc_info=True, extra={'force_alert': True})
        threading.Thread(target=_bg, daemon=True).start()

    def _decide(ok):
        def handler(ack, body, client):
            ack()
            rid = int(body['actions'][0]['value'])
            uid = body['user']['id']
            if uid not in at.manager_ids():
                se._ephemeral(client, body, '경영지원만 승인할 수 있어요.')
                return

            def _bg():
                msg = decide(client, rid, ok, uid)
                if msg:
                    se._ephemeral(client, body, msg)
            threading.Thread(target=_bg, daemon=True).start()
        return handler

    app.action(at.A_APPROVE)(_decide(True))
    app.action(at.A_REJECT)(_decide(False))
