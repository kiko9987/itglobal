# -*- coding: utf-8 -*-
"""슬랙 휴가 신청 (2026-11-05) — 휴가 봇(전자결재와 따로, SLACK_LEAVE_*). 장부·규칙은 services/leave.py.

입구: #휴가_일정 워크플로 버튼(단계 leave_open → 신청 창 바로) · 봇 홈 탭 [🌴 휴가 신청] [📅 내 휴가].
흐름: 신청 창 → 요청자 DM 카드 + JW DM [승인/반려] → JW 승인이면 YG DM →
YG 승인이면 요청자에게 알리고 #휴가_일정에 '누가 언제 휴가/반차'만 올린다. 채널에는 신청 카드를 올리지 않는다.
창·워크플로 단계 처리 도우미(_push_kw·_ack_completing_step 등)는 전자결재 것을 그대로 쓴다(봇과 무관한 함수).
"""
import json
import logging
import os
import re
import threading
from datetime import date
from typing import Optional

from dashboard.services import leave as lv

logger = logging.getLogger(__name__)
_LOG = '[SLACK/휴가]'


def _eap():
    from dashboard.blueprints import slack_eapproval as se
    return se


def leave_client():
    """휴가 봇 WebClient (없으면 None)."""
    tok = os.getenv('SLACK_LEAVE_BOT_TOKEN', '').strip()
    if not tok:
        return None
    from slack_sdk import WebClient
    return WebClient(token=tok)


def _client():
    return leave_client()


# ── 신청 창 ──────────────────────────────────────────────────────────
def _plain(t: str) -> dict:
    return {'type': 'plain_text', 'text': t, 'emoji': True}


def _opt(value: str, text: str) -> dict:
    return {'text': _plain(text[:75]), 'value': value}


def _select(block_id: str, label: str, opts, initial: str = '', dispatch: bool = False,
            optional: bool = False, hint: str = '') -> dict:
    el = {'type': 'static_select', 'action_id': block_id, 'options': [_opt(v, t) for v, t in opts],
          'placeholder': _plain('고르기')}
    init = next((o for o in el['options'] if o['value'] == initial), None)
    if init:
        el['initial_option'] = init
    b = {'type': 'input', 'block_id': block_id, 'label': _plain(label), 'element': el, 'optional': optional}
    if dispatch:
        b['dispatch_action'] = True
    if hint:
        b['hint'] = _plain(hint)
    return b


def _datepicker(block_id: str, label: str, initial: str = '', optional: bool = False, hint: str = '') -> dict:
    el = {'type': 'datepicker', 'action_id': block_id, 'placeholder': _plain('날짜 고르기')}
    if initial:
        el['initial_date'] = initial
    b = {'type': 'input', 'block_id': block_id, 'label': _plain(label), 'element': el, 'optional': optional}
    if hint:
        b['hint'] = _plain(hint)
    return b


def leave_view(meta: dict, vals: Optional[dict] = None) -> dict:
    """meta: name(신청하는 본인). vals: 고른 값(다시 그릴 때). 휴가는 본인만 신청(대신 신청 없음, 11-05)."""
    vals = vals or {}
    name = meta.get('name', '')
    ltype = vals.get('type') or lv.ANNUAL
    blocks = []
    left = lv.balance(name) if name else 0
    blocks.append({'type': 'context', 'elements': [{'type': 'mrkdwn', 'text':
                   f'*{name}* 님 오늘 기준 남은 연차: *{lv.fmt_days(left)}*'}]})
    # 칸 id 는 lv_ 를 붙인다 — 다른 봇 양식 칸 이름과 헷갈리지 않게
    blocks.append(_select('lv_type', '휴가 종류', lv.type_options(), initial=ltype, dispatch=True))
    if ltype == lv.EVENT:
        blocks.append(_select('lv_sub', '경조사 종류', [(k, f'{v[0]} ({v[1]}일)') for k, v in lv.EVENT_SUBS.items()],
                              initial=vals.get('sub', '')))
    if ltype in lv.HALF_TYPES:
        blocks.append(_datepicker('lv_start', '날짜', vals.get('start', '')))
    else:
        blocks.append(_datepicker('lv_start', '시작일', vals.get('start', '')))
        blocks.append(_datepicker('lv_end', '마지막 날', vals.get('end', ''), optional=True,
                                  hint='하루만 쉬면 비워 두세요. 주말·공휴일은 일수에서 빠져요.'))
    reason = {'type': 'plain_text_input', 'action_id': 'lv_reason', 'placeholder': _plain('예: 개인 사정')}
    if vals.get('reason'):
        reason['initial_value'] = vals['reason']
    blocks.append({'type': 'input', 'block_id': 'lv_reason', 'optional': False, 'label': _plain('사유'),   # 필수(11-05)
                   'element': reason, 'hint': _plain('승인자(JW·YG)에게만 보여요. #휴가_일정에는 날짜만 올라가요.')})
    return {
        'type': 'modal', 'callback_id': lv.CB_SUBMIT,
        'title': _plain('휴가 신청'), 'submit': _plain('신청'), 'close': _plain('취소'), 'notify_on_close': True,
        'private_metadata': json.dumps(meta, ensure_ascii=False),
        'blocks': blocks,
    }


def read_view(view: dict) -> dict:
    st = (view.get('state') or {}).get('values') or {}

    def _v(key):
        el = (st.get(f'lv_{key}') or {}).get(f'lv_{key}') or {}
        return ((el.get('selected_option') or {}).get('value') or el.get('selected_date')
                or (el.get('value') or '').strip() or '')
    return {k: _v(k) for k in ('type', 'sub', 'start', 'end', 'reason')}


def _errors(errs: dict) -> dict:
    """lv.validate 칸 이름 → 창 block_id."""
    return {f'lv_{k}': v for k, v in errs.items()}


def _not_listed_view(name: str) -> dict:
    return {'type': 'modal', 'title': _plain('휴가 신청'), 'close': _plain('닫기'), 'notify_on_close': True,
            'blocks': [{'type': 'section', 'text': {'type': 'mrkdwn', 'text':
                        f':warning: {name or "이 계정"} 님은 휴가 명단에 아직 없어요. 경영지원에 문의해 주세요.'}}]}


def _not_open_view() -> dict:
    return {'type': 'modal', 'title': _plain('휴가 신청'), 'close': _plain('닫기'), 'notify_on_close': True,
            'blocks': [{'type': 'section', 'text': {'type': 'mrkdwn', 'text':
                        '🌴 슬랙 휴가 신청은 준비 중이에요. 지금은 플렉스에서 신청해 주세요.'}}]}


def start_view(uid: str) -> dict:
    """누른 사람에게 맞는 첫 창 — 열기 전(LEAVE_ENABLED 전 시험자 아님)·명단에 없음·신청 창."""
    if not lv.is_open(uid):
        return _not_open_view()
    me = lv.staff_by_slack(uid)
    return leave_view({'name': me['name']}) if me else _not_listed_view(_eap()._name(uid))


def open_leave_modal(client, body) -> None:
    se = _eap()
    view = start_view((body.get('user') or {}).get('id', ''))
    push = se._push_kw(body)
    if push:
        client.views_push(**push, view=view)
    else:
        client.views_open(trigger_id=body['trigger_id'], view=view)


# ── DM ────────────────────────────────────────────────────────────────
def _requester_ids(row: dict) -> list:
    """요청자 카드 받을 사람 — 본인(신청한 사람과 같다)."""
    out = []
    for u in (row.get('slack_id'), row.get('submitter_id')):
        if u and u not in out:
            out.append(u)
    return out


def _render(row: dict, role: str):
    left = lv.balance(row['name'], lv._d(row['start'])) if lv.deducts(row['ltype']) else None
    blocks = lv.card_blocks(row, lv.steps(row['id']), role, left)
    return f"휴가 신청 — {row['name']} {lv.period_text(row, False)}", blocks


def _send_card(client, row: dict, uid: str, role: str) -> None:
    try:
        ch = client.conversations_open(users=uid)['channel']['id']
        text, blocks = _render(row, role)
        r = client.chat_postMessage(channel=ch, text=text, blocks=blocks)
        lv.add_dm_ref(row['id'], ch, r['ts'], role, uid)
    except Exception as exc:
        logger.warning(f'{_LOG} DM 실패 (#{row["id"]} → {uid}, {role}): {exc}')


def _refresh(client, row: dict) -> None:
    """보낸 카드 전부 지금 상태로. 승인 카드는 그 사람이 지금 승인할 차례일 때만 버튼."""
    cur = lv.current_approver(row)
    for ref in lv.dm_refs(row):
        role = ref['role']
        if role == 'approver' and ref['user'] != cur:
            role = 'info'
        try:
            text, blocks = _render(row, role)
            client.chat_update(channel=ref['ch'], ts=ref['ts'], text=text, blocks=blocks)
        except Exception as exc:
            logger.warning(f'{_LOG} 카드 고치기 실패 (#{row["id"]}): {exc}')


def _notify(client, uid: str, text: str) -> None:
    _eap()._notify(client, uid, text)


def _announce(client, row: dict) -> None:
    ch = lv.announce_channel()
    try:
        r = client.chat_postMessage(channel=ch, text=lv.announce_text(row))
        lv.set_announce(row['id'], ch, r['ts'])
    except Exception as exc:
        logger.error(f'{_LOG} #휴가_일정 게시 실패 (#{row["id"]}): {exc}', exc_info=True,
                     extra={'force_alert': True})


def _next_round(client, row: dict) -> None:
    """지금 승인할 사람에게 카드(이미 받았으면 건너뜀). 확정이면 요청자 알림 + #휴가_일정."""
    cur = lv.current_approver(row)
    if cur and (cur, 'approver') not in {(r['user'], r['role']) for r in lv.dm_refs(row)}:
        _send_card(client, row, cur, 'approver')
    if row['status'] == lv.APPROVED:
        for u in _requester_ids(row):
            _notify(client, u, f"✅ {row['name']} 님 휴가가 승인됐어요: {lv.period_text(row)} "
                               f"({lv.type_label(row['ltype'], row.get('sub') or '')})")
        _announce(client, row)


def start(row: dict) -> None:
    client = _client()
    for u in _requester_ids(row):
        _send_card(client, row, u, 'requester')
    _next_round(client, lv.get(row['id']))


def apply_action(action: str, rid: int, uid: str, reason: str = '') -> Optional[str]:
    """버튼 처리. 못 하면 사용자에게 보여 줄 문구."""
    client = _client()
    row = lv.get(rid)
    if not row:
        return '없는 신청이에요.'
    if not lv.can_act(action, row, uid):
        if action == 'cancel':
            return '시작한 휴가는 경영지원만 취소할 수 있어요.' if row['status'] == lv.APPROVED else '이미 처리된 신청이에요.'
        return '지금 승인할 차례가 아니에요.' if row['status'] in lv.OPEN else '이미 처리된 신청이에요.'
    actor = _eap()._name(uid)
    new = lv.transition(rid, action, uid, actor, reason)
    if not new:
        return '이미 처리된 신청이에요.'
    _refresh(client, new)
    if action == 'approve':
        _next_round(client, new)
    elif action == 'reject':
        for u in _requester_ids(new):
            _notify(client, u, f"❌ {new['name']} 님 휴가 신청이 반려됐어요: {lv.period_text(new)}\n사유: {reason}")
    elif action == 'cancel':
        if new.get('announce_ts'):
            try:
                client.chat_update(channel=new['announce_ch'], ts=new['announce_ts'],
                                   text=lv.cancelled_announce_text(new))
            except Exception as exc:
                logger.warning(f'{_LOG} #휴가_일정 취소 표시 실패 (#{rid}): {exc}')
        for u in _requester_ids(new):
            if u != uid:
                _notify(client, u, f"⚪ {new['name']} 님 휴가가 취소됐어요 ({actor}): {lv.period_text(new)}")
    logger.info(f'{_LOG} #{rid} {action} by {actor} → {new["status"]}')
    return None


# ── 내 휴가 ──────────────────────────────────────────────────────────
def mine_view(text: str) -> dict:
    return {
        'type': 'modal', 'title': _plain('내 휴가'), 'close': _plain('닫기'), 'notify_on_close': True,
        'blocks': [{'type': 'section', 'text': {'type': 'mrkdwn', 'text': text[:2900]}}],
    }


def home_view(uid: str) -> dict:
    """봇 홈 탭(왼쪽 앱 목록 → 휴가) — 남은 연차와 버튼 두 개."""
    me = lv.staff_by_slack(uid)
    head = (f'*{me["name"]}* 님 오늘 기준 남은 연차: *{lv.fmt_days(lv.balance(me["name"]))}*' if me
            else '휴가 명단에 아직 없어요. 경영지원에 문의해 주세요.')
    blocks = [{'type': 'section', 'text': {'type': 'mrkdwn', 'text': head}}]
    if lv.is_open(uid):
        blocks.append({'type': 'actions', 'elements': [
            {'type': 'button', 'action_id': lv.A_NEW, 'style': 'primary', 'text': _plain('🌴 휴가 신청')},
            {'type': 'button', 'action_id': lv.A_MINE, 'text': _plain('📅 내 휴가')}]})
    else:
        blocks.append({'type': 'context', 'elements': [{'type': 'mrkdwn', 'text':
                       '슬랙 휴가 신청은 준비 중이에요. 지금은 플렉스에서 신청해 주세요.'}]})
    blocks.append({'type': 'context', 'elements': [{'type': 'mrkdwn', 'text':
                   'JW → YG 승인 후 #휴가_일정에 날짜만 올라가요. 종류·사유는 승인자만 봐요.'}]})
    return {'type': 'home', 'blocks': blocks}


def _mine_text(uid: str) -> str:
    me = lv.staff_by_slack(uid)
    if not me:
        return ':warning: 휴가 명단에 아직 없어요. 경영지원에 문의해 주세요.'
    return lv.mine_text(me['name'])


# ── 핸들러 ────────────────────────────────────────────────────────────
def register_leave_handlers(app) -> None:
    se = _eap()

    @app.action(lv.A_NEW)
    def handle_leave_open(ack, body, client):
        ack()
        try:
            open_leave_modal(client, body)
        except Exception as exc:
            logger.error(f'{_LOG} 신청 창 열기 실패: {exc}', exc_info=True)

    def _redraw(ack, body, client):
        """종류를 고르면 창을 다시 그린다(경조사 세부·반차 날짜 칸)."""
        ack()
        view = body.get('view') or {}
        meta = json.loads(view.get('private_metadata') or '{}')
        vals = read_view(view)
        try:
            client.views_update(view_id=view['id'], hash=view.get('hash'), view=leave_view(meta, vals))
        except Exception as exc:
            logger.warning(f'{_LOG} 창 다시 그리기 실패: {exc}')

    @app.function(lv.FN_OPEN_LEAVE)
    def handle_open_leave_step(body, inputs, fail):
        """워크플로 단계 「휴가 신청 열기」 — #휴가_일정 [휴가 신청] 버튼 → 신청 창 바로(11-05).
        메뉴 단계와 같이 여기서 끝내지 않는다 — 창을 닫거나(view_closed) 제출에 성공할 때 끝낸다."""
        from slack_sdk import WebClient
        inter = (inputs or {}).get('interactivity') or {}
        uid = (inter.get('interactor') or {}).get('id', '')
        token = (body.get('event') or {}).get('bot_access_token', '')
        try:
            WebClient(token=token).views_open(interactivity_pointer=inter.get('interactivity_pointer', ''),
                                              view=start_view(uid))
        except Exception as exc:
            logger.error(f'{_LOG} 워크플로 단계(휴가 신청 열기) 실패: {exc}', exc_info=True)
            fail(error='휴가 신청 창을 열지 못했어요. 잠시 뒤 다시 눌러 주세요.')

    app.action('lv_type')(_redraw)

    @app.view(lv.CB_SUBMIT)
    def handle_leave_submit(ack, body, view):
        uid = (body.get('user') or {}).get('id', '')
        vals = read_view(view)
        person = lv.staff_by_slack(uid)       # 휴가는 본인만 — 누른 사람의 슬랙 계정으로 정한다
        if not person:
            ack(response_action='errors', errors={'lv_type': '휴가 명단에서 찾지 못했어요. 경영지원에 문의해 주세요.'})
            return
        name = person['name']
        if not vals.get('reason'):
            ack(response_action='errors', errors={'lv_reason': '사유를 적어 주세요.'})
            return
        ltype, sub = vals.get('type', ''), vals.get('sub', '')
        start = lv._d(vals.get('start'))
        end = start if ltype in lv.HALF_TYPES else (lv._d(vals.get('end')) or start)
        errs = lv.validate(name, ltype, sub, start, end)
        if errs:
            ack(response_action='errors', errors=_errors(errs))
            return
        row = lv.create(name, uid, uid, name, ltype, sub, start, end,
                        vals.get('reason', ''))
        if not row:
            ack(response_action='errors', errors={'lv_start': '그 사이 남은 연차가 바뀌었어요. 다시 확인해 주세요.'})
            return
        se._ack_completing_step(ack, body)()

        def _bg():
            try:
                start_row = lv.get(row['id'])
                start(start_row)
                logger.info(f'{_LOG} #{row["id"]} 신청 {name} {ltype} {row["start"]}~{row["end"]} ({row["days"]}일)')
            except Exception as exc:
                logger.error(f'{_LOG} 신청 처리 실패 (#{row["id"]}): {exc}', exc_info=True,
                             extra={'force_alert': True})
        threading.Thread(target=_bg, daemon=True).start()

    def _bg_action(action):
        def handler(ack, body, client):
            ack()
            rid = int(body['actions'][0]['value'])
            uid = (body.get('user') or {}).get('id', '')

            def _bg():
                try:
                    msg = apply_action(action, rid, uid)
                    if msg:
                        se._ephemeral(client, body, f'🔒 {msg}')
                except Exception as exc:
                    logger.error(f'{_LOG} {action} 실패 (#{rid}): {exc}', exc_info=True, extra={'force_alert': True})
                    se._ephemeral(client, body, ':warning: 처리 중 오류가 났어요. 잠시 뒤 다시 눌러 주세요.')
            threading.Thread(target=_bg, daemon=True).start()
        return handler

    app.action(lv.A_APPROVE)(_bg_action('approve'))
    app.action(lv.A_CANCEL)(_bg_action('cancel'))

    @app.action(lv.A_REJECT)
    def handle_leave_reject_open(ack, body, client):
        ack()
        rid = int(body['actions'][0]['value'])
        row = lv.get(rid)
        uid = (body.get('user') or {}).get('id', '')
        if not row or not lv.can_act('reject', row, uid):
            se._ephemeral(client, body, '🔒 지금 반려할 수 있는 신청이 아니에요.')
            return
        try:
            client.views_open(trigger_id=body['trigger_id'], view={
                'type': 'modal', 'callback_id': lv.CB_REJECT, 'title': _plain('휴가 반려'),
                'submit': _plain('반려'), 'close': _plain('취소'),
                'private_metadata': json.dumps({'rid': rid}),
                'blocks': [{'type': 'input', 'block_id': 'reason', 'label': _plain('반려 사유'),
                            'element': {'type': 'plain_text_input', 'action_id': 'reason', 'multiline': True,
                                        'placeholder': _plain('요청자에게 그대로 전달돼요.')}}]})
        except Exception as exc:
            logger.error(f'{_LOG} 반려 창 열기 실패: {exc}', exc_info=True)

    @app.view(lv.CB_REJECT)
    def handle_leave_reject_submit(ack, body, view):
        reason = ((view['state']['values'].get('reason') or {}).get('reason') or {}).get('value', '').strip()
        if not reason:
            ack(response_action='errors', errors={'reason': '반려 사유를 적어 주세요.'})
            return
        ack()
        rid = int(json.loads(view.get('private_metadata') or '{}').get('rid') or 0)
        uid = (body.get('user') or {}).get('id', '')

        def _bg():
            try:
                msg = apply_action('reject', rid, uid, reason)
                if msg:
                    _notify(_client(), uid, f'🔒 {msg}')
            except Exception as exc:
                logger.error(f'{_LOG} 반려 실패 (#{rid}): {exc}', exc_info=True, extra={'force_alert': True})
        threading.Thread(target=_bg, daemon=True).start()

    @app.action(lv.A_MINE)
    def handle_leave_mine(ack, body, client):
        ack()
        uid = (body.get('user') or {}).get('id', '')
        view_id = se._hub_view_id(body)
        if not view_id and se._from_home(body):
            try:
                view_id = client.views_open(trigger_id=body['trigger_id'],
                                            view=se._loading_view('내 휴가'))['view']['id']
            except Exception as exc:
                logger.error(f'{_LOG} 내 휴가 창 열기 실패: {exc}', exc_info=True)
                return

        def _bg():
            try:
                text = _mine_text(uid)
            except Exception as exc:
                logger.error(f'{_LOG} 내 휴가 조회 실패: {exc}', exc_info=True)
                text = ':warning: 휴가 내역을 불러오지 못했어요. 잠시 뒤 다시 눌러 주세요.'
            if view_id:
                try:
                    client.views_update(view_id=view_id, view=mine_view(text))
                except Exception as exc:
                    logger.error(f'{_LOG} 내 휴가 창 바꾸기 실패: {exc}', exc_info=True)
            else:
                se._ephemeral(client, body, text)
        threading.Thread(target=_bg, daemon=True).start()

    @app.view_closed(re.compile(r'.*'))
    def handle_leave_view_closed(ack, body):
        """워크플로로 연 신청 창을 닫으면 그 단계를 끝낸다(안 끝내면 버튼이 막힌 채로 남음)."""
        ack()
        if body.get('is_cleared') or not (body.get('view') or {}).get('previous_view_id'):
            se._complete_step(body)

    @app.event('app_home_opened')
    def handle_leave_home(event, client):
        if event.get('tab') != 'home':
            return
        try:
            client.views_publish(user_id=event['user'], view=home_view(event['user']))
        except Exception as exc:
            logger.error(f'{_LOG} 홈 탭 게시 실패: {exc}', exc_info=True)


def run_accrual(today: Optional[date] = None) -> int:
    """매일 — 입사일 기준 새 연차를 넣고 본인에게 알린다. 넣은 줄 수."""
    made = lv.accrue(today)
    if not made:
        return 0
    client = _client()
    for g in made:
        p = lv.staff_by_name(g['name']) or {}
        if client and p.get('slack_id'):
            _notify(client, p['slack_id'],
                    f"🌴 연차 {lv.fmt_days(g['days'])}이 생겼어요 ({g['note']}). "
                    f"{lv.kdate(g['valid_to'])}까지 쓸 수 있어요. 지금 남은 연차: {lv.fmt_days(lv.balance(g['name']))}")
    logger.info(f'{_LOG} 연차 자동 발생 {len(made)}줄: ' + ', '.join(f"{g['name']} {g['days']}" for g in made))
    return len(made)
