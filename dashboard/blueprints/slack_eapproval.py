"""전자결재 봇 슬랙 처리 (2026-10-30).

- 결재 카드는 전부 전자결재 봇이 #전자_결재(SLACK_EAPPROVAL_CHANNEL)에 올린다.
  버튼은 메시지를 올린 봇에게만 오므로, 어디서 시작한 결재든 승인·반려·이체 완료는 이 봇이 받는다.
- 다음 처리자에게 DM(전자결재 봇 im:write). 상태가 바뀌면 채널 카드와 그 DM 들을 함께 고친다.
- 계산서 글(#영업_관리)에서 시작한 결재는 그 스레드에도 결과를 남긴다(계산서봇으로 — 그 채널에 있는 봇).

입구:
  ① 계산서봇 [📍 현장 지정] 모달의 「이체 필요」 → start_from_invoice()   (slack_bot.py 에서 호출)
  ② #전자_결재 메뉴 [💸 자금집행 요청] / `/자금집행` → 새 요청 모달(전자결재 봇)
  ③ 공사확정 카드 [💸 자금집행 요청] → 새 요청 모달(공사봇이 열고, 제출도 공사봇이 받아 여기로)

장부·상태·카드 모양: services/e_approval.py. 설계: docs/e_approval_design.md
"""
from __future__ import annotations

import json
import logging
import os
import threading
from typing import Optional

from dashboard.services import e_approval as ea

logger = logging.getLogger(__name__)

_LOG = '[SLACK/전자결재]'


# ── 클라이언트 ────────────────────────────────────────────────────────
def eap_client():
    """전자결재 봇 WebClient (없으면 None)."""
    tok = os.getenv('SLACK_EAPPROVAL_BOT_TOKEN', '').strip()
    if not tok:
        return None
    from slack_sdk import WebClient
    return WebClient(token=tok)


def is_configured() -> bool:
    return bool(os.getenv('SLACK_EAPPROVAL_BOT_TOKEN', '').strip() and ea.channel_id())


def _invoice_client():
    tok = os.getenv('SLACK_INVOICE_BOT_TOKEN', '').strip()
    if not tok:
        return None
    from slack_sdk import WebClient
    return WebClient(token=tok)


def _name(uid: str) -> str:
    try:
        from dashboard.blueprints.slack_bot import _resolve_intake_manager_name
        return _resolve_intake_manager_name(uid) or uid
    except Exception:
        return uid


# ── 프로젝트 표시 ────────────────────────────────────────────────────
def _project_rec(code: str) -> Optional[dict]:
    from dashboard.services.purchase_invoice_assign import SPECIAL_LABELS
    if not code or code in SPECIAL_LABELS:
        return None
    try:
        from dashboard.services.project_service import get_project_records
        for r in get_project_records() or []:
            if str(r.get('프로젝트 코드') or '').strip() == code:
                return r
    except Exception as exc:
        logger.warning(f'{_LOG} 프로젝트 조회 실패 ({code}): {exc}')
    return None


def _project_line_and_quote(code: str):
    from dashboard.services.purchase_invoice_assign import project_display
    rec = _project_rec(code)
    quote = ea.parse_amount((rec or {}).get('총액 1')) if rec else 0
    return project_display(code, rec), quote


def render(row: dict, role: str = 'all'):
    """(text, blocks). 채널 카드(all)=내용+버튼 전부. DM=알림 한 줄+요약+카드 링크만(11-01) —
    승인·반려·이체/구매 완료는 채널 카드 한 곳에서 누른다."""
    head = f'{ea.title(row)} — {ea.status_label(row)[1]}'
    line, quote = _project_line_and_quote(row.get('project_code'))
    if role == 'all':
        return head, ea.card_blocks(row, ea.steps(row['id']), line, quote, role)
    return head, ea.dm_blocks(row, role, line, _card_link(row))


def _card_link(row: dict) -> str:
    if not (row.get('card_channel') and row.get('card_ts')):
        return ''
    try:
        return eap_client().chat_getPermalink(channel=row['card_channel'],
                                              message_ts=row['card_ts']).get('permalink', '')
    except Exception:
        return ''


# ── DM ────────────────────────────────────────────────────────────────
def _send_dm(client, aid: int, uid: str, role: str) -> None:
    row = ea.get(aid)
    try:
        ch = client.conversations_open(users=uid)['channel']['id']
        text, blocks = render(row, role)
        r = client.chat_postMessage(channel=ch, text=text, blocks=blocks)
        ea.add_dm_ref(aid, ch, r['ts'], role, uid)
    except Exception as exc:
        logger.warning(f'{_LOG} DM 실패 ({ea.title(row)} → {uid}, {role}): {exc}')


def _notify(client, uid: str, text: str) -> None:
    """알림용 짧은 DM (기록 안 함)."""
    if not uid:
        return
    try:
        ch = client.conversations_open(users=uid)['channel']['id']
        client.chat_postMessage(channel=ch, text=text)
    except Exception as exc:
        logger.warning(f'{_LOG} 알림 DM 실패 ({uid}): {exc}')


def _dm_round(client, row: dict) -> None:
    """지금 상태에서 처리할 사람에게 DM. 같은 사람이 이미 그 역할 DM 을 받았으면 건너뜀."""
    have = {(r['user'], r['role']) for r in ea.dm_refs(row)}
    if row['status'] == ea.PENDING:
        targets = [(u, 'approver') for u in ea.approver_ids()]
    elif row['status'] == ea.APPROVED:
        targets = [(u, 'executor') for u in ea.executor_dm_ids()]
    else:
        targets = []
    for uid, role in targets:
        if (uid, role) not in have:
            _send_dm(client, row['id'], uid, role)


# ── 시작 ──────────────────────────────────────────────────────────────
def start(row: dict) -> dict:
    """새 결재 → 채널 카드 게시·고정 → 처리자 DM → 요청자 DM. 갱신된 row.

    카드를 못 올리면 결재를 '취소'로 닫고 예외를 다시 던진다(카드 없는 결재가 장부에 남지 않게)."""
    client = eap_client()
    ch = ea.channel_id()
    text, blocks = render(row, 'all')
    try:
        r = client.chat_postMessage(channel=ch, text=text, blocks=blocks)
    except Exception:
        ea.transition(row['id'], 'cancel', 'system', '시스템', '결재 카드 게시 실패')
        raise
    ea.set_card(row['id'], ch, r['ts'])
    try:
        client.pins_add(channel=ch, timestamp=r['ts'])
    except Exception as exc:
        logger.debug(f'{_LOG} 카드 고정 실패(무시): {exc}')
    row = ea.get(row['id'])
    _dm_round(client, row)
    row = ea.get(row['id'])
    req = row.get('requester_id')
    if req and req not in {d['user'] for d in ea.dm_refs(row)}:
        _send_dm(client, row['id'], req, 'requester')
    logger.info(f'{_LOG} 시작: {ea.title(row)} {row.get("project_code")} {row.get("amount"):,}원 '
                f'({row.get("requester_name")}, {row["status"]})')
    return ea.get(row['id'])


def refresh(row: dict) -> None:
    """채널 카드 + 기록된 DM 전부 지금 상태로. 끝난 결재는 카드 고정 해제."""
    client = eap_client()
    if row.get('card_channel') and row.get('card_ts'):
        text, blocks = render(row, 'all')
        try:
            client.chat_update(channel=row['card_channel'], ts=row['card_ts'], text=text, blocks=blocks)
        except Exception as exc:
            logger.warning(f'{_LOG} 카드 갱신 실패 ({ea.title(row)}): {exc}')
        if row['status'] not in ea.OPEN_STATUSES:
            try:
                client.pins_remove(channel=row['card_channel'], timestamp=row['card_ts'])
            except Exception as exc:
                if 'no_pin' not in str(exc):
                    logger.debug(f'{_LOG} 카드 고정 해제 실패(무시): {exc}')
    for d in ea.dm_refs(row):
        text, blocks = render(row, d['role'])
        try:
            client.chat_update(channel=d['ch'], ts=d['ts'], text=text, blocks=blocks)
        except Exception as exc:
            logger.warning(f'{_LOG} DM 갱신 실패 ({ea.title(row)} {d}): {exc}')


def _thread_log(row: dict, text: str, broadcast: bool = False) -> None:
    if not (row.get('card_channel') and row.get('card_ts')):
        return
    try:
        eap_client().chat_postMessage(channel=row['card_channel'], thread_ts=row['card_ts'],
                                      text=text, reply_broadcast=broadcast)
    except Exception as exc:
        logger.warning(f'{_LOG} 스레드 기록 실패 ({ea.title(row)}): {exc}')


def _origin_log(row: dict, text: str) -> None:
    """계산서 글(#영업_관리) 스레드에 결과 한 줄 — 계산서봇으로."""
    if not (row.get('origin_channel') and row.get('origin_ts')):
        return
    c = _invoice_client()
    if not c:
        return
    try:
        c.chat_postMessage(channel=row['origin_channel'], thread_ts=row['origin_ts'], text=text)
    except Exception as exc:
        logger.warning(f'{_LOG} 계산서 글 스레드 기록 실패 ({ea.title(row)}): {exc}')


def apply_action(action: str, aid: int, uid: str, reason: str = '') -> Optional[str]:
    """승인·반려·이체 완료·취소 처리. 실패면 사용자에게 보여 줄 문구, 성공이면 None."""
    row = ea.get(aid)
    if not row:
        return '결재를 찾지 못했어요.'
    w = ea.step2_word(row)
    if not ea.can_act(action, row, uid):
        return {
            'approve': '승인은 승인자만 할 수 있어요.',
            'reject': '반려는 승인자만 할 수 있어요.',
            'transfer': f'{w} 완료는 경영지원 담당만 누를 수 있어요.',
            'cancel': '요청 취소는 요청한 사람이나 승인자만 할 수 있어요.',
        }[action]
    name = _name(uid)
    new = ea.transition(aid, action, uid, name, reason)
    if not new:
        return f"이미 처리된 결재예요 (지금: {ea.status_label(ea.get(aid) or row)[1]})."
    client = eap_client()
    amt = f"{int(new.get('amount') or 0):,}원"
    t = ea.title(new)
    if action == 'approve':
        _thread_log(new, f'✅ {name} 승인')
        _dm_round(client, new)
    elif action == 'transfer' and new.get('form') == ea.FORM_SUPPLY:
        _thread_log(new, f'🛒 {name} 구매 완료 — {new.get("content")} {amt}', broadcast=True)
        _notify(client, new.get('requester_id'), f'🛒 {t} 구매가 끝났어요. ({new.get("content")} {amt})')
    elif action == 'transfer':
        _thread_log(new, f'💸 {name} 이체 완료 — {new.get("partner")} {amt}', broadcast=True)
        _notify(client, new.get('requester_id'), f'💸 {t} 이체가 끝났어요. ({new.get("partner")} {amt})')
        _origin_log(new, f'💸 {t} 이체 완료 ({name})')
    elif action == 'reject':
        _thread_log(new, f'❌ {name} 반려 — {reason}')
        _notify(client, new.get('requester_id'), f'❌ {t} 반려됐어요. 사유: {reason}')
        _origin_log(new, f'❌ {t} 반려 ({name}) — {reason}')
    elif action == 'cancel':
        _thread_log(new, f'⚪ {name} 요청 취소')
        if uid != new.get('requester_id'):
            _notify(client, new.get('requester_id'), f'⚪ {t} 요청이 취소됐어요 ({name}).')
    refresh(ea.get(aid))
    _audit(new, action, name, reason)
    logger.info(f'{_LOG} {t} {action} ({name})')
    return None


def _audit(row: dict, action: str, name: str, reason: str = '') -> None:
    try:
        from dashboard.utils.user_database import get_audit_repository
        get_audit_repository().log_action(
            user_email=f'slack:{name}', action=f'EAPPROVAL_{action.upper()}',
            details=f"{ea.title(row)} {action} — {row.get('partner') or row.get('content')} "
                    f"{int(row.get('amount') or 0):,}원"
                    + (f' / {reason}' if reason else ''),
            project_code=row.get('project_code'), field_name='전자결재',
            old_value=None, new_value=row.get('status'), ip_address=None)
    except Exception as exc:
        logger.warning(f'{_LOG} 감사 로그 실패: {exc}')


# ── 입구 ① 계산서 글 ─────────────────────────────────────────────────
def start_from_invoice(uid: str, name: str, origin_ch: str, origin_ts: str, parsed: dict,
                       project_code: str, cost: str, memo: str, evidence: str,
                       payee_name: str, bank: str, account: str, urgency: str) -> dict:
    """[📍 현장 지정]에서 「이체 필요」로 제출 → 결재 생성·시작."""
    account = ea.normalize_account(account)
    ea.remember_payee(parsed.get('partner', ''), payee_name, bank, account)
    row = ea.create(ea.FORM_FUND, uid, name, {
        'project_code': project_code, 'cost_category': cost, 'content': memo, 'evidence': evidence,
        'grn': parsed.get('grn'), 'partner': parsed.get('partner'), 'payee_name': payee_name,
        'bank': bank, 'account': account, 'amount': parsed.get('amount'), 'urgency': urgency,
        'memo': '', 'origin_channel': origin_ch, 'origin_ts': origin_ts,
    })
    _audit(row, 'request', name)
    return start(row)


# ── 입구 ②③ 새 요청 모달 ─────────────────────────────────────────────
def open_new_modal(client, trigger_id: str, uid: str, project_code: str = '', origin: str = '') -> None:
    """새 요청 모달. project_code 가 있으면 현장 고정(공사확정 카드)."""
    # trigger_id 는 3초 — 시트 조회 없이 바로 연다(현장 주소는 카드에 나온다).
    meta = {'project': project_code, 'origin': origin}
    client.views_open(trigger_id=trigger_id, view=ea.new_request_view(meta, project_code or None, project_code))


def validate_new_submit(view: dict):
    """(fields, errors) — 제출 직후 동기 검사(ack 에 errors 를 실어야 해서)."""
    f = ea.read_new_request(view)
    errs = ea.complete_new_request(f)
    if not errs and not is_configured():
        errs['amount'] = '전자결재 봇이 아직 설정되지 않았어요. 관리자에게 알려 주세요.'
    return f, errs


def process_new_submit(uid: str, f: dict) -> dict:
    name = _name(uid)
    ea.remember_payee(f.get('partner', ''), f.get('payee_name'), f.get('bank'), f.get('account'))
    row = ea.create(ea.FORM_FUND, uid, name, f)
    _audit(row, 'request', name)
    return start(row)


def handle_new_submit(ack, body, view) -> None:
    """새 요청 모달 제출 — 전자결재 봇·공사봇 공용 핸들러 본문."""
    f, errs = validate_new_submit(view)
    if errs:
        ack(response_action='errors', errors=errs)
        return
    ack()
    uid = (body.get('user') or {}).get('id', '')

    def _bg():
        try:
            process_new_submit(uid, f)
        except Exception as exc:
            logger.error(f'{_LOG} 새 요청 처리 실패: {exc}', exc_info=True, extra={'force_alert': True})
            c = eap_client()
            if c:
                _notify(c, uid, ':warning: 자금집행 요청이 올라가지 않았어요. 잠시 뒤 다시 해 주세요.')
    threading.Thread(target=_bg, daemon=True).start()


# ── 비품 신청 (11-01) ─────────────────────────────────────────────────
def process_supply_submit(uid: str, f: dict) -> dict:
    name = _name(uid)
    row = ea.create(ea.FORM_SUPPLY, uid, name, f)
    _audit(row, 'request', name)
    return start(row)


def handle_supply_submit(ack, body, view) -> None:
    f = ea.read_supply_request(view)
    errs = ea.complete_supply_request(f)
    if not errs and not is_configured():
        errs['item'] = '전자결재 봇이 아직 설정되지 않았어요. 관리자에게 알려 주세요.'
    if errs:
        ack(response_action='errors', errors=errs)
        return
    ack()
    uid = (body.get('user') or {}).get('id', '')

    def _bg():
        try:
            process_supply_submit(uid, f)
        except Exception as exc:
            logger.error(f'{_LOG} 비품 신청 처리 실패: {exc}', exc_info=True, extra={'force_alert': True})
            c = eap_client()
            if c:
                _notify(c, uid, ':warning: 비품 신청이 올라가지 않았어요. 잠시 뒤 다시 해 주세요.')
    threading.Thread(target=_bg, daemon=True).start()


def _ephemeral(client, body, text: str) -> None:
    uid = (body.get('user') or {}).get('id', '')
    ch = (body.get('channel') or {}).get('id', '')
    try:
        if ch:
            client.chat_postEphemeral(channel=ch, user=uid, text=text)
        else:
            _notify(client, uid, text)
    except Exception:
        _notify(client, uid, text)


def _reject_view(aid: int, body: dict) -> dict:
    return {
        'type': 'modal', 'callback_id': ea.CB_REJECT,
        'title': {'type': 'plain_text', 'text': '반려'},
        'submit': {'type': 'plain_text', 'text': '반려'},
        'close': {'type': 'plain_text', 'text': '취소'},
        'private_metadata': json.dumps({'aid': aid}),
        'blocks': [{'type': 'input', 'block_id': 'reason', 'label': {'type': 'plain_text', 'text': '반려 사유'},
                    'element': {'type': 'plain_text_input', 'action_id': 'reason', 'multiline': True,
                                'placeholder': {'type': 'plain_text', 'text': '요청자에게 그대로 전달됩니다.'}}}],
    }


def register_eapproval_handlers(app) -> None:
    """전자결재 봇 핸들러."""

    def _bg_action(action):
        def handler(ack, body, client):
            ack()
            aid = int(body['actions'][0]['value'])
            uid = (body.get('user') or {}).get('id', '')

            def _bg():
                try:
                    msg = apply_action(action, aid, uid)
                    if msg:
                        _ephemeral(client, body, f'🔒 {msg}')
                except Exception as exc:
                    logger.error(f'{_LOG} {action} 실패 (#{aid}): {exc}', exc_info=True,
                                 extra={'force_alert': True})
                    _ephemeral(client, body, ':warning: 처리 중 오류가 났어요. 잠시 뒤 다시 눌러 주세요.')
            threading.Thread(target=_bg, daemon=True).start()
        return handler

    app.action(ea.A_APPROVE)(_bg_action('approve'))
    app.action(ea.A_TRANSFER)(_bg_action('transfer'))
    app.action(ea.A_CANCEL)(_bg_action('cancel'))

    @app.action(ea.A_REJECT)
    def handle_reject_open(ack, body, client):
        ack()
        aid = int(body['actions'][0]['value'])
        row = ea.get(aid)
        uid = (body.get('user') or {}).get('id', '')
        if not row or not ea.can_act('reject', row, uid):
            _ephemeral(client, body, '🔒 반려는 승인자만 할 수 있어요.')
            return
        try:
            client.views_open(trigger_id=body['trigger_id'], view=_reject_view(aid, body))
        except Exception as exc:
            logger.error(f'{_LOG} 반려 창 열기 실패: {exc}', exc_info=True)

    @app.view(ea.CB_REJECT)
    def handle_reject_submit(ack, body, view):
        reason = ((view['state']['values'].get('reason') or {}).get('reason') or {}).get('value', '').strip()
        if not reason:
            ack(response_action='errors', errors={'reason': '반려 사유를 적어 주세요.'})
            return
        ack()
        aid = json.loads(view.get('private_metadata') or '{}').get('aid')
        uid = (body.get('user') or {}).get('id', '')

        def _bg():
            try:
                msg = apply_action('reject', int(aid), uid, reason)
                if msg:
                    _notify(eap_client(), uid, f'🔒 {msg}')
            except Exception as exc:
                logger.error(f'{_LOG} 반려 실패 (#{aid}): {exc}', exc_info=True, extra={'force_alert': True})
        threading.Thread(target=_bg, daemon=True).start()

    @app.action(ea.A_NEW)
    def handle_new_open(ack, body, client):
        ack()
        try:
            open_new_modal(client, body['trigger_id'], (body.get('user') or {}).get('id', ''), origin='menu')
        except Exception as exc:
            logger.error(f'{_LOG} 새 요청 창 열기 실패: {exc}', exc_info=True)

    @app.command('/자금집행')
    def handle_new_command(ack, body, client):
        ack()
        try:
            open_new_modal(client, body['trigger_id'], body.get('user_id', ''), origin='command')
        except Exception as exc:
            logger.error(f'{_LOG} /자금집행 실패: {exc}', exc_info=True)

    @app.options(ea.A_PROJECT)
    def handle_project_options(ack, body):
        try:
            from dashboard.services import purchase_invoice_assign as pinv
            from dashboard.services.project_service import get_project_records
            from dashboard.blueprints.slack_bot import _intake_option_label
            query = (body.get('value') or '').strip()
            mgr = None if query else _name((body.get('user') or {}).get('id', ''))
            ack(options=pinv.build_project_options(
                get_project_records() or [], query, manager_name=mgr, label_fn=_intake_option_label,
                specials=(pinv.COMMON_CODE,)))
        except Exception as exc:
            logger.error(f'{_LOG} 프로젝트 목록 실패: {exc}', exc_info=True)
            try:
                ack(options=[])
            except Exception:
                pass

    @app.action(ea.A_PROJECT)
    def handle_project_pick(ack):
        ack()

    @app.view(ea.CB_NEW)
    def handle_new_submit_eap(ack, body, view):
        handle_new_submit(ack, body, view)

    @app.action(ea.A_NEW_SUPPLY)
    def handle_supply_open(ack, body, client):
        ack()
        try:
            client.views_open(trigger_id=body['trigger_id'], view=ea.supply_request_view({'origin': 'menu'}))
        except Exception as exc:
            logger.error(f'{_LOG} 비품 신청 창 열기 실패: {exc}', exc_info=True)

    @app.view(ea.CB_SUPPLY)
    def handle_supply_submit_eap(ack, body, view):
        handle_supply_submit(ack, body, view)


def register_project_bot_handlers(app) -> None:
    """공사봇 — 공사확정 카드 [💸 자금집행 요청] (공사봇 카드라 공사봇이 모달을 열고 제출도 받는다)."""

    @app.action(ea.A_NEW_FROM_PROJECT)
    def handle_new_from_project(ack, body, client):
        ack()
        code = body['actions'][0].get('value', '')
        try:
            open_new_modal(client, body['trigger_id'], (body.get('user') or {}).get('id', ''),
                           project_code=code, origin='project_card')
        except Exception as exc:
            logger.error(f'{_LOG} 공사 카드에서 새 요청 창 열기 실패 ({code}): {exc}', exc_info=True)

    @app.view(ea.CB_NEW)
    def handle_new_submit_project(ack, body, view):
        handle_new_submit(ack, body, view)


def post_menu() -> str:
    """#전자_결재에 메뉴 메시지를 올리고 고정(설치 후 한 번, scripts 에서 호출). ts 반환."""
    c = eap_client()
    ch = ea.channel_id()
    r = c.chat_postMessage(channel=ch, text='전자결재 요청', blocks=ea.menu_blocks())
    try:
        c.pins_add(channel=ch, timestamp=r['ts'])
    except Exception:
        pass
    return r['ts']


def update_menu(ts: str) -> None:
    """이미 올린 메뉴 메시지를 지금 버튼 구성으로 고친다(새 양식 추가 때)."""
    eap_client().chat_update(channel=ea.channel_id(), ts=ts, text='전자결재 요청', blocks=ea.menu_blocks())
