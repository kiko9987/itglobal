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
import re
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
    from dashboard.services.purchase_invoice_assign import NO_SITE_CODE, SPECIAL_LABELS
    if not code or code in SPECIAL_LABELS or code == NO_SITE_CODE:
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
    if not code:
        return '현장 무관', 0
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
    """새 결재 → 채널 카드 게시·고정 → 처리자 DM. 갱신된 row.

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
    # 신청자에게는 시작 DM 을 보내지 않는다(11-01) — 끝났을 때(완료·반려·남이 취소)만 알림.
    # 진행 상황은 채널 카드에서 보고, 요청 취소도 카드에서 한다.
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
    ini, ex = ea.initial(name), ea.executor_label(name)   # 화면엔 이니셜, ② 는 '경영지원 (SB)'
    if action == 'approve':
        _thread_log(new, f'✅ {ini} 승인')
        _dm_round(client, new)
    elif action == 'transfer' and new.get('form') in ea.TASK_FORMS:
        _thread_log(new, f'✅ {ex} 처리 완료', broadcast=True)
        _notify(client, new.get('requester_id'), f'✅ {t} 처리가 끝났어요.')
    elif action == 'transfer' and new.get('form') == ea.forms.FORM_EXPENSE:
        _thread_log(new, f'💸 {ex} 이체 완료 — 개인경비 {amt}', broadcast=True)
        _notify(client, new.get('requester_id'), f'💸 {t} 이체가 끝났어요. ({amt})')
    elif action == 'transfer' and new.get('form') == ea.FORM_SUPPLY:
        _thread_log(new, f'🛒 {ex} 구매 완료 — {new.get("content")} {amt}', broadcast=True)
        _notify(client, new.get('requester_id'), f'🛒 {t} 구매가 끝났어요. ({new.get("content")} {amt})')
    elif action == 'transfer':
        _thread_log(new, f'💸 {ex} 이체 완료 — {new.get("partner")} {amt}', broadcast=True)
        _notify(client, new.get('requester_id'), f'💸 {t} 이체가 끝났어요. ({new.get("partner")} {amt})')
        _origin_log(new, f'💸 {t} 이체 완료 ({ex})')
    elif action == 'reject':
        _thread_log(new, f'❌ {ini} 반려 — {reason}')
        _notify(client, new.get('requester_id'), f'❌ {t} 반려됐어요. 사유: {reason}')
        _origin_log(new, f'❌ {t} 반려 ({ini}) — {reason}')
    elif action == 'cancel':
        _thread_log(new, f'⚪ {ini} 요청 취소')
        if uid != new.get('requester_id'):
            _notify(client, new.get('requester_id'), f'⚪ {t} 요청이 취소됐어요 ({ini}).')
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
    fields = {
        'project_code': project_code, 'cost_category': cost, 'content': memo, 'evidence': evidence,
        'grn': parsed.get('grn'), 'partner': parsed.get('partner'), 'payee_name': payee_name,
        'bank': bank, 'account': account, 'amount': parsed.get('amount'), 'urgency': urgency,
        'memo': '', 'origin_channel': origin_ch, 'origin_ts': origin_ts,
    }
    ea.mark_payee_saved(fields)   # 기억해 둔 계좌 그대로면 카드에 표시 — remember 전에 비교
    ea.remember_payee(parsed.get('partner', ''), payee_name, bank, account)
    row = ea.create(ea.FORM_FUND, uid, name, fields)
    _audit(row, 'request', name)
    return start(row)


# ── 입구 ②③ 새 요청 모달 ─────────────────────────────────────────────
def open_new_modal(client, trigger_id: str, uid: str, project_code: str = '', origin: str = '',
                   push: Optional[dict] = None) -> None:
    """새 요청 모달. project_code 가 있으면 현장 고정(공사확정 카드).

    push(_push_kw) 가 있으면 메뉴 창 위에 쌓는다 — [취소]하면 메뉴로 돌아간다(11-05)."""
    # trigger_id 는 3초 — 시트 조회 없이 바로 연다(현장 주소는 카드에 나온다).
    meta = {'project': project_code, 'origin': origin}
    if push:
        client.views_push(**push, view=_new_view(meta))
    else:
        client.views_open(trigger_id=trigger_id, view=_new_view(meta))


def _new_view(meta: dict) -> dict:
    code = meta.get('project') or ''
    return ea.new_request_view(meta, code or None, code)


# ── 「계산서」 고르기 — #영업_관리에 고정된(=처리 전) 매입 계산서 글 중 현장 미지정 ──
_INV_CACHE = {'at': 0.0, 'items': []}


def open_invoice_posts(force: bool = False) -> list:
    """[{ch, ts, parsed}] 최신순. 고정 목록에서 계산서 양식이고, 현장 지정·진행 중 결재가 없는 글.

    목록을 누를 때마다 부르므로 1분 캐시(제출 때는 중복 검사를 다시 한다)."""
    import time
    if not force and time.time() - _INV_CACHE['at'] < 60:
        return _INV_CACHE['items']
    from dashboard.services import pin_remind as pr
    from dashboard.services import purchase_invoice_assign as pinv
    c, ch = _invoice_client(), os.getenv('SLACK_INVOICE_CHANNEL_ID', '').strip()
    items = []
    if c and ch:
        try:
            for it in (c.pins_list(channel=ch).get('items') or []):
                m = it.get('message') or {}
                text = pr._msg_full_text(m)
                if not m.get('ts') or not pr.is_invoice_request(text) or pr.is_deposit(text):
                    continue
                parsed = pinv.parse_invoice_post(text)
                if not parsed or pinv.get_assignment(ch, m['ts']):
                    continue
                ap = ea.find_by_origin(ch, m['ts'])
                if ap and ap['status'] not in (ea.REJECTED, ea.CANCELLED):
                    continue
                items.append({'ch': ch, 'ts': m['ts'], 'parsed': parsed})
        except Exception as exc:
            logger.warning(f'{_LOG} 계산서 글 목록 조회 실패: {exc}')
    items.sort(key=lambda x: float(x['ts']), reverse=True)
    _INV_CACHE.update(at=time.time(), items=items)
    return items


def invoice_options(query: str) -> list:
    from dashboard.services.purchase_invoice_assign import post_summary
    q = (query or '').strip().lower()
    opts = [ea.invoice_option(i) for i in open_invoice_posts()
            if not q or q in post_summary(i['parsed']).lower()]
    return [ea.invoice_option(None)] + opts[:99]


def _register_invoice_pick(app) -> None:
    """「계산서」 목록·선택 — 모달을 연 봇(전자결재 봇·공사봇)마다 등록."""

    @app.options(ea.A_INVOICE)
    def handle_invoice_options(ack, body):
        try:
            ack(options=invoice_options(body.get('value') or ''))
        except Exception as exc:
            logger.error(f'{_LOG} 계산서 목록 실패: {exc}', exc_info=True)
            try:
                ack(options=[ea.invoice_option(None)])
            except Exception:
                pass

    @app.action(ea.A_INVOICE)
    def handle_invoice_pick(ack, body, client):
        """계산서를 고르면 거래처·금액 칸을 숨기고 그 글 값으로(다른 칸에 넣은 값은 block_id 가 같아 유지)."""
        ack()
        try:
            view = body['view']
            meta = json.loads(view.get('private_metadata') or '{}')
            sel = (body['actions'][0].get('selected_option') or {}).get('value') or ''
            meta['inv'] = ea.invoice_from_value(sel)
            meta['no_inv'] = sel == ea.INVOICE_NONE
            partner = (meta['inv'] or {}).get('parsed', {}).get('partner') or meta.get('partner', '')
            _set_payee_prefill(meta, partner)
            client.views_update(view_id=view['id'], hash=view.get('hash'), view=_new_view(meta))
        except Exception as exc:
            logger.error(f'{_LOG} 계산서 선택 반영 실패: {exc}', exc_info=True)

    @app.options(ea.A_PARTNER)
    def handle_partner_options(ack, body):
        try:
            ack(**ea.partner_options(body.get('value') or ''))
        except Exception as exc:
            logger.error(f'{_LOG} 거래처 목록 실패: {exc}', exc_info=True)
            try:
                ack(options=[])
            except Exception:
                pass

    @app.action(ea.A_PARTNER)
    def handle_partner_pick(ack, body, client):
        """거래처를 고르면 예전에 보낸 계좌를 창에 채운다(확인 안내와 함께)."""
        ack()
        try:
            view = body['view']
            meta = json.loads(view.get('private_metadata') or '{}')
            sel = (body['actions'][0].get('selected_option') or {}).get('value') or ''
            meta['partner'] = ea.partner_from_value(sel)
            _set_payee_prefill(meta, meta['partner'])
            client.views_update(view_id=view['id'], hash=view.get('hash'), view=_new_view(meta))
        except Exception as exc:
            logger.error(f'{_LOG} 거래처 선택 반영 실패: {exc}', exc_info=True)


def _set_payee_prefill(meta: dict, partner: str) -> None:
    """거래처가 정해지면 기억한 계좌를 meta 에(없으면 비움). 칸 block_id 를 바꿔야 새 값이 보인다(pv+1)."""
    saved = ea.lookup_payee(partner) if partner else None
    meta['payee'] = ({k: saved[k] for k in ('payee_name', 'bank', 'account')} if saved else None)
    meta['pv'] = int(meta.get('pv') or 0) + 1


def validate_new_submit(view: dict):
    """(fields, errors) — 제출 직후 동기 검사(ack 에 errors 를 실어야 해서)."""
    f = ea.read_new_request(view)
    errs = ea.complete_new_request(f)
    if not errs and not is_configured():
        errs['content'] = '전자결재 봇이 아직 설정되지 않았어요. 관리자에게 알려 주세요.'
    # 오류는 화면에 있는 block_id 로(계좌 칸은 다시 채우면 payee_name_2 처럼 바뀐다)
    values = view['state']['values']
    errs = {ea.field_block_id(values, k) if k in ('payee_name', 'bank', 'account') else k: m
            for k, m in errs.items()}
    return f, errs


def process_new_submit(uid: str, f: dict) -> dict:
    name = _name(uid)
    ea.remember_payee(f.get('partner', ''), f.get('payee_name'), f.get('bank'), f.get('account'))
    row = ea.create(ea.FORM_FUND, uid, name, f)
    _audit(row, 'request', name)
    row = start(row)
    if row.get('origin_ts'):   # 계산서를 골랐으면 그 글의 현장 지정까지(계산서 글 버튼과 같은 결과)
        _INV_CACHE['at'] = 0.0
        try:
            from dashboard.blueprints.slack_bot import finish_pinv_assignment
            c = _invoice_client()
            if c:
                parsed = f.get('_inv_parsed') or {'date_md': '', 'grn': row.get('grn'),
                                                  'amount': row.get('amount'), 'partner': row.get('partner')}
                finish_pinv_assignment(
                    c, row['origin_channel'], row['origin_ts'], None, parsed, row.get('project_code') or '',
                    row.get('cost_category') or '', row.get('content') or '', uid, name,
                    extra=f"💸 {ea.title(row)} 결재 올림 — {ea.status_label(row)[1]}")
        except Exception as exc:
            logger.error(f'{_LOG} 계산서 글 현장 지정 실패 ({ea.title(row)}): {exc}', exc_info=True,
                         extra={'force_alert': True})
    return row


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


def reset_home(client, body) -> None:
    """홈 탭 고르기 칸 비우기 — 메뉴를 새 block_id 로 다시 게시(reset_select 의 홈 탭판)."""
    import time
    view = ea.home_view()
    for b in view['blocks']:
        if b['type'] == 'actions':
            b['block_id'] = f'eap_sel_{int(time.time() * 1000)}'
    try:
        client.views_publish(user_id=(body.get('user') or {}).get('id', ''), view=view)
    except Exception as exc:
        logger.warning(f'{_LOG} 홈 탭 고르기 칸 비우기 실패: {exc}')


def reset_select(client, body) -> None:
    """메시지의 고르기 칸을 처음 상태('📋 결재 요청')로 되돌린다 (11-01).

    슬랙은 고른 값을 칸에 남겨 두어 다시 고를 수 없게 보인다(같은 항목은 다시 골라도 이벤트가 안 옴).
    그 고르기가 든 actions 블록의 block_id 만 새로 바꿔 메시지를 고치면 칸이 비워진다 — 나머지 블록은 받은 그대로."""
    msg = body.get('message') or {}
    ch = (body.get('channel') or {}).get('id') or (body.get('container') or {}).get('channel_id')
    bid = (body.get('actions') or [{}])[0].get('block_id')
    if not (msg.get('ts') and ch and bid and msg.get('blocks')):
        return
    import time
    blocks = msg['blocks']
    for b in blocks:
        if b.get('block_id') == bid:
            b['block_id'] = f'eap_sel_{int(time.time() * 1000)}'
    kw = {'channel': ch, 'ts': msg['ts'], 'text': msg.get('text') or '', 'blocks': blocks}
    if msg.get('attachments'):
        kw['attachments'] = msg['attachments']
    try:
        client.chat_update(**kw)
    except Exception as exc:
        logger.warning(f'{_LOG} 고르기 칸 되돌리기 실패: {exc}')


# ── 나머지 양식 (11-01: 계약서·보증보험·카드결제·법인서류·법인카드 지출결의·개인경비) ──
def _loading_view(title: str) -> dict:
    return {'type': 'modal', 'title': {'type': 'plain_text', 'text': title[:24]},
            'close': {'type': 'plain_text', 'text': '취소'}, 'notify_on_close': True,
            'blocks': [{'type': 'section', 'text': {'type': 'mrkdwn', 'text': ':hourglass_flowing_sand: 불러오는 중...'}}]}


def _form_values(form_key: str, uid: str, project_code: str = '') -> dict:
    """창에 미리 넣을 값 — 현장(공사현황)·본인 기억(법인카드 번호·개인경비 계좌)."""
    from dashboard.services.purchase_invoice_assign import project_display
    form = ea.forms.get(form_key)
    rec = _project_rec(project_code) if project_code else None
    memory = ea.personal_memory(uid) if form_key in (ea.forms.FORM_CARDEXP, ea.forms.FORM_EXPENSE) else {}
    vals = ea.forms.prefill(form, rec, memory)
    # 이메일 — 공사현황 발주처 이메일이 비면 거래처 탭(사업자명) 이메일로
    mail_key = {ea.forms.FORM_CONTRACT: 'contact', ea.forms.FORM_INSURANCE: 'email'}.get(form_key)
    if rec and mail_key and not vals.get(mail_key):
        try:
            from dashboard.services.partner_status_sync import get_cached_partner_email
            mail = get_cached_partner_email(str(rec.get('사업자명') or ''))
            if mail:
                vals[mail_key] = mail
        except Exception:
            pass
    if project_code:
        vals['project'] = project_code
        vals['_project_label'] = project_display(project_code, rec)
    return vals


_FLEX_CONTRACT_CODES = None


def _flex_contract_codes() -> set:
    """Flex 1년치 내보내기에서 계약서 작성 요청이 끝난 공사코드 (보증보험 안내용, 한 번 읽고 기억)."""
    global _FLEX_CONTRACT_CODES
    if _FLEX_CONTRACT_CODES is None:
        codes = set()
        try:
            import re as _re
            root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
            path = os.path.join(root, 'instance', 'flex_export', 'flex_workflow_export_20261101.json')
            with open(path, encoding='utf-8') as fh:
                for d in json.load(fh).get('docs', []):
                    if '계약서' in d.get('t', '') and d.get('s') == 'DONE':
                        codes.update(_re.findall(r'[GR]\d{4}-[A-Z]{2,3}', str((d.get('f') or {}).get('프로젝트 코드', ''))))
        except Exception as exc:
            logger.info(f'{_LOG} Flex 계약서 기록 없음(안내만 생략): {exc}')
        _FLEX_CONTRACT_CODES = codes
    return _FLEX_CONTRACT_CODES


def contract_note(code: str) -> str:
    """보증보험 창 안내 — 이 현장 계약서에 적힌 보증. 입찰보증은 계약 전이라 막지 않고 안내만."""
    if not code:
        return ''
    rows = ea.find_by_project(ea.forms.FORM_CONTRACT, code)
    if rows:
        r = rows[0]
        g = ea.forms.guarantee_summary(ea.extra(r).get('f') or {})
        if g:
            return f'📑 {ea.title(r)}에 적힌 보증: *{g}*'
        return (f'⚠️ {ea.title(r)}에 보증보험이 적혀 있지 않아요. 입찰보증이 아니라면 계약서에 보증 내용이 '
                '있어야 신청할 수 있어요.')
    if code in _flex_contract_codes():
        return '📑 Flex로 올린 계약서 작성 요청이 있어요. 계약서에 이 보증 내용이 있는지 확인해 주세요.'
    return ('⚠️ 이 현장의 계약서 작성 요청이 아직 없어요. 입찰보증이 아니라면 계약서를 먼저 요청하거나 '
            '계약서 파일을 카드 스레드에 올려 주세요.')


def license_note(code: str) -> str:
    """계약서 창 안내 — 사업자등록증 판정(세금계산서 요청과 같은 규칙: 공사 폴더 > 같은 사업자 재사용 > 거래처 탭)."""
    if not code:
        return ''
    rec = _project_rec(code) or {}
    biz = str(rec.get('사업자명') or '').strip()
    if biz in ('', '-', 'None'):
        return ('ℹ️ 프로젝트에 사업자명이 없어요. 사업자 고객이면 사업자등록증을 첨부해 주세요. '
                '개인 고객은 이름 · 주소 · 연락처 · 주민번호가 필요합니다.')
    try:
        from dashboard.services.business_license_handler import get_license_state
        st = get_license_state(code) or {}
    except Exception as exc:
        logger.warning(f'{_LOG} 등록증 상태 조회 실패 ({code}): {exc}')
        return ''
    link = f" · <{st['view_url']}|보기>" if st.get('view_url') else ''
    if st.get('source') == 'own':
        return f'✅ 사업자등록증 있음 ({biz}){link}'
    if st.get('source') == 'reuse':
        return f'✅ 같은 사업자({biz})의 사업자등록증이 있어요{link}'
    if st.get('source') == 'partner':
        try:
            from dashboard.services.partner_status_sync import get_partner_info
            bno = (get_partner_info(biz) or {}).get('bno') or ''
        except Exception:
            bno = ''
        return f'✅ 거래처 탭에 등록된 거래처예요 ({biz}{" · " + bno if bno else ""}). 등록증 파일 없이 진행할 수 있어요.'
    return f'⚠️ {biz}의 사업자등록증이 없어요. 공사 카드 스레드나 PM에서 올려 주세요.'


def _form_notes(form_key: str, code: str) -> str:
    if not code:
        return ''
    if form_key == ea.forms.FORM_INSURANCE:
        return contract_note(code)
    if form_key == ea.forms.FORM_CONTRACT:
        return license_note(code)
    return ''


def open_form_modal(client, trigger_id: str, uid: str, form_key: str, project_code: str = '',
                    push: Optional[dict] = None) -> None:
    """양식 창 — 불러오는 창을 먼저 열고(3초 trigger) 공사현황·기억 값을 채워 바꾼다.

    push 가 있으면 메뉴 창 위에 쌓는다([취소] → 메뉴)."""
    form = ea.forms.get(form_key)
    if push:
        view_id = client.views_push(**push, view=_loading_view(form.label))['view']['id']
    else:
        view_id = client.views_open(trigger_id=trigger_id, view=_loading_view(form.label))['view']['id']

    def _bg():
        try:
            vals = _form_values(form_key, uid, project_code)
            meta = {'form': form_key, 'project': project_code, 'pv': 1}
            note = _form_notes(form_key, project_code)
            if note:
                meta['extra_note'] = note
            client.views_update(view_id=view_id, view=ea.forms.form_view(
                form_key, meta, vals, project_label=vals.get('_project_label', '')))
        except Exception as exc:
            logger.error(f'{_LOG} 양식 창 채우기 실패 ({form_key}): {exc}', exc_info=True)
    threading.Thread(target=_bg, daemon=True).start()


def validate_form_submit(view: dict):
    form_key, vals = ea.forms.read_form(view)
    errs = ea.forms.validate(form_key, vals)
    if not errs and not is_configured():
        errs[ea.forms.get(form_key).fields[-1].key] = '전자결재 봇이 아직 설정되지 않았어요. 관리자에게 알려 주세요.'
    values = view['state']['values']
    return form_key, vals, {ea.field_block_id(values, k): m for k, m in errs.items()}


def process_form_submit(uid: str, form_key: str, vals: dict) -> dict:
    from dashboard.services.purchase_invoice_assign import COST_CATEGORIES
    name = _name(uid)
    vals = {k: v for k, v in vals.items() if not k.startswith('_')}
    fields = {
        'project_code': vals.get('project') or None,
        'amount': ea.forms.amount_of(form_key, vals),
        'grn': ea.forms.grn_of(form_key, vals),
        'content': ea.forms.summary(form_key, vals),
        'urgency': '일반',
        'extra': {'f': vals, 'title': ea.forms.doc_title(form_key, vals, name)},
    }
    if form_key == ea.forms.FORM_EXPENSE:
        cost = {label: v for v, label in COST_CATEGORIES}.get(vals.get('cost'), vals.get('cost'))
        acct = ea.normalize_account(vals.get('account'))
        vals['account'] = acct
        fields.update(cost_category=cost, payee_name=name, bank=vals.get('bank'), account=acct)
        ea.remember_personal('staff', uid, name, vals.get('bank', ''), acct)
    if form_key == ea.forms.FORM_CARDEXP:
        ea.remember_personal('card', uid, name, '', vals.get('card_no', ''))
    row = ea.create(form_key, uid, name, fields)
    _audit(row, 'request', name)
    return start(row)


def handle_form_submit(ack, body, view) -> None:
    """양식 창 제출 — 전자결재 봇·공사봇 공용."""
    form_key, vals, errs = validate_form_submit(view)
    if errs:
        ack(response_action='errors', errors=errs)
        return
    ack()
    uid = (body.get('user') or {}).get('id', '')

    def _bg():
        try:
            process_form_submit(uid, form_key, vals)
        except Exception as exc:
            logger.error(f'{_LOG} {form_key} 요청 처리 실패: {exc}', exc_info=True, extra={'force_alert': True})
            c = eap_client()
            if c:
                _notify(c, uid, f':warning: {ea.forms.get(form_key).label}이 올라가지 않았어요. 잠시 뒤 다시 해 주세요.')
    threading.Thread(target=_bg, daemon=True).start()


def _register_form_handlers(app) -> None:
    """양식 창의 현장 목록·현장 선택·제출 — 창을 연 봇(전자결재 봇·공사봇)마다."""

    @app.options(ea.forms.A_FORM_PROJECT)
    def handle_form_project_options(ack, body):
        try:
            from dashboard.services import purchase_invoice_assign as pinv
            from dashboard.services.project_service import get_project_records
            from dashboard.blueprints.slack_bot import _intake_option_label
            meta = json.loads(((body.get('view') or {}).get('private_metadata')) or '{}')
            query = (body.get('value') or '').strip()
            mgr = None if query else _name((body.get('user') or {}).get('id', ''))
            # 개인경비는 프로젝트0·공통비도, 법인서류는 「현장 없음」(공통비는 비용 말이라 안 맞음, 11-05),
            # 나머지(계약서 등)는 실제 현장만
            specials = {ea.forms.FORM_EXPENSE: None, ea.forms.FORM_CORPDOCS: (pinv.NO_SITE_CODE,)}.get(
                meta.get('form'), ())
            ack(options=pinv.build_project_options(get_project_records() or [], query, manager_name=mgr,
                                                   label_fn=_intake_option_label, specials=specials))
        except Exception as exc:
            logger.error(f'{_LOG} 양식 현장 목록 실패: {exc}', exc_info=True)
            try:
                ack(options=[])
            except Exception:
                pass

    @app.action(ea.forms.A_FORM_PROJECT)
    def handle_form_project_pick(ack, body, client):
        """현장을 고르면 총액·인수자·연락처 등을 공사현황 값으로 채운다(그 칸 block_id 를 바꿔 새 값이 보이게).

        등록증 조회(Drive)가 느릴 수 있어 응답 뒤 백그라운드에서."""
        ack()
        view = body['view']
        code = (body['actions'][0].get('selected_option') or {}).get('value') or ''
        uid = (body.get('user') or {}).get('id', '')

        def _bg():
            try:
                meta = json.loads(view.get('private_metadata') or '{}')
                vals = _form_values(meta['form'], uid, code)
                vals.update({k: v for k, v in ea.forms.read_form(view)[1].items()
                             if v and k not in vals and k != 'project'})   # 펼쳐 둔 칸(보증 등) 유지
                meta['pv'] = int(meta.get('pv') or 0) + 1
                meta['extra_note'] = _form_notes(meta['form'], code)
                client.views_update(view_id=view['id'], view=ea.forms.form_view(meta['form'], meta, vals))
            except Exception as exc:
                logger.error(f'{_LOG} 양식 현장 반영 실패: {exc}', exc_info=True)
        threading.Thread(target=_bg, daemon=True).start()

    @app.action('guarantees')
    def handle_guarantees_pick(ack, body, client):
        """계약서 「보증보험」 고르면 고른 보증마다 보증금율·기간 칸을 펼친다(넣은 값은 block_id 가 같아 유지)."""
        ack()
        try:
            view = body['view']
            meta = json.loads(view.get('private_metadata') or '{}')
            vals = ea.forms.read_form(view)[1]
            if vals.get('project') and not meta.get('project'):
                from dashboard.services.purchase_invoice_assign import project_display
                vals['_project_label'] = project_display(vals['project'], _project_rec(vals['project']))
            client.views_update(view_id=view['id'], hash=view.get('hash'),
                                view=ea.forms.form_view(meta['form'], meta, vals))
        except Exception as exc:
            logger.error(f'{_LOG} 보증보험 칸 펼치기 실패: {exc}', exc_info=True)

    @app.view(ea.forms.CB_FORM)
    def handle_form_submit_view(ack, body, view):
        handle_form_submit(_ack_completing_step(ack, body), body, view)


def post_cardexp_reminder(today=None) -> Optional[str]:
    """매달 첫 영업일 — #전자_결재에 법인카드 지출결의서 안내 + [올리기] 버튼. 보낸 달은 다시 안 보냄. ts 반환."""
    from datetime import date as _date
    from dashboard.services.absent_remind import _is_business_day
    d = today or _date.today()
    first = _date(d.year, d.month, 1)
    while not _is_business_day(first):
        first = first.replace(day=first.day + 1)
    if d != first or not is_configured():
        return None
    try:
        from dashboard.utils.redis_client import get_redis_client
        if not get_redis_client().redis.set(f'eap_cardexp_remind:{d:%Y-%m}', '1', nx=True, ex=40 * 86400):
            return None
    except Exception as exc:
        logger.warning(f'{_LOG} 법인카드 안내 중복 확인 실패(보내지 않음): {exc}')
        return None
    prev = ea.forms.prev_month(d)[0]
    blocks = [
        {'type': 'section', 'text': {'type': 'mrkdwn', 'text':
            f'🧾 *{prev.month}월 법인카드 지출결의서*를 올려 주세요.\n'
            '카드 번호와 사용 기간은 채워져 있어요. 올린 뒤 카드 스레드에 영수증을 붙여 주세요.'}},
        {'type': 'actions', 'elements': [{
            'type': 'button', 'action_id': ea.A_FORM_OPEN, 'value': ea.forms.FORM_CARDEXP, 'style': 'primary',
            'text': {'type': 'plain_text', 'text': '🧾 지출결의서 올리기', 'emoji': True}}]},
    ]
    r = eap_client().chat_postMessage(channel=ea.channel_id(), text=f'{prev.month}월 법인카드 지출결의서', blocks=blocks)
    logger.info(f'{_LOG} 법인카드 지출결의 안내 게시 ({prev:%Y-%m})')
    return r.get('ts')


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


def _hub_view_id(body) -> str:
    """버튼을 바로가기 「전자결재」 창(모달) 안에서 눌렀으면 그 창 id. 채널 메뉴 메시지·앱 홈 탭이면 ''(새 창)."""
    view = body.get('view') or {}
    if (body.get('container') or {}).get('type') == 'view' and view.get('type') == 'modal':
        return view.get('id', '')
    return ''


def _complete_step(body) -> None:
    """워크플로 단계 「전자결재 열기」로 연 창이면 그 단계를 끝낸다(창 닫힘·제출 성공 때). 아니면 아무것도 안 함."""
    exec_id = (body.get('function_data') or {}).get('execution_id')
    token = body.get('bot_access_token')
    if not (exec_id and token):
        return

    def _bg():
        from slack_sdk import WebClient
        try:
            WebClient(token=token).functions_completeSuccess(function_execution_id=exec_id, outputs={})
            logger.info(f'{_LOG} 워크플로 단계 끝냄 ({body.get("type")})')
        except Exception as exc:
            logger.warning(f'{_LOG} 워크플로 단계 끝내기 실패: {exc}')
    threading.Thread(target=_bg, daemon=True).start()


def _push_kw(body) -> Optional[dict]:
    """메뉴 창(모달) 안 버튼이면 그 위에 쌓을 views_push 인자. 워크플로 창 요청엔 trigger_id 대신 pointer 만 온다."""
    if not _hub_view_id(body):
        return None
    if body.get('trigger_id'):
        return {'trigger_id': body['trigger_id']}
    pointer = (body.get('interactivity') or {}).get('interactivity_pointer')
    return {'interactivity_pointer': pointer} if pointer else None


def _ack_completing_step(ack, body):
    """제출 ack 감싸기 — 오류 없이 닫히는 ack 면

    ① 메뉴 위에 쌓인 창이면 메뉴까지 한 번에 닫고(clear — 제출 뒤 메뉴로 돌아가지 않게)
    ② 워크플로 단계도 끝낸다. 오류 ack 면 창이 남으니 그대로."""
    stacked = bool((body.get('view') or {}).get('previous_view_id'))

    def wrapped(*args, **kwargs):
        if not args and not kwargs and stacked:
            kwargs = {'response_action': 'clear'}
        r = ack(*args, **kwargs)
        if kwargs.get('response_action') != 'errors':
            _complete_step(body)
        return r
    return wrapped


def _from_home(body) -> bool:
    return (body.get('view') or {}).get('type') == 'home'


def show_my_requests_in_view(client, view_id: str, uid: str) -> None:
    """바로가기 창의 [📂 내 요청] — DM 대신 그 창 안에 목록."""
    try:
        opened, closed = ea.my_requests(uid)
        links = {r['id']: _card_link(r) for r in opened + closed}
        text = ea.my_requests_text(opened, closed, links)
    except Exception as exc:
        logger.error(f'{_LOG} 내 요청 목록 실패: {exc}', exc_info=True)
        text = ':warning: 내 요청을 불러오지 못했어요. 잠시 뒤 다시 눌러 주세요.'
    try:
        client.views_update(view_id=view_id, view=ea.my_requests_view(text))
    except Exception as exc:
        logger.error(f'{_LOG} 내 요청 창 바꾸기 실패: {exc}', exc_info=True)


def post_my_requests(client, body, uid: str) -> None:
    """[📂 내 요청] — 누른 사람에게만 보이는 진행 중·최근 끝난 요청 목록(11-03). 채널에 남지 않는다."""
    try:
        opened, closed = ea.my_requests(uid)
        links = {r['id']: _card_link(r) for r in opened + closed}
        _ephemeral(client, body, ea.my_requests_text(opened, closed, links))
    except Exception as exc:
        logger.error(f'{_LOG} 내 요청 목록 실패: {exc}', exc_info=True)
        _ephemeral(client, body, ':warning: 내 요청을 불러오지 못했어요. 잠시 뒤 다시 눌러 주세요.')


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
        push = _push_kw(body)
        try:
            open_new_modal(client, body.get('trigger_id', ''), (body.get('user') or {}).get('id', ''),
                           origin='shortcut' if push else 'menu', push=push)
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
            # 프로젝트0·공통비 둘 다(11-02) — 프로젝트0 자재 구매는 비품보다 자금집행이 더 많다
            ack(options=pinv.build_project_options(
                get_project_records() or [], query, manager_name=mgr, label_fn=_intake_option_label))
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
        handle_new_submit(_ack_completing_step(ack, body), body, view)

    _register_invoice_pick(app)
    _register_form_handlers(app)

    def _open_form_from(body, client, form_key):
        try:
            open_form_modal(client, body.get('trigger_id', ''), (body.get('user') or {}).get('id', ''), form_key,
                            push=_push_kw(body))
        except Exception as exc:
            logger.error(f'{_LOG} 양식 창 열기 실패 ({form_key}): {exc}', exc_info=True)

    @app.action(ea.A_FORM_PICK)
    def handle_form_pick(ack, body, client):
        """메뉴 「📋 다른 결재 요청」 — 고른 양식 창."""
        ack()
        form_key = (body['actions'][0].get('selected_option') or {}).get('value') or ''
        if ea.forms.get(form_key):
            _open_form_from(body, client, form_key)
        if _from_home(body):                  # 홈 탭도 고른 값이 남아 같은 양식을 다시 못 고른다 → 새 block_id 로 다시 게시
            threading.Thread(target=reset_home, args=(client, body), daemon=True).start()
        elif not _hub_view_id(body):          # 채널 메뉴 메시지만 고른 값을 되돌린다(창은 이미 양식으로 바뀜)
            threading.Thread(target=reset_select, args=(client, body), daemon=True).start()

    @app.action(re.compile(rf'^{ea.A_FORM_OPEN}(_.+)?$'))
    def handle_form_open(ack, body, client):
        """버튼 value=양식 key (법인카드 월 안내 [지출결의서 올리기]·메뉴 창 양식 버튼 eap_form_open_<key>)."""
        ack()
        form_key = body['actions'][0].get('value') or ''
        if ea.forms.get(form_key):
            _open_form_from(body, client, form_key)

    @app.action(ea.A_MY_REQUESTS)
    def handle_my_requests(ack, body, client):
        ack()
        uid = (body.get('user') or {}).get('id', '')
        view_id = _hub_view_id(body)
        if not view_id and _from_home(body):   # 앱 홈 탭 — 창을 먼저 열고(3초 trigger) 목록으로 바꾼다
            try:
                view_id = client.views_open(trigger_id=body['trigger_id'],
                                            view=_loading_view('내 요청'))['view']['id']
            except Exception as exc:
                logger.error(f'{_LOG} 내 요청 창 열기 실패: {exc}', exc_info=True)
                return
        if view_id:                            # 바로가기 창·홈 탭 — 목록을 창 안에 보여 준다
            threading.Thread(target=show_my_requests_in_view, args=(client, view_id, uid), daemon=True).start()
        else:
            threading.Thread(target=post_my_requests, args=(client, body, uid), daemon=True).start()

    @app.action(ea.A_HUB_BACK)
    def handle_hub_back(ack, body, client):
        ack()
        try:
            client.views_update(view_id=_hub_view_id(body), view=ea.hub_view())
        except Exception as exc:
            logger.error(f'{_LOG} 바로가기 처음으로 실패: {exc}', exc_info=True)

    @app.action(ea.A_NEW_SUPPLY)
    def handle_supply_open(ack, body, client):
        ack()
        push = _push_kw(body)
        try:
            if push:                           # 메뉴 창 위에 쌓기 — [취소]하면 메뉴로
                client.views_push(**push, view=ea.supply_request_view({'origin': 'shortcut'}))
            else:
                client.views_open(trigger_id=body['trigger_id'], view=ea.supply_request_view({'origin': 'menu'}))
        except Exception as exc:
            logger.error(f'{_LOG} 비품 신청 창 열기 실패: {exc}', exc_info=True)

    @app.view(ea.CB_SUPPLY)
    def handle_supply_submit_eap(ack, body, view):
        handle_supply_submit(_ack_completing_step(ack, body), body, view)

    # 전역 바로가기 「전자결재」 하나 — 어느 채널·DM 에서든 입력창 [+]로 메뉴 창을 연다(11-05).
    # 창 안 버튼은 채널 메뉴와 같은 action_id 이고, 핸들러가 창이면 그 창을 바꾼다(_hub_view_id).
    @app.shortcut(ea.SC_HUB)
    def handle_hub_shortcut(ack, body, client):
        ack()
        try:
            client.views_open(trigger_id=body['trigger_id'], view=ea.hub_view())
        except Exception as exc:
            logger.error(f'{_LOG} 바로가기(전자결재) 실패: {exc}', exc_info=True)

    @app.function(ea.FN_OPEN_MENU)
    def handle_open_menu_step(body, inputs, fail):
        """워크플로 단계 「전자결재 열기」 — #전자_결재 워크플로 버튼 → 메뉴 창(11-05).

        단계가 끝나면 슬랙이 그 창의 버튼을 막는다(⚠ 만 뜨고 무반응) → 여기서 끝내지 않고,
        창을 닫거나(view_closed) 제출에 성공할 때(_ack_completing_step) 끝낸다."""
        from slack_sdk import WebClient
        pointer = ((inputs or {}).get('interactivity') or {}).get('interactivity_pointer', '')
        token = (body.get('event') or {}).get('bot_access_token', '')
        try:
            WebClient(token=token).views_open(interactivity_pointer=pointer, view=ea.hub_view())
            logger.info(f'{_LOG} 워크플로 단계(전자결재 열기) 창 열림')
        except Exception as exc:
            logger.error(f'{_LOG} 워크플로 단계(전자결재 열기) 실패: {exc}', exc_info=True)
            fail(error='전자결재 창을 열지 못했어요. 잠시 뒤 다시 눌러 주세요.')

    @app.view_closed(re.compile(r'.*'))
    def handle_view_closed(ack, body):
        """전자결재 창 [닫기]·[취소] — 워크플로로 연 창이 전부 닫힐 때만 단계를 끝낸다.

        메뉴 위에 쌓인 창의 [취소]는 메뉴로 돌아가는 것뿐이라(previous_view_id 있음·is_cleared 아님) 끝내면 안 된다
        — 끝내면 남은 메뉴 창 버튼이 막힌다."""
        ack()
        if body.get('is_cleared') or not (body.get('view') or {}).get('previous_view_id'):
            _complete_step(body)

    @app.event('app_home_opened')
    def handle_home_opened(event, client):
        """왼쪽 앱 목록 → 전자결재 → 홈 탭을 열 때마다 메뉴를 게시(새로 온 직원도 따로 할 일 없음)."""
        if event.get('tab') != 'home':
            return
        try:
            client.views_publish(user_id=event['user'], view=ea.home_view())
        except Exception as exc:
            logger.error(f'{_LOG} 홈 탭 게시 실패: {exc}', exc_info=True)


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

    _register_invoice_pick(app)
    _register_form_handlers(app)

    @app.action(ea.A_PROJECT_FORM_PICK)
    def handle_project_form_pick(ack, body, client):
        """공사확정 카드 「📋 결재 요청」 — 자금집행이면 자금집행 창, 나머지는 현장 고정 양식 창(11-01)."""
        ack()
        sel = (body['actions'][0].get('selected_option') or {}).get('value') or ''
        form_key, _, code = sel.partition('|')   # 선택지 value = '양식|공사코드'
        uid = (body.get('user') or {}).get('id', '')
        try:
            if form_key == ea.FORM_FUND:
                open_new_modal(client, body['trigger_id'], uid, project_code=code, origin='project_card')
            elif ea.forms.get(form_key):
                open_form_modal(client, body['trigger_id'], uid, form_key, project_code=code)
        except Exception as exc:
            logger.error(f'{_LOG} 공사 카드 결재 요청 창 실패 ({sel}): {exc}', exc_info=True)
        threading.Thread(target=reset_select, args=(client, body), daemon=True).start()


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
