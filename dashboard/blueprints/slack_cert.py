# -*- coding: utf-8 -*-
"""증명서 신청 (2026-11-06) — 전자결재 봇, DM 만. 장부·창·카드는 services/certificates.py."""
import json
import logging
import threading

from dashboard.services import certificates as ct

logger = logging.getLogger(__name__)
_LOG = '[SLACK/증명서]'


def _se():
    from dashboard.blueprints import slack_eapproval as se
    return se


def _send(client, row: dict, uid: str, role: str) -> None:
    try:
        ch = client.conversations_open(users=uid)['channel']['id']
        r = client.chat_postMessage(channel=ch, text=f"증명서 신청 — {row['requester_name']} {row['ctype']}",
                                    blocks=ct.card_blocks(row, role))
        ct.add_ref(row['id'], ch, r['ts'], role)
    except Exception as exc:
        logger.warning(f'{_LOG} DM 실패 (#{row["id"]} → {uid}, {role}): {exc}')


def _refresh(client, row: dict) -> None:
    for ref in ct.refs(row):
        try:
            client.chat_update(channel=ref['ch'], ts=ref['ts'], text=f"증명서 신청 — {row['requester_name']}",
                               blocks=ct.card_blocks(row, ref['role']))
        except Exception as exc:
            logger.warning(f'{_LOG} 카드 고치기 실패 (#{row["id"]}): {exc}')


def submit(client, uid: str, vals: dict) -> dict:
    se = _se()
    row = ct.create(uid, se._name(uid), vals)
    _send(client, row, uid, 'requester')
    for h in ct.handler_ids():
        _send(client, row, h, 'handler')
    logger.info(f"{_LOG} 신청 #{row['id']} {row['requester_name']} {row['ctype']} ({row['purpose']})")
    return row


def deliver_files(client, row: dict, files: list) -> list:
    """경영지원이 올린 PDF → 신청자 DM. 보낸 파일 이름들(못 보낸 건 빠짐)."""
    se = _se()
    uploads = []
    for f in files:
        data = se._download(f.get('url', ''))
        if data:
            uploads.append({'content': data, 'filename': f.get('name') or 'file', 'title': f.get('name') or 'file'})
    if not uploads:
        return []
    ch = client.conversations_open(users=row['requester_id'])['channel']['id']
    client.files_upload_v2(channel=ch, file_uploads=uploads,
                           initial_comment=f"📜 {row['ctype']}가 발급됐어요. ({row['purpose']})")
    return [u['filename'] for u in uploads]


def issue(client, rid: int, uid: str, files: list, note: str) -> str:
    se = _se()
    row = ct.get(rid)
    if not row or row['status'] != ct.PENDING:
        return '이미 처리된 신청이에요.'
    sent = []
    if files:
        try:
            sent = deliver_files(client, row, files)
        except Exception as exc:
            logger.error(f'{_LOG} 파일 전달 실패 (#{rid}): {exc}', exc_info=True)
        if not sent:
            return '파일을 신청자에게 보내지 못했어요. 다시 [📎 발급 완료]를 눌러 올려 주세요.'
    new = ct.finish(rid, ct.ISSUED, se._name(uid), note, sent)
    if not new:
        return '이미 처리된 신청이에요.'
    _refresh(client, new)
    if not sent:   # 원본만 주는 경우
        se._notify(client, new['requester_id'], f"📜 {new['ctype']}가 준비됐어요. 경영지원에서 원본을 받아 가세요."
                   + (f" ({note})" if note else ''))
    elif note:
        se._notify(client, new['requester_id'], f"📜 경영지원: {note}")
    logger.info(f'{_LOG} 발급 #{rid} by {new["handled_by"]} 파일 {len(sent)}개')
    return ''


def register_cert_handlers(app) -> None:
    se = _se()

    @app.action(ct.A_OPEN)
    def handle_cert_open(ack, body, client):
        ack()
        uid = (body.get('user') or {}).get('id', '')
        view = ct.request_view(se._name(uid))
        try:
            push = se._push_kw(body)
            if push:
                client.views_push(**push, view=view)
            else:
                client.views_open(trigger_id=body['trigger_id'], view=view)
        except Exception as exc:
            logger.error(f'{_LOG} 신청 창 열기 실패: {exc}', exc_info=True)

    @app.view(ct.CB_REQ)
    def handle_cert_submit(ack, body, view):
        vals = ct.read_request(view)
        errs = ct.validate(vals)
        if errs:
            ack(response_action='errors', errors=errs)
            return
        se._ack_completing_step(ack, body)()
        uid = body['user']['id']

        def _bg():
            try:
                submit(se.eap_client(), uid, vals)
            except Exception as exc:
                logger.error(f'{_LOG} 신청 처리 실패: {exc}', exc_info=True, extra={'force_alert': True})
                se._notify(se.eap_client(), uid, '⚠️ 증명서 신청이 저장되지 않았어요. 다시 신청하거나 경영지원에 말해 주세요.')
        threading.Thread(target=_bg, daemon=True).start()

    def _guard(body, client) -> int:
        uid = body['user']['id']
        if not ct.can_handle(uid):
            se._ephemeral(client, body, '경영지원만 처리할 수 있어요.')
            return 0
        return int(body['actions'][0]['value'])

    @app.action(ct.A_ISSUE)
    def handle_cert_issue_open(ack, body, client):
        ack()
        rid = _guard(body, client)
        row = ct.get(rid) if rid else None
        if not row:
            return
        if row['status'] != ct.PENDING:
            se._ephemeral(client, body, '이미 처리된 신청이에요.')
            return
        try:
            client.views_open(trigger_id=body['trigger_id'], view=ct.issue_view(row))
        except Exception as exc:
            logger.error(f'{_LOG} 발급 창 열기 실패: {exc}', exc_info=True)

    @app.view(ct.CB_ISSUE)
    def handle_cert_issue_submit(ack, body, view):
        ack()
        rid = json.loads(view.get('private_metadata') or '{}').get('id')
        uid = body['user']['id']
        files = ct.read_files(view)
        note = ((view['state']['values'].get('i_note', {}).get('i_note') or {}).get('value') or '').strip()

        def _bg():
            client = se.eap_client()
            msg = issue(client, int(rid), uid, files, note)
            if msg:
                se._notify(client, uid, f'⚠️ {msg}')
        threading.Thread(target=_bg, daemon=True).start()

    @app.action(ct.A_REJECT)
    def handle_cert_reject_open(ack, body, client):
        ack()
        rid = _guard(body, client)
        row = ct.get(rid) if rid else None
        if row and row['status'] == ct.PENDING:
            try:
                client.views_open(trigger_id=body['trigger_id'], view=ct.reject_view(row))
            except Exception as exc:
                logger.error(f'{_LOG} 반려 창 열기 실패: {exc}', exc_info=True)

    @app.view(ct.CB_REJECT)
    def handle_cert_reject_submit(ack, body, view):
        reason = ((view['state']['values'].get('r_reason', {}).get('r_reason') or {}).get('value') or '').strip()
        if not reason:
            ack(response_action='errors', errors={'r_reason': '사유를 적어 주세요.'})
            return
        ack()
        rid = json.loads(view.get('private_metadata') or '{}').get('id')
        uid = body['user']['id']

        def _bg():
            client = se.eap_client()
            new = ct.finish(int(rid), ct.REJECTED, se._name(uid), reason)
            if not new:
                se._notify(client, uid, '⚠️ 이미 처리된 신청이에요.')
                return
            _refresh(client, new)
            se._notify(client, new['requester_id'], f"❌ {new['ctype']} 신청이 반려됐어요: {reason}")
        threading.Thread(target=_bg, daemon=True).start()

    @app.action(ct.A_CANCEL)
    def handle_cert_cancel(ack, body, client):
        ack()
        rid = int(body['actions'][0]['value'])
        uid = body['user']['id']
        row = ct.get(rid)
        if not row or row['requester_id'] != uid:
            return

        def _bg():
            new = ct.finish(rid, ct.CANCELLED, se._name(uid), '신청자 취소')
            if not new:
                se._ephemeral(client, body, '이미 처리된 신청이에요.')
                return
            _refresh(client, new)
        threading.Thread(target=_bg, daemon=True).start()
