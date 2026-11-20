# -*- coding: utf-8 -*-
"""증명서 신청 (2026-11-06) — 전자결재 봇, DM 만. 장부·창·카드는 services/certificates.py."""
import json
import logging
import threading

from dashboard.services import certificates as ct

logger = logging.getLogger(__name__)
_LOG = '[SLACK/증명서]'
_AUTO_BUSY = set()                 # 바로 발급 두 번 눌림 막기(같은 신청을 동시에 두 번 만들지 않게)
_AUTO_LOCK = threading.Lock()


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


def auto_issue(client, rid: int, uid: str) -> str:
    """[📄 바로 발급] — 인사카드로 직인 PDF 를 만들어 신청자 DM 으로(11-20). 문제가 있으면 경영지원에게 보낼 문구."""
    from dashboard.services import hr
    se = _se()
    row = ct.get(rid)
    if not row or row['status'] != ct.PENDING:
        return '이미 처리된 신청이에요.'
    person = hr.find_by_slack(row['requester_id'])
    if not person:
        return f"{row['requester_name']} 님 인사카드에 슬랙 연결이 없어 바로 발급을 못 해요. [📎 파일 올려 발급]을 써 주세요."
    me = hr.find_by_slack(uid) or {}
    by = (me.get('login_email') or me.get('work_email') or se._name(uid) or uid).strip()
    opts = ct.request_opts(row)
    if not opts['company']:      # 두 법인 선택 전(11-20 이전) 신청 — 회사를 모르니 바로 발급하지 않는다
        return '발급 회사가 정해지지 않은 신청이에요. [📎 파일 올려 발급]을 쓰거나 다시 신청해 달라고 해 주세요.'
    try:
        res = ct.issue_pdf(person['id'], row['ctype'], opts['purpose'], by, ssn=opts['ssn'], request_id=rid,
                           company_key=opts['company'])
    except (ValueError, RuntimeError, LookupError) as exc:
        return str(exc)
    ch = client.conversations_open(users=row['requester_id'])['channel']['id']
    client.files_upload_v2(channel=ch, file_uploads=[{'content': res['pdf'], 'filename': res['file_name'],
                                                      'title': res['file_name']}],
                           initial_comment=f"📜 {ct.company_label(opts['company'])} {row['ctype']}가 발급됐어요. ({row['purpose']}) · 발급번호 {res['doc_no']}")
    original = (row.get('receive') or '').startswith('직인')
    if original:   # 원본을 원한 신청 — 경영지원이 뽑아 줄 수 있게 같은 파일을 경영지원 DM 에도
        try:
            hch = client.conversations_open(users=uid)['channel']['id']
            client.files_upload_v2(channel=hch, file_uploads=[{'content': res['pdf'], 'filename': res['file_name'],
                                                               'title': res['file_name']}],
                                   initial_comment=f"🖨 {row['requester_name']} 님이 원본({row.get('copies') or 1}부)을 원해요. 뽑아서 전해 주세요.")
        except Exception as exc:
            logger.warning(f'{_LOG} 원본용 파일 전달 실패 (#{rid}): {exc}')
    new = ct.finish(rid, ct.ISSUED, se._name(uid), f"바로 발급 {res['doc_no']}", [res['file_name']])
    if new:
        _refresh(client, new)
        if original:
            se._notify(client, new['requester_id'], f"📜 원본은 경영지원에서 받아 가세요.")
    logger.info(f"{_LOG} 바로 발급 #{rid} {row['requester_name']} {row['ctype']} {res['doc_no']} by {by}")
    return ''


def register_cert_handlers(app) -> None:
    se = _se()

    @app.action(ct.A_AUTO)
    def handle_cert_auto(ack, body, client):
        ack()
        rid = _guard(body, client)
        if not rid:
            return
        uid = body['user']['id']

        with _AUTO_LOCK:
            if rid in _AUTO_BUSY:
                return
            _AUTO_BUSY.add(rid)

        def _bg():
            c = se.eap_client()
            try:
                msg = auto_issue(c, rid, uid)
            except Exception as exc:
                logger.error(f'{_LOG} 바로 발급 실패 (#{rid}): {exc}', exc_info=True, extra={'force_alert': True})
                msg = '바로 발급 중 오류가 났어요. 다시 누르거나 [📎 파일 올려 발급]을 써 주세요.'
            finally:
                with _AUTO_LOCK:
                    _AUTO_BUSY.discard(rid)
            if msg:
                se._notify(c, uid, f'⚠️ {msg}')
        threading.Thread(target=_bg, daemon=True).start()

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
            try:
                msg = issue(client, int(rid), uid, files, note)
            except Exception as exc:
                logger.error(f'{_LOG} 발급 처리 실패 (#{rid}): {exc}', exc_info=True, extra={'force_alert': True})
                msg = '발급 처리 중 오류가 났어요. 다시 [📎 발급 완료]를 눌러 주세요.'
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
