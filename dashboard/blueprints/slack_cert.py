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


def _send(client, row: dict, uid: str, role: str) -> bool:
    """카드 DM 한 장. 보냈으면 True — 부르는 쪽이 경영지원에게 한 장도 안 갔는지 알 수 있게(11-24 조용한 실패 점검)."""
    try:
        ch = client.conversations_open(users=uid)['channel']['id']
        r = client.chat_postMessage(channel=ch, text=f"증명서 신청 — {row['requester_name']} {row['ctype']}",
                                    blocks=ct.card_blocks(row, role))
        ct.add_ref(row['id'], ch, r['ts'], role)
        return True
    except Exception as exc:
        logger.warning(f'{_LOG} DM 실패 (#{row["id"]} → {uid}, {role}): {exc}')
        return False


def _refresh(client, row: dict) -> None:
    for ref in ct.refs(row):
        try:
            client.chat_update(channel=ref['ch'], ts=ref['ts'], text=f"증명서 신청 — {row['requester_name']}",
                               blocks=ct.card_blocks(row, ref['role']))
        except Exception as exc:
            logger.warning(f'{_LOG} 카드 고치기 실패 (#{row["id"]}): {exc}')


NOT_NOTIFIED_MSG = '신청은 저장했는데 경영지원에 알림이 안 갔어요. 경영지원에 직접 말해 주세요.'


def submit(client, uid: str, vals: dict) -> dict:
    """신청 저장 + 신청자·경영지원 카드. 돌려주는 row 의 handler_notified 가 False 면 경영지원 카드가 한 장도 안 간 것.
    경영지원 화면엔 대기 중 신청 목록이 없어 카드가 안 가면 아무도 모른 채 영구 대기 → ERROR 알림 +
    부르는 쪽(웹·슬랙)이 신청자에게 직접 말하라고 안내(11-24 조용한 실패 점검). 신청자 카드 실패는 WARNING 그대로."""
    se = _se()
    row = ct.create(uid, se._name(uid), vals)
    _send(client, row, uid, 'requester')
    handlers = tuple(ct.handler_ids() or ())
    sent = sum(1 for h in handlers if _send(client, row, h, 'handler'))
    row['handler_notified'] = sent > 0
    if not sent:
        why = '받을 경영지원이 설정돼 있지 않음' if not handlers else f'경영지원 {len(handlers)}명 모두 DM 실패'
        logger.error(f"{_LOG} 경영지원 카드 못 보냄 — 신청 #{row['id']} {row['requester_name']} {row['ctype']} "
                     f"({why}). 신청자에게 직접 확인해 처리해 주세요.", extra={'force_alert': True})
    logger.info(f"{_LOG} 신청 #{row['id']} {row['requester_name']} {row['ctype']} ({row['purpose']})"
                f" 경영지원 카드 {sent}/{len(handlers)}")
    return row


def submit_from_slack(client, uid: str, vals: dict) -> None:
    """슬랙 신청 창 제출(뒤에서). 저장 실패·경영지원 카드 0장이면 신청자에게 DM 으로 알린다."""
    se = _se()
    try:
        row = submit(client, uid, vals)
    except Exception as exc:
        logger.error(f'{_LOG} 신청 처리 실패: {exc}', exc_info=True, extra={'force_alert': True})
        se._notify(client, uid, '⚠️ 증명서 신청이 저장되지 않았어요. 다시 신청하거나 경영지원에 말해 주세요.')
        return
    if row.get('handler_notified') is False:    # 저장은 됐는데 경영지원에 안 닿음(11-24 조용한 실패 점검)
        se._notify(client, uid, f'⚠️ {NOT_NOTIFIED_MSG}')


def deliver_files(client, row: dict, files: list, failed: list = None) -> list:
    """경영지원이 올린 PDF → 신청자 DM. 보낸 파일 이름들(못 보낸 건 빠짐).
    failed 목록을 주면 슬랙에서 못 받은 파일 이름을 거기에 담는다 — 일부만 가도 '발급 완료'로만 보이던 것
    (11-24 조용한 실패 점검)."""
    se = _se()
    uploads = []
    for f in files:
        name = f.get('name') or 'file'
        data = se._download(f.get('url', ''))
        if data:
            uploads.append({'content': data, 'filename': name, 'title': name})
        else:
            logger.warning(f"{_LOG} 첨부 받기 실패 (#{row['id']}): {name}")
            if failed is not None:
                failed.append(name)
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
    sent, failed = [], []
    if files:
        try:
            sent = deliver_files(client, row, files, failed)
        except Exception as exc:
            logger.error(f'{_LOG} 파일 전달 실패 (#{rid}): {exc}', exc_info=True)
        if not sent:
            return '파일을 신청자에게 보내지 못했어요. 다시 [📎 발급 완료]를 눌러 올려 주세요.'
    new = ct.finish(rid, ct.ISSUED, se._name(uid), note, sent)
    if not new:
        return '이미 처리된 신청이에요.'
    _refresh(client, new)
    if not sent:   # 원본만 주는 경우 — 이 DM 말고는 알 길이 없음(11-24 조용한 실패 점검: 못 보내면 ERROR)
        se._notify_important(client, new['requester_id'],
                             f"📜 {new['ctype']}가 준비됐어요. 경영지원에서 원본을 받아 가세요." + (f" ({note})" if note else ''))
    elif note:
        se._notify(client, new['requester_id'], f"📜 경영지원: {note}")
    logger.info(f'{_LOG} 발급 #{rid} by {new["handled_by"]} 파일 {len(sent)}개')
    if failed:     # 일부만 감 — 발급 완료는 그대로, 못 간 파일은 경영지원이 직접(11-24 조용한 실패 점검)
        logger.warning(f'{_LOG} 발급 #{rid} 파일 {len(failed)}개는 신청자에게 못 보냄: {", ".join(failed)}')
        return (f"발급 완료로 처리했어요. 다만 이 파일은 신청자에게 못 보냈어요: {', '.join(failed)} — "
                f"신청자에게 직접 보내 주세요.")
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
    # 전에 만들어 두고 DM 만 못 보낸 신청이면 새로 발급하지 않고 같은 PDF 를 다시 보낸다 — 예전엔 다시 누르면
    #   발급번호·인사카드 서류가 하나 더 생겼다(11-24 조용한 실패 점검)
    res = ct.issued_for_request(rid)
    resent = bool(res)
    if not res:
        try:
            res = ct.issue_pdf(person['id'], row['ctype'], opts['purpose'], by, ssn=opts['ssn'], request_id=rid,
                               company_key=opts['company'])
        except ct.CertInputError as exc:          # 사람이 고칠 수 있는 문제 — 문구 그대로
            return str(exc)
        except hr.SsnKeyMissing:
            logger.error(f'{_LOG} 바로 발급 실패 (#{rid}) — 주민번호 열쇠(GW_HR_KEY)가 서버에 없음',
                         extra={'force_alert': True})
            return ct.SSN_KEY_MISSING_MSG
        except Exception as exc:                  # PDF·글꼴·저장 실패, 코드 오류 — 원문 대신 일반 문구 + ERROR
            logger.error(f'{_LOG} 바로 발급 실패 (#{rid}): {type(exc).__name__}: {exc}', exc_info=True,
                         extra={'force_alert': True})
            return '바로 발급하지 못했어요. [📎 파일 올려 발급]을 쓰거나 부장님께 말해 주세요(오류 내용은 알림으로 갔어요).'
    try:
        ch = client.conversations_open(users=row['requester_id'])['channel']['id']
        client.files_upload_v2(channel=ch, file_uploads=[{'content': res['pdf'], 'filename': res['file_name'],
                                                          'title': res['file_name']}],
                               initial_comment=f"📜 {ct.company_label(opts['company'])} {row['ctype']}가 발급됐어요. ({row['purpose']}) · 발급번호 {res['doc_no']}")
    except Exception as exc:
        logger.error(f"{_LOG} 바로 발급 #{rid} {res['doc_no']} 은 만들었는데 신청자 DM 을 못 보냄: {exc}", exc_info=True,
                     extra={'force_alert': True})
        return (f"증명서(발급번호 {res['doc_no']})는 만들었는데 신청자에게 보내지 못했어요. "
                "[📄 바로 발급]을 다시 누르면 새로 만들지 않고 보내기만 해요.")
    if resent:
        se._notify(client, uid, f"📜 발급번호 {res['doc_no']}는 이미 만들어 둔 증명서예요 — 새로 만들지 않고 "
                                f"{row['requester_name']} 님에게 다시 보내기만 했어요.")
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
            se._notify_important(client, new['requester_id'], "📜 원본은 경영지원에서 받아 가세요.")
    logger.info(f"{_LOG} 바로 발급 #{rid} {row['requester_name']} {row['ctype']} {res['doc_no']} by {by}"
                + (' (다시 보내기)' if resent else ''))
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

        threading.Thread(target=lambda: submit_from_slack(se.eap_client(), uid, vals), daemon=True).start()

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
            # 예전엔 try 가 없어 예외가 스레드 밖으로 사라졌고, 반려 DM 도 _notify(WARNING)라 못 가도 아무도 몰랐다
            #   (11-24 조용한 실패 점검) — 신청자는 이 DM 말고는 반려를 알 길이 없어 _notify_important(ERROR)
            client = None
            try:
                client = se.eap_client()
                new = ct.finish(int(rid), ct.REJECTED, se._name(uid), reason)
                if not new:
                    se._notify(client, uid, '⚠️ 이미 처리된 신청이에요.')
                    return
                _refresh(client, new)
                se._notify_important(client, new['requester_id'], f"❌ {new['ctype']} 신청이 반려됐어요: {reason}")
            except Exception as exc:
                logger.error(f'{_LOG} 반려 처리 실패 (#{rid}): {exc}', exc_info=True, extra={'force_alert': True})
                if client is not None:
                    se._notify(client, uid, '⚠️ 반려 처리 중 오류가 났어요. 신청 카드를 확인하고, 반려가 안 됐으면 다시 해 주세요.')
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
            try:    # 예전엔 try 가 없어 예외가 스레드 밖으로 사라짐(11-24 조용한 실패 점검)
                new = ct.finish(rid, ct.CANCELLED, se._name(uid), '신청자 취소')
                if not new:
                    se._ephemeral(client, body, '이미 처리된 신청이에요.')
                    return
                _refresh(client, new)
            except Exception as exc:
                logger.error(f'{_LOG} 신청 취소 처리 실패 (#{rid}): {exc}', exc_info=True, extra={'force_alert': True})
                se._ephemeral(client, body, '⚠️ 신청 취소 중 오류가 났어요. 카드가 그대로면 다시 눌러 주세요.')
        threading.Thread(target=_bg, daemon=True).start()
