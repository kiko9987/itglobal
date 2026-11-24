# -*- coding: utf-8 -*-
"""입사·퇴사 할 일 (2026-11-20) — 전자결재 봇 DM. 장부·카드 글은 services/hr_tasks.py.

- 인사카드를 저장하면(그룹웨어) 바로, 그리고 매일 09:00 에 맞춰 본다 → 새 할 일은 경영지원 DM 카드([✅ 완료]·[인사카드 열기]).
- 평일 09:00 에는 기한이 3일 안이거나 지난 열린 일을 한 번에 묶어 알린다.
"""
import logging
import threading
from datetime import date
from typing import List, Optional

from dashboard.services import hr_tasks as ht

logger = logging.getLogger(__name__)
_LOG = '[SLACK/인사 할 일]'


def _se():
    from dashboard.blueprints import slack_eapproval as se
    return se


def notify_new(client, rows: List[dict]) -> int:
    sent = 0
    for row in rows:
        uids = ht.handler_ids()
        row_sent = 0
        for uid in uids:
            try:
                ch = client.conversations_open(users=uid)['channel']['id']
                r = client.chat_postMessage(channel=ch, text=f"📋 {ht.title(row)} — {row.get('name') or ''}",
                                            blocks=ht.card_blocks(row))
                ht.add_ref(row['id'], ch, r['ts'])
                sent += 1
                row_sent += 1
            except Exception as exc:
                logger.warning(f'{_LOG} DM 실패 (#{row["id"]} → {uid}): {exc}')
        if uids and not row_sent:
            # 받을 사람 누구에게도 카드가 안 갔다 — 할 일은 생겼는데 아무도 모른다(11-24 조용한 실패 점검)
            logger.error(f"{_LOG} 새 할 일 DM 을 아무에게도 못 보냄 — #{row['id']} {row.get('name') or ''} "
                         f'{ht.title(row)}. 인사카드에서 확인해 주세요', extra={'force_alert': True})
    return sent


def _refresh(client, row: dict) -> None:
    for ref in ht.refs(row):
        try:
            client.chat_update(channel=ref['ch'], ts=ref['ts'], text=f"📋 {ht.title(row)} — {row.get('name') or ''}",
                               blocks=ht.card_blocks(row))
        except Exception as exc:
            logger.warning(f'{_LOG} 카드 고치기 실패 (#{row["id"]}): {exc}')


def sync_and_notify(client=None) -> int:
    """인사카드를 보고 새 할 일을 만들어 DM. 새로 만든 수."""
    rows = ht.sync()
    if rows:
        client = client or _se().eap_client()
        if client:
            notify_new(client, rows)
        elif ht.handler_ids():
            # 봇 토큰이 없어 DM 을 하나도 못 보냄 — 할 일은 생겼는데 아무도 모른다(11-24 조용한 실패 점검)
            logger.error(f'{_LOG} 전자결재 봇 토큰이 없어 새 할 일 DM 을 못 보냄 — '
                         + ', '.join(f"#{r['id']} {r['name']} {ht.title(r)}" for r in rows)
                         + '. 인사카드에서 확인해 주세요', extra={'force_alert': True})
        logger.info(f'{_LOG} 새 할 일 {len(rows)}건: ' + ', '.join(f"{r['name']} {ht.title(r)}" for r in rows))
    return len(rows)


def sync_async() -> None:
    """그룹웨어 인사카드 저장 뒤 — 화면을 기다리게 하지 않게 뒤에서."""
    def _bg():
        try:
            sync_and_notify()
        except Exception as exc:
            logger.error(f'{_LOG} 맞추기 실패: {exc}', exc_info=True)
    threading.Thread(target=_bg, daemon=True).start()


def run_daily(today: Optional[date] = None, client=None) -> None:
    """매일 09:00 — 맞추기 + (평일) 기한 다가온 일 묶어 알리기."""
    today = today or date.today()
    client = client or _se().eap_client()
    sync_and_notify(client)
    from dashboard.services import leave
    if not leave.is_workday(today):
        return
    rows = ht.reminder_rows(today)
    if not rows:
        return
    uids = [u for u in ht.handler_ids() if u]
    if not client:
        if uids:
            # 기한이 다가온·지난 할 일이 있는데 알릴 봇이 없다(11-24 조용한 실패 점검)
            logger.error(f'{_LOG} 전자결재 봇 토큰이 없어 기한 알림 {len(rows)}건을 못 보냄. 인사카드에서 확인해 주세요',
                         extra={'force_alert': True})
        return
    text = ht.reminder_text(rows, today)
    ok = sum(1 for uid in uids if _se()._notify(client, uid, text))
    if uids and not ok:
        # 받을 사람 누구에게도 안 갔다 — 사람별 실패는 _notify 가 WARNING 으로 남김(11-24 조용한 실패 점검)
        logger.error(f'{_LOG} 기한 알림 {len(rows)}건을 아무에게도 못 보냄. 인사카드에서 확인해 주세요',
                     extra={'force_alert': True})
        return
    logger.info(f'{_LOG} 기한 알림 {len(rows)}건')


def register_hr_task_handlers(app) -> None:
    se = _se()

    @app.action(ht.A_DONE)
    def handle_hr_task_done(ack, body, client):
        ack()
        uid = body['user']['id']
        from dashboard.services import certificates as ct
        if not ct.can_handle(uid):
            se._ephemeral(client, body, '경영지원만 처리할 수 있어요.')
            return
        tid = int(body['actions'][0]['value'])

        def _bg():
            # 뒤 스레드라 예외가 stderr 에만 찍히고 누른 사람은 아무 안내도 못 받던 것 — 잡아서 ERROR + 본인 안내
            # (11-24 조용한 실패 점검 25). 완료가 이미 저장됐으면 다시 누를 때 카드만 고친다(complete 는 두 번 눌러도 한 번).
            row = None
            try:
                c = se.eap_client() or client
                row = ht.complete(tid, se._name(uid))
                if row:
                    _refresh(c, row)
                    logger.info(f"{_LOG} 완료 #{tid} {row['name']} {ht.title(row)} by {row['done_by']}")
                else:
                    cur = ht.get(tid)
                    if cur:
                        _refresh(c, cur)
            except Exception as exc:
                logger.error(f'{_LOG} 완료 처리 실패 (#{tid}, {uid}, 완료 저장 {"됨" if row else "안 됨"}): {exc}',
                             exc_info=True, extra={'force_alert': True})
                msg = ('완료는 저장했는데 카드를 못 고쳤어요. 잠시 뒤 다시 눌러 주세요.' if row
                       else '완료 처리를 못 했어요. 잠시 뒤 다시 눌러 주세요.')
                try:
                    se._ephemeral(client, body, msg)      # 위 '경영지원만' 안내와 같은 방식(전자결재 봇 앱의 client)
                except Exception as exc2:
                    logger.error(f'{_LOG} 완료 실패 안내도 못 보냄 (#{tid}, {uid}): {exc2}', extra={'force_alert': True})
        threading.Thread(target=_bg, daemon=True).start()

    @app.action('hr_task_open')
    def handle_hr_task_open(ack):
        ack()      # 링크 단추 — 슬랙이 이벤트를 보내므로 받기만
