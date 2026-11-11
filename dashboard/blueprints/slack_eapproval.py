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
import time
from typing import List, Optional

from dashboard.services import e_approval as ea
from dashboard.services import handwritten_invoice as hw

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
        # 승인·이체 담당 DM 이 안 가면 아무도 모른 채 결재가 멈춘다(퇴사 계정 등) — 관리자 알림(11-06 검수)
        must = role in ('approver', 'final', 'executor')
        log = logger.error if must else logger.warning
        log(f'{_LOG} DM 실패 ({ea.title(row)} → {uid}, {role}): {exc}',
            extra={'force_alert': True} if must else None)


def _notify(client, uid: str, text: str) -> None:
    """알림용 짧은 DM (기록 안 함)."""
    if not uid:
        return
    try:
        ch = client.conversations_open(users=uid)['channel']['id']
        client.chat_postMessage(channel=ch, text=text)
    except Exception as exc:
        logger.warning(f'{_LOG} 알림 DM 실패 ({uid}): {exc}')


def _targets(row: dict) -> list:
    """지금 상태에서 처리할 사람 [(uid, 역할)]."""
    if row['status'] == ea.PENDING:   # KiKO + JW·YG(KiKO 휴가일 수 있어 11-09) — 요청자 본인은 빼고
        out = [(u, 'approver') for u in ea.approver_ids() if u != row.get('requester_id')]
        if ea.needs_final(row.get('form')):
            out += [(u, 'final') for u in ea.final_approver_ids()
                    if u != row.get('requester_id') and u not in ea.approver_ids()]
        return out
    if row['status'] == ea.FINAL:   # 최종 승인 — JW·YG 둘 다 받고 한 명만 누르면 된다(11-09)
        return [(u, 'final') for u in ea.final_approver_ids()]
    if row['status'] == ea.APPROVED:
        return ([(u, 'executor') for u in ea.executor_dm_ids(row)]
                + [(u, 'notice') for u in ea.notice_ids(row)])   # N통장 → SB 에게 알림만(11-06)
    return []


def _dm_round(client, row: dict) -> None:
    """지금 상태에서 처리할 사람에게 DM. 같은 사람이 이미 그 역할 DM 을 받았으면 건너뜀."""
    have = {(r['user'], r['role']) for r in ea.dm_refs(row)}
    targets = _targets(row)
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
            'approve': (f'승인은 {ea.first_names()}만 할 수 있어요.' if uid != row.get('requester_id')
                        else '본인이 올린 결재는 본인이 승인할 수 없어요.'),
            'final': f'최종 승인은 {ea.final_names()}만 할 수 있어요.',
            'reject': '반려는 승인자만 할 수 있어요.',
            'transfer': f'{w} 완료는 경영지원 담당만 누를 수 있어요.',
            'cancel': '요청 취소는 요청한 사람이나 승인자만 할 수 있어요.',
            'revise': '수정 요청은 승인자·경영지원만 할 수 있어요.',
        }[action]
    name = _name(uid)
    new = ea.transition(aid, action, uid, name, reason)
    if not new:
        return f"이미 처리된 결재예요 (지금: {ea.status_label(ea.get(aid) or row)[1]})."
    client = eap_client()
    amt = f"{int(new.get('amount') or 0):,}원"
    t = ea.title(new)
    ini, ex = ea.initial(name), ea.executor_label(name)   # 화면엔 이니셜, ② 는 '경영지원 (SB)'
    # 승인·반려·수정 요청·취소도 경영지원 명단이 누르면 '경영지원 (KiKO)'로 — 이체 완료 줄과 같은 모양(11-09 KiKO)
    who = ex if (uid in ea.approver_ids() or uid in ea.executor_ids()) else ini
    if row.get('status') in (ea.PENDING, ea.FINAL) and uid in ea.final_approver_ids():   # 승인자(JW·YG)로 누른 건 이니셜만
        who = ini
    if action in ('approve', 'final'):
        if action == 'final':
            _thread_log(new, f'✅ {ini} 최종 승인')
        elif any(s['step'] == 'final' and s['result'] == 'auto' and s['actor_id'] == uid for s in ea.steps(aid)):
            _thread_log(new, f'✅ {who} 승인 (최종 승인 겸)')
        else:
            _thread_log(new, f'✅ {who} 승인')
        if new['status'] == ea.FINAL:
            # JW·YG 는 ① 때 받은 DM 이 최종 승인 버튼으로 바뀐다 — 새 DM 대신 그 DM 스레드에 한 줄(알림)
            for r in ea.dm_refs(new):
                if r['role'] == 'final':
                    try:
                        client.chat_postMessage(channel=r['ch'], thread_ts=r['ts'],
                                                text=f'🔔 {who} 승인 — 이제 최종 승인 차례예요. 위 메시지에서 [최종 승인]이나 [반려]를 눌러 주세요.')
                    except Exception as exc:
                        logger.warning(f'{_LOG} 최종 승인 알림 실패 ({t} → {r["user"]}): {exc}')
        if any(s['step'] == 'resubmit' for s in ea.steps(aid)):
            # 수정해서 다시 올린 건 — 예전에 같은 역할 DM 을 받은 사람은 _dm_round 가 건너뛰니 다시 알린다(11-05 점검)
            have = {(r['user'], r['role']) for r in ea.dm_refs(new)}
            link = _card_link(new)
            todo = f'{w}해 주세요.' if new['status'] == ea.APPROVED else '최종 승인해 주세요.'
            for u, role in _targets(new):
                if (u, role) in have and role != 'notice' and not (role == 'final' and new['status'] == ea.FINAL):
                    _notify(client, u, f'🔔 수정해서 다시 올라온 {t}이 승인됐어요. {todo}'
                            + (f' <{link}|카드 보기>' if link else ''))
        _dm_round(client, new)
    elif action == 'transfer' and new.get('form') in ea.TASK_FORMS:
        _thread_log(new, f'✅ {ex} 처리 완료')
        _notify(client, new.get('requester_id'), f'✅ {t} 처리가 끝났어요.')
    elif action == 'transfer' and new.get('form') == ea.forms.FORM_EXPENSE:
        _thread_log(new, f'💸 {ex} 이체 완료 — 개인경비 {amt}')
        _notify(client, new.get('requester_id'), f'💸 {t} 이체가 끝났어요. ({amt})')
    elif action == 'transfer' and new.get('form') == ea.FORM_SUPPLY:
        _thread_log(new, f'🛒 {ex} 구매 완료 — {new.get("content")} {amt}')
        _notify(client, new.get('requester_id'), f'🛒 {t} 구매가 끝났어요. ({new.get("content")} {amt})')
    elif action == 'transfer':
        _thread_log(new, f'💸 {ex} 이체 완료 — {new.get("partner")} {amt}')
        # 실제로 보낸 계좌만 다음 요청에 채운다(11-11 점검)
        ea.remember_payee(new.get('partner') or '', new.get('payee_name'), new.get('bank'), new.get('account'))
        _notify(client, new.get('requester_id'), f'💸 {t} 이체가 끝났어요. ({new.get("partner")} {amt})')
        _origin_log(new, f'💸 {t} 이체 완료 ({ex})')
        if (new.get('evidence') or '') == hw.EVIDENCE:
            threading.Thread(target=save_handwritten, args=(new,), daemon=True).start()
    elif action == 'reject':
        _thread_log(new, f'❌ {who} 반려 — {reason}')
        _notify(client, new.get('requester_id'), f'❌ {t} 반려됐어요. 사유: {reason}')
        _origin_log(new, f'❌ {t} 반려 ({who}) — {reason}')
    elif action == 'cancel':
        _thread_log(new, f'⚪ {ini if uid == new.get("requester_id") else who} 요청 취소')
        if uid != new.get('requester_id'):
            _notify(client, new.get('requester_id'), f'⚪ {t} 요청이 취소됐어요 ({ini}).')
    elif action == 'revise':
        _thread_log(new, f'✏️ {who} 수정 요청 — {reason}')
        _origin_log(new, f'✏️ {t} 수정 요청 ({who}) — {reason}')
        # 요청자 DM 에 [✏️ 수정하기] — 두 번째 수정 요청이면 예전 DM 은 refresh 가 다시 버튼을 띄운다
        if (new.get('requester_id'), 'requester') not in {(r['user'], r['role']) for r in ea.dm_refs(new)}:
            _send_dm(client, aid, new.get('requester_id'), 'requester')
        else:
            _notify(client, new.get('requester_id'), f'✏️ {t} 수정 요청이 왔어요: {reason}')
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


# ── 첨부 파일 (11-05) — 창 「첨부 파일」에 올린 것을 결재 카드 스레드로 옮겨 올린다 ──
def _download(url: str) -> Optional[bytes]:
    """슬랙 파일 받기 — 창을 연 봇(전자결재·공사봇)·계산서 글 봇 토큰 중 되는 것으로."""
    import requests
    for env in ('SLACK_EAPPROVAL_BOT_TOKEN', 'SLACK_PROJECT_BOT_TOKEN', 'SLACK_INVOICE_BOT_TOKEN'):
        tok = os.getenv(env, '').strip()
        if not (tok and url):
            continue
        try:
            r = requests.get(url, headers={'Authorization': f'Bearer {tok}'}, timeout=60)
            # 권한이 없으면 슬랙이 200 + 로그인 HTML 을 준다 — 내용 형식으로 거른다
            if r.ok and 'text/html' not in r.headers.get('Content-Type', ''):
                return r.content
        except Exception as exc:
            logger.warning(f'{_LOG} 첨부 받기 실패 ({env}): {exc}')
    return None


def post_files(row: dict, files: list, who: str, label: str = '') -> int:
    """첨부를 카드 스레드에 올린다. 올린 개수. 못 올린 게 있으면 스레드에 '직접 올려 달라' 안내."""
    if not files or not (row.get('card_channel') and row.get('card_ts')):
        return 0
    uploads, failed, ok_refs = [], [], []
    for f in files:
        data = _download(f.get('url', ''))
        if data:
            uploads.append({'content': data, 'filename': f.get('name') or 'file', 'title': f.get('name') or 'file'})
            ok_refs.append(f)
        else:
            failed.append(f.get('name') or 'file')
    try:   # 이체 완료 때 수기계산서를 드라이브에 넣으려고 원본 위치를 기억(11-09)
        ea.add_extra_files(row['id'], ok_refs)
    except Exception as exc:
        logger.warning(f'{_LOG} 첨부 위치 기억 실패 ({ea.title(row)}): {exc}')
    if uploads:
        try:
            eap_client().files_upload_v2(channel=row['card_channel'], thread_ts=row['card_ts'],
                                         file_uploads=uploads, initial_comment=label or f'📎 {ea.initial(who)} 첨부')
        except Exception as exc:
            logger.warning(f'{_LOG} 첨부 올리기 실패 ({ea.title(row)}): {exc}')
            failed += [u['filename'] for u in uploads]
            uploads = []
    if failed:
        _thread_log(row, f"⚠️ 첨부 {len(failed)}개를 옮기지 못했어요 ({', '.join(failed)}). 이 스레드에 직접 올려 주세요.")
        _notify(eap_client(), row.get('requester_id'),
                f"⚠️ {ea.title(row)} 첨부 {len(failed)}개를 카드에 붙이지 못했어요. 카드 스레드에 직접 올려 주세요.")
    return len(uploads)


def save_handwritten(row: dict) -> List[str]:
    """수기계산서 자금집행 이체 완료 → 계산서 사진을 ★운반비 수기세금계산서\\{연도}년 에(기존 이름 규칙).
    사진 = 창에서 올린 첨부 + #영업_관리 계산서 글 사진. 결과는 카드 스레드에 한 줄."""
    try:
        refs = list(ea.extra(row).get('files') or [])
        seen = {f.get('url') for f in refs}
        refs += [f for f in invoice_post_files(row.get('origin_channel') or '', row.get('origin_ts') or '')
                 if f.get('url') not in seen]
        files = []
        for f in refs:
            data = _download(f.get('url', ''))
            if data:
                files.append({'name': f.get('name') or 'file', 'content': data})
        md = ''
        if row.get('origin_ts'):
            from dashboard.services import purchase_invoice_assign as pinv
            md = (pinv.get_assignment(row['origin_channel'], row['origin_ts']) or {}).get('date_md') or ''
        d = hw.invoice_date(md, hw.base_date_of(row))
        saved = hw.save(row.get('grn') or '', d, row.get('partner') or '', files)
    except Exception as exc:
        logger.error(f'{_LOG} 수기계산서 드라이브 저장 실패 ({ea.title(row)}): {exc}', exc_info=True,
                     extra={'force_alert': True})
        _thread_log(row, '⚠️ 수기계산서를 드라이브에 넣지 못했어요. ★운반비 수기세금계산서 폴더에 직접 넣어 주세요.')
        return []
    if saved:
        _thread_log(row, f"🗂 수기계산서 드라이브 저장 — ★운반비 수기세금계산서 › {d.year}년 › {', '.join(saved)}\n"
                         '엑셀 장부(수기세금계산서.xlsx)는 직접 적어 주세요.')
    elif (row.get('grn') or '')[:1].upper() in ('G', 'R'):
        _thread_log(row, '⚠️ 수기계산서 사진을 찾지 못해 드라이브에 넣지 못했어요. 이 스레드에 사진을 올리고 폴더에 직접 넣어 주세요.')
    return saved


def _post_files_bg(row: dict, files: list, who: str) -> None:
    if files:
        threading.Thread(target=post_files, args=(row, files, who), daemon=True).start()


def invoice_post_files(ch: str, ts: str) -> list:
    """#영업_관리 계산서 글에 붙은 파일(계산서 사진) [{name, url}] — 계산서 봇으로 읽는다."""
    c = _invoice_client()
    if not (c and ch and ts):
        return []
    try:
        msgs = c.conversations_history(channel=ch, latest=ts, inclusive=True, limit=1).get('messages') or []
    except Exception as exc:
        logger.warning(f'{_LOG} 계산서 글 읽기 실패 ({ch} {ts}): {exc}')
        return []
    m = next((x for x in msgs if x.get('ts') == ts), None) or {}
    return [{'name': f.get('name') or 'invoice', 'url': f.get('url_private_download') or f.get('url_private') or ''}
            for f in m.get('files') or [] if f.get('url_private') or f.get('url_private_download')]


def post_invoice_files(row: dict) -> int:
    """자금집행 — 고른 계산서 글의 사진을 카드 스레드에 그대로 붙인다(요청자가 따로 첨부 안 해도, 11-05)."""
    files = invoice_post_files(row.get('origin_channel') or '', row.get('origin_ts') or '')
    return post_files(row, files, '', label='📄 계산서 (#영업_관리 계산서 글에서)') if files else 0


# ── 수정 요청 → 수정하기 (11-05) ─────────────────────────────────────
_FIELD_NAMES = {'project_code': '현장', 'cost_category': '원가 분류', 'content': '내용', 'evidence': '증빙',
                'grn': '사업자 구분', 'partner': '거래처', 'payee_name': '예금주', 'bank': '은행',
                'account': '계좌', 'amount': '금액', 'urgency': '집행 시기', 'memo': '비고',
                'link': '상품 링크', 'qty': '수량', 'unit_price': '단가', 'reason': '사유', 'title': '결재 이름'}
# 양식 결재는 같은 변경이 장부 칸과 양식 칸(f)에 둘 다 잡힌다 — 장부 칸 쪽은 기록에서 뺀다(11-05 점검)
_FORM_DUP_COLUMNS = ('project_code', 'amount', 'content', 'grn', 'cost_category', 'payee_name', 'bank', 'account')
_SECRET_KEYS = ('bank', 'account')   # 개인경비 계좌 — 채널 기록엔 '바뀜'만(경영지원 DM 에만 보인다)


def _changes_text(form: str, changed: dict) -> str:
    """'금액 50,000 → 55,000 · 계좌 …' — 스레드 기록용. 양식 칸은 양식 이름표로."""
    labels = dict(_FIELD_NAMES)
    field_keys = set()
    if form in ea.forms.FORMS:
        field_keys = {f.key for f in ea.forms.FORMS[form].fields}
        labels.update({f.key: f.label.split(' (')[0] for f in ea.forms.FORMS[form].fields})
    parts = []
    for k, (a, b) in changed.items():
        # 장부 칸 중복은 양식에 같은 이름 칸이 없을 때만 뺀다(개인경비 amount·bank·account 는 양식 칸과 이름이 같다)
        dup = form in ea.forms.FORMS and k in _FORM_DUP_COLUMNS and k not in field_keys
        if k in ('extra', 'payee_saved') or k.startswith('_') or dup:
            continue
        if form == ea.forms.FORM_EXPENSE and k in _SECRET_KEYS:
            parts.append(f'{labels.get(k, k)} 바뀜 (경영지원 DM 에서 확인)')
            continue

        def fmt(v):
            if k in ('amount', 'qty', 'unit_price'):
                return f'{int(v or 0):,}'
            if isinstance(v, list):
                return ', '.join(map(str, v)) or '-'
            return str(v or '-')
        parts.append(f'{labels.get(k, k)} {fmt(a)} → {fmt(b)}')
    return ' · '.join(dict.fromkeys(parts[:8])) or '바뀐 칸 없음'


def open_edit_modal(client, trigger_id: str, aid: int, uid: str) -> Optional[str]:
    """[✏️ 수정하기] — 올린 값이 채워진 같은 양식 창. 못 열면 사용자에게 보여 줄 문구."""
    row = ea.get(aid)
    if not row or row['status'] != ea.REVISE:
        return '지금 고칠 수 있는 요청이 아니에요.'
    if not ea.can_act('edit', row, uid):
        return '요청한 사람만 고칠 수 있어요.'
    if row['form'] not in (ea.FORM_FUND, ea.FORM_SUPPLY) and row['form'] not in ea.forms.FORMS:
        return '이 요청은 고칠 수 없어요. 취소하고 다시 올려 주세요.'
    # 현장 이름은 공사현황 조회라 3초(trigger)를 넘길 수 있다 — 불러오는 창을 먼저 열고 바꾼다
    view_id = client.views_open(trigger_id=trigger_id, view=_loading_view('수정'))['view']['id']

    def _bg():
        try:
            client.views_update(view_id=view_id, view=edit_view(row))
        except Exception as exc:
            logger.error(f'{_LOG} 수정 창 채우기 실패 ({ea.title(row)}): {exc}', exc_info=True)
            try:   # '불러오는 중'에 갇히지 않게 안내로 바꾼다(11-05 점검)
                client.views_update(view_id=view_id, view={
                    'type': 'modal', 'title': {'type': 'plain_text', 'text': '수정'},
                    'close': {'type': 'plain_text', 'text': '닫기'}, 'notify_on_close': True,
                    'blocks': [{'type': 'section', 'text': {'type': 'mrkdwn', 'text':
                                ':warning: 수정 창을 불러오지 못했어요. 닫고 다시 눌러 주세요. '
                                '계속 안 되면 카드 스레드에 고칠 내용을 남겨 주세요.'}}]})
            except Exception:
                pass
    threading.Thread(target=_bg, daemon=True).start()
    return None


def edit_view(row: dict) -> dict:
    """수정 창 — 양식별로 올린 값을 채운 같은 창."""
    from dashboard.services.purchase_invoice_assign import project_display
    aid, form = row['id'], row['form']
    code = row.get('project_code') or ''
    if form == ea.FORM_FUND:
        meta = {'edit': aid, 'origin': 'edit', 'partner': row.get('partner') or '',
                'payee': {k: row.get(k) or '' for k in ('payee_name', 'bank', 'account')}}
        if row.get('origin_ts'):   # 계산서 글로 올린 결재 — 거래처·금액은 그 계산서 값 그대로
            meta['inv'] = {'ch': row['origin_channel'], 'ts': row['origin_ts'],
                           'parsed': {'date_md': '', 'grn': row.get('grn') or '', 'amount': int(row.get('amount') or 0),
                                      'partner': row.get('partner') or ''}}
        prefill = {k: row.get(k) for k in ea.EDIT_COLUMNS}
        label = project_display(code, _project_rec(code)) if code else ''
        prefill['_project_label'] = label
        if row.get('origin_ts') and code:   # 계산서 글 현장 지정과 어긋나지 않게 현장도 고정(11-05 점검)
            meta['project'] = code
            return ea.new_request_view(meta, code, label, prefill=prefill)
        return ea.new_request_view(meta, prefill=prefill)
    if form == ea.FORM_SUPPLY:
        return ea.supply_request_view({'edit': aid}, ea.supply_prefill(row))
    vals = dict(ea.extra(row).get('f') or {})
    if code:
        vals['project'] = code
        vals['_project_label'] = project_display(code, _project_rec(code))
    return ea.forms.form_view(form, {'form': form, 'edit': aid, 'pv': 1}, vals,
                              project_label=vals.get('_project_label', ''))


def process_edit(uid: str, aid: int, fields: dict, files: list, docs: Optional[list] = None) -> Optional[dict]:
    """고친 값으로 다시 올림 → 카드·DM 갱신, 스레드에 바뀐 칸, 승인자·경영지원에게 다시 알림."""
    name = _name(uid)
    fields = {k: v for k, v in fields.items() if not k.startswith('_') and k not in ('origin_channel', 'origin_ts')}
    res = ea.resubmit(aid, uid, name, fields)
    if not res:
        _notify(eap_client(), uid, '⚠️ 이미 처리된 요청이라 고친 내용을 올리지 못했어요. '
                                   '같이 올린 첨부 파일도 붙지 않았어요.')
        return None
    row, changed = res
    client = eap_client()
    _thread_log(row, f'✏️ {ea.initial(name)} 수정해서 다시 올림 — {_changes_text(row["form"], changed)}')
    refresh(row)
    who = [x for x in _targets(row) if x[1] != 'notice']
    have = {(r['user'], r['role']) for r in ea.dm_refs(row)}
    link = _card_link(row)
    for u, role in who:
        if (u, role) in have:   # 이미 받은 DM 은 refresh 로 고쳐졌다 — 다시 왔다는 알림만
            _notify(client, u, f'✏️ {ea.title(row)} 수정돼서 다시 올라왔어요.' + (f' <{link}|카드 보기>' if link else ''))
        else:
            _send_dm(client, aid, u, role)
    _audit(row, 'resubmit', name, _changes_text(row['form'], changed))
    _post_folder_docs_bg(row, docs or [])   # 수정 창에서 고른 공사 폴더 서류도(11-05 점검)
    _post_files_bg(row, files, name)
    logger.info(f'{_LOG} {ea.title(row)} 수정 후 다시 올림 ({name}): {changed}')
    return row


def _revise_view(aid: int) -> dict:
    row = ea.get(aid) or {}
    return {
        'type': 'modal', 'callback_id': ea.CB_REVISE,
        'title': {'type': 'plain_text', 'text': '수정 요청'},
        'submit': {'type': 'plain_text', 'text': '돌려보내기'},
        'close': {'type': 'plain_text', 'text': '취소'},
        'private_metadata': json.dumps({'aid': aid}),
        'blocks': [
            {'type': 'context', 'elements': [{'type': 'mrkdwn', 'text':
             f"{ea.title(row)} · {ea.initial(row.get('requester_name'))}님에게 돌려보내요. "
             '고쳐서 다시 올리면 같은 카드가 바뀌어요.'}]},
            {'type': 'input', 'block_id': 'reason', 'label': {'type': 'plain_text', 'text': '고칠 내용'},
             'element': {'type': 'plain_text_input', 'action_id': 'reason', 'multiline': True, 'max_length': 500,
                         'placeholder': {'type': 'plain_text', 'text': '예: 계좌번호가 계산서와 달라요. 다시 확인해 주세요.'}}}],
    }


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
    ea.mark_payee_saved(fields)   # 기억해 둔 계좌 그대로면 카드에 표시
    # 거래처 계좌는 이체 완료 때 기억한다(11-11 점검) — 잘못 친 계좌가 반려돼도 다음 요청에 채워지던 것
    row = ea.create(ea.FORM_FUND, uid, name, fields)
    _audit(row, 'request', name)
    row = start(row)
    threading.Thread(target=post_invoice_files, args=(row,), daemon=True).start()   # 계산서 사진 → 카드 스레드
    return row


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
                # 계산서 없이 먼저 올린 같은 건이 있으면 목록에서 뺀다 — [📍 현장 지정]이 그 결재에 연결한다(11-06)
                if ea.find_unlinked_match(parsed.get('grn') or '', int(parsed.get('amount') or 0),
                                          parsed.get('partner') or ''):
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
            meta.pop('inv_link', None)
            meta.pop('inv_files', None)
            if meta['inv']:   # 창에 '계산서 사진은 자동으로 붙어요' + 글 링크(첨부 칸은 미리 채울 수 없다, 11-05)
                inv = meta['inv']
                meta['inv_files'] = [f['name'] for f in invoice_post_files(inv['ch'], inv['ts'])][:3]
                try:
                    meta['inv_link'] = _invoice_client().chat_getPermalink(
                        channel=inv['ch'], message_ts=inv['ts']).get('permalink', '')
                except Exception as exc:
                    logger.debug(f'{_LOG} 계산서 글 링크 실패(무시): {exc}')
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


_GRN_WARNED: dict = {}   # (창 id, 현장, 사업자) → 시각. 서버 한 개라 메모리로 충분(Redis 장애에도 동작)


def _grn_warn_once(view_id: str, code: str, grn: str) -> bool:
    """처음이면 True(경고를 띄움), 같은 창에서 같은 조합을 다시 내면 False(통과)."""
    now = time.time()
    for k in [k for k, t in _GRN_WARNED.items() if now - t > 3600]:
        _GRN_WARNED.pop(k, None)
    key = (view_id, code or '', grn or '')
    if not view_id or key in _GRN_WARNED:
        return False
    _GRN_WARNED[key] = now
    return True


def validate_new_submit(view: dict):
    """(fields, errors) — 제출 직후 동기 검사(ack 에 errors 를 실어야 해서)."""
    f = ea.read_new_request(view)
    errs = ea.complete_new_request(f)
    if not errs and not is_configured():
        errs['content'] = '전자결재 봇이 아직 설정되지 않았어요. 관리자에게 알려 주세요.'
    values = view['state']['values']
    if not errs:
        # 현장 코드 앞글자(G/R)와 고른 사업자가 다르면 한 번 묻는다 — 5일 12건 중 3건이 이것(11-11 점검).
        # 정말 다른 사업자로 보내는 경우가 있어 막지는 않고, 같은 창에서 한 번 더 누르면 통과
        msg = ea.grn_mismatch(f.get('project_code'), f.get('grn'))
        if msg and _grn_warn_once(view.get('id', ''), f.get('project_code'), f.get('grn')):
            bid = 'grn' if 'grn' in values else ('invoice' if 'invoice' in values else 'project')
            errs[bid] = msg
    # 오류는 화면에 있는 block_id 로(계좌 칸은 다시 채우면 payee_name_2 처럼 바뀐다)
    errs = {ea.field_block_id(values, k) if k in ('payee_name', 'bank', 'account') else k: m
            for k, m in errs.items()}
    return f, errs


def process_new_submit(uid: str, f: dict) -> dict:
    name = _name(uid)   # 거래처 계좌 기억은 이체 완료 때(apply_action) — 승인 전 계좌가 기억되지 않게(11-11)
    if f.get('_edit'):   # [✏️ 수정하기] 로 고친 것 — 같은 결재·카드(계산서 글 현장 지정은 처음에 끝났다)
        # '예전 계좌 자동 채움' 표시는 고친 계좌 기준으로 다시 — 다른 계좌로 고쳤는데 표시가 남지 않게(11-05 점검)
        f['extra'] = {**(f.get('extra') or {}), 'payee_saved': bool((f.get('extra') or {}).get('payee_saved'))}
        return process_edit(uid, f['_edit'], f, f.get('_files') or [])
    row = ea.create(ea.FORM_FUND, uid, name, f)
    _audit(row, 'request', name)
    row = start(row)
    if row.get('origin_ts'):
        threading.Thread(target=post_invoice_files, args=(row,), daemon=True).start()
    _post_files_bg(row, f.get('_files') or [], name)
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
    if f.get('_edit'):
        return process_edit(uid, f['_edit'], f, f.get('_files') or [])
    row = ea.create(ea.FORM_SUPPLY, uid, name, f)
    _audit(row, 'request', name)
    row = start(row)
    _post_files_bg(row, f.get('_files') or [], name)
    return row


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
        if form_key == ea.forms.FORM_INSURANCE:
            fill_insurance_from_contract(vals, project_code)
        if form_key in ea.forms.FOLDER_DOC_FORMS:
            vals['_docs'] = folder_docs(project_code, precheck=(form_key == ea.forms.FORM_CONTRACT))
    return vals


# ── 공사 폴더 서류 (11-05) — 「견적서 및 계약서 폴더」 파일을 창에서 골라 카드 스레드에 링크로 ──
_FINAL_RE = re.compile(r'확정|최종')
_QUOTE_RE = re.compile(r'견적|내역서')
_DOC_EXT = ('.pdf', '.png', '.jpg', '.jpeg')


def _doc_rank(name: str) -> int:
    """0 = 확정·최종 견적서, 1 = 확정·최종(도면 등), 2 = 나머지. 같은 순위 안은 최근 고친 순."""
    if _FINAL_RE.search(name):
        # 견적·내역서 글자가 있거나, 도면·사진이 아닌 PDF 면 확정 견적서로 본다('최종_현장사진.jpg' 는 아님, 11-05 점검)
        if _QUOTE_RE.search(name) or (name.lower().endswith('.pdf') and not re.search(r'도면|배치도|사진', name)):
            return 0
        return 1
    return 2


def folder_docs(code: str, precheck: bool = False) -> list:
    """공사 폴더의 PDF·사진 [{id, name, note, checked}] — 확정·최종 붙은 것 먼저, 그다음 최근 고친 순.
    precheck 면 확정·최종 붙은 첫 파일과 사업자등록증을 미리 고른다(계약서 작성 요청)."""
    try:
        from dashboard.services import business_license_handler as bl
        fid = bl._project_folder_id_for(code)
        if not fid:
            return []
        items = bl._get_drive().files().list(
            q=f"'{fid}' in parents and trashed=false and mimeType!='application/vnd.google-apps.folder'",
            fields='files(id,name,modifiedTime)', pageSize=100, orderBy='modifiedTime desc',
            supportsAllDrives=True, includeItemsFromAllDrives=True).execute().get('files', [])
    except Exception as exc:
        logger.warning(f'{_LOG} 공사 폴더 서류 조회 실패 ({code}): {exc}')
        return []
    files = [f for f in items if f['name'].lower().endswith(_DOC_EXT) and '사업자등록증' not in f['name']]
    files.sort(key=lambda f: _doc_rank(f['name']))   # 안정 정렬 — 같은 순위 안은 최근 고친 순 그대로
    docs = []
    try:
        from dashboard.services.business_license_handler import get_license_state
        st = get_license_state(code) or {}
        if st.get('view_url') and st.get('source') in ('own', 'reuse'):
            docs.append({'id': ea.forms.LICENSE_DOC, 'name': '사업자등록증',
                         'note': '이 공사 폴더' if st['source'] == 'own' else '같은 사업자 다른 공사 폴더',
                         'checked': precheck})
    except Exception as exc:
        logger.warning(f'{_LOG} 등록증 상태 조회 실패 ({code}): {exc}')
    picked = False
    for f in files[:10 - len(docs)]:
        pick = precheck and not picked and _doc_rank(f['name']) == 0   # 확정·최종 견적서 하나만 미리 고름
        docs.append({'id': f['id'], 'name': f['name'], 'note': (f.get('modifiedTime') or '')[:10], 'checked': pick})
        picked = picked or pick
    return docs


_DOC_MAX_BYTES = 50 * 1024 * 1024   # 이보다 큰 파일은 붙이지 않고 이름만(폴더에서 열기)


def _folder_doc_file(code: str, d: dict) -> Optional[dict]:
    """고른 공사 폴더 서류 하나를 드라이브에서 받아 {'content','filename','title'}. 못 받으면 None."""
    import re as _re
    from dashboard.services import business_license_handler as bl
    fid = d.get('id') or ''
    if fid == ea.forms.LICENSE_DOC:   # 사업자등록증은 등록증 상태의 열람 링크에서 파일 ID 를 꺼낸다
        m = _re.search(r'/d/([^/]+)', (bl.get_license_state(code) or {}).get('view_url') or '')
        fid = m.group(1) if m else ''
    if not fid:
        return None
    drive = bl._get_drive()
    meta = drive.files().get(fileId=fid, fields='name,size', supportsAllDrives=True).execute()
    if int(meta.get('size') or 0) > _DOC_MAX_BYTES:
        return None
    name = meta.get('name') or d.get('name') or 'file'
    return {'content': bl._download_drive_file(drive, fid), 'filename': name, 'title': name}


def post_folder_docs(row: dict, docs: list) -> None:
    """고른 공사 폴더 서류를 카드 스레드에 실제 파일로 붙인다(11-09 — 휴대폰에서도 바로 열리게).
    맨 위 글은 탐색기로 여는 폴더 링크(itgfolder://, 공사확정 카드와 같은 방식). 못 받은 파일은 이름만 남긴다."""
    if not docs:
        return
    code = row.get('project_code') or ''
    folder = ''
    try:
        from dashboard.services import business_license_handler as bl
        folder = bl._project_folder_id_for(code) or ''
    except Exception as exc:
        logger.warning(f'{_LOG} 공사 폴더 찾기 실패 ({code}): {exc}')
    head = f'📁 공사 폴더 서류 · <itgfolder://{folder}|탐색기에서 열기>' if folder else '📁 공사 폴더 서류'
    uploads, failed = [], []
    for d in docs:
        label = '사업자등록증' if d['id'] == ea.forms.LICENSE_DOC else d['name']
        try:
            f = _folder_doc_file(code, d)
        except Exception as exc:
            logger.warning(f'{_LOG} 공사 폴더 서류 받기 실패 ({code} {label}): {exc}')
            f = None
        (uploads.append(f) if f else failed.append(label))
    note = ('\n' + '\n'.join(f'• {n} — 못 붙였어요. 폴더에서 열어 주세요.' for n in failed)) if failed else ''
    if uploads and row.get('card_channel') and row.get('card_ts'):
        try:
            eap_client().files_upload_v2(channel=row['card_channel'], thread_ts=row['card_ts'],
                                         file_uploads=uploads, initial_comment=head + note)
            return
        except Exception as exc:
            logger.warning(f'{_LOG} 공사 폴더 서류 올리기 실패 ({ea.title(row)}): {exc}')
            failed = [u['filename'] for u in uploads] + failed
    # 하나도 못 붙였으면 예전처럼 폴더 링크 + 파일 이름
    _thread_log(row, head + '\n' + '\n'.join(f'• {n}' for n in failed))


def _post_folder_docs_bg(row: dict, docs: list) -> None:
    if docs:   # 드라이브에서 받느라 몇 초 걸릴 수 있어 창 제출을 붙잡지 않게
        threading.Thread(target=post_folder_docs, args=(row, docs), daemon=True).start()


def post_contract_link(row: dict) -> None:
    """보증보험 신청 — 그 현장 계약서 결재 카드 링크(경영지원이 완성한 계약서가 그 카드에 있다)."""
    rows = ea.find_by_project(ea.forms.FORM_CONTRACT, row.get('project_code') or '') if row.get('project_code') else []
    if rows:
        link = _card_link(rows[0])
        if link:
            _thread_log(row, f'📑 이 현장 계약서: <{link}|{ea.title(rows[0])}>')


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


def contract_terms(code: str) -> dict:
    """이 현장 계약서(전자결재)에 적은 보증 조건 {보증 종류: (보증금율, 기간)} — 보증보험 신청 칸 채우기용(11-05)."""
    if not code:
        return {}
    rows = ea.find_by_project(ea.forms.FORM_CONTRACT, code)
    f = (ea.extra(rows[0]).get('f') or {}) if rows else {}
    out = {}
    for g in f.get('guarantees') or []:
        k = ea.forms.GUARANTEE_KEYS.get(g)
        if k:
            out[g] = (str(f.get(f'{k}_rate') or '').strip(), str(f.get(f'{k}_years') or '').strip())
    return out


def fill_insurance_from_contract(vals: dict, code: str, override: bool = False) -> None:
    """보증보험 신청 — 계약서에 적은 보증금율·기간을 채운다. 종류를 안 골랐고 계약서 보증이 하나면 종류도.
    override=True(보증 종류를 바꿔 고름)면 이미 채운 값도 그 종류 값으로 바꾼다."""
    terms = contract_terms(code)
    if not terms:
        return
    if not vals.get('kind') and len(terms) == 1:
        vals['kind'] = next(iter(terms))
    rate, years = terms.get(vals.get('kind') or '', ('', ''))
    for key, v in (('rate', rate), ('years', years)):
        if v and (override or not vals.get(key)):
            vals[key] = v


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
            try:   # '불러오는 중'에 멈춰 있지 않게(11-06 검수 — 수정 창과 같게)
                client.views_update(view_id=view_id, view={
                    'type': 'modal', 'title': {'type': 'plain_text', 'text': form.label[:24]},
                    'close': {'type': 'plain_text', 'text': '닫기'}, 'notify_on_close': True,
                    'blocks': [{'type': 'section', 'text': {'type': 'mrkdwn', 'text':
                                ':warning: 창을 불러오지 못했어요. 닫고 다시 눌러 주세요.'}}]})
            except Exception:
                pass
    threading.Thread(target=_bg, daemon=True).start()


def validate_form_submit(view: dict):
    form_key, vals = ea.forms.read_form(view)
    errs = ea.forms.validate(form_key, vals)
    if not errs and not is_configured():
        errs[ea.forms.get(form_key).fields[-1].key] = '전자결재 봇이 아직 설정되지 않았어요. 관리자에게 알려 주세요.'
    values = view['state']['values']
    meta = json.loads(view.get('private_metadata') or '{}')
    vals['_files'] = ea.forms.read_attachments(view)   # '_' 칸은 장부에 안 들어간다(process_form_submit)
    vals['_folder_docs'] = ea.forms.read_folder_docs(view)
    if meta.get('edit'):
        vals['_edit'] = int(meta['edit'])
    return form_key, vals, {ea.field_block_id(values, k): m for k, m in errs.items()}


def process_form_submit(uid: str, form_key: str, vals: dict) -> dict:
    from dashboard.services.purchase_invoice_assign import COST_CATEGORIES
    name = _name(uid)
    files, edit_aid, docs = vals.get('_files') or [], vals.get('_edit'), vals.get('_folder_docs') or []
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
    if edit_aid:
        return process_edit(uid, edit_aid, fields, files, docs)
    row = ea.create(form_key, uid, name, fields)
    _audit(row, 'request', name)
    row = start(row)
    if form_key == ea.forms.FORM_INSURANCE:
        post_contract_link(row)
    _post_folder_docs_bg(row, docs)
    _post_files_bg(row, files, name)
    return row


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
                # 공사 폴더 서류는 새 현장 것으로(_form_values 가 채움)
                if meta['form'] == ea.forms.FORM_INSURANCE:   # 먼저 고른 보증 종류의 계약서 조건
                    fill_insurance_from_contract(vals, code)
                meta['pv'] = int(meta.get('pv') or 0) + 1
                meta['extra_note'] = _form_notes(meta['form'], code)
                client.views_update(view_id=view['id'], view=ea.forms.form_view(meta['form'], meta, vals))
            except Exception as exc:
                logger.error(f'{_LOG} 양식 현장 반영 실패: {exc}', exc_info=True)
        threading.Thread(target=_bg, daemon=True).start()

    @app.action('kind')
    def handle_insurance_kind_pick(ack, body, client):
        """보증보험 「보증 종류」를 고르면 이 현장 계약서에 적은 그 보증의 보증금율·기간을 채운다(11-05)."""
        ack()
        try:
            view = body['view']
            meta = json.loads(view.get('private_metadata') or '{}')
            if meta.get('form') != ea.forms.FORM_INSURANCE:
                return
            vals = ea.forms.read_form(view)[1]
            vals['_docs_block'] = ea.forms.carry_folder_docs(view)   # 고른 공사 폴더 서류 그대로
            code = vals.get('project') or meta.get('project') or ''
            before = (vals.get('rate'), vals.get('years'))
            fill_insurance_from_contract(vals, code, override=True)
            if (vals.get('rate'), vals.get('years')) == before:
                return                           # 계약서 조건이 없으면 창을 다시 그리지 않음(입력 유지)
            if code and not meta.get('project'):
                from dashboard.services.purchase_invoice_assign import project_display
                vals['_project_label'] = project_display(code, _project_rec(code))
            meta['pv'] = int(meta.get('pv') or 0) + 1   # 채운 칸 block_id 를 바꿔 새 값이 보이게
            client.views_update(view_id=view['id'], view=ea.forms.form_view(meta['form'], meta, vals))
        except Exception as exc:
            logger.error(f'{_LOG} 보증 종류 반영 실패: {exc}', exc_info=True)

    @app.action('guarantees')
    def handle_guarantees_pick(ack, body, client):
        """계약서 「보증보험」 고르면 고른 보증마다 보증금율·기간 칸을 펼친다(넣은 값은 block_id 가 같아 유지)."""
        ack()
        try:
            view = body['view']
            meta = json.loads(view.get('private_metadata') or '{}')
            vals = ea.forms.read_form(view)[1]
            vals['_docs_block'] = ea.forms.carry_folder_docs(view)   # 고른 공사 폴더 서류 그대로
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
    blocks = ea.spaced(
        f'🧾 *{prev.month}월 법인카드 지출결의서*를 올려 주세요.\n'
        '카드 번호와 사용 기간은 채워져 있어요. 영수증은 창 맨 아래 「첨부 파일」에 올려 주세요.',
        [{'type': 'button', 'action_id': ea.A_FORM_OPEN, 'value': ea.forms.FORM_CARDEXP, 'style': 'primary',
          'text': {'type': 'plain_text', 'text': '🧾 지출결의서 올리기', 'emoji': True}}])
    r = eap_client().chat_postMessage(channel=ea.channel_id(), text=f'{prev.month}월 법인카드 지출결의서', blocks=blocks)
    logger.info(f'{_LOG} 법인카드 지출결의 안내 게시 ({prev:%Y-%m})')
    return r.get('ts')


# ── 리마인드 (11-11 운영 점검) ───────────────────────────────────────
# 처음 DM 한 번뿐이라 비품 3만 원이 최종 승인 대기로 16시간, 수정 요청이 하루 넘게 머물렀다.
# 알림 피로를 줄이려 건마다 보내지 않고 사람마다 '내 차례 목록' 한 통으로 모아 보낸다(영업일만).
#   10:00·14:00  승인·최종 승인 차례(1시간 넘게 기다린 것) / 10:00 만 수정 요청 방치(20시간 넘은 것) → 요청자
#   15:30        이체·처리 대기(1시간 넘은 것) → 경영지원(N통장은 JW·YG)
_REMIND_WAIT_MIN = 60
_REVISE_WAIT_MIN = 15 * 60      # 전날 오후(19시 전)에 받은 수정 요청은 다음 날 10시에 — 20시간이면 17시 건이 이틀 밀렸다
_FINAL_WAIT_MIN = 20 * 60       # JW·YG 가 ① 대기를 받는 기준(KiKO 휴가 등으로 하루 가까이 머문 것)


def _wait_text(minutes: int) -> str:
    if minutes < 60:
        return f'{minutes}분째'
    if minutes < 24 * 60:
        return f'{minutes // 60}시간째'
    return f'{minutes // (24 * 60)}일째'


def remind_targets(kind: str, rows: List[dict], now=None) -> dict:
    """{uid: [(row, 기다린 분)]} — kind: 'approve'(오전·오후) / 'morning'(오전: approve + 수정 요청) / 'transfer'."""
    from datetime import datetime as _dt
    now = now or _dt.now()
    out: dict = {}

    def _add(uid, row, mins):
        if uid:
            out.setdefault(uid, []).append((row, mins))
    for row in rows:
        try:
            mins = int((now - _dt.strptime(str(row.get('updated_at'))[:19], '%Y-%m-%d %H:%M:%S')).total_seconds() // 60)
        except ValueError:
            continue
        st = row['status']
        if kind in ('approve', 'morning') and st in (ea.PENDING, ea.FINAL) and mins >= _REMIND_WAIT_MIN:
            for uid, role in _targets(row):
                # ① 대기는 KiKO 에게. JW·YG 는 하루 가까이 머문 것만(KiKO 휴가 등) — 매번 받으면 알림이 많다
                if role == 'final' and st == ea.PENDING and mins < _FINAL_WAIT_MIN:
                    continue
                _add(uid, row, mins)
        elif kind == 'morning' and st == ea.REVISE and mins >= _REVISE_WAIT_MIN:
            _add(row.get('requester_id'), row, mins)
        elif kind == 'transfer' and st == ea.APPROVED and mins >= _REMIND_WAIT_MIN:
            for uid, role in _targets(row):
                if role == 'executor':
                    _add(uid, row, mins)
    return out


def _remind_line(row: dict, mins: int) -> str:
    link = _card_link(row)
    t = ea.title(row)
    head = f'<{link}|{t}>' if link else t
    st = row['status']
    todo = {ea.PENDING: '승인', ea.FINAL: '최종 승인', ea.REVISE: '수정해서 다시 올리기',
            ea.APPROVED: f'{ea.step2_word(row)} 완료'}.get(st, '')
    who = '' if st == ea.REVISE else f" · {ea.initial(row.get('requester_name'))}"
    amt = int(row.get('amount') or 0)
    money = f' {amt:,}원' if amt else ''
    return f'• {head} — {ea.what_of(row)}{money}{who} · *{todo}* 기다린 지 {_wait_text(mins)}'


def post_reminders(kind: str, now=None) -> int:
    """리마인드 DM 보내기. 보낸 사람 수. 영업일·전자결재 설정·같은 회차 한 번만."""
    from datetime import datetime as _dt
    from dashboard.services.absent_remind import _is_business_day
    now = now or _dt.now()
    if os.getenv('EAPPROVAL_REMIND_ENABLED', 'true').strip().lower() == 'false':
        return 0
    if not is_configured() or not _is_business_day(now.date()):
        return 0
    try:   # 같은 회차 두 번 보내지 않게. Redis 가 안 되면 그냥 보낸다(스케줄러는 한 곳에서만 돈다)
        from dashboard.utils.redis_client import get_redis_client
        if not get_redis_client().redis.set(f'eap_remind:{kind}:{now:%Y-%m-%d-%H}', '1', nx=True, ex=86400):
            return 0
    except Exception as exc:
        logger.warning(f'{_LOG} 리마인드 중복 확인 실패(보냄): {exc}')
    targets = remind_targets(kind, ea.open_rows(), now)
    client = eap_client()
    for uid, items in targets.items():
        items.sort(key=lambda x: -x[1])
        mine_revise = [x for x in items if x[0]['status'] == ea.REVISE]
        mine_todo = [x for x in items if x[0]['status'] != ea.REVISE]
        parts = []
        if mine_todo:
            parts.append(f'⏰ *처리할 결재가 {len(mine_todo)}건 있어요.*\n'
                         + '\n'.join(_remind_line(r, m) for r, m in mine_todo))
        if mine_revise:
            parts.append(f'✏️ *수정 요청을 받은 결재가 {len(mine_revise)}건 있어요.* 카드의 [✏️ 수정하기]로 고쳐 올리거나, '
                         '필요 없으면 [요청 취소]를 눌러 주세요.\n'
                         + '\n'.join(_remind_line(r, m) for r, m in mine_revise))
        _notify(client, uid, '\n\n'.join(parts))
    if targets:
        logger.info(f'{_LOG} 리마인드({kind}) {len(targets)}명 · '
                    f'{sum(len(v) for v in targets.values())}건')
    return len(targets)


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
                    'element': {'type': 'plain_text_input', 'action_id': 'reason', 'multiline': True, 'max_length': 500,
                                'placeholder': {'type': 'plain_text', 'text': '요청자에게 그대로 전달됩니다.'}}}],
    }


def register_eapproval_handlers(app) -> None:
    """전자결재 봇 핸들러."""

    def _bg_action(action):
        def handler(ack, body, client):
            ack()
            aid, rev = ea.parse_button_value(body['actions'][0]['value'])
            uid = (body.get('user') or {}).get('id', '')

            def _bg():
                try:
                    if rev is not None and rev != ea.revision(aid):   # 그 사이 고쳐 다시 올린 결재(11-11)
                        _ephemeral(client, body, '🔒 이 결재는 그 사이 내용이 수정됐어요. '
                                                 '카드의 바뀐 내용을 확인하고 다시 눌러 주세요.')
                        refresh(ea.get(aid))
                        return
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
    app.action(ea.A_FINAL)(_bg_action('final'))
    app.action(ea.A_TRANSFER)(_bg_action('transfer'))
    app.action(ea.A_CANCEL)(_bg_action('cancel'))

    @app.action(ea.A_REVISE)
    def handle_revise_open(ack, body, client):
        """[✏️ 수정 요청] — 고칠 내용을 적는 창(플렉스에서 댓글로 돌려보내던 것, 11-05)."""
        ack()
        aid = int(body['actions'][0]['value'])
        row = ea.get(aid)
        uid = (body.get('user') or {}).get('id', '')
        if not row or not ea.can_act('revise', row, uid):
            _ephemeral(client, body, '🔒 수정 요청은 승인자·경영지원만 할 수 있어요.')
            return
        if row['status'] not in (ea.PENDING, ea.FINAL, ea.APPROVED):
            _ephemeral(client, body, f"🔒 지금은 수정 요청을 할 수 없어요 (지금: {ea.status_label(row)[1]}).")
            return
        try:
            client.views_open(trigger_id=body['trigger_id'], view=_revise_view(aid))
        except Exception as exc:
            logger.error(f'{_LOG} 수정 요청 창 열기 실패: {exc}', exc_info=True)

    @app.view(ea.CB_REVISE)
    def handle_revise_submit(ack, body, view):
        reason = ((view['state']['values'].get('reason') or {}).get('reason') or {}).get('value', '').strip()
        if not reason:
            ack(response_action='errors', errors={'reason': '무엇을 고쳐야 하는지 적어 주세요.'})
            return
        ack()
        aid = json.loads(view.get('private_metadata') or '{}').get('aid')
        uid = (body.get('user') or {}).get('id', '')

        def _bg():
            try:
                msg = apply_action('revise', int(aid), uid, reason)
                if msg:
                    _notify(eap_client(), uid, f'🔒 {msg}')
            except Exception as exc:
                logger.error(f'{_LOG} 수정 요청 실패 (#{aid}): {exc}', exc_info=True, extra={'force_alert': True})
        threading.Thread(target=_bg, daemon=True).start()

    @app.action(ea.A_EDIT)
    def handle_edit_open(ack, body, client):
        """[✏️ 수정하기] — 요청자만. 올린 값이 채워진 같은 양식 창(trigger 3초라 바로 연다)."""
        ack()
        aid = int(body['actions'][0]['value'])
        uid = (body.get('user') or {}).get('id', '')
        try:
            msg = open_edit_modal(client, body.get('trigger_id', ''), aid, uid)
            if msg:
                _ephemeral(client, body, f'🔒 {msg}')
        except Exception as exc:
            logger.error(f'{_LOG} 수정 창 열기 실패 (#{aid}): {exc}', exc_info=True)
            _ephemeral(client, body, ':warning: 수정 창을 열지 못했어요. 잠시 뒤 다시 눌러 주세요.')

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
    from dashboard.blueprints.slack_cert import register_cert_handlers
    register_cert_handlers(app)   # 증명서 신청(11-06) — DM 만

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
