"""
APScheduler 백그라운드 폴링 스케줄러

- 당근 자동 연동 시트: KARROT_SYNC_INTERVAL_MIN 분 주기로 sync_karrot()
- 향후: Gmail 폴링, 기타 채널 추가
"""

import os
import time
from typing import Optional

from apscheduler.schedulers.background import BackgroundScheduler

from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)

_scheduler: Optional[BackgroundScheduler] = None

# 관리자 알림 dedup (스팸 방지) — {key: last_sent_ts}
_admin_alert_last_sent: dict = {}
_ADMIN_ALERT_COOLDOWN_SEC = 30 * 60  # 같은 알림 30분 1회만


def _notify_admin(key: str, text: str) -> None:
    """슬랙 DM + 이메일로 운영 알림 발송 (Redis 다운, sync 실패 등).
    같은 key는 30분 cooldown — 스팸 방지.
    외부에 있을 때 대비 — 슬랙 + 이메일 모두 발송.
    """
    now = time.time()
    last = _admin_alert_last_sent.get(key, 0)
    if now - last < _ADMIN_ALERT_COOLDOWN_SEC:
        return
    _admin_alert_last_sent[key] = now

    # 1) 슬랙 DM
    target = os.getenv('SLACK_ADMIN_CHANNEL', '').strip()
    token = os.getenv('SLACK_BOT_TOKEN', '').strip()
    if target and token:
        try:
            import json
            import urllib.request
            req = urllib.request.Request(
                'https://slack.com/api/chat.postMessage',
                data=json.dumps({'channel': target, 'text': text}).encode('utf-8'),
                headers={
                    'Authorization': f'Bearer {token}',
                    'Content-Type': 'application/json; charset=utf-8',
                },
            )
            with urllib.request.urlopen(req, timeout=5) as r:
                pass
        except Exception as exc:
            logger.warning(f'[NOTIFY] 슬랙 DM 발송 실패: {exc}')
    else:
        logger.warning(f'[NOTIFY] SLACK_ADMIN_CHANNEL 미설정 — slack skip')

    # 2) 이메일 (외부에서도 받을 수 있게)
    admin_email = os.getenv('ADMIN_NOTIFY_EMAIL', '').strip()
    if admin_email:
        try:
            _send_admin_email(admin_email, f'[ITG 운영 알림] {key}', text)
        except Exception as exc:
            logger.warning(f'[NOTIFY] 이메일 발송 실패: {exc}')


def _send_admin_email(to_email: str, subject: str, body: str) -> None:
    """Gmail API로 운영 알림 메일 발송 (서비스 계정 도메인 위임).
    HOMEPAGE_MAIL_USER로 위임된 계정에서 to_email로 발송.
    필요 scope: https://www.googleapis.com/auth/gmail.send
    """
    import base64
    from email.mime.text import MIMEText
    from google.oauth2 import service_account
    from googleapiclient.discovery import build

    creds_file = (
        os.getenv('GOOGLE_CREDENTIALS_FILE')
        or os.getenv('GOOGLE_APPLICATION_CREDENTIALS')
        or 'credentials.json'
    )
    delegated_user = os.getenv('HOMEPAGE_MAIL_USER', '').strip() or to_email
    credentials = service_account.Credentials.from_service_account_file(
        creds_file,
        scopes=['https://www.googleapis.com/auth/gmail.send'],
    )
    delegated_creds = credentials.with_subject(delegated_user)
    svc = build('gmail', 'v1', credentials=delegated_creds, cache_discovery=False)

    msg = MIMEText(body, 'plain', 'utf-8')
    msg['to'] = to_email
    msg['from'] = delegated_user
    msg['subject'] = subject
    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode('utf-8')
    svc.users().messages().send(userId='me', body={'raw': raw}).execute()


def _redis_healthy() -> bool:
    """Redis ping 체크 — 죽었으면 모든 sync 작업 skip (폭주 방지 circuit breaker).
    Redis가 죽으면 dedup 키 조회 실패 → 모든 리드를 신규로 인식 → 슬랙 메시지 폭주.
    Redis 다운 시 관리자 슬랙 DM 알림 (30분 1회).
    """
    try:
        from dashboard.utils.redis_client import get_redis_client
        ok = get_redis_client().ping()
        if not ok:
            _notify_admin(
                'redis_down',
                ':rotating_light: *Redis 다운* — 모든 sync 작업 정지됨. '
                'Docker container 상태 확인 필요 (`docker ps | grep redis`).'
            )
        return ok
    except Exception as exc:
        logger.error(f'[SCHED] Redis 헬스체크 실패: {exc}')
        _notify_admin(
            'redis_down',
            f':rotating_light: *Redis 헬스체크 실패* — `{exc}`. '
            f'모든 sync 작업 정지됨. Docker 확인 바랍니다.'
        )
        return False


def start_scheduler():
    """앱 시작 시 1회 호출. Flask debug reload 시 이중 시작 방지."""
    global _scheduler
    if _scheduler is not None:
        logger.debug('[SCHED] 이미 시작됨 - 무시')
        return

    if os.getenv('SLACK_BOT_ENABLED', 'false').lower() != 'true':
        logger.info('[SCHED] SLACK_BOT_ENABLED=false - 스케줄러 비활성화')
        return

    if not os.getenv('KARROT_AUTO_SHEET_ID', '').strip():
        logger.info('[SCHED] KARROT_AUTO_SHEET_ID 미설정 - 스케줄러 비활성화')
        return

    _scheduler = BackgroundScheduler(
        timezone='Asia/Seoul',
        job_defaults={'coalesce': True, 'max_instances': 1, 'misfire_grace_time': 60},
    )

    # 당근 시트 동기화
    karrot_interval = int(os.getenv('KARROT_SYNC_INTERVAL_MIN', '2'))
    _scheduler.add_job(
        _safe_karrot_sync,
        'interval',
        minutes=karrot_interval,
        id='karrot_sync',
    )
    jobs = [f'당근 {karrot_interval}분']

    # 홈페이지 메일 동기화 (HOMEPAGE_MAIL_USER 설정된 경우만)
    if os.getenv('HOMEPAGE_MAIL_USER', '').strip():
        homepage_interval = int(os.getenv('HOMEPAGE_MAIL_SYNC_INTERVAL_MIN', '2'))
        _scheduler.add_job(
            _safe_homepage_sync,
            'interval',
            minutes=homepage_interval,
            id='homepage_mail_sync',
        )
        jobs.append(f'홈페이지 메일 {homepage_interval}분')

    # 채널톡 미배정 알림 체크 (1분 주기 — 5분 경과한 미응답 채팅 찾기)
    if os.getenv('CHANNELTALK_ACCESS_KEY', '').strip():
        _scheduler.add_job(
            _safe_channeltalk_pending_check,
            'interval',
            minutes=1,
            id='channeltalk_pending_check',
        )
        jobs.append('채널톡 미배정 알림 1분')

    # 슬랙 워크플로우 → 메인 시트 직접 추가 전화 lead 보정 (10초 주기 — 사용자 체감 즉시)
    _scheduler.add_job(
        _safe_workflow_phone_sync,
        'interval',
        seconds=10,
        id='workflow_phone_sync',
    )
    jobs.append('워크플로 전화 lead 10초')

    # 수금 관리 알림 — 공사 현황 시트 U/V/W 입금 메모 변경 감지 (30초 주기)
    if os.getenv('SLACK_PAYMENT_BOT_TOKEN', '').strip():
        _scheduler.add_job(
            _safe_payment_sync,
            'interval',
            seconds=30,
            id='payment_sync',
        )
        jobs.append('수금 관리 30초')
        # 일일 요약/미수금 자동 발송은 disable — /수금 슬래시 명령으로만 조회

    # 미발송 슬랙 알림 재발송 (SSL 에러 등으로 누락된 lead 자동 복구) — 5분 주기
    _scheduler.add_job(
        _safe_retry_pending_slack,
        'interval',
        minutes=5,
        id='retry_pending_slack',
    )
    jobs.append('미발송 슬랙 재발송 5분')

    # SSL 인증서 만료 체크 — 매일 09시 (Caddy 자동 갱신 실패 안전망)
    if os.getenv('SLACK_PUBLIC_HOST', '').strip():
        _scheduler.add_job(
            _safe_cert_expiry_check,
            'cron',
            hour=9, minute=0,
            id='cert_expiry_check',
        )
        # Flask 시작 시 1회 즉시 체크 — 재시작 직후 만료 임박 즉시 감지
        import threading as _th
        _th.Timer(30, _safe_cert_expiry_check).start()
        jobs.append('SSL 인증서 체크 매일 09시 + 시작 30초 후 1회')

    _scheduler.start()
    logger.info(f'[SCHED] 백그라운드 스케줄러 시작 ({" / ".join(jobs)} 주기)')


def stop_scheduler():
    global _scheduler
    if _scheduler is not None:
        try:
            _scheduler.shutdown(wait=False)
        except Exception:
            pass
        _scheduler = None
        logger.info('[SCHED] 스케줄러 정지')


def _is_transient_error(exc: Exception) -> bool:
    """일시 에러 (SSL/timeout/connection) 판정 — 다음 폴링에 자동 재시도되므로 알림 불필요."""
    s = str(exc).lower()
    return (
        'ssl' in s or 'wrong_version' in s or 'decryption_failed' in s
        or 'timeout' in s or 'connection' in s or 'bad_record_mac' in s
        or 'eof' in s or 'unreachable' in s or 'reset' in s
    )


def _safe_karrot_sync():
    """예외가 스케줄러를 중단시키지 않도록 wrapper"""
    if not _redis_healthy():
        logger.warning('[SCHED] Redis 다운 — 당근 sync skip (폭주 방지)')
        return
    try:
        from dashboard.services.lead_sync import sync_karrot
        sync_karrot()
    except Exception as exc:
        logger.error(f'[SCHED] 당근 sync 실행 실패: {exc}', exc_info=True)
        if not _is_transient_error(exc):
            _notify_admin('karrot_sync_fail',
                          f':warning: 당근 sync 실패 — `{exc}`. 로그 확인 필요.')


def _safe_workflow_phone_sync():
    """슬랙 워크플로우가 메인 시트에 직접 추가한 전화 lead 보정.

    PHONE_WORKFLOW_SYNC_ENABLED=true 일 때만 활성화.
    기본 disable — 슬랙 도입 전 매니저가 시트 수동 입력할 때 자동 발번/카드 방지.
    Apps Script onEdit이 대신 lead_no 발번.
    """
    if os.getenv('PHONE_WORKFLOW_SYNC_ENABLED', 'false').lower() != 'true':
        return
    if not _redis_healthy():
        logger.warning('[SCHED] Redis 다운 — 워크플로 전화 lead 보정 skip')
        return
    try:
        from dashboard.services.lead_sync import sync_workflow_phone_leads
        sync_workflow_phone_leads()
    except Exception as exc:
        logger.error(f'[SCHED] 워크플로 전화 lead 보정 실패: {exc}', exc_info=True)
        if not _is_transient_error(exc):
            _notify_admin('workflow_phone_fail',
                          f':warning: 전화 lead 보정 실패 — `{exc}`.')


def _safe_retry_pending_slack():
    """미발송 슬랙 알림 자동 재발송 (SSL 에러 등으로 누락된 lead 복구)"""
    if not _redis_healthy():
        logger.warning('[SCHED] Redis 다운 — pending slack 재발송 skip')
        return
    try:
        from dashboard.services.lead_sync import retry_pending_slack_notifications
        retry_pending_slack_notifications()
    except Exception as exc:
        logger.error(f'[SCHED] 미발송 슬랙 재발송 실패: {exc}', exc_info=True)
        if _is_transient_error(exc):
            return
        _notify_admin('retry_pending_fail',
                      f':warning: 미발송 슬랙 재발송 실패 — `{exc}`.')


def _safe_payment_sync():
    """공사 현황 시트 U/V/W 입금 메모 변경 감지 → #수금_관리 채널 발송"""
    if not _redis_healthy():
        logger.warning('[SCHED] Redis 다운 — 수금 sync skip')
        return
    try:
        from dashboard.services.payment_sync import sync_payments
        sync_payments()
    except Exception as exc:
        logger.error(f'[SCHED] 수금 알림 sync 실패: {exc}', exc_info=True)
        if not _is_transient_error(exc):
            _notify_admin('payment_sync_fail',
                          f':warning: 수금 알림 sync 실패 — `{exc}`.')


def _safe_payment_daily_report():
    """매일 18시 — 일일 요약 + 미수금 30일 경과 리포트 #수금_관리 채널 발송"""
    try:
        import os
        from slack_sdk import WebClient
        from dashboard.services.payment_sync import (
            daily_payment_summary, build_overdue_message,
        )
        token = os.getenv('SLACK_PAYMENT_BOT_TOKEN', '').strip()
        channel = os.getenv('SLACK_PAYMENT_CHANNEL', '').strip()
        if not token or not channel:
            return
        slack = WebClient(token=token)
        summary = daily_payment_summary()
        if summary:
            slack.chat_postMessage(channel=channel, text=summary)
        overdue = build_overdue_message(days=30)
        if overdue:
            slack.chat_postMessage(channel=channel, text=overdue)
    except Exception as exc:
        logger.error(f'[SCHED] 수금 일일 리포트 실패: {exc}', exc_info=True)


def _safe_homepage_sync():
    """예외가 스케줄러를 중단시키지 않도록 wrapper"""
    if not _redis_healthy():
        logger.warning('[SCHED] Redis 다운 — 홈페이지 메일 sync skip')
        return
    try:
        from dashboard.services.homepage_mail_sync import sync_homepage_email
        sync_homepage_email()
    except Exception as exc:
        logger.error(f'[SCHED] 홈페이지 메일 sync 실행 실패: {exc}', exc_info=True)
        if not _is_transient_error(exc):
            _notify_admin('homepage_sync_fail',
                          f':warning: 홈페이지 메일 sync 실패 — `{exc}`.')


def _safe_channeltalk_pending_check():
    """5분 경과한 미배정 채팅에 @here 알림 발송 (1분 주기)"""
    try:
        import time
        import json
        import urllib.request
        from dashboard.services import channeltalk_threads as _t

        channel = os.getenv('SLACK_CHANNELTALK_CHANNEL', '').strip()
        token = os.getenv('SLACK_BOT_TOKEN', '').strip()
        if not channel or not token:
            return

        threshold_sec = int(os.getenv('CHANNELTALK_REMIND_AFTER_MIN', '5')) * 60
        now = int(time.time())

        for entry in _t.list_pending():
            if entry.get('reminded'):
                continue
            elapsed = now - int(entry.get('created_at', 0))
            if elapsed < threshold_sec:
                continue

            # 알림 발송
            thread_ts = entry.get('thread_ts')
            chat_id = entry.get('chat_id')
            if not thread_ts:
                continue

            # 메시지 존재 확인 — 카드가 삭제됐으면 알림 skip + pending 제거
            try:
                check_req = urllib.request.Request(
                    f'https://slack.com/api/conversations.replies'
                    f'?channel={channel}&ts={thread_ts}&limit=1',
                    headers={'Authorization': f'Bearer {token}'},
                )
                with urllib.request.urlopen(check_req, timeout=5) as r:
                    check_resp = json.loads(r.read())
                if not check_resp.get('ok') or not check_resp.get('messages'):
                    logger.info(
                        f'[SCHED] 채널톡 카드 삭제됨 — 미응답 알림 skip (chat_id={chat_id})'
                    )
                    _t.remove_pending(chat_id)
                    continue
            except Exception as exc:
                logger.warning(f'[SCHED] 메시지 존재 확인 실패: {exc}')
                # 확인 실패 시에도 알림은 발송 (보수적)

            text = (
                f':alarm_clock: <!here> *고객이 {elapsed // 60}분째 답변을 기다리고 있어요*\n'
                f'_이 thread에서 답변 부탁드립니다._'
            )
            try:
                req = urllib.request.Request(
                    'https://slack.com/api/chat.postMessage',
                    data=json.dumps({
                        'channel': channel,
                        'thread_ts': thread_ts,
                        'text': text,
                        'reply_broadcast': False,  # thread에만 — 메인 채널 중복 노출 방지
                    }).encode('utf-8'),
                    headers={
                        'Authorization': f'Bearer {token}',
                        'Content-Type': 'application/json; charset=utf-8',
                    },
                )
                with urllib.request.urlopen(req, timeout=5) as r:
                    resp = json.loads(r.read())
                if resp.get('ok'):
                    _t.mark_reminded(chat_id)
                    logger.info(f'[SCHED] 채널톡 미배정 알림 발송 (chat_id={chat_id}, {elapsed // 60}분 경과)')
                else:
                    logger.warning(f'[SCHED] 미배정 알림 발송 실패: {resp.get("error")}')
            except Exception as exc:
                logger.warning(f'[SCHED] 미배정 알림 예외: {exc}')
    except Exception as exc:
        logger.error(f'[SCHED] 채널톡 미배정 체크 실패: {exc}', exc_info=True)


def _safe_cert_expiry_check():
    """SLACK_PUBLIC_HOST 도메인의 SSL 인증서 만료일 체크.
    30일 이내 만료면 관리자 알림 (Caddy 자동 갱신 실패 대비).
    """
    host = os.getenv('SLACK_PUBLIC_HOST', '').strip()
    if not host:
        return
    try:
        import socket
        import ssl
        from datetime import datetime as _dt
        ctx = ssl.create_default_context()
        with socket.create_connection((host, 443), timeout=10) as sock:
            with ctx.wrap_socket(sock, server_hostname=host) as ssock:
                cert = ssock.getpeercert()
        not_after = cert.get('notAfter')
        if not not_after:
            return
        # 'Jun 30 09:00:00 2026 GMT' 양식
        expiry = _dt.strptime(not_after, '%b %d %H:%M:%S %Y %Z')
        days_left = (expiry - _dt.utcnow()).days
        if days_left <= 30:
            _notify_admin(
                f'cert_expiry_{host}',
                f':warning: *SSL 인증서 만료 임박* — `{host}` 인증서가 *{days_left}일* 후 만료. '
                f'Caddy 자동 갱신 상태 확인 필요 (`docker logs caddy --tail 50`).'
            )
        else:
            logger.debug(f'[SCHED] SSL 인증서 만료까지 {days_left}일 ({host})')
    except Exception as exc:
        logger.warning(f'[SCHED] 인증서 만료 체크 실패: {exc}')


def trigger_karrot_sync_now() -> dict:
    """수동 트리거 (테스트/디버깅용)"""
    from dashboard.services.lead_sync import sync_karrot
    return sync_karrot()


def trigger_homepage_sync_now() -> dict:
    """수동 트리거"""
    from dashboard.services.homepage_mail_sync import sync_homepage_email
    return sync_homepage_email()
