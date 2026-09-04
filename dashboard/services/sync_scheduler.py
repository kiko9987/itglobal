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

        # 매니저 실수 감지 일일 요약 — 매일 오전 9시 (평일만, 매니저 출근 시각)
        # 2026-07-17 사용자 요청으로 기본 disable. 활성 원하면 env
        # PAYMENT_ALERT_DAILY_ENABLED=true 로 재활성.
        if os.getenv('PAYMENT_ALERT_DAILY_ENABLED', '').strip().lower() in ('1', 'true', 'yes', 'on'):
            _scheduler.add_job(
                _safe_payment_alert_daily,
                'cron',
                day_of_week='mon-fri',
                hour=9, minute=0,
                id='payment_alert_daily',
            )
            jobs.append('수금 매니저 실수 요약 평일 09시')
        else:
            logger.info('[SCHED] payment_alert_daily 비활성 (PAYMENT_ALERT_DAILY_ENABLED=false)')

        # 수금봇 토큰 health check — 매 1시간 (P1-4, 2026-07-10)
        # 토큰 만료·비활성화 시 관리자 DM 알림
        _scheduler.add_job(
            _safe_payment_bot_health,
            'interval', hours=1,
            id='payment_bot_health',
        )
        jobs.append('수금봇 health check 1시간')

    # 미발송 슬랙 알림 재발송 (SSL 에러 등으로 누락된 lead 자동 복구) — 5분 주기
    _scheduler.add_job(
        _safe_retry_pending_slack,
        'interval',
        minutes=5,
        id='retry_pending_slack',
    )
    jobs.append('미발송 슬랙 재발송 5분')

    # 미발송 방문 카드 재발송 (전화WF sync SSL 안전망) — 5분 주기
    _scheduler.add_job(
        _safe_retry_pending_visit_notice,
        'interval',
        minutes=5,
        id='retry_pending_visit_notice',
    )
    jobs.append('미발송 방문 카드 재발송 5분')

    # 방문 사진 배치 복구 (hang·재시작으로 중단된 업로드의 남은 사진 재저장) — 5분 주기
    _scheduler.add_job(
        _safe_recover_photo_batches,
        'interval',
        minutes=5,
        id='recover_photo_batches',
    )
    jobs.append('방문 사진 배치 복구 5분')

    # 방문 캔버스1 정기 재생성 (이벤트 훅 누락·write-behind 레이스 안전망) — 10분 주기.
    # 캔버스1 방문 누락은 업무상 치명적인데 캔버스1은 등록·완료·확정 이벤트에서만
    # 재생성 → 훅 누락/레이스로 빠지면 다음 이벤트까지 영구 stale (L-03565 계기).
    # 주기 재생성으로 최대 10분 내 자가 치유. rebuild_canvas 는 env 없으면 no-op.
    _scheduler.add_job(
        _safe_rebuild_visit_canvas,
        'interval',
        minutes=10,
        id='rebuild_visit_canvas',
    )
    jobs.append('방문 캔버스1 정기 재생성 10분')

    # 고아 리드 감지 (시트에 있는데 슬랙 카드 흔적 없음 = Flask 재시작 등으로 발송 유실)
    # — 5분 주기, pending 큐에도 없는 케이스가 대상
    _scheduler.add_job(
        _safe_recover_orphan_leads,
        'interval',
        minutes=5,
        id='recover_orphan_leads',
    )
    jobs.append('고아 리드 재발송 5분')

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

    # 2026-07-09 매일 새벽 03:15 자동 백업 (users.db + Redis)
    _scheduler.add_job(
        _safe_daily_backup,
        'cron',
        hour=3, minute=15,
        id='daily_backup',
        replace_existing=True,
    )
    jobs.append('일 백업 매일 03:15')

    # 2026-07-23 매일 아침 9시 부재중/미완료 리마인드 (온라인 문의 채널)
    _scheduler.add_job(
        _safe_absent_remind_daily,
        'cron',
        hour=9, minute=0,
        id='absent_remind_daily',
        replace_existing=True,
    )
    jobs.append('부재중 리마인드 매일 09:00')

    # 2026-07-28 미처리 정산 핀 리마인드 — 매일 오후 1시 #영업_관리 (세금계산서 관리 알림 봇).
    # 경영지원이 고정한 입금내역·세금계산서(미처리) 요약. pins:read 필요.
    if os.getenv('SLACK_INVOICE_BOT_TOKEN', '').strip() and os.getenv('SLACK_INVOICE_CHANNEL_ID', '').strip():
        _scheduler.add_job(
            _safe_pin_remind_daily,
            'cron',
            hour=13, minute=0,
            id='pin_remind_daily',
            replace_existing=True,
        )
        jobs.append('정산 핀 리마인드 매일 13:00')
        # 2026-08-18 저녁 재확인 — 하루 처리 후 남은 미처리 다시 리마인드 (동일 내용).
        _scheduler.add_job(
            _safe_pin_remind_daily,
            'cron',
            hour=17, minute=0,
            id='pin_remind_evening',
            replace_existing=True,
        )
        jobs.append('정산 핀 리마인드 매일 17:00')

    # 2026-07-28 거래처 탭 국세청 상태 갱신. NTS_SERVICE_KEY 있을 때만.
    if os.getenv('NTS_SERVICE_KEY', '').strip():
        # 주간 풀갱신 (매주 월 04:00) — 기존 거래처 폐업/휴업 상태 변경 감지.
        _scheduler.add_job(
            _safe_partner_status_refresh,
            'cron',
            day_of_week='mon', hour=4, minute=0,
            id='partner_status_refresh',
            replace_existing=True,
        )
        jobs.append('거래처 상태 주간갱신 월 04:00')
        # 데일리 증분 (매일 07:00) — 신규 추가 거래처(J 빈 행)만 채움.
        _scheduler.add_job(
            _safe_partner_status_fill_new,
            'cron',
            hour=7, minute=0,
            id='partner_status_fill_new',
            replace_existing=True,
        )
        jobs.append('거래처 신규 증분 매일 07:00')

    # 2026-09-02 사업자등록증 상태 캐시 워밍 — PM '첫 조회' 지연 제거 (Redis 영구캐시 + 워머).
    #   미캐시 비취소 프로젝트만 계산·저장(화면 밖 백그라운드). 이미 캐시된 건은 즉시 skip.
    _scheduler.add_job(
        _safe_warm_license_states,
        'interval',
        hours=6,
        id='warm_license_states',
        replace_existing=True,
    )
    # Flask 시작 60초 후 1회 즉시 워밍 (재시작 직후에도 캐시 채워둠 → 첫 조회 즉시)
    import threading as _th_warm
    _th_warm.Timer(60, _safe_warm_license_states).start()
    jobs.append('등록증 상태 캐시 워밍 6시간 + 시작 60초 후 1회')

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


def _safe_rebuild_visit_canvas():
    """방문 일정 캔버스1 정기 재생성 안전망 (2026-08-06).

    캔버스1은 방문 등록·완료·확정 이벤트 훅에서만 재생성돼, 훅 누락이나
    write-behind(시트 반영 지연) 레이스로 방문이 빠지면 다음 이벤트까지 stale.
    방문 누락은 업무상 치명적이라 10분 주기로 강제 재생성해 자가 치유.
    rebuild_canvas 는 SLACK_VISIT_CANVAS_ID/토큰 없으면 자체적으로 no-op.
    """
    try:
        from dashboard.services.visit_canvas_sync import rebuild_canvas
        res = rebuild_canvas()
        if not res.get('ok') and res.get('reason'):
            logger.warning(f"[SCHED] 캔버스1 정기 재생성 skip/실패: {res.get('reason')}")
    except Exception as exc:
        logger.error(f'[SCHED] 캔버스1 정기 재생성 예외: {exc}', exc_info=True)
        if not _is_transient_error(exc):
            _notify_admin('rebuild_canvas_fail',
                          f':warning: 방문 캔버스1 정기 재생성 실패 — `{exc}`.')


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


def _safe_retry_pending_visit_notice():
    """미발송 방문 카드 자동 재발송 (전화WF sync SSL 에러 등 복구)"""
    if not _redis_healthy():
        logger.warning('[SCHED] Redis 다운 — pending visit notice 재발송 skip')
        return
    try:
        from dashboard.services.lead_sync import retry_pending_visit_notices
        retry_pending_visit_notices()
    except Exception as exc:
        logger.error(f'[SCHED] 미발송 방문 카드 재발송 실패: {exc}', exc_info=True)
        if _is_transient_error(exc):
            return
        _notify_admin('retry_visit_fail',
                      f':warning: 미발송 방문 카드 재발송 실패 — `{exc}`.')


def _safe_recover_photo_batches():
    """중단된 방문 사진 배치(photo_batch:*)의 남은 파일 재저장 — hang·재시작 안전망"""
    if not _redis_healthy():
        return
    try:
        from dashboard.blueprints.slack_bot import _recover_photo_batches
        _recover_photo_batches()
    except Exception as exc:
        logger.error(f'[SCHED] 방문 사진 배치 복구 실패: {exc}', exc_info=True)


def _safe_recover_orphan_leads():
    """시트에 등록됐지만 슬랙 카드 흔적 없는 고아 리드 자동 재발송.

    Flask 재시작 등으로 시트 append 후 pending 큐 등록 전에 프로세스가 죽어
    retry_pending_slack도 못 잡는 케이스 대응.
    """
    if not _redis_healthy():
        logger.warning('[SCHED] Redis 다운 — 고아 리드 재발송 skip')
        return
    try:
        from dashboard.services.lead_sync import recover_orphan_lead_notifications
        recover_orphan_lead_notifications()
    except Exception as exc:
        logger.error(f'[SCHED] 고아 리드 재발송 실패: {exc}', exc_info=True)
        if _is_transient_error(exc):
            return
        _notify_admin('recover_orphan_fail',
                      f':warning: 고아 리드 재발송 실패 — `{exc}`.')


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


def _safe_payment_bot_health():
    """수금봇 토큰 health check — 매 1시간 (P1-4, 2026-07-10).

    auth.test 실패 시 관리자 DM 알림. 봇 토큰 만료·비활성화 즉시 감지.
    """
    token = os.getenv('SLACK_PAYMENT_BOT_TOKEN', '').strip()
    if not token:
        return
    try:
        from slack_sdk import WebClient
        from dashboard.blueprints.slack_bot import _verify_bot_token
        client = WebClient(token=token)
        ok = _verify_bot_token(client, '수금봇')
        if not ok:
            _notify_admin('payment_bot_health_fail',
                          ':warning: 수금봇 토큰 검증 실패 — 토큰 회전 필요 (`SLACK_PAYMENT_BOT_TOKEN`)')
    except Exception as exc:
        logger.warning(f'[SCHED] 수금봇 health check 예외: {exc}')


def _safe_payment_alert_daily():
    """매일 오전 9시 (평일) — 매니저 실수 전체 스캔 + 일일 요약 채널 발송 (2026-07-10)."""
    if not _redis_healthy():
        logger.warning('[SCHED] Redis 다운 — 수금 매니저 요약 skip')
        return
    try:
        from dashboard.services.payment_alert_daily import run_daily_scan_and_summary
        run_daily_scan_and_summary()
    except Exception as exc:
        logger.error(f'[SCHED] 수금 매니저 요약 실패: {exc}', exc_info=True)
        if not _is_transient_error(exc):
            _notify_admin('payment_alert_daily_fail',
                          f':warning: 수금 매니저 실수 요약 실패 — `{exc}`.')


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


def _safe_absent_remind_daily():
    """어제 미처리 문의 리마인드 카드 발송 (매일 아침 9시)."""
    try:
        from dashboard.services.absent_remind import send_daily_remind
        result = send_daily_remind()
        logger.info(f'[SCHED] 부재중 리마인드 실행 결과: {result}')
    except Exception as exc:
        logger.error(f'[SCHED] 부재중 리마인드 실패: {exc}', exc_info=True)


def _safe_pin_remind_daily():
    """미처리 정산 핀 리마인드 발송 (매일 아침 9시 #영업_관리)."""
    try:
        from dashboard.services.pin_remind import send_pin_remind
        result = send_pin_remind()
        logger.info(f'[SCHED] 정산 핀 리마인드 실행 결과: {result}')
    except Exception as exc:
        logger.error(f'[SCHED] 정산 핀 리마인드 실패: {exc}', exc_info=True)


def _safe_partner_status_refresh():
    """거래처 탭 국세청 사업자등록 상태 주기 갱신 (주 1회). 폐업/휴업 최신화.

    조회 0건(키 미설정/API 실패)이면 refresh_partner_status 내부 가드로
    시트 쓰기 skip → 기존 값 보존.
    """
    try:
        from dashboard.services.partner_status_sync import refresh_partner_status
        result = refresh_partner_status(dry_run=False)
        s = result.get('summary', {})
        if result.get('skipped_write'):
            logger.warning('[SCHED] 거래처 상태 갱신 — 조회 0건으로 쓰기 skip (기존 보존)')
        else:
            logger.info(
                f"[SCHED] 거래처 상태 갱신 완료 — 계속 {s.get('계속사업자')}, "
                f"휴업 {s.get('휴업자')}, 폐업 {s.get('폐업자')}, 조회안됨 {s.get('조회안됨')}"
            )
    except Exception as exc:
        logger.error(f'[SCHED] 거래처 상태 갱신 실패: {exc}', exc_info=True)


def _safe_partner_status_fill_new():
    """신규 추가 거래처 J/K 증분 채움 (매일). J가 빈 행만 조회·기록.

    주간 풀갱신(_safe_partner_status_refresh)은 상태 변경 감지용으로 별도 유지.
    """
    try:
        from dashboard.services.partner_status_sync import (
            refresh_partner_status, rebuild_partner_caches,
        )
        result = refresh_partner_status(dry_run=False, only_blank=True)
        if not result.get('no_blank'):
            s = result.get('summary', {})
            logger.info(
                f"[SCHED] 거래처 신규 증분 채움 — {result.get('total_rows')}행 "
                f"(계속 {s.get('계속사업자', 0)}, 휴업 {s.get('휴업자', 0)}, "
                f"폐업 {s.get('폐업자', 0)}, 조회안됨 {s.get('조회안됨', 0)})"
            )
        # 거래처 캐시(상호→이메일, 번호→상호) 매일 재구성
        rebuild_partner_caches()
    except Exception as exc:
        logger.error(f'[SCHED] 거래처 신규 증분/캐시 실패: {exc}', exc_info=True)


def _safe_warm_license_states():
    """사업자등록증 상태 캐시 워밍 (백그라운드, 2026-09-02).

    PM 아코디언 '사업자등록증' 상태 조회는 Drive 폴더 조회(~1초)가 필요해 첫 조회가 느리다.
    Redis 영구캐시(30일) + 이 워머로 미리 채워 사용자 첫 조회도 즉시 응답되게 한다.
    캐시된 건은 즉시 skip(Drive 호출 X), 비취소·미캐시 건만 계산. 화면 밖 백그라운드 처리.
    """
    try:
        from dashboard.services.business_license_handler import warm_license_states
        warm_license_states()
    except Exception as exc:
        logger.error(f'[SCHED] 등록증 상태 캐시 워밍 실패: {exc}', exc_info=True)


def _safe_daily_backup():
    """매일 자동 백업 — users.db + Redis dump.rdb."""
    try:
        import sys as _sys, subprocess as _sp
        from pathlib import Path as _P
        script = _P(__file__).resolve().parent.parent.parent / 'scripts' / 'backup_daily.py'
        if not script.exists():
            logger.warning('[BACKUP] 스크립트 파일 없음')
            return
        result = _sp.run(
            [_sys.executable, '-X', 'utf8', str(script)],
            capture_output=True, text=True, timeout=180,
        )
        # result.stdout/stderr None 가능 (Windows 서비스 컨텍스트에서 capture 실패 케이스)
        _out = (result.stdout or '').strip()
        _err = (result.stderr or '').strip()
        if result.returncode == 0:
            logger.info(f'[BACKUP] 완료:\n{_out}' if _out else '[BACKUP] 완료 (stdout 없음)')
        else:
            logger.error(f'[BACKUP] 실패 (exit={result.returncode}):\n{_err}')
            _notify_admin(
                'daily_backup_fail',
                f':warning: 일 백업 실패 (exit={result.returncode})\n```{(_err or "(stderr 없음)")[:1000]}```',
            )
    except Exception as exc:
        logger.error(f'[BACKUP] 예외: {exc}', exc_info=True)


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
                f'Caddy가 만료 ~30일 전 자동 갱신하므로 보통 곧 해결됨. '
                f'며칠 뒤에도 이 알림이 계속 오면 자동 갱신 실패 → Caddy(네이티브 프로세스) 로그·재시작 점검.'
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
