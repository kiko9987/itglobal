"""
APScheduler 백그라운드 폴링 스케줄러

- 당근 자동 연동 시트: KARROT_SYNC_INTERVAL_MIN 분 주기로 sync_karrot()
- 향후: Gmail 폴링, 기타 채널 추가
"""

import os
from typing import Optional

from apscheduler.schedulers.background import BackgroundScheduler

from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)

_scheduler: Optional[BackgroundScheduler] = None


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

    # 슬랙 워크플로우 → 메인 시트 직접 추가 전화 lead 보정 (2분 주기)
    _scheduler.add_job(
        _safe_workflow_phone_sync,
        'interval',
        minutes=2,
        id='workflow_phone_sync',
    )
    jobs.append('워크플로 전화 lead 2분')

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


def _safe_karrot_sync():
    """예외가 스케줄러를 중단시키지 않도록 wrapper"""
    try:
        from dashboard.services.lead_sync import sync_karrot
        sync_karrot()
    except Exception as exc:
        logger.error(f'[SCHED] 당근 sync 실행 실패: {exc}', exc_info=True)


def _safe_workflow_phone_sync():
    """슬랙 워크플로우가 메인 시트에 직접 추가한 전화 lead 보정"""
    try:
        from dashboard.services.lead_sync import sync_workflow_phone_leads
        sync_workflow_phone_leads()
    except Exception as exc:
        logger.error(f'[SCHED] 워크플로 전화 lead 보정 실패: {exc}', exc_info=True)


def _safe_homepage_sync():
    """예외가 스케줄러를 중단시키지 않도록 wrapper"""
    try:
        from dashboard.services.homepage_mail_sync import sync_homepage_email
        sync_homepage_email()
    except Exception as exc:
        logger.error(f'[SCHED] 홈페이지 메일 sync 실행 실패: {exc}', exc_info=True)


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


def trigger_karrot_sync_now() -> dict:
    """수동 트리거 (테스트/디버깅용)"""
    from dashboard.services.lead_sync import sync_karrot
    return sync_karrot()


def trigger_homepage_sync_now() -> dict:
    """수동 트리거"""
    from dashboard.services.homepage_mail_sync import sync_homepage_email
    return sync_homepage_email()
