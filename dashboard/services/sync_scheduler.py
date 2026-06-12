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


def _safe_homepage_sync():
    """예외가 스케줄러를 중단시키지 않도록 wrapper"""
    try:
        from dashboard.services.homepage_mail_sync import sync_homepage_email
        sync_homepage_email()
    except Exception as exc:
        logger.error(f'[SCHED] 홈페이지 메일 sync 실행 실패: {exc}', exc_info=True)


def trigger_karrot_sync_now() -> dict:
    """수동 트리거 (테스트/디버깅용)"""
    from dashboard.services.lead_sync import sync_karrot
    return sync_karrot()


def trigger_homepage_sync_now() -> dict:
    """수동 트리거"""
    from dashboard.services.homepage_mail_sync import sync_homepage_email
    return sync_homepage_email()
