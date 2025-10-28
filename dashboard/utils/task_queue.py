"""
RQ (Redis Queue) 작업 큐 시스템
비동기 백그라운드 작업 처리 (보고서 생성, 대량 데이터 처리 등)

Note:
    - Windows에서는 Worker가 제한적으로 동작할 수 있습니다.
    - 프로덕션에서는 Linux/Docker 환경에서 워커를 실행하세요.
    - Windows에서 테스트 시: SimpleWorker 사용 (signal 대신 polling)
"""

import os
import sys
import logging
from typing import Optional, Any, Callable
from datetime import timedelta

# Windows 호환성을 위한 signal 처리
if sys.platform == 'win32':
    # Windows에서 SIGUSR1이 없는 문제 해결
    import signal
    if not hasattr(signal, 'SIGUSR1'):
        signal.SIGUSR1 = None
    if not hasattr(signal, 'SIGUSR2'):
        signal.SIGUSR2 = None

from rq import Queue, Worker, SimpleWorker
from rq.job import Job
from dashboard.utils.redis_client import get_redis_client

logger = logging.getLogger(__name__)

# 전역 큐 인스턴스 (싱글톤)
_default_queue: Optional[Queue] = None
_high_priority_queue: Optional[Queue] = None
_low_priority_queue: Optional[Queue] = None


def get_task_queue(priority: str = 'default') -> Queue:
    """
    RQ 작업 큐 가져오기 (우선순위별)

    Args:
        priority: 'high', 'default', 'low'

    Returns:
        Queue 인스턴스
    """
    global _default_queue, _high_priority_queue, _low_priority_queue

    try:
        redis_client = get_redis_client()
        redis_conn = redis_client.redis

        if priority == 'high':
            if _high_priority_queue is None:
                _high_priority_queue = Queue('high', connection=redis_conn, default_timeout=600)
                logger.info("High priority 작업 큐 초기화 완료")
            return _high_priority_queue
        elif priority == 'low':
            if _low_priority_queue is None:
                _low_priority_queue = Queue('low', connection=redis_conn, default_timeout=1800)
                logger.info("Low priority 작업 큐 초기화 완료")
            return _low_priority_queue
        else:
            if _default_queue is None:
                _default_queue = Queue('default', connection=redis_conn, default_timeout=1200)
                logger.info("Default 작업 큐 초기화 완료")
            return _default_queue

    except Exception as e:
        logger.error(f"작업 큐 초기화 실패: {e}")
        raise


def enqueue_task(
    func: Callable,
    *args,
    priority: str = 'default',
    job_timeout: Optional[int] = None,
    result_ttl: int = 3600,
    failure_ttl: int = 86400,
    **kwargs
) -> Job:
    """
    작업을 큐에 추가

    Args:
        func: 실행할 함수
        *args: 함수 인자
        priority: 우선순위 ('high', 'default', 'low')
        job_timeout: 작업 타임아웃 (초)
        result_ttl: 결과 보관 시간 (초, 기본 1시간)
        failure_ttl: 실패 기록 보관 시간 (초, 기본 24시간)
        **kwargs: 함수 키워드 인자

    Returns:
        Job 인스턴스
    """
    try:
        queue = get_task_queue(priority)

        job = queue.enqueue(
            func,
            *args,
            **kwargs,
            job_timeout=job_timeout,
            result_ttl=result_ttl,
            failure_ttl=failure_ttl
        )

        logger.info(f"작업 큐 추가 완료: {func.__name__} (job_id={job.id}, priority={priority})")
        return job

    except Exception as e:
        logger.error(f"작업 큐 추가 실패: {func.__name__} - {e}")
        raise


def get_job(job_id: str) -> Optional[Job]:
    """
    Job ID로 작업 조회

    Args:
        job_id: 작업 ID

    Returns:
        Job 인스턴스 또는 None
    """
    try:
        redis_client = get_redis_client()
        job = Job.fetch(job_id, connection=redis_client.redis)
        return job
    except Exception as e:
        logger.warning(f"작업 조회 실패: {job_id} - {e}")
        return None


def get_job_status(job_id: str) -> dict:
    """
    작업 상태 조회

    Args:
        job_id: 작업 ID

    Returns:
        상태 정보 딕셔너리
    """
    job = get_job(job_id)

    if job is None:
        return {
            'status': 'not_found',
            'message': '작업을 찾을 수 없습니다.'
        }

    return {
        'status': job.get_status(),
        'result': job.result if job.is_finished else None,
        'exc_info': job.exc_info if job.is_failed else None,
        'created_at': job.created_at.isoformat() if job.created_at else None,
        'started_at': job.started_at.isoformat() if job.started_at else None,
        'ended_at': job.ended_at.isoformat() if job.ended_at else None,
        'meta': job.meta
    }


def cancel_job(job_id: str) -> bool:
    """
    작업 취소

    Args:
        job_id: 작업 ID

    Returns:
        성공 여부
    """
    try:
        job = get_job(job_id)
        if job:
            job.cancel()
            logger.info(f"작업 취소 완료: {job_id}")
            return True
        return False
    except Exception as e:
        logger.error(f"작업 취소 실패: {job_id} - {e}")
        return False


def get_queue_stats(priority: str = 'default') -> dict:
    """
    큐 통계 조회

    Args:
        priority: 우선순위 ('high', 'default', 'low')

    Returns:
        통계 딕셔너리
    """
    try:
        queue = get_task_queue(priority)

        return {
            'name': queue.name,
            'count': len(queue),  # 대기 중인 작업 수
            'started_count': queue.started_job_registry.count,  # 실행 중인 작업 수
            'finished_count': queue.finished_job_registry.count,  # 완료된 작업 수
            'failed_count': queue.failed_job_registry.count,  # 실패한 작업 수
            'scheduled_count': queue.scheduled_job_registry.count,  # 예약된 작업 수
            'deferred_count': queue.deferred_job_registry.count  # 지연된 작업 수
        }

    except Exception as e:
        logger.error(f"큐 통계 조회 실패: {priority} - {e}")
        return {}


def start_worker(queues: list = None, burst: bool = False) -> Worker:
    """
    RQ 워커 시작

    Args:
        queues: 처리할 큐 리스트 (기본값: ['high', 'default', 'low'])
        burst: True면 큐가 비면 종료, False면 계속 대기

    Returns:
        Worker 인스턴스

    Note:
        프로덕션에서는 별도 프로세스로 워커를 실행해야 합니다:
        $ rq worker high default low --url redis://localhost:6379/0

        Windows에서는 SimpleWorker를 사용합니다 (signal 미사용).
    """
    if queues is None:
        queues = ['high', 'default', 'low']

    try:
        redis_client = get_redis_client()
        redis_conn = redis_client.redis

        queue_objects = [Queue(q, connection=redis_conn) for q in queues]

        # Windows에서는 SimpleWorker 사용 (signal 없이 polling으로 동작)
        if sys.platform == 'win32':
            worker = SimpleWorker(queue_objects, connection=redis_conn)
            logger.info(f"RQ SimpleWorker 시작 (Windows): queues={queues}, burst={burst}")
        else:
            worker = Worker(queue_objects, connection=redis_conn)
            logger.info(f"RQ 워커 시작: queues={queues}, burst={burst}")

        worker.work(burst=burst)
        return worker

    except Exception as e:
        logger.error(f"워커 시작 실패: {e}")
        raise


# ===== 예제 백그라운드 작업들 =====

def example_generate_report(project_id: str, user_email: str) -> dict:
    """
    예제: 프로젝트 보고서 생성 (시간이 오래 걸리는 작업)

    실제 사용 예시:
    >>> from dashboard.utils.task_queue import enqueue_task, example_generate_report
    >>> job = enqueue_task(example_generate_report, 'PRJ-001', 'user@example.com', priority='high')
    >>> # 나중에 결과 확인
    >>> status = get_job_status(job.id)
    """
    import time
    from datetime import datetime

    logger.info(f"보고서 생성 시작: project_id={project_id}, user={user_email}")

    # 시뮬레이션: 실제로는 데이터 수집, 분석, PDF 생성 등
    time.sleep(5)

    result = {
        'project_id': project_id,
        'user_email': user_email,
        'report_url': f'/reports/{project_id}.pdf',
        'generated_at': datetime.now().isoformat(),
        'status': 'completed'
    }

    logger.info(f"보고서 생성 완료: {result}")
    return result


def example_bulk_data_sync(sheet_ids: list) -> dict:
    """
    예제: 여러 Google Sheets 대량 동기화

    실제 사용 예시:
    >>> job = enqueue_task(example_bulk_data_sync, ['sheet1', 'sheet2', 'sheet3'], priority='low')
    """
    import time
    from datetime import datetime

    logger.info(f"대량 데이터 동기화 시작: {len(sheet_ids)}개 시트")

    results = []
    for sheet_id in sheet_ids:
        # 시뮬레이션: 실제로는 Google Sheets API 호출
        time.sleep(2)
        results.append({
            'sheet_id': sheet_id,
            'status': 'synced',
            'row_count': 100
        })

    logger.info(f"대량 데이터 동기화 완료: {len(results)}개 시트")
    return {
        'total_sheets': len(sheet_ids),
        'synced_sheets': len(results),
        'completed_at': datetime.now().isoformat(),
        'results': results
    }


def example_send_notification_batch(user_emails: list, message: str) -> dict:
    """
    예제: 대량 알림 발송
    """
    import time
    from datetime import datetime

    logger.info(f"대량 알림 발송 시작: {len(user_emails)}명")

    sent_count = 0
    for email in user_emails:
        # 시뮬레이션: 실제로는 이메일/SMS 발송
        time.sleep(0.5)
        sent_count += 1

    logger.info(f"대량 알림 발송 완료: {sent_count}명")
    return {
        'total_recipients': len(user_emails),
        'sent_count': sent_count,
        'message': message,
        'completed_at': datetime.now().isoformat()
    }
