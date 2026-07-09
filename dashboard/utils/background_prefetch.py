"""
백그라운드 프리패치 시스템
캐시 만료 전에 미리 데이터를 갱신하여 API 한도 보호와 실시간성을 동시에 확보
"""

import threading
import time
import logging
import os
import requests
from typing import Optional, Callable, Dict, Any
from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor, TimeoutError as FutureTimeoutError
from .smart_cache_manager import get_smart_cache, CacheStrategy, smart_get, smart_set

logger = logging.getLogger(__name__)

def classify_error(error: Exception) -> str:
    """
    에러를 분류하여 적절한 재시도 전략을 결정

    Returns:
        'temporary': 일시적 에러 (재시도 권장)
        'rate_limit': API 제한 에러 (긴 백오프 필요)
        'permanent': 영구적 에러 (재시도 비권장)
    """
    error_str = str(error).lower()
    error_type = type(error).__name__

    # API 제한 관련 에러
    if any(keyword in error_str for keyword in ['rate limit', 'quota', 'too many requests', '429']):
        return 'rate_limit'

    # 네트워크 관련 일시적 에러
    if isinstance(error, (requests.ConnectionError, requests.Timeout, OSError)):
        return 'temporary'

    if any(keyword in error_str for keyword in ['timeout', 'connection', 'network', 'socket']):
        return 'temporary'

    # HTTP 상태 코드 기반 분류
    if hasattr(error, 'response') and hasattr(error.response, 'status_code'):
        status_code = error.response.status_code
        if status_code in [429, 502, 503, 504]:  # 일시적 서버 에러
            return 'temporary'
        elif status_code in [401, 403, 404]:  # 인증/권한/리소스 문제
            return 'permanent'

    # 설정 관련 영구적 에러
    if any(keyword in error_str for keyword in ['authentication', 'authorization', 'permission', 'credential']):
        return 'permanent'

    # 기본값: 일시적 에러로 처리
    return 'temporary'

def calculate_backoff_time(error_type: str, attempt_count: int, max_backoff: int = 300) -> int:
    """
    에러 타입과 시도 횟수에 따른 백오프 시간 계산

    Args:
        error_type: classify_error()의 반환값
        attempt_count: 연속 에러 횟수
        max_backoff: 최대 백오프 시간 (초)

    Returns:
        백오프 시간 (초)
    """
    if error_type == 'rate_limit':
        # API 제한: 더 긴 백오프 (5분부터 시작)
        base_time = 300  # 5분
        backoff = min(base_time * (2 ** min(attempt_count, 3)), max_backoff * 2)  # 최대 10분
    elif error_type == 'temporary':
        # 일시적 에러: 표준 지수 백오프
        base_time = 30  # 30초
        backoff = min(base_time * (2 ** min(attempt_count, 4)), max_backoff)
    else:  # permanent
        # 영구적 에러: 매우 긴 백오프 (1시간)
        backoff = max_backoff * 12  # 1시간

    return backoff

class BackgroundPrefetch:
    """백그라운드에서 캐시 데이터를 미리 갱신하는 시스템"""

    def __init__(self):
        self._running = False
        self._thread: Optional[threading.Thread] = None
        self._lock = threading.RLock()
        self._prefetch_jobs: Dict[str, Dict] = {}
        self._error_threshold = 5  # 연속 에러 허용 횟수
        self._max_backoff = 300   # 최대 백오프 시간 (5분)
        self._shutdown_event = threading.Event()

    def register_prefetch_job(self,
                            cache_key: str,
                            data_fetcher: Callable[[], Any],
                            strategy: CacheStrategy = CacheStrategy.CRITICAL_DATA,
                            prefetch_threshold: float = 0.8):
        """
        프리패치 작업 등록

        Args:
            cache_key: 캐시 키
            data_fetcher: 데이터를 가져오는 함수
            strategy: 캐시 전략
            prefetch_threshold: 캐시 TTL의 몇 %에서 프리패치할지 (0.8 = 80%)
        """
        with self._lock:
            self._prefetch_jobs[cache_key] = {
                'data_fetcher': data_fetcher,
                'strategy': strategy,
                'prefetch_threshold': prefetch_threshold,
                'last_prefetch': None,
                'prefetch_count': 0,
                'error_count': 0,
                'consecutive_errors': 0,
                'last_error': None,
                'is_disabled': False,
                'backoff_until': None,
                'created_at': datetime.now()
            }
            logger.info(f"[PREFETCH] 작업 등록: {cache_key}")

    def start(self):
        """백그라운드 프리패치 시작"""
        with self._lock:
            if self._running:
                return

            self._running = True
            self._thread = threading.Thread(target=self._prefetch_loop, daemon=True)
            self._thread.start()
            logger.info("[PREFETCH] 백그라운드 프리패치 시작")

    def stop(self):
        """백그라운드 프리패치 중지"""
        with self._lock:
            if not self._running:
                return

            self._running = False
            self._shutdown_event.set()

            if self._thread and self._thread.is_alive():
                self._thread.join(timeout=10)  # 더 긴 대기시간

                # 강제 종료가 필요한 경우
                if self._thread.is_alive():
                    logger.warning("[PREFETCH] 정상 종료 실패, 강제 종료 시도")

            logger.info("[PREFETCH] 백그라운드 프리패치 중지")

    def _prefetch_loop(self):
        """프리패치 메인 루프"""
        loop_error_count = 0

        while self._running and not self._shutdown_event.is_set():
            try:
                self._check_and_prefetch()
                loop_error_count = 0  # 성공 시 에러 카운트 리셋

                # 인터럽트 가능한 대기
                if self._shutdown_event.wait(timeout=30):  # 30초마다 체크
                    break

            except KeyboardInterrupt:
                logger.info("[PREFETCH] 키보드 인터럽트로 중지")
                break
            except Exception as e:
                loop_error_count += 1
                backoff_time = min(60 * (2 ** min(loop_error_count - 1, 4)), 300)  # 지수 백오프, 최대 5분

                logger.error(f"[PREFETCH] 프리패치 루프 오류 #{loop_error_count}: {e}")
                logger.info(f"[PREFETCH] {backoff_time}초 후 재시도")

                if self._shutdown_event.wait(timeout=backoff_time):
                    break

        logger.info("[PREFETCH] 프리패치 루프 종료")

    def _check_and_prefetch(self):
        """캐시 상태를 확인하고 필요시 프리패치 실행"""
        cache = get_smart_cache()
        current_time = time.time()

        for cache_key, job_info in list(self._prefetch_jobs.items()):
            try:
                # 비활성화된 작업이나 백오프 중인 작업 스킵
                if job_info.get('is_disabled', False):
                    continue

                backoff_until = job_info.get('backoff_until')
                if backoff_until and current_time < backoff_until:
                    continue

                # 캐시 만료 시간 확인
                cache_info = cache._get_cache_item_info(cache_key)
                if not cache_info:
                    # 캐시가 없으면 즉시 프리패치
                    self._execute_prefetch(cache_key, job_info)
                    continue

                # TTL 기반 프리패치 판단
                ttl = cache_info.get('ttl', 300)
                created_time = cache_info.get('created', time.time())

                elapsed_ratio = (current_time - created_time) / ttl
                threshold = job_info['prefetch_threshold']

                if elapsed_ratio >= threshold:
                    self._execute_prefetch(cache_key, job_info)

            except Exception as e:
                logger.error(f"[PREFETCH] 캐시 체크 오류 ({cache_key}): {e}")
                self._handle_job_error(cache_key, job_info, e)

    def _execute_prefetch(self, cache_key: str, job_info: Dict):
        """실제 프리패치 실행 (무효화 마커 체크 포함 - 레이스 컨디션 방지)"""
        start_time = time.time()
        data_fetch_start = None

        try:
            data_fetcher = job_info['data_fetcher']
            strategy = job_info['strategy']

            logger.debug(f"[PREFETCH] 데이터 갱신 시작: {cache_key}")

            # 스레드 안전한 타임아웃 구현 (30초)
            # signal.SIGALRM은 메인 스레드에서만 작동하므로 ThreadPoolExecutor 사용
            # wait=False로 타임아웃 시 즉시 종료 (wait=True면 작업 완료까지 대기)
            executor = ThreadPoolExecutor(max_workers=1)
            future = None
            try:
                # 🔒 데이터 가져오기 시작 시각 기록 (무효화 마커 비교용)
                data_fetch_start = time.time()

                future = executor.submit(data_fetcher)
                try:
                    # 타임아웃 30초로 데이터 가져오기
                    new_data = future.result(timeout=30)
                except FutureTimeoutError:
                    future.cancel()  # 명시적으로 future 취소

                    # 🔧 타임아웃 발생 시 Google Sheets 서비스 재초기화
                    # OpenSSL 세션이 손상되었을 수 있으므로 서비스 객체 재생성
                    logger.warning(f"[PREFETCH] 타임아웃 발생 ({cache_key}), Google Sheets 서비스 재초기화 시도")
                    try:
                        from dashboard.utils.google_sheets import GoogleSheetsManager
                        sheets_manager = GoogleSheetsManager()
                        if sheets_manager.reset_service():
                            logger.info("[PREFETCH] Google Sheets 서비스 재초기화 성공")
                        else:
                            logger.error("[PREFETCH] Google Sheets 서비스 재초기화 실패")
                    except Exception as reset_err:
                        logger.error(f"[PREFETCH] 서비스 재초기화 중 오류: {reset_err}")

                    raise TimeoutError(f"프리패치 타임아웃: {cache_key}")

                # 데이터 유효성 검증
                if new_data is None:
                    raise ValueError("데이터 가져오기 실패: None 반환")

                # 🔒 레이스 컨디션 방지: 무효화 마커 체크 (1차 검증)
                from dashboard.utils.smart_cache_manager import smart_get_invalidation_marker
                invalidation_time = smart_get_invalidation_marker(cache_key)

                if invalidation_time and data_fetch_start < invalidation_time:
                    logger.warning(
                        f"[PREFETCH] 캐시 쓰기 스킵: {cache_key} - 데이터 가져오기 시작이 무효화 마커보다 이전 "
                        f"(시작: {data_fetch_start:.2f}, 무효화: {invalidation_time:.2f})"
                    )
                    # 오래된 데이터는 캐시에 쓰지 않고 무시
                    return

                # 캐시 업데이트 (smart_set 내부에서도 마커 체크함 - 이중 보호)
                # 🔒 fetched_at 전달하여 set() 메서드에서도 마커 체크
                smart_set(cache_key, new_data, strategy, fetched_at=data_fetch_start)
            finally:
                # 미완료 future가 있다면 취소 시도
                if future and not future.done():
                    future.cancel()
                executor.shutdown(wait=False)  # 타임아웃 발생 시에도 즉시 종료

            # 성공 시 통계 업데이트 및 에러 상태 리셋
            execution_time = time.time() - start_time
            job_info['last_prefetch'] = datetime.now()
            job_info['prefetch_count'] += 1
            job_info['consecutive_errors'] = 0  # 성공 시 연속 에러 리셋
            job_info['backoff_until'] = None    # 백오프 해제
            job_info['is_disabled'] = False     # 재활성화

            logger.info(f"[PREFETCH] 데이터 갱신 완료: {cache_key} ({execution_time:.2f}초)")

        except Exception as e:
            self._handle_job_error(cache_key, job_info, e)

    def _handle_job_error(self, cache_key: str, job_info: Dict, error: Exception):
        """프리패치 작업 에러 처리 (향상된 에러 분류 및 백오프 전략)"""
        job_info['error_count'] += 1
        job_info['consecutive_errors'] += 1
        job_info['last_error'] = str(error)

        # 에러 타입 분류
        error_type = classify_error(error)
        job_info['last_error_type'] = error_type

        # 에러 타입별 로깅
        if error_type == 'rate_limit':
            logger.warning(f"[PREFETCH] API 제한 에러 ({cache_key}): {error}")
        elif error_type == 'temporary':
            logger.warning(f"[PREFETCH] 일시적 에러 ({cache_key}): {error}")
        else:  # permanent
            logger.error(f"[PREFETCH] 영구적 에러 ({cache_key}): {error}")

        # 연속 에러 처리
        if job_info['consecutive_errors'] >= self._error_threshold:
            # 에러 타입별 백오프 시간 계산
            backoff_time = calculate_backoff_time(
                error_type,
                job_info['consecutive_errors'] - self._error_threshold + 1,
                self._max_backoff
            )
            job_info['backoff_until'] = time.time() + backoff_time

            logger.warning(f"[PREFETCH] 작업 일시 비활성화 ({cache_key}): {backoff_time}초 후 재시도 (에러타입: {error_type})")

            # 영구적 에러나 너무 많은 에러 발생 시 작업 비활성화
            disable_threshold = self._error_threshold * 2
            if error_type == 'permanent' or job_info['consecutive_errors'] >= disable_threshold:
                job_info['is_disabled'] = True
                reason = "영구적 에러" if error_type == 'permanent' else f"연속 {job_info['consecutive_errors']}회 오류"
                logger.error(f"[PREFETCH] 작업 비활성화 ({cache_key}): {reason}")

    def get_stats(self) -> Dict[str, Any]:
        """프리패치 통계 반환"""
        with self._lock:
            current_time = time.time()
            active_jobs = 0
            disabled_jobs = 0
            backoff_jobs = 0

            stats = {
                'running': self._running,
                'jobs_count': len(self._prefetch_jobs),
                'jobs': {}
            }

            for cache_key, job_info in self._prefetch_jobs.items():
                # 작업 상태 분류
                if job_info.get('is_disabled', False):
                    disabled_jobs += 1
                    status = 'disabled'
                elif job_info.get('backoff_until') and current_time < job_info.get('backoff_until', 0):
                    backoff_jobs += 1
                    status = 'backoff'
                else:
                    active_jobs += 1
                    status = 'active'

                stats['jobs'][cache_key] = {
                    'status': status,
                    'prefetch_count': job_info['prefetch_count'],
                    'error_count': job_info['error_count'],
                    'consecutive_errors': job_info.get('consecutive_errors', 0),
                    'last_prefetch': job_info['last_prefetch'].isoformat() if job_info['last_prefetch'] else None,
                    'last_error': job_info.get('last_error'),
                    'backoff_until': datetime.fromtimestamp(job_info['backoff_until']).isoformat() if job_info.get('backoff_until') else None,
                    'created_at': job_info.get('created_at', datetime.now()).isoformat()
                }

            stats.update({
                'active_jobs': active_jobs,
                'disabled_jobs': disabled_jobs,
                'backoff_jobs': backoff_jobs
            })

            return stats

    def reset_job_errors(self, cache_key: str) -> bool:
        """특정 작업의 에러 상태 리셋"""
        with self._lock:
            if cache_key not in self._prefetch_jobs:
                return False

            job_info = self._prefetch_jobs[cache_key]
            job_info['consecutive_errors'] = 0
            job_info['error_count'] = 0
            job_info['backoff_until'] = None
            job_info['is_disabled'] = False
            job_info['last_error'] = None

            logger.info(f"[PREFETCH] 작업 에러 상태 리셋: {cache_key}")
            return True

    def disable_job(self, cache_key: str) -> bool:
        """특정 작업 비활성화"""
        with self._lock:
            if cache_key not in self._prefetch_jobs:
                return False

            self._prefetch_jobs[cache_key]['is_disabled'] = True
            logger.info(f"[PREFETCH] 작업 비활성화: {cache_key}")
            return True

    def enable_job(self, cache_key: str) -> bool:
        """특정 작업 활성화"""
        with self._lock:
            if cache_key not in self._prefetch_jobs:
                return False

            job_info = self._prefetch_jobs[cache_key]
            job_info['is_disabled'] = False
            job_info['backoff_until'] = None

            logger.info(f"[PREFETCH] 작업 활성화: {cache_key}")
            return True

# 전역 프리패치 인스턴스
_global_prefetch = BackgroundPrefetch()

def get_background_prefetch() -> BackgroundPrefetch:
    """글로벌 백그라운드 프리패치 인스턴스 반환"""
    return _global_prefetch

def start_background_prefetch():
    """백그라운드 프리패치 시작"""
    _global_prefetch.start()

def stop_background_prefetch():
    """백그라운드 프리패치 중지"""
    _global_prefetch.stop()

def init_background_prefetch(max_retries=3, retry_delay=5):
    """
    백그라운드 프리패치 시스템 초기화 (재시도 로직 포함)

    Args:
        max_retries: 최대 재시도 횟수
        retry_delay: 재시도 간격 (초)
    """
    for attempt in range(max_retries + 1):
        try:
            logger.info(f"[PREFETCH] 백그라운드 프리패치 시스템 초기화 시도 {attempt + 1}/{max_retries + 1}")

            # 1단계: 백그라운드 프리패치 시작
            start_background_prefetch()
            logger.debug("[PREFETCH] 백그라운드 프리패치 시작 완료")

            # 2단계: 프리패치 작업 등록
            register_sheets_prefetch()
            logger.debug("[PREFETCH] 프리패치 작업 등록 완료")

            logger.info("[PREFETCH] 백그라운드 프리패치 시스템 초기화 완료")
            return True

        except Exception as e:
            logger.warning(f"[PREFETCH] 초기화 시도 {attempt + 1} 실패: {e}")

            if attempt < max_retries:
                logger.info(f"[PREFETCH] {retry_delay}초 후 재시도...")
                time.sleep(retry_delay)
                retry_delay = min(retry_delay * 1.5, 30)  # 지수 백오프, 최대 30초
            else:
                logger.error("[PREFETCH] 모든 초기화 시도 실패 - 백그라운드 프리패치 비활성화")
                return False

def register_sheets_prefetch():
    """Google Sheets 데이터 프리패치 등록"""
    # 순환 import 방지를 위해 지연 import
    def get_fresh_data():
        from ..services.project_service import _fetch_fresh_data
        return _fetch_fresh_data()

    prefetch = get_background_prefetch()
    prefetch.register_prefetch_job(
        cache_key="current_sheet_data",
        data_fetcher=get_fresh_data,
        strategy=CacheStrategy.CRITICAL_DATA,
        # 2026-07-09 매니저는 관리 사이트만 사용 → Google 지연이 UX에 영향 없도록
        # 공격적으로 미리 갱신. CRITICAL_DATA TTL 20분 기준 0.15 = 약 3분 시점에 리프레시.
        # 시트가 실질적으로 절대 만료되지 않도록 유지.
        prefetch_threshold=0.15,
    )
    logger.info("[PREFETCH] Google Sheets 프리패치 등록 완료 (threshold=0.15, ≈3분 주기)")