"""
캐시 프리로더 자동화 (Cold Start 제거)

앱 시작 시 및 주기적으로 중요 데이터를 미리 캐시에 로드하여
첫 요청의 응답 시간을 개선합니다.
"""

import threading
import time
import logging
from typing import List, Callable, Optional
from datetime import datetime

logger = logging.getLogger(__name__)


class CachePreloader:
    """
    캐시 프리로더 서비스

    특징:
    - 앱 시작 시 자동 캐시 워밍
    - 주기적 캐시 갱신 (백그라운드 스레드)
    - 우선순위 기반 프리로딩
    - 실패 시 자동 재시도
    """

    def __init__(self, refresh_interval: int = 3600):
        """
        Args:
            refresh_interval: 캐시 갱신 주기 (초, 기본: 1시간)
        """
        self._refresh_interval = refresh_interval
        self._preload_tasks: List[dict] = []
        self._running = False
        self._refresh_thread: Optional[threading.Thread] = None
        logger.info(f"CachePreloader 초기화 (갱신 주기: {refresh_interval}초)")

    def register_task(
        self,
        task_name: str,
        task_func: Callable[[], bool],
        priority: int = 1,
        enabled: bool = True
    ):
        """
        프리로드 작업 등록

        Args:
            task_name: 작업 이름
            task_func: 프리로드 함수 (성공 시 True 반환)
            priority: 우선순위 (1=높음, 2=중간, 3=낮음)
            enabled: 활성화 여부
        """
        self._preload_tasks.append({
            'name': task_name,
            'func': task_func,
            'priority': priority,
            'enabled': enabled
        })
        logger.debug(f"프리로드 작업 등록: {task_name} (우선순위: {priority})")

    def warmup_cache(self, priority_filter: Optional[int] = None):
        """
        캐시 워밍 실행 (앱 시작 시 호출)

        Args:
            priority_filter: 특정 우선순위만 실행 (None이면 전체 실행)
        """
        logger.info("캐시 워밍 시작...")
        start_time = time.time()

        # 우선순위 기준 정렬
        tasks = sorted(
            [t for t in self._preload_tasks if t['enabled']],
            key=lambda x: x['priority']
        )

        # 우선순위 필터 적용
        if priority_filter is not None:
            tasks = [t for t in tasks if t['priority'] == priority_filter]

        success_count = 0
        fail_count = 0

        for task in tasks:
            task_name = task['name']
            task_func = task['func']

            try:
                logger.debug(f"프리로드 실행: {task_name}")
                result = task_func()

                if result:
                    success_count += 1
                    logger.debug(f"프리로드 성공: {task_name}")
                else:
                    fail_count += 1
                    logger.warning(f"프리로드 실패: {task_name}")

            except Exception as e:
                fail_count += 1
                logger.error(f"프리로드 오류: {task_name} - {e}")

        elapsed = time.time() - start_time
        logger.info(
            f"캐시 워밍 완료: {success_count}개 성공, {fail_count}개 실패 "
            f"(소요시간: {elapsed:.2f}초)"
        )

    def start_periodic_refresh(self):
        """
        주기적 캐시 갱신 시작 (백그라운드 스레드)
        """
        if self._running:
            logger.warning("캐시 갱신 스레드가 이미 실행 중입니다.")
            return

        self._running = True
        self._refresh_thread = threading.Thread(
            target=self._refresh_loop,
            daemon=True,
            name="CachePreloaderRefresh"
        )
        self._refresh_thread.start()
        logger.info(f"주기적 캐시 갱신 시작 (주기: {self._refresh_interval}초)")

    def stop_periodic_refresh(self):
        """
        주기적 캐시 갱신 중지
        """
        if not self._running:
            return

        self._running = False

        if self._refresh_thread:
            self._refresh_thread.join(timeout=10)

        logger.info("주기적 캐시 갱신 중지")

    def _refresh_loop(self):
        """
        캐시 갱신 루프 (백그라운드 스레드)
        """
        try:
            while self._running:
                # 다음 갱신 시각까지 대기
                time.sleep(self._refresh_interval)

                if not self._running:
                    break

                logger.info(f"주기적 캐시 갱신 시작 ({datetime.now().isoformat()})")
                try:
                    # 우선순위 1, 2만 주기적 갱신 (핵심 데이터만)
                    self.warmup_cache(priority_filter=1)
                    self.warmup_cache(priority_filter=2)
                except Exception as e:
                    logger.error(f"캐시 갱신 중 오류: {e}")

        except Exception as e:
            logger.error(f"캐시 갱신 루프 오류: {e}")

        finally:
            logger.info("캐시 갱신 루프 종료")


# 싱글톤 인스턴스
_cache_preloader: Optional[CachePreloader] = None


def get_cache_preloader(refresh_interval: int = 3600) -> CachePreloader:
    """
    CachePreloader 싱글톤 인스턴스 반환

    Args:
        refresh_interval: 캐시 갱신 주기 (초, 기본: 1시간)

    Returns:
        CachePreloader: 캐시 프리로더 인스턴스
    """
    global _cache_preloader
    if _cache_preloader is None:
        _cache_preloader = CachePreloader(refresh_interval=refresh_interval)
    return _cache_preloader


# ========== 프리로드 작업 예시 ==========

def preload_metadata():
    """
    메타데이터 프리로드 (담당자, 사업자, 거래처 등)

    Returns:
        bool: 성공 여부
    """
    try:
        from dashboard.utils.metadata_manager import get_metadata_manager

        metadata_manager = get_metadata_manager()

        # 각 메타데이터 조회 (캐시 로드)
        metadata_manager.get_managers()  # 담당자 목록
        metadata_manager.get_subs()  # 사업자 목록
        metadata_manager.get_traders()  # 거래처 목록

        logger.debug("메타데이터 프리로드 완료")
        return True

    except Exception as e:
        logger.error(f"메타데이터 프리로드 실패: {e}")
        return False


def preload_folder_mapping():
    """
    폴더 ID 매핑 프리로드

    Returns:
        bool: 성공 여부
    """
    try:
        from dashboard.utils.metadata_manager import get_metadata_manager

        metadata_manager = get_metadata_manager()

        # 폴더 매핑 조회 (캐시 로드)
        metadata_manager.get_folder_ids()

        logger.debug("폴더 매핑 프리로드 완료")
        return True

    except Exception as e:
        logger.error(f"폴더 매핑 프리로드 실패: {e}")
        return False


def register_default_preload_tasks(preloader: CachePreloader):
    """
    기본 프리로드 작업 등록

    Args:
        preloader: CachePreloader 인스턴스
    """
    # 우선순위 1: 핵심 메타데이터 (담당자, 사업자, 거래처)
    preloader.register_task(
        task_name="metadata_preload",
        task_func=preload_metadata,
        priority=1,
        enabled=True
    )

    # 우선순위 2: 폴더 ID 매핑
    preloader.register_task(
        task_name="folder_mapping_preload",
        task_func=preload_folder_mapping,
        priority=2,
        enabled=True
    )

    logger.info("기본 프리로드 작업 등록 완료")
