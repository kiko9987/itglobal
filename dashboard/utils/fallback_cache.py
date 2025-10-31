"""
Fallback 메모리 캐시 (Redis 장애 시 대체용)

Redis가 다운되었을 때 최소한의 기능을 유지하기 위한 인메모리 캐시
"""

import time
import threading
from typing import Any, Optional, Dict
import logging

logger = logging.getLogger(__name__)


class FallbackMemoryCache:
    """
    Redis 장애 시 사용할 메모리 캐시

    특징:
    - TTL 지원
    - 스레드 안전
    - 자동 만료 처리
    - 최대 크기 제한 (LRU 방식)
    """

    def __init__(self, max_size: int = 1000):
        """
        Args:
            max_size: 최대 캐시 항목 수 (초과 시 LRU 제거)
        """
        self._cache: Dict[str, tuple] = {}  # {key: (value, expire_at)}
        self._lock = threading.RLock()
        self._max_size = max_size
        self._access_order = []  # LRU 추적용
        logger.warning(f"FallbackMemoryCache 초기화 (max_size={max_size})")
        logger.warning("⚠️ Redis 대신 메모리 캐시 사용 중 - 재시작 시 데이터 손실")

    def get(self, key: str) -> Optional[Any]:
        """캐시에서 값 가져오기"""
        with self._lock:
            if key not in self._cache:
                return None

            value, expire_at = self._cache[key]

            # 만료 확인
            if expire_at and time.time() > expire_at:
                del self._cache[key]
                if key in self._access_order:
                    self._access_order.remove(key)
                return None

            # LRU 업데이트
            if key in self._access_order:
                self._access_order.remove(key)
            self._access_order.append(key)

            return value

    def set(self, key: str, value: Any, ex: Optional[int] = None) -> bool:
        """캐시에 값 저장 (TTL 지원)"""
        with self._lock:
            # 최대 크기 초과 시 LRU 제거
            if len(self._cache) >= self._max_size and key not in self._cache:
                if self._access_order:
                    oldest_key = self._access_order.pop(0)
                    if oldest_key in self._cache:
                        del self._cache[oldest_key]
                        logger.debug(f"LRU 제거: {oldest_key}")

            # 만료 시간 계산
            expire_at = None
            if ex is not None:
                expire_at = time.time() + ex

            self._cache[key] = (value, expire_at)

            # LRU 업데이트
            if key in self._access_order:
                self._access_order.remove(key)
            self._access_order.append(key)

            return True

    def delete(self, *keys: str) -> int:
        """키 삭제"""
        with self._lock:
            deleted_count = 0
            for key in keys:
                if key in self._cache:
                    del self._cache[key]
                    if key in self._access_order:
                        self._access_order.remove(key)
                    deleted_count += 1
            return deleted_count

    def exists(self, *keys: str) -> int:
        """키 존재 여부 확인"""
        with self._lock:
            count = 0
            for key in keys:
                if key in self._cache:
                    _, expire_at = self._cache[key]
                    # 만료되지 않은 키만 카운트
                    if not expire_at or time.time() <= expire_at:
                        count += 1
            return count

    def expire(self, key: str, seconds: int) -> bool:
        """키 TTL 설정"""
        with self._lock:
            if key not in self._cache:
                return False

            value, _ = self._cache[key]
            expire_at = time.time() + seconds
            self._cache[key] = (value, expire_at)
            return True

    def ttl(self, key: str) -> int:
        """키의 남은 TTL 조회"""
        with self._lock:
            if key not in self._cache:
                return -2  # 키 없음

            _, expire_at = self._cache[key]
            if not expire_at:
                return -1  # 만료 없음

            remaining = int(expire_at - time.time())
            return max(0, remaining)

    def keys(self, pattern: str = '*') -> list:
        """패턴에 맞는 키 목록 조회 (간단한 와일드카드 지원)"""
        with self._lock:
            import fnmatch
            current_time = time.time()

            # 만료된 키 제거
            expired_keys = []
            for key, (_, expire_at) in self._cache.items():
                if expire_at and current_time > expire_at:
                    expired_keys.append(key)

            for key in expired_keys:
                del self._cache[key]
                if key in self._access_order:
                    self._access_order.remove(key)

            # 패턴 매칭
            matching_keys = []
            for key in self._cache.keys():
                if fnmatch.fnmatch(key, pattern):
                    matching_keys.append(key)

            return matching_keys

    def flushdb(self) -> bool:
        """모든 캐시 삭제"""
        with self._lock:
            self._cache.clear()
            self._access_order.clear()
            logger.warning("FallbackMemoryCache 전체 삭제")
            return True

    def get_stats(self) -> Dict[str, Any]:
        """캐시 통계"""
        with self._lock:
            return {
                'total_items': len(self._cache),
                'max_size': self._max_size,
                'lru_queue_length': len(self._access_order),
                'type': 'fallback_memory'
            }


# 전역 fallback 캐시 인스턴스
_fallback_cache: Optional[FallbackMemoryCache] = None


def get_fallback_cache() -> FallbackMemoryCache:
    """Fallback 캐시 싱글톤 인스턴스 반환"""
    global _fallback_cache
    if _fallback_cache is None:
        # max_size를 2000으로 증가 (프로젝트 2830개 + 메타데이터 고려)
        _fallback_cache = FallbackMemoryCache(max_size=2000)
    return _fallback_cache
