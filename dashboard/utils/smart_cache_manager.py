"""
스마트 캐시 관리자
데이터 타입과 상황에 따라 적절한 캐시 전략을 자동으로 적용
"""

import time
import threading
from typing import Any, Optional, Dict, List
from datetime import datetime, timedelta
import logging
from dataclasses import dataclass
from enum import Enum

logger = logging.getLogger(__name__)

class CacheStrategy(Enum):
    """캐시 전략 유형"""
    CRITICAL_DATA = "critical_data"      # 구글 시트 데이터 (30초 TTL)
    STATIC_CONFIG = "static_config"      # 설정 데이터 (1시간 TTL)
    FOLDER_MAPPING = "folder_mapping"    # 폴더 ID 매핑 (1일 TTL)
    UI_STATE = "ui_state"               # UI 상태 (5분 TTL)
    TEMPORARY = "temporary"             # 임시 데이터 (1분 TTL)

@dataclass
class CacheConfig:
    """캐시 설정"""
    ttl: int
    auto_refresh: bool = False
    version_key: Optional[str] = None
    priority: int = 1  # 1=높음, 2=보통, 3=낮음

class SmartCacheManager:
    """스마트 캐시 관리자"""

    # 전략별 기본 설정
    STRATEGY_CONFIGS = {
        CacheStrategy.CRITICAL_DATA: CacheConfig(ttl=30, auto_refresh=True, priority=1),
        CacheStrategy.STATIC_CONFIG: CacheConfig(ttl=3600, auto_refresh=False, priority=2),
        CacheStrategy.FOLDER_MAPPING: CacheConfig(ttl=86400, auto_refresh=False, priority=3),
        CacheStrategy.UI_STATE: CacheConfig(ttl=300, auto_refresh=False, priority=2),
        CacheStrategy.TEMPORARY: CacheConfig(ttl=60, auto_refresh=False, priority=3)
    }

    def __init__(self, max_memory_mb: int = 100):
        self._cache: Dict[str, Dict] = {}
        self._lock = threading.RLock()
        self._max_memory_bytes = max_memory_mb * 1024 * 1024
        self._last_cleanup = time.time()
        self._access_log: Dict[str, List[float]] = {}

    def get(self, key: str, strategy: CacheStrategy = CacheStrategy.CRITICAL_DATA) -> Optional[Any]:
        """캐시에서 값 가져오기"""
        with self._lock:
            self._cleanup_if_needed()

            if key not in self._cache:
                logger.debug(f"캐시 미스: {key}")
                return None

            item = self._cache[key]
            current_time = time.time()

            # TTL 체크
            if current_time > item['expires']:
                del self._cache[key]
                logger.debug(f"캐시 만료: {key}")
                return None

            # 접근 로그 기록
            self._record_access(key, current_time)

            # 자동 새로고침 체크 (만료 10초 전)
            config = self.STRATEGY_CONFIGS[strategy]
            if (config.auto_refresh and
                current_time > item['expires'] - 10 and
                'refreshing' not in item):
                item['refreshing'] = True
                logger.debug(f"자동 새로고침 필요: {key}")

            logger.debug(f"캐시 히트: {key}")
            return item['value']

    def set(self, key: str, value: Any, strategy: CacheStrategy = CacheStrategy.CRITICAL_DATA,
            custom_ttl: Optional[int] = None) -> None:
        """캐시에 값 저장"""
        with self._lock:
            config = self.STRATEGY_CONFIGS[strategy]
            ttl = custom_ttl or config.ttl
            current_time = time.time()

            # 메모리 사용량 체크
            if self._should_evict_memory():
                self._evict_by_priority()

            self._cache[key] = {
                'value': value,
                'expires': current_time + ttl,
                'created': current_time,
                'last_accessed': current_time,
                'strategy': strategy,
                'priority': config.priority,
                'size_estimate': self._estimate_size(value)
            }

            # 접근 로그 초기화
            self._access_log[key] = [current_time]

            logger.debug(f"캐시 저장: {key}, 전략: {strategy.value}, TTL: {ttl}초")

    def delete(self, key: str) -> bool:
        """캐시에서 키 삭제"""
        with self._lock:
            if key in self._cache:
                del self._cache[key]
                if key in self._access_log:
                    del self._access_log[key]
                logger.debug(f"캐시 삭제: {key}")
                return True
            return False

    def invalidate_by_pattern(self, pattern: str) -> int:
        """패턴에 맞는 캐시 무효화"""
        with self._lock:
            keys_to_delete = [key for key in self._cache.keys() if pattern in key]
            for key in keys_to_delete:
                self.delete(key)
            logger.info(f"패턴 '{pattern}'으로 {len(keys_to_delete)}개 캐시 무효화")
            return len(keys_to_delete)

    def clear_by_strategy(self, strategy: CacheStrategy) -> int:
        """특정 전략의 캐시만 삭제"""
        with self._lock:
            keys_to_delete = [
                key for key, item in self._cache.items()
                if item.get('strategy') == strategy
            ]
            for key in keys_to_delete:
                self.delete(key)
            logger.info(f"전략 '{strategy.value}'로 {len(keys_to_delete)}개 캐시 삭제")
            return len(keys_to_delete)

    def get_cache_info(self) -> Dict[str, Any]:
        """캐시 상태 정보"""
        with self._lock:
            current_time = time.time()
            active_count = 0
            expired_count = 0
            total_memory = 0
            strategy_stats = {}

            for key, item in self._cache.items():
                if current_time > item['expires']:
                    expired_count += 1
                else:
                    active_count += 1

                total_memory += item.get('size_estimate', 0)

                strategy = item.get('strategy', CacheStrategy.CRITICAL_DATA)
                if strategy not in strategy_stats:
                    strategy_stats[strategy] = 0
                strategy_stats[strategy] += 1

            return {
                'total_items': len(self._cache),
                'active_items': active_count,
                'expired_items': expired_count,
                'memory_usage_mb': total_memory / (1024 * 1024),
                'strategy_distribution': {s.value: count for s, count in strategy_stats.items()},
                'hit_rate': self._calculate_hit_rate()
            }

    def _record_access(self, key: str, timestamp: float) -> None:
        """접근 로그 기록 (최근 10회만 유지)"""
        if key not in self._access_log:
            self._access_log[key] = []

        self._access_log[key].append(timestamp)
        if len(self._access_log[key]) > 10:
            self._access_log[key] = self._access_log[key][-10:]

    def _calculate_hit_rate(self) -> float:
        """캐시 히트율 계산"""
        total_accesses = sum(len(accesses) for accesses in self._access_log.values())
        if total_accesses == 0:
            return 0.0

        # 간단한 히트율 추정 (실제 구현시 더 정확한 로직 필요)
        cache_hits = len(self._cache)
        return min(cache_hits / total_accesses, 1.0) * 100

    def _should_evict_memory(self) -> bool:
        """메모리 제한 초과 체크"""
        total_memory = sum(
            item.get('size_estimate', 0)
            for item in self._cache.values()
        )
        return total_memory > self._max_memory_bytes

    def _evict_by_priority(self) -> None:
        """우선순위 기반 캐시 제거"""
        # 우선순위가 낮고 접근 빈도가 낮은 항목부터 제거
        items_by_priority = []
        current_time = time.time()

        for key, item in self._cache.items():
            if current_time > item['expires']:
                continue  # 만료된 항목은 cleanup에서 처리

            priority = item.get('priority', 3)
            last_accessed = item['last_accessed']
            access_frequency = len(self._access_log.get(key, []))

            # 점수가 낮을수록 먼저 제거
            score = priority * 1000 - access_frequency * 10 - (current_time - last_accessed)
            items_by_priority.append((score, key))

        # 점수 기준 정렬 후 하위 25% 제거
        items_by_priority.sort()
        remove_count = max(1, len(items_by_priority) // 4)

        for _, key in items_by_priority[:remove_count]:
            self.delete(key)

        logger.info(f"메모리 압박으로 {remove_count}개 캐시 제거")

    def _cleanup_if_needed(self) -> None:
        """주기적 정리"""
        current_time = time.time()

        if current_time - self._last_cleanup < 30:  # 30초마다
            return

        expired_keys = []
        for key, item in self._cache.items():
            if current_time > item['expires']:
                expired_keys.append(key)

        for key in expired_keys:
            self.delete(key)

        if expired_keys:
            logger.debug(f"만료된 캐시 정리: {len(expired_keys)}개 항목")

        self._last_cleanup = current_time

    def _estimate_size(self, value: Any) -> int:
        """값의 크기 추정 (바이트)"""
        try:
            import sys
            return sys.getsizeof(value)
        except:
            # 대략적 추정
            if isinstance(value, str):
                return len(value.encode('utf-8'))
            elif isinstance(value, (list, dict)):
                return len(str(value)) * 2
            else:
                return 1024  # 기본값

# 전역 스마트 캐시 인스턴스
_smart_cache = SmartCacheManager()

def get_smart_cache() -> SmartCacheManager:
    """스마트 캐시 인스턴스 반환"""
    return _smart_cache

# 편의 함수들
def smart_get(key: str, strategy: CacheStrategy = CacheStrategy.CRITICAL_DATA) -> Optional[Any]:
    """스마트 캐시에서 값 가져오기"""
    return _smart_cache.get(key, strategy)

def smart_set(key: str, value: Any, strategy: CacheStrategy = CacheStrategy.CRITICAL_DATA,
              ttl: Optional[int] = None) -> None:
    """스마트 캐시에 값 저장"""
    _smart_cache.set(key, value, strategy, ttl)

def smart_delete(key: str) -> bool:
    """스마트 캐시에서 키 삭제"""
    return _smart_cache.delete(key)

def smart_invalidate(pattern: str) -> int:
    """패턴 기반 캐시 무효화"""
    return _smart_cache.invalidate_by_pattern(pattern)

def smart_clear_strategy(strategy: CacheStrategy) -> int:
    """전략별 캐시 삭제"""
    return _smart_cache.clear_by_strategy(strategy)