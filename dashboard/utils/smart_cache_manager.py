"""
간단한 TTL 캐시 관리자
기존 스마트 캐시의 복잡성을 제거하고 실제 필요한 기능만 제공
"""

import time
import threading
from typing import Any, Optional, Dict
import logging
from enum import Enum

logger = logging.getLogger(__name__)

class CacheStrategy(Enum):
    """캐시 전략 유형 (실시간성과 API 효율성 균형)"""
    CRITICAL_DATA = "critical_data"      # 구글 시트 데이터 (60초 TTL - API 호출 최적화)
    STATIC_CONFIG = "static_config"      # 설정 데이터 (1시간 TTL)
    FOLDER_MAPPING = "folder_mapping"    # 폴더 ID 매핑 (1일 TTL)
    UI_STATE = "ui_state"               # UI 상태 (5분 TTL)
    TEMPORARY = "temporary"             # 임시 데이터 (1분 TTL)
    METADATA = "metadata"               # 메타데이터 (담당자, 사업자, 거래처) (10분 TTL)

class SimpleCache:
    """간단한 TTL 기반 캐시"""

    # 전략별 TTL (초) - 실시간성과 API 효율성 균형
    TTL_MAP = {
        CacheStrategy.CRITICAL_DATA: 60,   # 60초 (협업 환경 최적 - 프리패치 75회/시간)
        CacheStrategy.STATIC_CONFIG: 3600, # 1시간 (설정 데이터)
        CacheStrategy.FOLDER_MAPPING: 86400, # 1일 (폴더 매핑)
        CacheStrategy.UI_STATE: 300,       # 5분 (UI 상태)
        CacheStrategy.TEMPORARY: 60,       # 1분 (임시 데이터)
        CacheStrategy.METADATA: 600        # 10분 (메타데이터 - 담당자, 사업자, 거래처)
    }

    def __init__(self):
        self._cache: Dict[str, Dict] = {}
        self._lock = threading.RLock()
        self._invalidation_markers: Dict[str, float] = {}  # 무효화 시각 추적 (레이스 컨디션 방지)

    def get(self, key: str, strategy: CacheStrategy = CacheStrategy.CRITICAL_DATA) -> Optional[Any]:
        """캐시에서 값 가져오기"""
        with self._lock:
            self._cleanup_expired()

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

            logger.debug(f"캐시 히트: {key}")
            return item['value']

    def set(self, key: str, value: Any, strategy: CacheStrategy = CacheStrategy.CRITICAL_DATA,
            custom_ttl: Optional[int] = None, fetched_at: Optional[float] = None) -> None:
        """캐시에 값 저장 (무효화 마커 체크 포함 - 레이스 컨디션 방지)

        Args:
            key: 캐시 키
            value: 저장할 값
            strategy: 캐시 전략
            custom_ttl: 커스텀 TTL (선택사항)
            fetched_at: 데이터 수집 시작 시각 (Unix timestamp) - 없으면 현재 시각 사용
        """
        with self._lock:
            ttl = custom_ttl or self.TTL_MAP[strategy]
            current_time = time.time()

            # 🔒 레이스 컨디션 방지: 무효화 마커보다 오래된 데이터는 캐시에 쓰지 않음
            # fetched_at이 제공되면 데이터 수집 시작 시각으로 비교 (정확한 보호)
            # 없으면 현재 시각으로 폴백 (하위 호환성)
            data_timestamp = fetched_at if fetched_at is not None else current_time
            invalidation_time = self._invalidation_markers.get(key)

            if invalidation_time and data_timestamp < invalidation_time:
                logger.warning(
                    f"캐시 쓰기 거부: {key} - 데이터 수집이 무효화 마커보다 이전 "
                    f"(수집 시작: {data_timestamp:.2f}, 무효화: {invalidation_time:.2f})"
                )
                return

            self._cache[key] = {
                'value': value,
                'expires': current_time + ttl,
                'strategy': strategy,
                'created_at': current_time  # 실제 데이터 생성 시간 추적
            }

            logger.debug(f"캐시 저장: {key}, 전략: {strategy.value}, TTL: {ttl}초, "
                        f"생성시간: {current_time}, 수집시각: {data_timestamp}")

    def delete(self, key: str, set_marker: bool = True) -> bool:
        """캐시에서 키 삭제 (무효화 마커 설정 옵션)

        Args:
            key: 캐시 키
            set_marker: True이면 무효화 마커를 설정하여 레이스 컨디션 방지
        """
        with self._lock:
            if key in self._cache:
                del self._cache[key]

                # 무효화 마커 설정 (레이스 컨디션 방지)
                if set_marker:
                    self._invalidation_markers[key] = time.time()
                    logger.debug(f"캐시 삭제 + 무효화 마커 설정: {key}")
                else:
                    logger.debug(f"캐시 삭제: {key}")
                return True
            return False

    def invalidate_by_pattern(self, pattern: str, set_marker: bool = True) -> int:
        """패턴에 맞는 캐시 무효화 (무효화 마커 설정 옵션)

        Args:
            pattern: 캐시 키 패턴
            set_marker: True이면 무효화 마커를 설정하여 레이스 컨디션 방지
        """
        with self._lock:
            keys_to_delete = [key for key in self._cache.keys() if pattern in key]
            for key in keys_to_delete:
                self.delete(key, set_marker=set_marker)
            logger.info(f"패턴 '{pattern}'으로 {len(keys_to_delete)}개 캐시 무효화 (마커: {set_marker})")
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

    def clear(self) -> int:
        """전체 캐시 삭제"""
        with self._lock:
            count = len(self._cache)
            self._cache.clear()
            logger.info(f"전체 캐시 삭제: {count}개 항목")
            return count

    def _cleanup_expired(self) -> None:
        """만료된 캐시 정리"""
        current_time = time.time()
        expired_keys = [
            key for key, item in self._cache.items()
            if current_time > item['expires']
        ]

        for key in expired_keys:
            del self._cache[key]

        if expired_keys:
            logger.debug(f"만료된 캐시 정리: {len(expired_keys)}개 항목")

    def _get_cache_item_info(self, key: str) -> Optional[Dict[str, Any]]:
        """특정 키의 캐시 정보 반환 (백그라운드 프리패치용)"""
        with self._lock:
            if key not in self._cache:
                return None

            item = self._cache[key]
            current_time = time.time()

            return {
                'ttl': self.TTL_MAP.get(item.get('strategy', CacheStrategy.CRITICAL_DATA), 300),
                'created': item['expires'] - self.TTL_MAP.get(item.get('strategy', CacheStrategy.CRITICAL_DATA), 300),
                'expires': item['expires'],
                'strategy': item.get('strategy', CacheStrategy.CRITICAL_DATA)
            }

    def get_cache_info(self) -> Dict[str, Any]:
        """캐시 상태 정보 (간소화됨)"""
        with self._lock:
            current_time = time.time()
            active_count = 0
            expired_count = 0

            for item in self._cache.values():
                if current_time > item['expires']:
                    expired_count += 1
                else:
                    active_count += 1

            return {
                'total_items': len(self._cache),
                'active_items': active_count,
                'expired_items': expired_count,
                'invalidation_markers': len(self._invalidation_markers)
            }

    def set_invalidation_marker(self, key: str, timestamp: Optional[float] = None) -> None:
        """무효화 마커 설정 (레이스 컨디션 방지용)

        Args:
            key: 캐시 키
            timestamp: 무효화 시각 (기본값: 현재 시각)
        """
        with self._lock:
            marker_time = timestamp or time.time()
            self._invalidation_markers[key] = marker_time
            logger.debug(f"무효화 마커 설정: {key} = {marker_time}")

    def clear_invalidation_marker(self, key: str) -> bool:
        """무효화 마커 제거

        Args:
            key: 캐시 키

        Returns:
            제거 성공 여부
        """
        with self._lock:
            if key in self._invalidation_markers:
                del self._invalidation_markers[key]
                logger.debug(f"무효화 마커 제거: {key}")
                return True
            return False

    def get_invalidation_marker(self, key: str) -> Optional[float]:
        """무효화 마커 조회

        Args:
            key: 캐시 키

        Returns:
            무효화 시각 (Unix timestamp), 없으면 None
        """
        with self._lock:
            return self._invalidation_markers.get(key)

# 전역 간단 캐시 인스턴스
_simple_cache = SimpleCache()

def get_smart_cache() -> SimpleCache:
    """캐시 인스턴스 반환 (호환성 유지)"""
    return _simple_cache

# 편의 함수들 (기존 인터페이스 유지)
def smart_get(key: str, strategy: CacheStrategy = CacheStrategy.CRITICAL_DATA) -> Optional[Any]:
    """캐시에서 값 가져오기"""
    return _simple_cache.get(key, strategy)

def smart_set(key: str, value: Any, strategy: CacheStrategy = CacheStrategy.CRITICAL_DATA,
              ttl: Optional[int] = None, fetched_at: Optional[float] = None) -> None:
    """캐시에 값 저장

    Args:
        key: 캐시 키
        value: 저장할 값
        strategy: 캐시 전략
        ttl: 커스텀 TTL (선택사항)
        fetched_at: 데이터 수집 시작 시각 (Unix timestamp) - 레이스 컨디션 방지용
    """
    _simple_cache.set(key, value, strategy, ttl, fetched_at)

def smart_delete(key: str) -> bool:
    """캐시에서 키 삭제"""
    return _simple_cache.delete(key)

def smart_invalidate(pattern: str) -> int:
    """패턴 기반 캐시 무효화"""
    return _simple_cache.invalidate_by_pattern(pattern)

def smart_clear_strategy(strategy: CacheStrategy) -> int:
    """전략별 캐시 삭제"""
    return _simple_cache.clear_by_strategy(strategy)

def smart_register_refresh_callback(key: str, callback: callable) -> None:
    """자동 새로고침 콜백 함수 등록 (호환성 유지, 실제 동작 안함)"""
    logger.warning(f"자동 새로고침 기능이 제거됨: {key}")

def smart_get_timestamp(key: str) -> Optional[float]:
    """캐시 항목의 생성 시간 반환 (Unix timestamp)

    Args:
        key: 캐시 키

    Returns:
        캐시된 데이터의 생성 시간 (Unix timestamp), 캐시 미스 시 None
    """
    with _simple_cache._lock:
        if key not in _simple_cache._cache:
            return None

        item = _simple_cache._cache[key]
        current_time = time.time()

        # TTL 체크
        if current_time > item['expires']:
            return None

        return item.get('created_at')

def cache_clear() -> int:
    """전체 캐시 삭제 (레거시 호환)"""
    return _simple_cache.clear()

def cache_stats() -> Dict[str, Any]:
    """캐시 통계 (레거시 호환)"""
    return _simple_cache.get_cache_info()

# 무효화 마커 관리 편의 함수들
def smart_set_invalidation_marker(key: str, timestamp: Optional[float] = None) -> None:
    """무효화 마커 설정 (레이스 컨디션 방지용)

    Args:
        key: 캐시 키
        timestamp: 무효화 시각 (기본값: 현재 시각)
    """
    _simple_cache.set_invalidation_marker(key, timestamp)

def smart_clear_invalidation_marker(key: str) -> bool:
    """무효화 마커 제거

    Args:
        key: 캐시 키

    Returns:
        제거 성공 여부
    """
    return _simple_cache.clear_invalidation_marker(key)

def smart_get_invalidation_marker(key: str) -> Optional[float]:
    """무효화 마커 조회

    Args:
        key: 캐시 키

    Returns:
        무효화 시각 (Unix timestamp), 없으면 None
    """
    return _simple_cache.get_invalidation_marker(key)