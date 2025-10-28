"""
Redis 기반 TTL 캐시 관리자
분산 환경에서 안전한 캐시 관리 제공

변경 이력:
- 2025-01: Dict → Redis 전환 (다중 프로세스 지원)
- 2025-01: pandas.DataFrame pickle 직렬화 지원 추가
- 인터페이스 유지 (호출부 수정 불필요)
"""

import time
import pickle
import threading
from typing import Any, Optional, Dict
import logging
from enum import Enum

from dashboard.utils.redis_client import get_redis_client
from dashboard.utils.exceptions import ServiceUnavailable

logger = logging.getLogger(__name__)

# pandas import (선택적 - 없어도 동작)
try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False
    pd = None

class CacheStrategy(Enum):
    """캐시 전략 유형 (실시간성과 API 효율성 균형)"""
    CRITICAL_DATA = "critical_data"      # 구글 시트 데이터 (60초 TTL - API 호출 최적화)
    STATIC_CONFIG = "static_config"      # 설정 데이터 (1시간 TTL)
    FOLDER_MAPPING = "folder_mapping"    # 폴더 ID 매핑 (1일 TTL)
    UI_STATE = "ui_state"               # UI 상태 (5분 TTL)
    TEMPORARY = "temporary"             # 임시 데이터 (1분 TTL)
    METADATA = "metadata"               # 메타데이터 (담당자, 사업자, 거래처) (10분 TTL)

class SimpleCache:
    """Redis 기반 TTL 캐시

    변경사항:
    - Dict → Redis로 전환
    - TTL은 Redis가 자동 관리
    - 무효화 마커도 Redis에 저장
    - 인터페이스는 그대로 유지
    """

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
        self.redis = get_redis_client()
        self._lock = threading.RLock()  # 로컬 동기화용 (Redis는 자체적으로 원자적)

    def get(self, key: str, strategy: CacheStrategy = CacheStrategy.CRITICAL_DATA) -> Optional[Any]:
        """캐시에서 값 가져오기

        Args:
            key: 캐시 키
            strategy: 캐시 전략 (사용 안 함, 하위 호환성)

        Returns:
            캐시된 값, 없으면 None

        Raises:
            ServiceUnavailable: Redis 장애 시

        Note:
            bytes 타입이면 pickle 역직렬화 시도 (DataFrame 등)
        """
        try:
            cache_key = f"cache:{key}"
            value = self.redis.get(cache_key)

            if value is None:
                logger.debug(f"캐시 미스: {key}")
                return None

            # bytes이면 pickle 역직렬화 시도
            if isinstance(value, bytes):
                try:
                    deserialized = pickle.loads(value)
                    logger.debug(f"캐시 히트 (pickle): {key}")
                    return deserialized
                except (pickle.PickleError, Exception) as e:
                    logger.warning(f"pickle 역직렬화 실패: {key}, error={e}")
                    # 실패 시 None 반환
                    return None

            logger.debug(f"캐시 히트: {key}")
            return value

        except ServiceUnavailable:
            logger.error(f"Redis 장애로 캐시 조회 실패: {key}")
            raise

    def set(self, key: str, value: Any, strategy: CacheStrategy = CacheStrategy.CRITICAL_DATA,
            custom_ttl: Optional[int] = None, fetched_at: Optional[float] = None) -> None:
        """캐시에 값 저장 (무효화 마커 체크 포함 - 레이스 컨디션 방지)

        Args:
            key: 캐시 키
            value: 저장할 값 (DataFrame은 자동으로 pickle 직렬화)
            strategy: 캐시 전략
            custom_ttl: 커스텀 TTL (선택사항)
            fetched_at: 데이터 수집 시작 시각 (Unix timestamp) - 없으면 현재 시각 사용

        Raises:
            ServiceUnavailable: Redis 장애 시

        Note:
            pandas.DataFrame은 pickle로 자동 직렬화됩니다.
        """
        try:
            ttl = custom_ttl or self.TTL_MAP[strategy]
            current_time = time.time()

            # 🔒 레이스 컨디션 방지: 무효화 마커보다 오래된 데이터는 캐시에 쓰지 않음
            # fetched_at이 제공되면 데이터 수집 시작 시각으로 비교 (정확한 보호)
            # 없으면 현재 시각으로 폴백 (하위 호환성)
            data_timestamp = fetched_at if fetched_at is not None else current_time

            # 무효화 마커 확인 (Redis에서 조회)
            marker_key = f"invalidation:{key}"
            invalidation_time_str = self.redis.get(marker_key)

            if invalidation_time_str is not None:
                try:
                    invalidation_time = float(invalidation_time_str)
                    if data_timestamp < invalidation_time:
                        logger.warning(
                            f"캐시 쓰기 거부: {key} - 데이터 수집이 무효화 마커보다 이전 "
                            f"(수집 시작: {data_timestamp:.2f}, 무효화: {invalidation_time:.2f})"
                        )
                        return
                except (ValueError, TypeError):
                    logger.warning(f"무효화 마커 파싱 실패: {key}")

            # DataFrame은 pickle로 직렬화
            if HAS_PANDAS and isinstance(value, pd.DataFrame):
                logger.debug(f"DataFrame 감지 - pickle 직렬화: {key} (shape: {value.shape})")
                try:
                    serialized_value = pickle.dumps(value, protocol=pickle.HIGHEST_PROTOCOL)
                    cache_key = f"cache:{key}"
                    self.redis.set(cache_key, serialized_value, ex=ttl)
                    logger.debug(f"캐시 저장 (pickle): {key}, 전략: {strategy.value}, TTL: {ttl}초")
                    return
                except (pickle.PickleError, Exception) as e:
                    logger.error(f"DataFrame pickle 직렬화 실패: {key}, error={e}")
                    raise ServiceUnavailable(f"DataFrame serialization failed: {e}")

            # 일반 값은 Redis client가 자동 처리 (JSON 직렬화)
            cache_key = f"cache:{key}"
            self.redis.set(cache_key, value, ex=ttl)

            logger.debug(f"캐시 저장: {key}, 전략: {strategy.value}, TTL: {ttl}초, "
                        f"생성시간: {current_time}, 수집시각: {data_timestamp}")

        except ServiceUnavailable:
            logger.error(f"Redis 장애로 캐시 저장 실패: {key}")
            raise

    def delete(self, key: str, set_marker: bool = True) -> bool:
        """캐시에서 키 삭제 (무효화 마커 설정 옵션)

        Args:
            key: 캐시 키
            set_marker: True이면 무효화 마커를 설정하여 레이스 컨디션 방지

        Returns:
            삭제 성공 여부

        Raises:
            ServiceUnavailable: Redis 장애 시
        """
        try:
            cache_key = f"cache:{key}"
            deleted_count = self.redis.delete(cache_key)

            # 무효화 마커 설정 (레이스 컨디션 방지)
            if set_marker:
                marker_key = f"invalidation:{key}"
                self.redis.set(marker_key, str(time.time()), ex=10)  # 10초 TTL
                logger.debug(f"캐시 삭제 + 무효화 마커 설정: {key}")
            else:
                logger.debug(f"캐시 삭제: {key}")

            return deleted_count > 0

        except ServiceUnavailable:
            logger.error(f"Redis 장애로 캐시 삭제 실패: {key}")
            raise

    def invalidate_by_pattern(self, pattern: str, set_marker: bool = True) -> int:
        """패턴에 맞는 캐시 무효화 (무효화 마커 설정 옵션)

        Args:
            pattern: 캐시 키 패턴
            set_marker: True이면 무효화 마커를 설정하여 레이스 컨디션 방지

        Returns:
            삭제된 키 개수

        Raises:
            ServiceUnavailable: Redis 장애 시

        Warning:
            KEYS 명령은 O(N)이므로 대량 키가 있을 때 주의
        """
        try:
            # Redis KEYS 명령으로 패턴 검색
            search_pattern = f"cache:*{pattern}*"
            matching_keys = self.redis.keys(search_pattern)

            # cache: 접두사 제거하여 원래 키 복원
            original_keys = [key.replace("cache:", "", 1) for key in matching_keys]

            # 각 키 삭제
            for key in original_keys:
                self.delete(key, set_marker=set_marker)

            logger.info(f"패턴 '{pattern}'으로 {len(original_keys)}개 캐시 무효화 (마커: {set_marker})")
            return len(original_keys)

        except ServiceUnavailable:
            logger.error(f"Redis 장애로 패턴 무효화 실패: {pattern}")
            raise

    def clear_by_strategy(self, strategy: CacheStrategy) -> int:
        """특정 전략의 캐시만 삭제

        Note:
            Redis 전환 후에는 전략 정보를 저장하지 않으므로
            전체 캐시를 삭제합니다. (하위 호환성 유지)

        Returns:
            삭제된 키 개수

        Raises:
            ServiceUnavailable: Redis 장애 시
        """
        logger.warning(f"clear_by_strategy는 Redis 전환 후 전체 캐시를 삭제합니다: {strategy.value}")
        return self.clear()

    def clear(self) -> int:
        """전체 캐시 삭제

        Returns:
            삭제된 키 개수

        Raises:
            ServiceUnavailable: Redis 장애 시

        Warning:
            프로덕션에서는 주의해서 사용
        """
        try:
            # cache:* 패턴으로 모든 캐시 키 조회
            cache_keys = self.redis.keys("cache:*")

            if cache_keys:
                self.redis.delete(*cache_keys)
                logger.info(f"전체 캐시 삭제: {len(cache_keys)}개 항목")
                return len(cache_keys)
            else:
                logger.info("삭제할 캐시 항목 없음")
                return 0

        except ServiceUnavailable:
            logger.error("Redis 장애로 전체 캐시 삭제 실패")
            raise

    def _get_cache_item_info(self, key: str) -> Optional[Dict[str, Any]]:
        """특정 키의 캐시 정보 반환 (백그라운드 프리패치용)

        Note:
            Redis 전환 후 TTL만 반환 가능 (strategy, created, expires 정보 없음)
        """
        try:
            cache_key = f"cache:{key}"
            ttl = self.redis.ttl(cache_key)

            # -2: 키 없음, -1: 만료 없음
            if ttl < 0:
                return None

            current_time = time.time()

            return {
                'ttl': ttl,
                'created': current_time - ttl,  # 대략적인 생성 시간
                'expires': current_time + ttl,
                'strategy': CacheStrategy.CRITICAL_DATA  # 기본값 (실제로는 알 수 없음)
            }

        except ServiceUnavailable:
            logger.error(f"Redis 장애로 캐시 정보 조회 실패: {key}")
            return None

    def get_cache_info(self) -> Dict[str, Any]:
        """캐시 상태 정보

        Note:
            Redis 전환 후 expired_items 정보 없음 (Redis가 자동 삭제)
        """
        try:
            cache_keys = self.redis.keys("cache:*")
            marker_keys = self.redis.keys("invalidation:*")

            return {
                'total_items': len(cache_keys),
                'active_items': len(cache_keys),  # Redis는 만료된 키를 자동 삭제
                'expired_items': 0,  # Redis가 자동 관리
                'invalidation_markers': len(marker_keys)
            }

        except ServiceUnavailable:
            logger.error("Redis 장애로 캐시 정보 조회 실패")
            return {
                'total_items': 0,
                'active_items': 0,
                'expired_items': 0,
                'invalidation_markers': 0
            }

    def set_invalidation_marker(self, key: str, timestamp: Optional[float] = None) -> None:
        """무효화 마커 설정 (레이스 컨디션 방지용)

        Args:
            key: 캐시 키
            timestamp: 무효화 시각 (기본값: 현재 시각)

        Raises:
            ServiceUnavailable: Redis 장애 시
        """
        try:
            marker_time = timestamp or time.time()
            marker_key = f"invalidation:{key}"
            self.redis.set(marker_key, str(marker_time), ex=10)  # 10초 TTL
            logger.debug(f"무효화 마커 설정: {key} = {marker_time}")

        except ServiceUnavailable:
            logger.error(f"Redis 장애로 무효화 마커 설정 실패: {key}")
            raise

    def clear_invalidation_marker(self, key: str) -> bool:
        """무효화 마커 제거

        Args:
            key: 캐시 키

        Returns:
            제거 성공 여부

        Raises:
            ServiceUnavailable: Redis 장애 시
        """
        try:
            marker_key = f"invalidation:{key}"
            deleted_count = self.redis.delete(marker_key)

            if deleted_count > 0:
                logger.debug(f"무효화 마커 제거: {key}")
                return True
            return False

        except ServiceUnavailable:
            logger.error(f"Redis 장애로 무효화 마커 제거 실패: {key}")
            raise

    def get_invalidation_marker(self, key: str) -> Optional[float]:
        """무효화 마커 조회

        Args:
            key: 캐시 키

        Returns:
            무효화 시각 (Unix timestamp), 없으면 None

        Raises:
            ServiceUnavailable: Redis 장애 시
        """
        try:
            marker_key = f"invalidation:{key}"
            value = self.redis.get(marker_key)

            if value is None:
                return None

            try:
                return float(value)
            except (ValueError, TypeError):
                logger.warning(f"무효화 마커 파싱 실패: {key}")
                return None

        except ServiceUnavailable:
            logger.error(f"Redis 장애로 무효화 마커 조회 실패: {key}")
            raise

# 전역 간단 캐시 인스턴스 (Lazy Loading)
_simple_cache: Optional[SimpleCache] = None

def get_smart_cache() -> SimpleCache:
    """캐시 인스턴스 반환 (Lazy Loading - 첫 호출 시 초기화)"""
    global _simple_cache
    if _simple_cache is None:
        _simple_cache = SimpleCache()
    return _simple_cache

# 편의 함수들 (기존 인터페이스 유지)
def smart_get(key: str, strategy: CacheStrategy = CacheStrategy.CRITICAL_DATA) -> Optional[Any]:
    """캐시에서 값 가져오기"""
    return get_smart_cache().get(key, strategy)

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
    get_smart_cache().set(key, value, strategy, ttl, fetched_at)

def smart_delete(key: str) -> bool:
    """캐시에서 키 삭제"""
    return get_smart_cache().delete(key)

def smart_invalidate(pattern: str) -> int:
    """패턴 기반 캐시 무효화"""
    return get_smart_cache().invalidate_by_pattern(pattern)

def smart_clear_strategy(strategy: CacheStrategy) -> int:
    """전략별 캐시 삭제"""
    return get_smart_cache().clear_by_strategy(strategy)

def smart_register_refresh_callback(key: str, callback: callable) -> None:
    """자동 새로고침 콜백 함수 등록 (호환성 유지, 실제 동작 안함)"""
    logger.warning(f"자동 새로고침 기능이 제거됨: {key}")

def smart_get_timestamp(key: str) -> Optional[float]:
    """캐시 항목의 생성 시간 반환 (Unix timestamp)

    Args:
        key: 캐시 키

    Returns:
        캐시된 데이터의 생성 시간 (Unix timestamp), 캐시 미스 시 None

    Note:
        Redis 전환 후 정확한 생성 시간은 저장하지 않음
        TTL 기반으로 대략적인 생성 시간 추정
    """
    try:
        cache = get_smart_cache()
        cache_key = f"cache:{key}"
        ttl = cache.redis.ttl(cache_key)

        # -2: 키 없음, -1: 만료 없음
        if ttl < 0:
            return None

        # 대략적인 생성 시간 = 현재 시간 - 남은 TTL
        # (실제 생성 시간보다 약간 부정확할 수 있음)
        current_time = time.time()
        return current_time - ttl

    except ServiceUnavailable:
        logger.error(f"Redis 장애로 타임스탬프 조회 실패: {key}")
        return None

def cache_clear() -> int:
    """전체 캐시 삭제 (레거시 호환)"""
    return get_smart_cache().clear()

def cache_stats() -> Dict[str, Any]:
    """캐시 통계 (레거시 호환)"""
    return get_smart_cache().get_cache_info()

# 무효화 마커 관리 편의 함수들
def smart_set_invalidation_marker(key: str, timestamp: Optional[float] = None) -> None:
    """무효화 마커 설정 (레이스 컨디션 방지용)

    Args:
        key: 캐시 키
        timestamp: 무효화 시각 (기본값: 현재 시각)
    """
    get_smart_cache().set_invalidation_marker(key, timestamp)

def smart_clear_invalidation_marker(key: str) -> bool:
    """무효화 마커 제거

    Args:
        key: 캐시 키

    Returns:
        제거 성공 여부
    """
    return get_smart_cache().clear_invalidation_marker(key)

def smart_get_invalidation_marker(key: str) -> Optional[float]:
    """무효화 마커 조회

    Args:
        key: 캐시 키

    Returns:
        무효화 시각 (Unix timestamp), 없으면 None
    """
    return get_smart_cache().get_invalidation_marker(key)