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
from dashboard.utils.cache_invalidation import get_cache_invalidation_service

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

    개선사항 (Priority 4):
    - 캐시 히트/미스 메트릭 추가 (Redis 기반)
    - 배치 연산 지원 (pipeline)
    - 캐시 워밍 기능 추가
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

    # 메트릭 키 (Redis에 저장)
    METRIC_HIT_KEY = "cache:metrics:hits"
    METRIC_MISS_KEY = "cache:metrics:misses"
    METRIC_RESET_KEY = "cache:metrics:reset_at"

    def __init__(self):
        self.redis = get_redis_client()
        self._lock = threading.RLock()  # 로컬 동기화용 (Redis는 자체적으로 원자적)
        self._use_fallback = False  # Fallback 모드 플래그
        self._fallback_cache = None  # Lazy loading
        self._initialize_metrics()

        # Redis 연결 확인
        try:
            if not self.redis.ping():
                logger.warning("Redis 연결 불가 - Fallback 메모리 캐시로 전환")
                self._enable_fallback()
        except Exception as e:
            logger.error(f"Redis 연결 오류 - Fallback 메모리 캐시로 전환: {e}")
            self._enable_fallback()

        # 캐시 무효화 서비스 등록 (Pub/Sub)
        try:
            invalidation_service = get_cache_invalidation_service()
            invalidation_service.register_handler(self._handle_invalidation_message)
            logger.info("캐시 무효화 Pub/Sub 핸들러 등록 완료")
        except Exception as e:
            logger.error(f"캐시 무효화 서비스 등록 실패: {e}")

    def _enable_fallback(self):
        """Fallback 메모리 캐시 활성화"""
        from dashboard.utils.fallback_cache import get_fallback_cache
        self._use_fallback = True
        self._fallback_cache = get_fallback_cache()
        logger.warning("⚠️ Fallback 메모리 캐시 모드 활성화 - 성능 저하 및 멀티 워커 비공유")

    def _handle_invalidation_message(self, message: dict):
        """
        Pub/Sub 무효화 메시지 처리 (다른 워커에서 발행)

        Args:
            message: 무효화 메시지 (dict)
        """
        action = message.get("action")

        try:
            if action == "invalidate":
                # 특정 키 무효화
                keys = message.get("keys", [])
                for key in keys:
                    self.delete(key, set_marker=False)  # 마커 중복 설정 방지
                logger.debug(f"Pub/Sub 무효화 처리: {len(keys)}개 키")

            elif action == "invalidate_pattern":
                # 패턴 기반 무효화
                pattern = message.get("pattern", "")
                self.invalidate_by_pattern(pattern, set_marker=False, broadcast=False)
                logger.debug(f"Pub/Sub 패턴 무효화 처리: {pattern}")

            elif action == "clear_all":
                # 전체 캐시 무효화
                self.clear(broadcast=False)
                logger.warning("Pub/Sub 전체 캐시 무효화 처리")

        except Exception as e:
            logger.error(f"Pub/Sub 무효화 메시지 처리 오류: {e}")

    def _initialize_metrics(self) -> None:
        """메트릭 초기화 (최초 1회만)"""
        try:
            if not self.redis.redis.exists(self.METRIC_RESET_KEY):
                self.redis.redis.set(self.METRIC_HIT_KEY, 0)
                self.redis.redis.set(self.METRIC_MISS_KEY, 0)
                self.redis.redis.set(self.METRIC_RESET_KEY, str(time.time()))
                logger.info("캐시 메트릭 초기화 완료")
        except Exception as e:
            logger.warning(f"메트릭 초기화 실패 (무시): {e}")

    def _record_hit(self) -> None:
        """캐시 히트 기록"""
        try:
            self.redis.redis.incr(self.METRIC_HIT_KEY)
        except Exception:
            pass  # 메트릭 실패는 무시

    def _record_miss(self) -> None:
        """캐시 미스 기록"""
        try:
            self.redis.redis.incr(self.METRIC_MISS_KEY)
        except Exception:
            pass  # 메트릭 실패는 무시

    def get(self, key: str, strategy: CacheStrategy = CacheStrategy.CRITICAL_DATA) -> Optional[Any]:
        """캐시에서 값 가져오기 (Fallback 지원)

        Args:
            key: 캐시 키
            strategy: 캐시 전략 (사용 안 함, 하위 호환성)

        Returns:
            캐시된 값, 없으면 None

        Raises:
            ServiceUnavailable: Redis 장애 시 (Fallback 없을 때)

        Note:
            bytes 타입이면 pickle 역직렬화 시도 (DataFrame 등)
            Redis 실패 시 자동으로 Fallback 메모리 캐시 사용
        """
        # Fallback 모드일 때
        if self._use_fallback and self._fallback_cache:
            cache_key = f"cache:{key}"
            value = self._fallback_cache.get(cache_key)
            if value is None:
                self._record_miss()
            else:
                self._record_hit()
            return value

        try:
            cache_key = f"cache:{key}"
            value = self.redis.get(cache_key)

            if value is None:
                logger.debug(f"캐시 미스: {key}")
                self._record_miss()
                return None

            # bytes이면 pickle 역직렬화 시도
            if isinstance(value, bytes):
                try:
                    deserialized = pickle.loads(value)
                    logger.debug(f"캐시 히트 (pickle): {key}")
                    self._record_hit()
                    return deserialized
                except (pickle.PickleError, Exception) as e:
                    logger.warning(f"pickle 역직렬화 실패: {key}, error={e}")
                    # 실패 시 None 반환
                    self._record_miss()
                    return None

            logger.debug(f"캐시 히트: {key}")
            self._record_hit()
            return value

        except ServiceUnavailable:
            # Redis 장애 시 Fallback으로 전환
            logger.error(f"Redis 장애 감지 - Fallback 모드로 전환: {key}")
            self._enable_fallback()
            # Fallback으로 재시도
            return self.get(key, strategy)

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

    def delete(self, key: str, set_marker: bool = True, broadcast: bool = True) -> bool:
        """캐시에서 키 삭제 (무효화 마커 설정 옵션 + Pub/Sub 브로드캐스트)

        Args:
            key: 캐시 키
            set_marker: True이면 무효화 마커를 설정하여 레이스 컨디션 방지
            broadcast: True이면 Pub/Sub으로 다른 워커에게 무효화 메시지 발행

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

            # Pub/Sub 브로드캐스트 (다른 워커에게 무효화 알림)
            if broadcast:
                try:
                    invalidation_service = get_cache_invalidation_service()
                    invalidation_service.publish_invalidation([key], reason="cache_delete")
                except Exception as e:
                    logger.debug(f"Pub/Sub 브로드캐스트 실패 (무시): {e}")

            return deleted_count > 0

        except ServiceUnavailable:
            logger.error(f"Redis 장애로 캐시 삭제 실패: {key}")
            raise

    def invalidate_by_pattern(self, pattern: str, set_marker: bool = True, broadcast: bool = True) -> int:
        """패턴에 맞는 캐시 무효화 (무효화 마커 설정 옵션 + Pub/Sub 브로드캐스트)

        Args:
            pattern: 캐시 키 패턴
            set_marker: True이면 무효화 마커를 설정하여 레이스 컨디션 방지
            broadcast: True이면 Pub/Sub으로 다른 워커에게 무효화 메시지 발행

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

            # Pub/Sub 브로드캐스트 (다른 워커에게 패턴 무효화 알림)
            if broadcast and pattern:
                try:
                    invalidation_service = get_cache_invalidation_service()
                    invalidation_service.publish_pattern_invalidation(pattern, reason="pattern_invalidation")
                except Exception as e:
                    logger.debug(f"Pub/Sub 브로드캐스트 실패 (무시): {e}")

            # 각 키 삭제 (broadcast=False로 중복 방지)
            for key in original_keys:
                self.delete(key, set_marker=set_marker, broadcast=False)

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

    def get_metrics(self) -> Dict[str, Any]:
        """캐시 메트릭 조회

        Returns:
            메트릭 딕셔너리 (hits, misses, hit_rate, reset_at)
        """
        try:
            hits = int(self.redis.redis.get(self.METRIC_HIT_KEY) or 0)
            misses = int(self.redis.redis.get(self.METRIC_MISS_KEY) or 0)
            reset_at_str = self.redis.redis.get(self.METRIC_RESET_KEY)

            total = hits + misses
            hit_rate = (hits / total * 100) if total > 0 else 0.0

            return {
                'hits': hits,
                'misses': misses,
                'total': total,
                'hit_rate': round(hit_rate, 2),
                'reset_at': reset_at_str,
                'timestamp': time.time()
            }

        except Exception as e:
            logger.error(f"메트릭 조회 실패: {e}")
            return {
                'hits': 0,
                'misses': 0,
                'total': 0,
                'hit_rate': 0.0,
                'reset_at': None,
                'timestamp': time.time()
            }

    def reset_metrics(self) -> None:
        """메트릭 초기화"""
        try:
            self.redis.redis.set(self.METRIC_HIT_KEY, 0)
            self.redis.redis.set(self.METRIC_MISS_KEY, 0)
            self.redis.redis.set(self.METRIC_RESET_KEY, str(time.time()))
            logger.info("캐시 메트릭 초기화 완료")
        except Exception as e:
            logger.error(f"메트릭 초기화 실패: {e}")

    def mget(self, keys: list, strategy: CacheStrategy = CacheStrategy.CRITICAL_DATA) -> Dict[str, Any]:
        """여러 키를 한 번에 조회 (배치 연산)

        Args:
            keys: 조회할 키 리스트
            strategy: 캐시 전략 (사용 안 함, 하위 호환성)

        Returns:
            {key: value} 딕셔너리 (존재하지 않는 키는 제외)
        """
        if not keys:
            return {}

        try:
            cache_keys = [f"cache:{key}" for key in keys]
            values = self.redis.redis.mget(cache_keys)

            result = {}
            for key, value in zip(keys, values):
                if value is not None:
                    # pickle 역직렬화 시도
                    if isinstance(value, bytes):
                        try:
                            result[key] = pickle.loads(value)
                            self._record_hit()
                        except (pickle.PickleError, Exception):
                            self._record_miss()
                    else:
                        result[key] = value
                        self._record_hit()
                else:
                    self._record_miss()

            logger.debug(f"배치 조회: {len(keys)}개 요청, {len(result)}개 히트")
            return result

        except ServiceUnavailable:
            logger.error(f"Redis 장애로 배치 조회 실패")
            raise

    def mset(self, items: Dict[str, Any], strategy: CacheStrategy = CacheStrategy.CRITICAL_DATA,
             custom_ttl: Optional[int] = None) -> None:
        """여러 키를 한 번에 저장 (배치 연산)

        Args:
            items: {key: value} 딕셔너리
            strategy: 캐시 전략
            custom_ttl: 커스텀 TTL (선택사항)

        Note:
            모든 키에 동일한 TTL이 적용됩니다.
            Redis Pipeline을 사용하여 성능 향상
        """
        if not items:
            return

        try:
            ttl = custom_ttl or self.TTL_MAP[strategy]

            # Redis pipeline 사용 (원자적 배치 연산)
            pipe = self.redis.redis.pipeline()

            for key, value in items.items():
                cache_key = f"cache:{key}"

                # DataFrame은 pickle로 직렬화
                if HAS_PANDAS and isinstance(value, pd.DataFrame):
                    try:
                        serialized_value = pickle.dumps(value, protocol=pickle.HIGHEST_PROTOCOL)
                        pipe.set(cache_key, serialized_value, ex=ttl)
                    except (pickle.PickleError, Exception) as e:
                        logger.error(f"DataFrame pickle 직렬화 실패: {key}, error={e}")
                        continue
                else:
                    pipe.set(cache_key, value, ex=ttl)

            pipe.execute()
            logger.debug(f"배치 저장: {len(items)}개 항목, 전략: {strategy.value}, TTL: {ttl}초")

        except ServiceUnavailable:
            logger.error("Redis 장애로 배치 저장 실패")
            raise

    def warm_cache(self, warm_items: Dict[str, tuple]) -> int:
        """캐시 워밍 (애플리케이션 시작 시 미리 데이터 로드)

        Args:
            warm_items: {key: (value, strategy)} 딕셔너리

        Returns:
            성공적으로 캐시된 항목 수

        Example:
            cache.warm_cache({
                'user_list': (get_users(), CacheStrategy.METADATA),
                'config': (load_config(), CacheStrategy.STATIC_CONFIG)
            })
        """
        success_count = 0

        try:
            pipe = self.redis.redis.pipeline()

            for key, (value, strategy) in warm_items.items():
                ttl = self.TTL_MAP.get(strategy, 300)
                cache_key = f"cache:{key}"

                # DataFrame은 pickle로 직렬화
                if HAS_PANDAS and isinstance(value, pd.DataFrame):
                    try:
                        serialized_value = pickle.dumps(value, protocol=pickle.HIGHEST_PROTOCOL)
                        pipe.set(cache_key, serialized_value, ex=ttl)
                        success_count += 1
                    except (pickle.PickleError, Exception) as e:
                        logger.error(f"캐시 워밍 실패 (pickle): {key}, error={e}")
                        continue
                else:
                    pipe.set(cache_key, value, ex=ttl)
                    success_count += 1

            pipe.execute()
            logger.info(f"캐시 워밍 완료: {success_count}/{len(warm_items)}개 항목")
            return success_count

        except ServiceUnavailable:
            logger.error("Redis 장애로 캐시 워밍 실패")
            return success_count

    def get_cache_info(self) -> Dict[str, Any]:
        """캐시 상태 정보

        Note:
            Redis 전환 후 expired_items 정보 없음 (Redis가 자동 삭제)
            메트릭 정보 포함
        """
        try:
            cache_keys = self.redis.keys("cache:*")
            # 메트릭 키 제외
            cache_keys = [k for k in cache_keys if not k.startswith("cache:metrics:")]

            marker_keys = self.redis.keys("invalidation:*")
            metrics = self.get_metrics()

            return {
                'total_items': len(cache_keys),
                'active_items': len(cache_keys),  # Redis는 만료된 키를 자동 삭제
                'expired_items': 0,  # Redis가 자동 관리
                'invalidation_markers': len(marker_keys),
                'metrics': metrics
            }

        except ServiceUnavailable:
            logger.error("Redis 장애로 캐시 정보 조회 실패")
            return {
                'total_items': 0,
                'active_items': 0,
                'expired_items': 0,
                'invalidation_markers': 0,
                'metrics': {
                    'hits': 0,
                    'misses': 0,
                    'total': 0,
                    'hit_rate': 0.0
                }
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