"""
Redis 클라이언트 헬퍼
- 캐시 저장/조회
- 분산 락 관리
- Fail Fast 에러 처리
"""

import os
import sys
import json
import logging
import time
from typing import Any, Optional

import redis
from redis.exceptions import RedisError, ConnectionError, TimeoutError
from redis.sentinel import Sentinel
from tenacity import (
    retry,
    stop_after_attempt,
    wait_exponential,
    retry_if_exception_type,
    before_sleep_log
)

from dashboard.utils.exceptions import ServiceUnavailable

logger = logging.getLogger(__name__)


class RedisClient:
    """
    Redis 연결 및 작업 헬퍼

    특징:
    - Fail Fast: 부팅 시 Redis 연결 실패 시 서비스 시작 불가
    - 운영 중 에러: ServiceUnavailable 예외 발생 → HTTP 503
    - JSON 자동 직렬화/역직렬화
    """

    _instance = None  # Singleton 패턴

    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance

    def __init__(self):
        # Singleton 패턴: 이미 초기화되었으면 재초기화하지 않음
        if hasattr(self, '_initialized') and self._initialized:
            return

        try:
            # 환경 변수에서 Redis 설정 읽기
            use_sentinel = os.getenv('REDIS_USE_SENTINEL', 'false').lower() == 'true'
            redis_password = os.getenv('REDIS_PASSWORD', None)
            socket_timeout = int(os.getenv('REDIS_SOCKET_TIMEOUT', 5))
            socket_connect_timeout = int(os.getenv('REDIS_SOCKET_CONNECT_TIMEOUT', 5))
            redis_db = int(os.getenv('REDIS_DB', 0))

            # 비밀번호가 빈 문자열이면 None으로 처리
            if redis_password == '':
                redis_password = None

            # Sentinel 모드 vs 단일 인스턴스 모드
            if use_sentinel:
                # Sentinel 모드 (고가용성)
                sentinel_hosts = os.getenv('REDIS_SENTINEL_HOSTS', 'localhost:26379,localhost:26380,localhost:26381')
                master_name = os.getenv('REDIS_MASTER_NAME', 'mymaster')

                # Sentinel 호스트 파싱 (예: "host1:port1,host2:port2,host3:port3")
                sentinel_list = []
                for host_port in sentinel_hosts.split(','):
                    host, port = host_port.strip().split(':')
                    sentinel_list.append((host, int(port)))

                logger.info(f"Redis Sentinel 모드로 연결 중: {sentinel_list}, master={master_name}")

                # Sentinel 연결 생성
                sentinel = Sentinel(
                    sentinel_list,
                    socket_timeout=socket_timeout,
                    socket_connect_timeout=socket_connect_timeout,
                    password=redis_password
                )

                # Master 연결 가져오기 (자동 Failover 지원)
                self.redis = sentinel.master_for(
                    master_name,
                    socket_timeout=socket_timeout,
                    socket_connect_timeout=socket_connect_timeout,
                    password=redis_password,
                    db=redis_db,
                    decode_responses=True
                )

                self.redis_binary = sentinel.master_for(
                    master_name,
                    socket_timeout=socket_timeout,
                    socket_connect_timeout=socket_connect_timeout,
                    password=redis_password,
                    db=redis_db,
                    decode_responses=False
                )

                # 연결 확인
                self.redis.ping()
                self.redis_binary.ping()
                logger.info(f"Redis Sentinel 연결 성공: master={master_name} (DB: {redis_db})")
                logger.info("  - 자동 Failover 활성화 (Master 장애 시 자동 전환)")
                logger.info("  - 문자열용 연결 (decode_responses=True): 캐시, 일반 데이터")
                logger.info("  - 바이너리용 연결 (decode_responses=False): Flask-Session, RQ")

            else:
                # 단일 인스턴스 모드 (기본)
                redis_host = os.getenv('REDIS_HOST', 'localhost')
                redis_port = int(os.getenv('REDIS_PORT', 6379))

                # 공통 연결 설정
                common_config = {
                    'host': redis_host,
                    'port': redis_port,
                    'db': redis_db,
                    'password': redis_password,
                    'socket_timeout': socket_timeout,
                    'socket_connect_timeout': socket_connect_timeout,
                    'socket_keepalive': True,
                    'health_check_interval': 30
                }

                # Redis 연결 #1: 문자열용 (캐시, 일반 데이터)
                self.redis = redis.Redis(
                    **common_config,
                    decode_responses=True  # 문자열로 자동 디코딩
                )

                # Redis 연결 #2: 바이너리용 (Flask-Session, RQ 등)
                self.redis_binary = redis.Redis(
                    **common_config,
                    decode_responses=False  # bytes 그대로 반환 (pickle 직렬화 지원)
                )

                # Graceful Degradation: 재시도 후 연결 확인
                connection_success = False
                max_retries = 3

                for attempt in range(max_retries):
                    try:
                        self.redis.ping()
                        self.redis_binary.ping()
                        logger.info(f"Redis 연결 성공: {redis_host}:{redis_port} (DB: {redis_db})")
                        logger.info("  - 문자열용 연결 (decode_responses=True): 캐시, 일반 데이터")
                        logger.info("  - 바이너리용 연결 (decode_responses=False): Flask-Session, RQ")
                        connection_success = True
                        break
                    except (ConnectionError, TimeoutError, RedisError) as retry_error:
                        wait_time = 2 ** attempt  # 1초, 2초, 4초
                        logger.warning(f"Redis 연결 실패 (시도 {attempt + 1}/{max_retries}): {retry_error}")
                        if attempt < max_retries - 1:
                            logger.info(f"{wait_time}초 후 재시도...")
                            import time
                            time.sleep(wait_time)
                        else:
                            logger.error(f"Redis 연결 최종 실패 ({max_retries}회 시도)")

                if not connection_success:
                    # Fallback 모드로 전환 (서버는 계속 실행)
                    logger.warning("⚠️ Redis 연결 불가 - Fallback 메모리 캐시로 전환")
                    logger.warning("⚠️ 다중 워커 환경에서 캐시가 공유되지 않습니다")
                    from dashboard.utils.fallback_cache import get_fallback_cache
                    self._use_fallback = True
                    self._fallback_cache = get_fallback_cache()

            self._initialized = True

        except Exception as e:
            # 예상치 못한 에러 발생 시에도 Fallback으로 전환
            logger.error(f"Redis 초기화 중 예상치 못한 오류: {e}")
            logger.warning("⚠️ Fallback 메모리 캐시로 전환")
            from dashboard.utils.fallback_cache import get_fallback_cache
            self._use_fallback = True
            self._fallback_cache = get_fallback_cache()
            self._initialized = True

    @retry(
        retry=retry_if_exception_type((ConnectionError, TimeoutError)),
        wait=wait_exponential(multiplier=1, min=1, max=10),
        stop=stop_after_attempt(3),
        before_sleep=before_sleep_log(logger, logging.WARNING)
    )
    def ping(self) -> bool:
        """
        Redis 연결 확인 (재연결 로직 포함)

        Returns:
            bool: 연결 성공 시 True, 실패 시 False

        Note:
            - 최대 3회 재시도 (1초, 2초, 4초 간격)
            - ConnectionError, TimeoutError 시 자동 재시도
        """
        # Fallback 모드일 때는 항상 False 반환 (Redis 연결 없음)
        if getattr(self, '_use_fallback', False):
            return False

        try:
            return self.redis.ping()
        except RedisError as e:
            logger.error(f"Redis ping 실패: {e}")
            return False

    @retry(
        retry=retry_if_exception_type((ConnectionError, TimeoutError)),
        wait=wait_exponential(multiplier=1, min=1, max=10),
        stop=stop_after_attempt(3),
        before_sleep=before_sleep_log(logger, logging.WARNING)
    )
    def get(self, key: str) -> Optional[Any]:
        """
        값 조회 (JSON 자동 역직렬화 또는 bytes 복원)

        Args:
            key: Redis 키

        Returns:
            저장된 값 (None이면 키가 없음)

        Raises:
            ServiceUnavailable: Redis 에러 시

        Note:
            - __BYTES__: 접두사가 있으면 base64 디코딩하여 bytes 반환
            - JSON: 역직렬화 시도
            - 기타: 문자열 그대로 반환
            - 재연결: 최대 3회 재시도 (exponential backoff)
            - Fallback 모드: 메모리 캐시 사용
        """
        # Fallback 모드일 때 메모리 캐시 사용
        if getattr(self, '_use_fallback', False) and self._fallback_cache:
            value = self._fallback_cache.get(key)
            if value is None:
                return None

            # Fallback 캐시는 JSON 직렬화 없이 저장하므로 그대로 반환
            return value

        try:
            value = self.redis.get(key)
            if value is None:
                return None

            # bytes로 저장된 값 복원 (base64 디코딩)
            if isinstance(value, str) and value.startswith("__BYTES__:"):
                import base64
                encoded = value[len("__BYTES__:"):]
                return base64.b64decode(encoded.encode('ascii'))

            # JSON 역직렬화 시도
            try:
                return json.loads(value)
            except (json.JSONDecodeError, TypeError):
                # JSON이 아니면 문자열 그대로 반환
                return value

        except (ConnectionError, TimeoutError) as e:
            logger.error(f"Redis 연결 오류 (재시도 중): key={key}, error={e}")
            raise  # tenacity가 재시도 처리
        except RedisError as e:
            logger.error(f"Redis get 실패: key={key}, error={e}")
            raise ServiceUnavailable(f"Cache unavailable: {e}")

    @retry(
        retry=retry_if_exception_type((ConnectionError, TimeoutError)),
        wait=wait_exponential(multiplier=1, min=1, max=10),
        stop=stop_after_attempt(3),
        before_sleep=before_sleep_log(logger, logging.WARNING)
    )
    def set(self, key: str, value: Any, ex: int = None, nx: bool = False) -> bool:
        """
        값 저장 (JSON 자동 직렬화 또는 bytes 그대로 저장)

        Args:
            key: Redis 키
            value: 저장할 값 (JSON 직렬화 가능한 객체 또는 bytes)
            ex: TTL (초 단위, None이면 만료 없음)
            nx: True이면 키가 없을 때만 설정 (SET NX)

        Returns:
            bool: 성공 시 True

        Raises:
            ServiceUnavailable: Redis 에러 시

        Note:
            - bytes 타입: 그대로 저장 (pickle 결과 등)
            - str 타입: 그대로 저장
            - 기타: JSON 직렬화
            - Fallback 모드: 메모리 캐시 사용
        """
        # Fallback 모드일 때 메모리 캐시 사용
        if getattr(self, '_use_fallback', False) and self._fallback_cache:
            # nx 옵션 처리 (키가 없을 때만 설정)
            if nx and self._fallback_cache.exists(key) > 0:
                return False
            # Fallback 캐시는 JSON 직렬화 없이 저장 (Python 객체 그대로)
            return self._fallback_cache.set(key, value, ex=ex)

        try:
            # bytes는 그대로 저장 (pickle 등)
            if isinstance(value, bytes):
                # Redis는 decode_responses=True이므로 bytes를 직접 저장하려면
                # decode_responses=False인 별도 연결 필요
                # 임시 방안: base64 인코딩하여 문자열로 저장
                import base64
                serialized_value = f"__BYTES__:{base64.b64encode(value).decode('ascii')}"
            # 문자열은 그대로
            elif isinstance(value, str):
                serialized_value = value
            # 기타는 JSON 직렬화
            else:
                serialized_value = json.dumps(value, ensure_ascii=False)

            return self.redis.set(key, serialized_value, ex=ex, nx=nx)

        except (RedisError, TypeError, json.JSONEncodeError) as e:
            logger.error(f"Redis set 실패: key={key}, error={e}")
            raise ServiceUnavailable(f"Cache unavailable: {e}")

    def delete(self, *keys: str) -> int:
        """
        키 삭제

        Args:
            *keys: 삭제할 키 목록

        Returns:
            int: 삭제된 키 개수

        Raises:
            ServiceUnavailable: Redis 에러 시
        """
        # Fallback 모드일 때 메모리 캐시 사용
        if getattr(self, '_use_fallback', False) and self._fallback_cache:
            return self._fallback_cache.delete(*keys)

        try:
            return self.redis.delete(*keys)
        except RedisError as e:
            logger.error(f"Redis delete 실패: keys={keys}, error={e}")
            raise ServiceUnavailable(f"Cache unavailable: {e}")

    def exists(self, *keys: str) -> int:
        """
        키 존재 여부 확인

        Args:
            *keys: 확인할 키 목록

        Returns:
            int: 존재하는 키 개수

        Raises:
            ServiceUnavailable: Redis 에러 시
        """
        # Fallback 모드일 때 메모리 캐시 사용
        if getattr(self, '_use_fallback', False) and self._fallback_cache:
            return self._fallback_cache.exists(*keys)

        try:
            return self.redis.exists(*keys)
        except RedisError as e:
            logger.error(f"Redis exists 실패: keys={keys}, error={e}")
            raise ServiceUnavailable(f"Cache unavailable: {e}")

    def expire(self, key: str, seconds: int) -> bool:
        """
        키 TTL 설정

        Args:
            key: Redis 키
            seconds: TTL (초)

        Returns:
            bool: 성공 시 True

        Raises:
            ServiceUnavailable: Redis 에러 시
        """
        # Fallback 모드일 때 메모리 캐시 사용
        if getattr(self, '_use_fallback', False) and self._fallback_cache:
            return self._fallback_cache.expire(key, seconds)

        try:
            return self.redis.expire(key, seconds)
        except RedisError as e:
            logger.error(f"Redis expire 실패: key={key}, error={e}")
            raise ServiceUnavailable(f"Cache unavailable: {e}")

    def ttl(self, key: str) -> int:
        """
        키의 남은 TTL 조회

        Args:
            key: Redis 키

        Returns:
            int: 남은 TTL (초), -1이면 만료 없음, -2이면 키 없음

        Raises:
            ServiceUnavailable: Redis 에러 시
        """
        # Fallback 모드일 때 메모리 캐시 사용
        if getattr(self, '_use_fallback', False) and self._fallback_cache:
            return self._fallback_cache.ttl(key)

        try:
            return self.redis.ttl(key)
        except RedisError as e:
            logger.error(f"Redis ttl 실패: key={key}, error={e}")
            raise ServiceUnavailable(f"Cache unavailable: {e}")

    def keys(self, pattern: str = '*') -> list:
        """
        패턴에 맞는 키 목록 조회

        Args:
            pattern: 패턴 (예: "cache:*")

        Returns:
            list: 키 목록

        Raises:
            ServiceUnavailable: Redis 에러 시

        Warning:
            KEYS 명령은 O(N)이므로 프로덕션에서는 주의해서 사용.
            대규모 키 조회 시 scan_iter() 사용 권장.
        """
        # Fallback 모드일 때 메모리 캐시 사용
        if getattr(self, '_use_fallback', False) and self._fallback_cache:
            return self._fallback_cache.keys(pattern)

        try:
            return self.redis.keys(pattern)
        except RedisError as e:
            logger.error(f"Redis keys 실패: pattern={pattern}, error={e}")
            raise ServiceUnavailable(f"Cache unavailable: {e}")

    def scan_iter(self, match: str = '*', count: int = 100) -> list:
        """
        패턴에 맞는 키를 SCAN으로 non-blocking iteration (2026-07-08 추가).

        KEYS(O(N) 블로킹) 대신 SCAN cursor를 사용해 서버 blocking 회피.
        내부적으로 여러 번 SCAN 호출하지만 사이사이 다른 요청 처리 가능.

        Args:
            match: 패턴 (예: "lock:*")
            count: 배치당 반환 수 힌트 (기본 100)

        Returns:
            list: 매칭된 키 전체 목록
        """
        if getattr(self, '_use_fallback', False) and self._fallback_cache:
            return self._fallback_cache.keys(match)

        try:
            return list(self.redis.scan_iter(match=match, count=count))
        except RedisError as e:
            logger.error(f"Redis scan_iter 실패: match={match}, error={e}")
            raise ServiceUnavailable(f"Cache unavailable: {e}")

    def flushdb(self) -> bool:
        """
        현재 DB의 모든 키 삭제

        Returns:
            bool: 성공 시 True

        Raises:
            ServiceUnavailable: Redis 에러 시

        Warning:
            프로덕션에서는 절대 사용하지 마세요!
        """
        # Fallback 모드일 때 메모리 캐시 사용
        if getattr(self, '_use_fallback', False) and self._fallback_cache:
            return self._fallback_cache.flushdb()

        try:
            return self.redis.flushdb()
        except RedisError as e:
            logger.error(f"Redis flushdb 실패: error={e}")
            raise ServiceUnavailable(f"Cache unavailable: {e}")

    # ========== 분산 락 관련 메서드 ==========

    def acquire_lock(self, key: str, timeout: int = 300, value: str = "locked") -> bool:
        """
        분산 락 획득 (SETNX + TTL)

        Args:
            key: 락 키
            timeout: 락 TTL (초)
            value: 락 값 (소유자 식별용)

        Returns:
            bool: 획득 성공 시 True, 실패 시 False

        Raises:
            ServiceUnavailable: Redis 에러 시
        """
        # Fallback 모드일 때 메모리 캐시 사용 (단, 분산 락 기능은 제한적)
        if getattr(self, '_use_fallback', False) and self._fallback_cache:
            logger.warning(f"⚠️ Fallback 모드: 분산 락이 단일 프로세스로 제한됨 - key={key}")
            # nx 옵션 처리 (키가 없을 때만 설정)
            if self._fallback_cache.exists(key) > 0:
                return False
            return self._fallback_cache.set(key, value, ex=timeout)

        try:
            return self.redis.set(key, value, nx=True, ex=timeout)
        except RedisError as e:
            logger.error(f"Redis 락 획득 실패: key={key}, error={e}")
            raise ServiceUnavailable(f"Lock unavailable: {e}")

    def release_lock(self, key: str) -> int:
        """
        분산 락 해제

        Args:
            key: 락 키

        Returns:
            int: 삭제된 키 개수 (1이면 성공, 0이면 이미 없음)

        Raises:
            ServiceUnavailable: Redis 에러 시
        """
        # delete() 메서드가 이미 fallback을 처리하므로 그대로 위임
        return self.delete(key)

    def extend_lock(self, key: str, additional_time: int) -> bool:
        """
        락 TTL 연장

        Args:
            key: 락 키
            additional_time: 추가 시간 (초)

        Returns:
            bool: 성공 시 True

        Raises:
            ServiceUnavailable: Redis 에러 시
        """
        # ttl() 과 expire() 메서드가 이미 fallback을 처리하므로 그대로 사용
        try:
            current_ttl = self.ttl(key)
            if current_ttl > 0:
                new_ttl = current_ttl + additional_time
                return self.expire(key, new_ttl)
            return False
        except RedisError as e:
            logger.error(f"Redis 락 연장 실패: key={key}, error={e}")
            raise ServiceUnavailable(f"Lock unavailable: {e}")


# Singleton 인스턴스 생성 (모듈 로드 시 자동 생성)
_redis_client = None


def get_redis_client() -> RedisClient:
    """
    RedisClient 싱글톤 인스턴스 반환

    Returns:
        RedisClient: Redis 클라이언트 인스턴스
    """
    global _redis_client
    if _redis_client is None:
        _redis_client = RedisClient()
    return _redis_client
