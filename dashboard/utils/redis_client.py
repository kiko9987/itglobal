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
from typing import Any, Optional

import redis
from redis.exceptions import RedisError, ConnectionError, TimeoutError

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
            redis_host = os.getenv('REDIS_HOST', 'localhost')
            redis_port = int(os.getenv('REDIS_PORT', 6379))
            redis_db = int(os.getenv('REDIS_DB', 0))
            redis_password = os.getenv('REDIS_PASSWORD', None)
            socket_timeout = int(os.getenv('REDIS_SOCKET_TIMEOUT', 5))
            socket_connect_timeout = int(os.getenv('REDIS_SOCKET_CONNECT_TIMEOUT', 5))

            # 비밀번호가 빈 문자열이면 None으로 처리
            if redis_password == '':
                redis_password = None

            # Redis 연결
            self.redis = redis.Redis(
                host=redis_host,
                port=redis_port,
                db=redis_db,
                password=redis_password,
                decode_responses=True,  # 문자열로 자동 디코딩
                socket_timeout=socket_timeout,
                socket_connect_timeout=socket_connect_timeout,
                socket_keepalive=True,
                health_check_interval=30  # 30초마다 연결 확인
            )

            # Fail Fast: 부팅 시 연결 확인
            self.redis.ping()
            logger.info(f"Redis 연결 성공: {redis_host}:{redis_port} (DB: {redis_db})")

            self._initialized = True

        except (ConnectionError, TimeoutError, RedisError) as e:
            logger.critical(f"Redis 연결 실패: {e}")
            logger.critical("서비스를 시작할 수 없습니다. Redis 서버를 확인하세요.")
            sys.exit(1)  # 프로세스 종료

    def ping(self) -> bool:
        """
        Redis 연결 확인

        Returns:
            bool: 연결 성공 시 True, 실패 시 False
        """
        try:
            return self.redis.ping()
        except RedisError:
            return False

    def get(self, key: str) -> Optional[Any]:
        """
        값 조회 (JSON 자동 역직렬화)

        Args:
            key: Redis 키

        Returns:
            저장된 값 (None이면 키가 없음)

        Raises:
            ServiceUnavailable: Redis 에러 시
        """
        try:
            value = self.redis.get(key)
            if value is None:
                return None

            # JSON 역직렬화 시도
            try:
                return json.loads(value)
            except (json.JSONDecodeError, TypeError):
                # JSON이 아니면 문자열 그대로 반환
                return value

        except RedisError as e:
            logger.error(f"Redis get 실패: key={key}, error={e}")
            raise ServiceUnavailable(f"Cache unavailable: {e}")

    def set(self, key: str, value: Any, ex: int = None, nx: bool = False) -> bool:
        """
        값 저장 (JSON 자동 직렬화)

        Args:
            key: Redis 키
            value: 저장할 값 (JSON 직렬화 가능한 객체)
            ex: TTL (초 단위, None이면 만료 없음)
            nx: True이면 키가 없을 때만 설정 (SET NX)

        Returns:
            bool: 성공 시 True

        Raises:
            ServiceUnavailable: Redis 에러 시
        """
        try:
            # JSON 직렬화 (문자열이면 그대로)
            if isinstance(value, str):
                serialized_value = value
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
            KEYS 명령은 O(N)이므로 프로덕션에서는 주의해서 사용
        """
        try:
            return self.redis.keys(pattern)
        except RedisError as e:
            logger.error(f"Redis keys 실패: pattern={pattern}, error={e}")
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
