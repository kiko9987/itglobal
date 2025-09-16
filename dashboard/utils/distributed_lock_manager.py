"""
분산 잠금 관리자 - Redis 기반
여러 서버 인스턴스 간 데이터 동시성 보장
"""

import redis
import uuid
import time
import json
from contextlib import contextmanager
from typing import Optional, Dict, Any
from datetime import datetime, timedelta
import logging

logger = logging.getLogger(__name__)

class DistributedLockManager:
    """Redis 기반 분산 잠금 관리자"""

    def __init__(self, redis_host='localhost', redis_port=6379, redis_db=0):
        self.redis_client = redis.Redis(
            host=redis_host,
            port=redis_port,
            db=redis_db,
            decode_responses=True
        )
        self.default_timeout = 300  # 5분

    @contextmanager
    def acquire_field_lock(self, project_code: str, field_name: str,
                          user_id: str, timeout: int = None):
        """필드별 분산 잠금 획득"""
        timeout = timeout or self.default_timeout
        lock_key = f"field_lock:{project_code}:{field_name}"
        lock_value = f"{user_id}:{uuid.uuid4()}"

        try:
            # 잠금 획득 시도
            if self.redis_client.set(lock_key, lock_value, nx=True, ex=timeout):
                logger.info(f"필드 잠금 획득: {lock_key} by {user_id}")
                yield lock_value
            else:
                # 잠금 실패 시 현재 잠금 정보 확인
                current_lock = self.redis_client.get(lock_key)
                if current_lock:
                    current_user = current_lock.split(':')[0]
                    raise FieldLockError(f"필드가 다른 사용자에 의해 편집 중입니다: {current_user}")
                else:
                    raise FieldLockError("잠금 획득에 실패했습니다")

        finally:
            # Lua 스크립트로 원자적 잠금 해제 (자신의 잠금만 해제)
            release_script = """
            if redis.call("get", KEYS[1]) == ARGV[1] then
                return redis.call("del", KEYS[1])
            else
                return 0
            end
            """
            result = self.redis_client.eval(release_script, 1, lock_key, lock_value)
            if result:
                logger.info(f"필드 잠금 해제: {lock_key} by {user_id}")

    def get_field_lock_info(self, project_code: str, field_name: str) -> Optional[Dict[str, Any]]:
        """필드 잠금 정보 조회"""
        lock_key = f"field_lock:{project_code}:{field_name}"
        lock_value = self.redis_client.get(lock_key)

        if lock_value:
            user_id, lock_id = lock_value.split(':', 1)
            ttl = self.redis_client.ttl(lock_key)

            return {
                'locked': True,
                'user_id': user_id,
                'lock_id': lock_id,
                'expires_in': ttl,
                'expires_at': (datetime.now() + timedelta(seconds=ttl)).isoformat()
            }

        return {'locked': False}

    def force_release_lock(self, project_code: str, field_name: str, admin_user: str) -> bool:
        """관리자 권한으로 강제 잠금 해제"""
        lock_key = f"field_lock:{project_code}:{field_name}"
        result = self.redis_client.delete(lock_key)

        if result:
            logger.warning(f"관리자 강제 잠금 해제: {lock_key} by {admin_user}")

            # 강제 해제 로그 기록
            log_key = f"force_release_log:{datetime.now().strftime('%Y%m%d')}"
            log_entry = {
                'timestamp': datetime.now().isoformat(),
                'project_code': project_code,
                'field_name': field_name,
                'admin_user': admin_user,
                'action': 'force_release'
            }
            self.redis_client.lpush(log_key, json.dumps(log_entry))
            self.redis_client.expire(log_key, 86400 * 30)  # 30일 보관

        return bool(result)

    def get_all_locks(self) -> Dict[str, Dict[str, Any]]:
        """현재 활성화된 모든 잠금 조회"""
        pattern = "field_lock:*"
        lock_keys = self.redis_client.keys(pattern)

        locks = {}
        for lock_key in lock_keys:
            _, project_code, field_name = lock_key.split(':', 2)
            lock_info = self.get_field_lock_info(project_code, field_name)

            if lock_info['locked']:
                locks[f"{project_code}:{field_name}"] = lock_info

        return locks

    def cleanup_expired_locks(self):
        """만료된 잠금 정리 (Redis가 자동으로 처리하지만 명시적 정리용)"""
        pattern = "field_lock:*"
        lock_keys = self.redis_client.keys(pattern)

        cleaned_count = 0
        for lock_key in lock_keys:
            ttl = self.redis_client.ttl(lock_key)
            if ttl == -1:  # TTL이 설정되지 않은 경우
                self.redis_client.delete(lock_key)
                cleaned_count += 1

        logger.info(f"만료된 잠금 정리 완료: {cleaned_count}개")
        return cleaned_count

    def extend_lock(self, project_code: str, field_name: str,
                   user_id: str, additional_time: int = 300) -> bool:
        """잠금 시간 연장"""
        lock_key = f"field_lock:{project_code}:{field_name}"
        lock_value = self.redis_client.get(lock_key)

        if lock_value and lock_value.startswith(f"{user_id}:"):
            current_ttl = self.redis_client.ttl(lock_key)
            new_ttl = max(current_ttl + additional_time, additional_time)

            result = self.redis_client.expire(lock_key, new_ttl)
            if result:
                logger.info(f"잠금 시간 연장: {lock_key} +{additional_time}초")
            return bool(result)

        return False

    def health_check(self) -> Dict[str, Any]:
        """Redis 연결 상태 확인"""
        try:
            # Redis 연결 테스트
            self.redis_client.ping()

            # 현재 잠금 통계
            lock_count = len(self.redis_client.keys("field_lock:*"))

            return {
                'status': 'healthy',
                'redis_connected': True,
                'active_locks': lock_count,
                'timestamp': datetime.now().isoformat()
            }
        except Exception as e:
            logger.error(f"Redis 연결 오류: {e}")
            return {
                'status': 'error',
                'redis_connected': False,
                'error': str(e),
                'timestamp': datetime.now().isoformat()
            }


class FieldLockError(Exception):
    """필드 잠금 관련 예외"""
    pass


# 전역 인스턴스 (설정에 따라 Redis 연결 정보 수정 필요)
try:
    distributed_lock_manager = DistributedLockManager()
except Exception as e:
    logger.warning(f"분산 잠금 관리자 초기화 실패, 로컬 모드로 폴백: {e}")
    distributed_lock_manager = None


def get_distributed_lock_manager() -> Optional[DistributedLockManager]:
    """분산 잠금 관리자 인스턴스 반환"""
    return distributed_lock_manager


# 편의 함수들
def acquire_field_lock(project_code: str, field_name: str, user_id: str, timeout: int = None):
    """필드 잠금 획득 (컨텍스트 매니저)"""
    if distributed_lock_manager:
        return distributed_lock_manager.acquire_field_lock(project_code, field_name, user_id, timeout)
    else:
        # 로컬 폴백 (기존 field_lock_manager 사용)
        from .field_lock_manager import field_lock_manager
        return field_lock_manager.acquire_lock(project_code, field_name, user_id, "Unknown")


def check_field_lock(project_code: str, field_name: str) -> Dict[str, Any]:
    """필드 잠금 상태 확인"""
    if distributed_lock_manager:
        return distributed_lock_manager.get_field_lock_info(project_code, field_name)
    else:
        # 로컬 폴백
        from .field_lock_manager import field_lock_manager
        return field_lock_manager.get_lock_status(project_code, field_name)