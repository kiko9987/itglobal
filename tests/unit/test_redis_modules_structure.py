"""
Redis 모듈 구조 검증 테스트 (Redis 서버 불필요)

Redis 연결 없이 모듈 구조, 인터페이스, 예외 클래스 등을 검증합니다.
"""

import pytest
from unittest.mock import Mock, patch, MagicMock
from datetime import datetime, timedelta

# Mock Redis를 사용하여 import 가능하도록 설정
with patch('dashboard.utils.redis_client.redis.Redis'):
    with patch('dashboard.utils.redis_client.RedisClient.__init__', lambda x: None):
        from dashboard.utils.exceptions import ServiceUnavailable, LockAcquisitionError, CacheError
        from dashboard.utils.smart_cache_manager import CacheStrategy
        from dashboard.utils.project_lock_manager import ProjectLock, GRACE_PERIOD_SECONDS


# ==================== 예외 클래스 테스트 ====================

def test_service_unavailable_exception():
    """ServiceUnavailable 예외 생성 및 메시지 테스트"""
    exc = ServiceUnavailable("Redis is down")
    assert exc.message == "Redis is down"
    assert str(exc) == "Redis is down"


def test_lock_acquisition_error_exception():
    """LockAcquisitionError 예외 생성 테스트"""
    exc = LockAcquisitionError("Cannot acquire lock")
    assert exc.message == "Cannot acquire lock"


def test_cache_error_exception():
    """CacheError 예외 생성 테스트"""
    exc = CacheError("Cache operation failed")
    assert exc.message == "Cache operation failed"


# ==================== CacheStrategy Enum 테스트 ====================

def test_cache_strategy_enum_values():
    """CacheStrategy Enum 값 검증"""
    assert CacheStrategy.CRITICAL_DATA.value == "critical_data"
    assert CacheStrategy.STATIC_CONFIG.value == "static_config"
    assert CacheStrategy.FOLDER_MAPPING.value == "folder_mapping"
    assert CacheStrategy.UI_STATE.value == "ui_state"
    assert CacheStrategy.TEMPORARY.value == "temporary"
    assert CacheStrategy.METADATA.value == "metadata"


def test_cache_strategy_ttl_mapping():
    """CacheStrategy TTL 매핑 검증"""
    from dashboard.utils.smart_cache_manager import SimpleCache

    ttl_map = SimpleCache.TTL_MAP

    assert ttl_map[CacheStrategy.CRITICAL_DATA] == 60  # 60초
    assert ttl_map[CacheStrategy.STATIC_CONFIG] == 3600  # 1시간
    assert ttl_map[CacheStrategy.FOLDER_MAPPING] == 86400  # 1일
    assert ttl_map[CacheStrategy.UI_STATE] == 300  # 5분
    assert ttl_map[CacheStrategy.TEMPORARY] == 60  # 1분
    assert ttl_map[CacheStrategy.METADATA] == 600  # 10분


# ==================== ProjectLock 데이터 클래스 테스트 ====================

def test_project_lock_creation():
    """ProjectLock 객체 생성 테스트"""
    now = datetime.now()
    expires = now + timedelta(minutes=5)

    lock = ProjectLock(
        project_code="TEST001",
        user_email="user@test.com",
        user_name="Test User",
        tab_id="tab123",
        locked_at=now,
        expires_at=expires
    )

    assert lock.project_code == "TEST001"
    assert lock.user_email == "user@test.com"
    assert lock.user_name == "Test User"
    assert lock.tab_id == "tab123"
    assert lock.locked_at == now
    assert lock.expires_at == expires


def test_project_lock_to_dict():
    """ProjectLock.to_dict() 메서드 테스트"""
    now = datetime.now()
    expires = now + timedelta(minutes=5)

    lock = ProjectLock(
        project_code="TEST002",
        user_email="user@test.com",
        user_name="Test User",
        tab_id="tab456",
        locked_at=now,
        expires_at=expires
    )

    lock_dict = lock.to_dict()

    assert lock_dict['project_code'] == "TEST002"
    assert lock_dict['user_email'] == "user@test.com"
    assert lock_dict['user_name'] == "Test User"
    assert lock_dict['tab_id'] == "tab456"
    assert lock_dict['locked_at'] == now.isoformat()
    assert lock_dict['expires_at'] == expires.isoformat()
    assert 'remaining_minutes' in lock_dict
    assert lock_dict['remaining_minutes'] >= 0


def test_project_lock_remaining_minutes_calculation():
    """ProjectLock remaining_minutes 계산 테스트"""
    now = datetime.now()
    expires = now + timedelta(minutes=3)

    lock = ProjectLock(
        project_code="TEST003",
        user_email="user@test.com",
        user_name="Test User",
        tab_id="tab789",
        locked_at=now,
        expires_at=expires
    )

    lock_dict = lock.to_dict()

    # 약 3분 남아있어야 함 (오차 허용)
    assert 2.9 <= lock_dict['remaining_minutes'] <= 3.1


def test_project_lock_expired():
    """만료된 ProjectLock의 remaining_minutes는 0"""
    now = datetime.now()
    expires = now - timedelta(minutes=1)  # 1분 전에 만료

    lock = ProjectLock(
        project_code="TEST004",
        user_email="user@test.com",
        user_name="Test User",
        tab_id="tab101",
        locked_at=now - timedelta(minutes=6),
        expires_at=expires
    )

    lock_dict = lock.to_dict()

    # 만료된 경우 0으로 표시
    assert lock_dict['remaining_minutes'] == 0


# ==================== Grace Period 상수 테스트 ====================

def test_grace_period_constant():
    """GRACE_PERIOD_SECONDS 상수 검증"""
    assert GRACE_PERIOD_SECONDS == 30  # 30초


# ==================== 모듈 인터페이스 검증 ====================

def test_smart_cache_manager_interface():
    """smart_cache_manager 모듈의 public 인터페이스 검증"""
    from dashboard.utils import smart_cache_manager

    # 전역 함수 존재 확인
    assert hasattr(smart_cache_manager, 'smart_get')
    assert hasattr(smart_cache_manager, 'smart_set')
    assert hasattr(smart_cache_manager, 'smart_delete')
    assert hasattr(smart_cache_manager, 'smart_invalidate')
    assert hasattr(smart_cache_manager, 'cache_clear')
    assert hasattr(smart_cache_manager, 'cache_stats')
    assert hasattr(smart_cache_manager, 'get_smart_cache')

    # 클래스 존재 확인
    assert hasattr(smart_cache_manager, 'SimpleCache')
    assert hasattr(smart_cache_manager, 'CacheStrategy')


def test_project_lock_manager_interface():
    """project_lock_manager 모듈의 public 인터페이스 검증"""
    from dashboard.utils import project_lock_manager

    # 클래스 존재 확인
    assert hasattr(project_lock_manager, 'ProjectLockManager')
    assert hasattr(project_lock_manager, 'ProjectLock')

    # 전역 함수 존재 확인
    assert hasattr(project_lock_manager, 'get_project_lock_manager')

    # 상수 존재 확인
    assert hasattr(project_lock_manager, 'GRACE_PERIOD_SECONDS')


def test_redis_client_interface():
    """redis_client 모듈의 public 인터페이스 검증"""
    from dashboard.utils import redis_client

    # 클래스 존재 확인
    assert hasattr(redis_client, 'RedisClient')

    # 전역 함수 존재 확인
    assert hasattr(redis_client, 'get_redis_client')


def test_exceptions_interface():
    """exceptions 모듈의 public 인터페이스 검증"""
    from dashboard.utils import exceptions

    # 예외 클래스 존재 확인
    assert hasattr(exceptions, 'ServiceUnavailable')
    assert hasattr(exceptions, 'LockAcquisitionError')
    assert hasattr(exceptions, 'CacheError')


# ==================== 인터페이스 시그니처 테스트 ====================

def test_cache_manager_method_signatures():
    """SimpleCache 메서드 시그니처 검증"""
    from dashboard.utils.smart_cache_manager import SimpleCache
    import inspect

    # get 메서드
    get_sig = inspect.signature(SimpleCache.get)
    assert 'key' in get_sig.parameters
    assert 'strategy' in get_sig.parameters

    # set 메서드
    set_sig = inspect.signature(SimpleCache.set)
    assert 'key' in set_sig.parameters
    assert 'value' in set_sig.parameters
    assert 'strategy' in set_sig.parameters
    assert 'custom_ttl' in set_sig.parameters
    assert 'fetched_at' in set_sig.parameters

    # delete 메서드
    delete_sig = inspect.signature(SimpleCache.delete)
    assert 'key' in delete_sig.parameters
    assert 'set_marker' in delete_sig.parameters


def test_lock_manager_method_signatures():
    """ProjectLockManager 메서드 시그니처 검증"""
    from dashboard.utils.project_lock_manager import ProjectLockManager
    import inspect

    # acquire_lock 메서드
    acquire_sig = inspect.signature(ProjectLockManager.acquire_lock)
    assert 'project_code' in acquire_sig.parameters
    assert 'user_email' in acquire_sig.parameters
    assert 'user_name' in acquire_sig.parameters
    assert 'tab_id' in acquire_sig.parameters

    # release_lock 메서드
    release_sig = inspect.signature(ProjectLockManager.release_lock)
    assert 'project_code' in release_sig.parameters
    assert 'user_email' in release_sig.parameters
    assert 'tab_id' in release_sig.parameters

    # force_release_lock 메서드
    force_sig = inspect.signature(ProjectLockManager.force_release_lock)
    assert 'project_code' in force_sig.parameters
    assert 'admin_email' in force_sig.parameters
    assert 'admin_permission' in force_sig.parameters


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
