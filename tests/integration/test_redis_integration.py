"""
Redis 통합 테스트
캐시 매니저와 락 매니저의 Redis 기반 동작 검증

테스트 시나리오:
1. Redis 연결 실패 시 Fail Fast (sys.exit)
2. 캐시 기본 동작 (get/set/delete/invalidate)
3. 캐시 레이스 컨디션 방지 (무효화 마커)
4. 분산 락 획득/해제
5. 락 TTL 자동 만료
6. Grace Period 복구
7. 동시 락 획득 시도
"""

import pytest
import time
import os
import sys
from datetime import datetime, timedelta
from unittest.mock import patch, MagicMock
from redis.exceptions import ConnectionError as RedisConnectionError

from dashboard.utils.redis_client import RedisClient, get_redis_client
from dashboard.utils.smart_cache_manager import SimpleCache, CacheStrategy
from dashboard.utils.project_lock_manager import ProjectLockManager, ProjectLock
from dashboard.utils.exceptions import ServiceUnavailable


# ==================== Fixtures ====================

@pytest.fixture
def redis_client():
    """Redis 클라이언트 픽스처.

    ⚠️ 2026-07-24 사고 반영: 프로덕션 Redis (db=0) 를 flushdb 하면 lead_card_msg,
    visit_auto_completed 등 운영 데이터 대량 손실. 반드시 별도 test DB (db=15) 사용.
    REDIS_TEST_DB env 로 override 가능. 미설정 시 기본 15.

    또 다중 안전장치:
    - db=0 이면 강제 skip (assertion) — 실수로 override 해도 프로덕션 안 지움
    - flushdb 대신 keys("test:*") 만 삭제하는 방식 대안도 있으나 fixture 격리 위해 flushdb 유지
    """
    from redis import Redis
    test_db = int(os.getenv('REDIS_TEST_DB', '15'))
    assert test_db != 0, (
        f'REDIS_TEST_DB={test_db} 은 프로덕션 DB — 통합 테스트에서 flushdb 금지. '
        f'별도 test DB (예: 15) 지정 필수.'
    )
    host = os.getenv('REDIS_HOST', 'localhost')
    port = int(os.getenv('REDIS_PORT', '6379'))
    client = Redis(host=host, port=port, db=test_db, decode_responses=False)
    # 테스트 전 Redis 정리 (test DB 만)
    client.flushdb()
    yield client
    # 테스트 후 Redis 정리
    client.flushdb()
    client.close()


@pytest.fixture
def cache_manager(redis_client):
    """캐시 매니저 픽스처"""
    cache = SimpleCache()
    return cache


@pytest.fixture
def lock_manager(redis_client):
    """락 매니저 픽스처"""
    manager = ProjectLockManager(lock_timeout_minutes=1)  # 테스트용 1분 타임아웃
    yield manager
    manager.shutdown()


# ==================== Redis 연결 테스트 ====================

def test_redis_connection_success(redis_client):
    """Redis 연결 성공 테스트"""
    assert redis_client.ping() is True


def test_redis_connection_failure_on_init():
    """Redis 연결 실패 시 Fail Fast 테스트

    Note:
        실제 sys.exit(1) 호출을 검증하므로 주의 필요
        CI 환경에서는 skip 권장
    """
    # Redis 연결 실패 시뮬레이션
    with patch('redis.Redis.ping', side_effect=RedisConnectionError("Connection refused")):
        with pytest.raises(SystemExit) as exc_info:
            # 새 RedisClient 인스턴스 생성 시도
            RedisClient._instance = None  # Singleton 리셋
            RedisClient()

        # sys.exit(1) 호출 확인
        assert exc_info.value.code == 1


# ==================== 캐시 매니저 테스트 ====================

def test_cache_basic_get_set(cache_manager):
    """캐시 기본 get/set 동작 테스트"""
    # Set
    cache_manager.set("test_key", {"data": "test_value"}, strategy=CacheStrategy.CRITICAL_DATA)

    # Get
    result = cache_manager.get("test_key")
    assert result is not None
    assert result["data"] == "test_value"


def test_cache_get_nonexistent_key(cache_manager):
    """존재하지 않는 키 조회 테스트"""
    result = cache_manager.get("nonexistent_key")
    assert result is None


def test_cache_ttl_expiry(cache_manager):
    """캐시 TTL 만료 테스트"""
    # 1초 TTL로 설정
    cache_manager.set("ttl_key", {"data": "ttl_test"}, custom_ttl=1)

    # 즉시 조회 - 존재해야 함
    result = cache_manager.get("ttl_key")
    assert result is not None

    # 2초 대기 후 조회 - 만료되어야 함
    time.sleep(2)
    result = cache_manager.get("ttl_key")
    assert result is None


def test_cache_delete(cache_manager):
    """캐시 삭제 테스트"""
    cache_manager.set("delete_key", {"data": "delete_test"})

    # 삭제 전 확인
    assert cache_manager.get("delete_key") is not None

    # 삭제
    deleted = cache_manager.delete("delete_key")
    assert deleted is True

    # 삭제 후 확인
    assert cache_manager.get("delete_key") is None


def test_cache_invalidate_by_pattern(cache_manager):
    """패턴 기반 캐시 무효화 테스트"""
    # 여러 키 생성
    cache_manager.set("user:1:profile", {"name": "User1"})
    cache_manager.set("user:2:profile", {"name": "User2"})
    cache_manager.set("product:1:info", {"name": "Product1"})

    # user 패턴으로 무효화
    deleted_count = cache_manager.invalidate_by_pattern("user")
    assert deleted_count == 2

    # 확인
    assert cache_manager.get("user:1:profile") is None
    assert cache_manager.get("user:2:profile") is None
    assert cache_manager.get("product:1:info") is not None


def test_cache_clear(cache_manager):
    """전체 캐시 삭제 테스트"""
    # 여러 키 생성
    cache_manager.set("key1", {"data": "value1"})
    cache_manager.set("key2", {"data": "value2"})
    cache_manager.set("key3", {"data": "value3"})

    # 전체 삭제
    deleted_count = cache_manager.clear()
    assert deleted_count == 3

    # 확인
    assert cache_manager.get("key1") is None
    assert cache_manager.get("key2") is None
    assert cache_manager.get("key3") is None


def test_cache_race_condition_prevention(cache_manager):
    """레이스 컨디션 방지 테스트 (무효화 마커)"""
    # 1. 데이터 수집 시작 시각 기록
    fetch_start_time = time.time()

    # 2. 캐시에 데이터 저장
    cache_manager.set("race_key", {"data": "old_value"}, fetched_at=fetch_start_time)

    # 3. 즉시 캐시 무효화 (마커 설정)
    cache_manager.delete("race_key", set_marker=True)

    # 4. 무효화 이전에 수집된 데이터로 캐시 쓰기 시도 → 거부되어야 함
    cache_manager.set("race_key", {"data": "stale_value"}, fetched_at=fetch_start_time)

    # 5. 캐시에 저장되지 않았는지 확인
    result = cache_manager.get("race_key")
    assert result is None

    # 6. 무효화 이후 새로 수집한 데이터는 저장 가능
    new_fetch_time = time.time()
    cache_manager.set("race_key", {"data": "fresh_value"}, fetched_at=new_fetch_time)

    result = cache_manager.get("race_key")
    assert result is not None
    assert result["data"] == "fresh_value"


def test_cache_info(cache_manager):
    """캐시 정보 조회 테스트"""
    # 여러 키 생성
    cache_manager.set("info_key1", {"data": "value1"})
    cache_manager.set("info_key2", {"data": "value2"})

    # 정보 조회
    info = cache_manager.get_cache_info()

    assert info['total_items'] == 2
    assert info['active_items'] == 2
    assert info['expired_items'] == 0


# ==================== 락 매니저 테스트 ====================

def test_lock_acquire_release(lock_manager):
    """락 획득 및 해제 테스트"""
    # 락 획득
    result = lock_manager.acquire_lock(
        project_code="TEST001",
        user_email="user@test.com",
        user_name="Test User",
        tab_id="tab123"
    )

    assert result['success'] is True
    assert result['message'] == '편집 모드로 전환되었습니다.'
    assert result['lock_info']['project_code'] == "TEST001"

    # 락 해제
    result = lock_manager.release_lock(
        project_code="TEST001",
        user_email="user@test.com",
        tab_id="tab123"
    )

    assert result['success'] is True
    assert result['message'] == '편집 모드가 종료되었습니다.'


def test_lock_acquire_by_different_user(lock_manager):
    """다른 사용자가 이미 잠긴 프로젝트 획득 시도 테스트"""
    # User1이 락 획득
    lock_manager.acquire_lock(
        project_code="TEST002",
        user_email="user1@test.com",
        user_name="User 1",
        tab_id="tab1"
    )

    # User2가 같은 프로젝트 락 획득 시도 → 실패해야 함
    result = lock_manager.acquire_lock(
        project_code="TEST002",
        user_email="user2@test.com",
        user_name="User 2",
        tab_id="tab2"
    )

    assert result['success'] is False
    assert "User 1" in result['message']
    assert result['same_user'] is False


def test_lock_acquire_different_tab_same_user(lock_manager):
    """같은 사용자가 다른 탭에서 획득 시도 테스트"""
    # Tab1에서 락 획득
    lock_manager.acquire_lock(
        project_code="TEST003",
        user_email="user@test.com",
        user_name="Test User",
        tab_id="tab1"
    )

    # 같은 사용자가 Tab2에서 획득 시도 → 실패해야 함
    result = lock_manager.acquire_lock(
        project_code="TEST003",
        user_email="user@test.com",
        user_name="Test User",
        tab_id="tab2"
    )

    assert result['success'] is False
    assert "다른 탭" in result['message']
    assert result['same_user'] is True


def test_lock_extend(lock_manager):
    """같은 사용자+탭이 락 재획득 시 연장 테스트"""
    # 락 획득
    result1 = lock_manager.acquire_lock(
        project_code="TEST004",
        user_email="user@test.com",
        user_name="Test User",
        tab_id="tab123"
    )

    expires_at_1 = result1['lock_info']['expires_at']

    # 1초 대기
    time.sleep(1)

    # 같은 사용자+탭이 재획득 → 연장되어야 함
    result2 = lock_manager.acquire_lock(
        project_code="TEST004",
        user_email="user@test.com",
        user_name="Test User",
        tab_id="tab123"
    )

    assert result2['success'] is True
    assert "연장" in result2['message']

    expires_at_2 = result2['lock_info']['expires_at']

    # 만료 시간이 연장되었는지 확인
    assert expires_at_2 > expires_at_1


def test_lock_ttl_expiry(lock_manager):
    """락 TTL 자동 만료 테스트

    Note:
        lock_manager 픽스처는 1분 타임아웃 설정
        실제 대기하면 테스트가 느려지므로 Redis TTL 직접 조작
    """
    # 락 획득
    lock_manager.acquire_lock(
        project_code="TEST005",
        user_email="user@test.com",
        user_name="Test User",
        tab_id="tab123"
    )

    # Redis에서 TTL을 1초로 강제 설정
    lock_manager.redis.expire("lock:TEST005", 1)

    # 2초 대기
    time.sleep(2)

    # 락이 만료되어 조회 시 None이어야 함
    status = lock_manager.get_lock_status("TEST005")
    assert status is None

    # 다른 사용자가 락 획득 가능해야 함
    result = lock_manager.acquire_lock(
        project_code="TEST005",
        user_email="other@test.com",
        user_name="Other User",
        tab_id="tab456"
    )

    assert result['success'] is True


def test_lock_grace_period_recovery(lock_manager):
    """Grace Period 내 락 복구 테스트"""
    # 1. 락 획득
    lock_manager.acquire_lock(
        project_code="TEST006",
        user_email="user@test.com",
        user_name="Test User",
        tab_id="tab123"
    )

    # 2. 락 강제 만료 (TTL을 1초로 설정)
    lock_manager.redis.expire("lock:TEST006", 1)
    time.sleep(2)

    # 3. 락이 만료되었는지 확인
    status = lock_manager.get_lock_status("TEST006")
    assert status is None

    # 4. Grace Period 마커가 아직 존재하는지 확인 (grace TTL은 더 길게 설정됨)
    grace_info = lock_manager._get_grace_info("TEST006")
    assert grace_info is not None

    # 5. 원래 사용자가 재획득 시도 → 복구되어야 함
    result = lock_manager.acquire_lock(
        project_code="TEST006",
        user_email="user@test.com",
        user_name="Test User",
        tab_id="tab123"
    )

    assert result['success'] is True
    assert result.get('recovered') is True
    assert "복구" in result['message']


def test_lock_force_release(lock_manager):
    """관리자 강제 락 해제 테스트"""
    # User1이 락 획득
    lock_manager.acquire_lock(
        project_code="TEST007",
        user_email="user@test.com",
        user_name="Regular User",
        tab_id="tab123"
    )

    # Admin이 강제 해제
    result = lock_manager.force_release_lock(
        project_code="TEST007",
        admin_email="admin@test.com",
        admin_permission="admin"
    )

    assert result['success'] is True
    assert result['forced'] is True
    assert result['previous_owner'] == "Regular User"

    # 락이 해제되었는지 확인
    status = lock_manager.get_lock_status("TEST007")
    assert status is None


def test_lock_force_release_without_permission(lock_manager):
    """권한 없는 사용자의 강제 해제 시도 테스트"""
    # User1이 락 획득
    lock_manager.acquire_lock(
        project_code="TEST008",
        user_email="user@test.com",
        user_name="Regular User",
        tab_id="tab123"
    )

    # 권한 없는 사용자가 강제 해제 시도
    result = lock_manager.force_release_lock(
        project_code="TEST008",
        admin_email="hacker@test.com",
        admin_permission="user"
    )

    assert result['success'] is False
    assert "권한" in result['message']


def test_lock_get_all_locks(lock_manager):
    """전체 락 조회 테스트"""
    # 여러 락 획득
    lock_manager.acquire_lock("TEST009", "user1@test.com", "User 1", "tab1")
    lock_manager.acquire_lock("TEST010", "user2@test.com", "User 2", "tab2")
    lock_manager.acquire_lock("TEST011", "user3@test.com", "User 3", "tab3")

    # 전체 락 조회
    all_locks = lock_manager.get_all_locks()

    assert len(all_locks) == 3
    project_codes = [lock['project_code'] for lock in all_locks]
    assert "TEST009" in project_codes
    assert "TEST010" in project_codes
    assert "TEST011" in project_codes


def test_lock_get_user_locks(lock_manager):
    """특정 사용자 락 조회 테스트"""
    # User1이 여러 프로젝트 잠금
    lock_manager.acquire_lock("TEST012", "user1@test.com", "User 1", "tab1")
    lock_manager.acquire_lock("TEST013", "user1@test.com", "User 1", "tab2")
    lock_manager.acquire_lock("TEST014", "user2@test.com", "User 2", "tab3")

    # User1의 락만 조회
    user1_locks = lock_manager.get_user_locks("user1@test.com")

    assert len(user1_locks) == 2
    project_codes = [lock['project_code'] for lock in user1_locks]
    assert "TEST012" in project_codes
    assert "TEST013" in project_codes


def test_lock_release_all_user_locks(lock_manager):
    """사용자의 모든 락 해제 테스트"""
    # User1이 여러 프로젝트 잠금
    lock_manager.acquire_lock("TEST015", "user1@test.com", "User 1", "tab1")
    lock_manager.acquire_lock("TEST016", "user1@test.com", "User 1", "tab2")
    lock_manager.acquire_lock("TEST017", "user2@test.com", "User 2", "tab3")

    # User1의 모든 락 해제
    released_count = lock_manager.release_all_user_locks("user1@test.com", reason="logout")

    assert released_count == 2

    # User1의 락이 모두 해제되었는지 확인
    user1_locks = lock_manager.get_user_locks("user1@test.com")
    assert len(user1_locks) == 0

    # User2의 락은 유지되는지 확인
    user2_locks = lock_manager.get_user_locks("user2@test.com")
    assert len(user2_locks) == 1


# ==================== 동시성 테스트 ====================

def test_concurrent_lock_acquisition(lock_manager):
    """동시 락 획득 시도 테스트 (시뮬레이션)"""
    import threading

    results = []

    def acquire_lock(user_id):
        result = lock_manager.acquire_lock(
            project_code="TEST018",
            user_email=f"user{user_id}@test.com",
            user_name=f"User {user_id}",
            tab_id=f"tab{user_id}"
        )
        results.append((user_id, result['success']))

    # 3명의 사용자가 동시에 락 획득 시도
    threads = []
    for i in range(1, 4):
        thread = threading.Thread(target=acquire_lock, args=(i,))
        threads.append(thread)
        thread.start()

    # 모든 스레드 종료 대기
    for thread in threads:
        thread.join()

    # 결과 확인: 정확히 1명만 성공해야 함
    successful_locks = [r for r in results if r[1] is True]
    failed_locks = [r for r in results if r[1] is False]

    assert len(successful_locks) == 1
    assert len(failed_locks) == 2


# ==================== 에러 처리 테스트 ====================

def test_cache_service_unavailable_handling(cache_manager):
    """Redis 장애 시 ServiceUnavailable 예외 테스트"""
    # Redis 연결 끊기 시뮬레이션
    with patch.object(cache_manager.redis, 'get', side_effect=ServiceUnavailable("Redis down")):
        with pytest.raises(ServiceUnavailable):
            cache_manager.get("any_key")


def test_lock_service_unavailable_handling(lock_manager):
    """Redis 장애 시 락 획득 에러 처리 테스트"""
    # Redis 연결 끊기 시뮬레이션
    with patch.object(lock_manager.redis, 'get', side_effect=ServiceUnavailable("Redis down")):
        result = lock_manager.acquire_lock(
            project_code="TEST019",
            user_email="user@test.com",
            user_name="Test User",
            tab_id="tab123"
        )

        # 에러 메시지가 반환되어야 함
        assert result['success'] is False
        assert '서버 오류' in result['message']


if __name__ == "__main__":
    pytest.main([__file__, "-v", "-s"])
