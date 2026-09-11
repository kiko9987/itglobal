"""프로젝트 루트 conftest — 전체 테스트 스위트 안전 격리.

목적
----
어떤 테스트도 프로덕션 Redis(db=0)·외부 API 에 닿지 못하게, 앱 모듈이 import 되기
'전'(= 이 파일 최상단)에 테스트 환경을 강제한다. 2026-07-24 flushdb 로 운영 1238키가
날아간 사고의 근본 재발 방지 + "파일 단위로만 돌린다"는 제약을 안전하게 해제한다.

방어 계층
--------
  1. import 시점 env 강제: REDIS_DB=15(테스트 DB), FLASK_ENV=testing, 알림 off
  2. RedisClient.flushdb() 자체 가드(redis_client.py) — db 0 이면 거부(별도 커밋)
  3. 세션 tripwire: 앱 Redis 싱글톤이 실제로 db≠0 에 묶였는지 종료 시 검증
  4. --run-integration 없으면 통합 테스트(프로덕션 리소스 접근) skip
"""
import os

# ── 앱 import 전에 테스트 환경 강제 (이 블록이 최우선 실행돼야 함) ──────
os.environ['FLASK_ENV'] = 'testing'
# 프로덕션 Redis(db 0) 절대 접근 금지 → 별도 테스트 DB (통합 fixture 와 동일 관례)
os.environ['REDIS_DB'] = os.environ.get('REDIS_TEST_DB', '15')
# 테스트 중 외부 발신 차단 (에러 슬랙 알림 등)
os.environ.setdefault('ERROR_SLACK_ALERTS_ENABLED', '0')

# 방어: 혹시라도 테스트 DB 가 0 이면 세션 자체를 중단 (import 실패로 전량 차단)
assert os.environ['REDIS_DB'] != '0', (
    'REDIS_DB=0(프로덕션)에서 테스트 금지. REDIS_TEST_DB 로 별도 DB(예: 15) 지정.'
)

import pytest


# ── Phase 3: 통합 테스트 opt-in (--run-integration) ───────────────────
def pytest_addoption(parser):
    parser.addoption(
        '--run-integration', action='store_true', default=False,
        help='통합 테스트(프로덕션 리소스 접근) 실행. 미지정 시 integration 마킹/경로 skip.',
    )


def pytest_collection_modifyitems(config, items):
    if config.getoption('--run-integration'):
        return
    skip_integration = pytest.mark.skip(reason='통합 테스트는 --run-integration 필요')
    for item in items:
        path = str(item.fspath).replace('\\', '/')
        if 'integration' in item.keywords or '/tests/integration/' in path:
            item.add_marker(skip_integration)


# ── 세션 tripwire: 앱 Redis 싱글톤이 비-프로덕션 DB 에 묶였는지 검증 ───
@pytest.fixture(autouse=True, scope='session')
def _assert_redis_isolated():
    # 검증은 teardown 시점 — 테스트들이 싱글톤을 만든 뒤 확인 (미가동 시 오탐 방지)
    yield
    try:
        from dashboard.utils.redis_client import RedisClient
        inst = getattr(RedisClient, '_instance', None)
        if inst is not None:
            db = getattr(inst, '_db', None)
            if db is not None:
                assert int(db) != 0, (
                    f'테스트가 프로덕션 Redis(db={db})에 연결됨 — 격리 실패.'
                )
    except AssertionError:
        raise
    except Exception:
        pass  # Redis 미가동 등은 개별 테스트에서 처리
