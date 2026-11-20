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

# 슬랙 토큰·웹훅 제거 — 아래 '_load_redis_conn_env' 가 .env 를 안 읽는 것만 믿고 있었는데,
#   .env 가 이미 로드된 셸에서 pytest 를 돌리면 토큰이 살아 있어 실제로 나갔다.
#   2026-11-17: 단위 테스트(test_lead_project_link)가 경영지원 4명에게 가짜 '공사 취소' DM 12건 발송.
for _k in [k for k in os.environ if k.startswith(('SLACK_', 'SLACK'))
           and ('TOKEN' in k.upper() or 'WEBHOOK' in k.upper())]:
    os.environ.pop(_k, None)


# Redis 접속 정보만 .env 에서 — 2026-10-29 Redis 에 비밀번호를 건 뒤로 테스트(.env 안 읽음)가 비밀번호
#   없이 붙어 'Authentication required' 로 깨지고 연결 재시도에 60초씩 묶였음(2026-11-09 발견).
#   주소·포트·비밀번호 세 값만 읽음(슬랙 토큰 등 다른 값은 안 읽어 테스트가 밖으로 못 나감). DB 는 위 15 고정.
#   워크트리엔 .env 가 없어 본 저장소(git 공용 폴더의 부모)의 .env 를 찾음.
def _load_redis_conn_env():
    keys = ('REDIS_HOST', 'REDIS_PORT', 'REDIS_PASSWORD')
    here = os.path.dirname(os.path.abspath(__file__))
    cands = [os.path.join(here, '.env')]
    try:
        import subprocess
        common = subprocess.run(['git', 'rev-parse', '--git-common-dir'], cwd=here,
                                capture_output=True, text=True, timeout=5).stdout.strip()
        if common:
            cands.append(os.path.join(os.path.dirname(os.path.abspath(os.path.join(here, common))), '.env'))
    except Exception:
        pass
    for path in cands:
        if not os.path.isfile(path):
            continue
        with open(path, encoding='utf-8') as fh:
            for line in fh:
                k, sep, v = line.strip().partition('=')
                k = k.strip()
                if sep and k in keys and not os.environ.get(k):
                    os.environ[k] = v.split(' #')[0].strip().strip('"').strip("'")
        return


_load_redis_conn_env()
# 테스트 로그는 임시 폴더로 — 운영 로그(dashboard/logs)에 가짜 발송·장애 기록이 섞여
#   진단을 오도하던 문제 방지 (2026-09-15·10-06: 테스트 기록을 실제 race/Redis 순단으로 오판).
import tempfile as _tempfile
os.environ['LOG_DIR'] = os.environ.get(
    'TEST_LOG_DIR', os.path.join(_tempfile.gettempdir(), 'itg_pytest_logs')
)

# 사용자 DB(users.db: 감사 로그·전자결재·인사 등)도 임시 파일로 — get_user_database()/get_audit_repository()
#   싱글톤이 기본 경로(instance/users.db = 운영 DB)를 잡아, 전자결재 테스트의 _audit 이 운영 audit_logs 에
#   가짜 'EAPPROVAL_*'(강원포장 533,500원) 9,815건을 남겼다(2026-10-30~11-20). UserDatabase 도 테스트 중
#   …/instance/users.db 연결을 거부한다(user_database._refuse_prod_db_in_tests).
os.environ['USER_DB_PATH'] = os.path.join(_tempfile.mkdtemp(prefix='itg_pytest_userdb_'), 'users.db')
_PROD_USER_DB = os.path.normcase(os.path.abspath(
    os.path.join(os.path.dirname(__file__), 'instance', 'users.db')))
assert os.path.normcase(os.path.abspath(os.environ['USER_DB_PATH'])) != _PROD_USER_DB, (
    '테스트 사용자 DB 가 운영 instance/users.db 를 가리킴.'
)

# 방어: 혹시라도 테스트 DB 가 0 이면 세션 자체를 중단 (import 실패로 전량 차단)
assert os.environ['REDIS_DB'] != '0', (
    'REDIS_DB=0(프로덕션)에서 테스트 금지. REDIS_TEST_DB 로 별도 DB(예: 15) 지정.'
)
# 방어: 테스트 로그 폴더가 운영 로그 폴더면 중단
_PROD_LOG_DIR = os.path.normcase(os.path.abspath(
    os.path.join(os.path.dirname(__file__), 'dashboard', 'logs')))
assert os.path.normcase(os.path.abspath(os.environ['LOG_DIR'])) != _PROD_LOG_DIR, (
    '테스트 로그가 운영 로그 폴더(dashboard/logs)를 가리킴 — TEST_LOG_DIR 를 다른 폴더로 지정.'
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


# ── 공사 현황 옛 45열 배치 (2026-10-22 거래처명 삽입 전) ───────────────────
@pytest.fixture
def sheet_layout_45(monkeypatch):
    """테스트 데이터를 옛 열 위치(계약금=U, 버전=AS 등)로 짠 테스트용 — 기본 배치를 45열로 고정.

    코드는 헤더 지도를 따라가므로 어떤 배치에서도 같은 논리를 검사한다. pytest 중 헤더 지도 =
    _STATIC_COLUMN_MAPPING 이라 이걸 45열(거래처명 빼고)로 바꾼다.
    """
    import dashboard.utils.google_sheets as gs
    fields = [f for _, f in sorted(gs._STATIC_COLUMN_MAPPING.items(), key=lambda kv: gs.column_index(kv[0]))
              if f != '거래처명']
    monkeypatch.setattr(gs, '_STATIC_COLUMN_MAPPING', {gs.column_letter(i): f for i, f in enumerate(fields)})


# ── 실제 슬랙 호출 차단 (2026-11-17 가짜 '공사 취소' DM 12건) ─────────
@pytest.fixture(autouse=True, scope='session')
def _block_real_slack(request):
    """테스트가 진짜 슬랙 API 를 부르면 거부한다 — 토큰 제거(위)가 뚫려도 막는 둘째 겹.

    계기: test_lead_project_link 가 apply_project_cancel_to_slack 를 실제로 부르는데
    _project_client 만 가짜로 막고 _dm_client(메인봇 토큰)는 안 막아, 경영지원 4명에게
    가짜 취소 DM 이 나갔다. 테스트는 가짜 클라이언트를 써야 하고, 진짜 호출은 버그다.
    통합 테스트(--run-integration)는 실제 호출이 목적이라 걸지 않는다.
    """
    if request.config.getoption('--run-integration'):
        yield
        return
    try:
        from slack_sdk.web.base_client import BaseClient
    except Exception:
        yield
        return
    _orig = BaseClient.api_call

    def _blocked(self, api_method, *a, **kw):
        raise RuntimeError(
            f'테스트에서 실제 슬랙 호출 금지 ({api_method}) — 가짜 클라이언트를 쓰세요.')

    BaseClient.api_call = _blocked
    try:
        yield
    finally:
        BaseClient.api_call = _orig


# ── 세션 tripwire: 사용자 DB 싱글톤이 운영 users.db 에 묶이지 않았는지 검증 ───
@pytest.fixture(autouse=True, scope='session')
def _assert_user_db_isolated():
    """핵심 차단은 UserDatabase 의 연결 거부 — 이건 그게 뚫렸는지 보는 둘째 겹(종료 시 검사)."""
    yield
    try:
        import dashboard.utils.user_database as udb
    except ImportError:
        return
    inst = udb._user_db_instance
    if inst is not None:
        assert os.path.normcase(os.path.abspath(inst.db_path)) != _PROD_USER_DB, (
            f'사용자 DB 싱글톤이 운영 DB 에 묶임: {inst.db_path}')


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
