# Redis 마이그레이션 계획

## 📋 목차
1. [배경](#배경)
2. [기술적 의사결정](#기술적-의사결정)
3. [전문가 피드백 요약](#전문가-피드백-요약)
4. [구현 범위](#구현-범위)
5. [구현 플랜 (3주)](#구현-플랜-3주)
6. [테스트 전략](#테스트-전략)
7. [배포 및 모니터링](#배포-및-모니터링)
8. [확장 계획](#확장-계획)

---

## 배경

### 현재 상황
- **환경**: 사내 PC 1대, 내부 20명 사용
- **문제**: 동시 편집 충돌 가능성
- **현재 구조**: 메모리 기반 lock/cache (workers=1 전용)

### 모바일 환경 대비 필요성
- **추후 계획**: 모바일 앱 출시 예정
- **요구사항**:
  - 외부 접속 지원
  - 다중 프로세스/서버 확장
  - 고가용성 보장

### Option 비교

| 항목 | Option A (workers=1) | Redis 로컬 |
|------|---------------------|-----------|
| 구현 기간 | 1주 | 3주 |
| 동시성 | 단일 프로세스만 | 다중 프로세스 가능 |
| 확장성 | 제한적 | 우수 |
| 총 작업량 | 5주 (나중 전환 포함) | 3주 |
| 모바일 대응 | 재작업 필요 (3주) | 설정만 변경 (1일) |

**결론: Redis 선택**
- 총 작업량 2주 절약
- 마이그레이션 리스크 제거
- 모바일 확장 시 설정만 변경

---

## 기술적 의사결정

### 1. Redis 로컬 설치
```bash
# Windows
choco install redis-64

# Linux
sudo apt install redis-server
```

**선택 이유:**
- 비용 0원 (오픈소스)
- 메모리 기반 (빠른 성능)
- Database lock 없음
- 확장성 우수

### 2. Fail Fast 에러 처리
```python
# 부팅 시: Redis 없으면 서비스 시작 불가
if not redis.ping():
    sys.exit(1)

# 운영 중: 503 Service Unavailable 반환
except RedisError:
    return jsonify({'error': 'Service unavailable'}), 503
```

**선택 이유:**
- "Redis 없이는 서비스 불가" 명확
- Graceful Degradation은 복잡도만 증가
- 빠른 장애 감지 가능

### 3. 재시작 시 초기화 (마이그레이션 불필요)
```python
# 캐시: 재시작 후 자동으로 다시 채워짐
# 락: 재시작 시 모두 해제 (정상 동작)
```

**선택 이유:**
- 마이그레이션 복잡도 제거
- 다운타임 허용 가능 (내부 환경)
- 구현 단순화

### 4. Workers 점진적 증가
```
배포: workers=1 (기존 유지)
1주 후: workers=2 (모니터링)
2주 후: workers=4 (조건부)
```

**선택 이유:**
- 로컬 PC: CPU 제한
- 과도한 워커 = 컨텍스트 스위칭 비용
- 실제 부하 확인 후 조정

---

## 전문가 피드백 요약

### ✅ 검증된 사항
1. **수정 필요 모듈**: 4개만 수정
   - `project_lock_manager.py`
   - `smart_cache_manager.py`
   - `background_prefetch.py`
   - `calendar_sync_scheduler.py`

2. **호출부 수정 불필요**: 10개 파일
   - 인터페이스 유지로 호출부 영향 없음

3. **Optimistic Lock 유지**: 이미 구현됨
   - Redis와 독립적으로 작동
   - 추가 수정 불필요

### ⚠️ 보완 사항
1. **Gunicorn 설정**: workers 점진적 증가
2. **환경 변수**: Redis 연결 설정
3. **에러 처리**: Fail Fast 전략
4. **테스트**: 자동화 3가지 시나리오

---

## 구현 범위

### 수정 필요 파일 (6개)

#### 1. 신규 생성
```
✅ dashboard/utils/redis_client.py         [Redis 연결 헬퍼]
```

#### 2. 필수 수정
```
✅ dashboard/utils/smart_cache_manager.py  [Dict → Redis]
✅ dashboard/utils/project_lock_manager.py [Dict → Redis]
✅ gunicorn.conf.py                        [workers 설정]
```

#### 3. 설정 파일
```
✅ .env                                    [Redis 환경변수]
✅ dashboard/config.py                     [Redis 설정 클래스]
```

### 수정 불필요 파일 (검증 완료)

**캐시 호출부 (10개):**
- `projects.py` - 인터페이스 유지
- `project_service.py` - 인터페이스 유지
- `background_prefetch.py` - 인터페이스 유지
- `metadata.py` - 인터페이스 유지
- `lead_service.py` - 인터페이스 유지
- `admin.py` - 인터페이스 유지
- `folders.py` - 인터페이스 유지
- `data_management.py` - 인터페이스 유지
- `analytics.py` - 인터페이스 유지
- `calendar_sync_scheduler.py` - 인터페이스 유지

**락 호출부 (3개):**
- `locks.py` - 인터페이스 유지
- `users.py` - 인터페이스 유지

**프론트엔드:**
- 모든 JS 파일 - API 변경 없음

---

## 구현 플랜 (3주)

### Week 1: Redis 설치 및 헬퍼 작성

#### Day 1-2: Redis 설치 및 환경 설정
```bash
# Windows
choco install redis-64
redis-server --service-install
redis-server --service-start

# 기본 설정 (redis.conf)
maxmemory 512mb
maxmemory-policy allkeys-lru
```

**환경 변수 설정:**
```bash
# .env
REDIS_HOST=localhost
REDIS_PORT=6379
REDIS_DB=0
REDIS_PASSWORD=  # optional
```

#### Day 3-5: redis_client.py 작성
```python
# dashboard/utils/redis_client.py
import redis
import json
import logging
from typing import Any, Optional

logger = logging.getLogger(__name__)

class RedisClient:
    """Redis 연결 및 작업 헬퍼"""

    def __init__(self):
        try:
            self.redis = redis.Redis(
                host=os.getenv('REDIS_HOST', 'localhost'),
                port=int(os.getenv('REDIS_PORT', 6379)),
                db=int(os.getenv('REDIS_DB', 0)),
                password=os.getenv('REDIS_PASSWORD', None),
                decode_responses=True,
                socket_connect_timeout=5,
                socket_timeout=5
            )
            # Fail Fast: 부팅 시 연결 확인
            self.redis.ping()
            logger.info("Redis 연결 성공")
        except redis.RedisError as e:
            logger.critical(f"Redis 연결 실패: {e}")
            sys.exit(1)

    def get(self, key: str) -> Optional[Any]:
        """값 조회"""
        try:
            value = self.redis.get(key)
            return json.loads(value) if value else None
        except redis.RedisError as e:
            logger.error(f"Redis get 실패: {key}, {e}")
            raise ServiceUnavailable("Cache unavailable")

    def set(self, key: str, value: Any, ex: int = None) -> bool:
        """값 저장 (TTL 옵션)"""
        try:
            return self.redis.set(key, json.dumps(value), ex=ex)
        except redis.RedisError as e:
            logger.error(f"Redis set 실패: {key}, {e}")
            raise ServiceUnavailable("Cache unavailable")

    def delete(self, key: str) -> int:
        """값 삭제"""
        try:
            return self.redis.delete(key)
        except redis.RedisError as e:
            logger.error(f"Redis delete 실패: {key}, {e}")
            raise ServiceUnavailable("Cache unavailable")

    def acquire_lock(self, key: str, timeout: int = 300) -> bool:
        """분산 락 획득 (SETNX + TTL)"""
        try:
            return self.redis.set(key, "locked", nx=True, ex=timeout)
        except redis.RedisError as e:
            logger.error(f"Redis 락 획득 실패: {key}, {e}")
            raise ServiceUnavailable("Lock unavailable")

    def release_lock(self, key: str) -> int:
        """분산 락 해제"""
        return self.delete(key)

    def ping(self) -> bool:
        """연결 확인"""
        try:
            return self.redis.ping()
        except redis.RedisError:
            return False
```

---

### Week 2: 캐시/락 외부화

#### Day 1-3: smart_cache_manager.py 수정
```python
# dashboard/utils/smart_cache_manager.py (수정 부분만)

from dashboard.utils.redis_client import RedisClient

class SimpleCache:
    """Redis 기반 TTL 캐시"""

    def __init__(self):
        self.redis = RedisClient()
        self._lock = threading.RLock()  # 로컬 동기화용

    def get(self, key: str, strategy: CacheStrategy = CacheStrategy.CRITICAL_DATA) -> Optional[Any]:
        """캐시에서 값 가져오기"""
        try:
            # Redis에서 조회 (TTL은 Redis가 자동 관리)
            value = self.redis.get(f"cache:{key}")
            if value:
                logger.debug(f"캐시 히트: {key}")
                return value
            logger.debug(f"캐시 미스: {key}")
            return None
        except ServiceUnavailable:
            # Fail Fast: Redis 장애 시 503 반환
            raise

    def set(self, key: str, value: Any, strategy: CacheStrategy = CacheStrategy.CRITICAL_DATA,
            fetched_at: Optional[float] = None) -> None:
        """캐시에 값 저장"""
        ttl = self.TTL_MAP.get(strategy, 300)

        # 무효화 마커 확인 (레이스 컨디션 방지)
        if fetched_at and self._is_invalidated(key, fetched_at):
            logger.warning(f"무효화 마커로 인해 캐시 저장 차단: {key}")
            return

        try:
            self.redis.set(f"cache:{key}", value, ex=ttl)
            logger.debug(f"캐시 저장: {key} (TTL: {ttl}초)")
        except ServiceUnavailable:
            raise

    def delete(self, key: str, set_marker: bool = True) -> None:
        """캐시 삭제"""
        try:
            self.redis.delete(f"cache:{key}")

            # 무효화 마커 설정 (레이스 컨디션 방지)
            if set_marker:
                marker_key = f"invalidation:{key}"
                self.redis.set(marker_key, time.time(), ex=10)
                logger.debug(f"캐시 삭제 + 무효화 마커 설정: {key}")
        except ServiceUnavailable:
            raise
```

#### Day 4-5: project_lock_manager.py 수정
```python
# dashboard/utils/project_lock_manager.py (수정 부분만)

from dashboard.utils.redis_client import RedisClient

class ProjectLockManager:
    """Redis 기반 프로젝트 잠금 관리자"""

    def __init__(self, lock_timeout_minutes=None):
        if lock_timeout_minutes is None:
            lock_timeout_minutes = int(os.getenv('LOCK_TIMEOUT_MINUTES', 5))

        self.redis = RedisClient()
        self.lock_timeout_minutes = lock_timeout_minutes
        self.lock_timeout_seconds = lock_timeout_minutes * 60
        self._lock = threading.RLock()  # 로컬 동기화용

        logger.info(f"ProjectLockManager 초기화 완료 (타임아웃: {lock_timeout_minutes}분)")

    def acquire_lock(self, project_code: str, user_email: str, user_name: str, tab_id: str) -> Dict:
        """프로젝트 잠금 획득"""
        with self._lock:
            lock_key = f"lock:project:{project_code}"

            # 기존 락 확인
            existing_lock_data = self.redis.get(lock_key)

            if existing_lock_data:
                existing_lock = json.loads(existing_lock_data)

                # 같은 사용자/탭이면 연장
                if (existing_lock['user_email'] == user_email and
                    existing_lock['tab_id'] == tab_id):
                    lock_data = {
                        'project_code': project_code,
                        'user_email': user_email,
                        'user_name': user_name,
                        'tab_id': tab_id,
                        'locked_at': datetime.now().isoformat()
                    }
                    self.redis.set(lock_key, json.dumps(lock_data), ex=self.lock_timeout_seconds)
                    logger.info(f"락 연장: {project_code} by {user_email}")
                    return {
                        'success': True,
                        'message': '잠금이 연장되었습니다.',
                        'lock_info': lock_data
                    }
                else:
                    # 다른 사용자가 보유 중
                    return {
                        'success': False,
                        'message': f"{existing_lock['user_name']}님이 편집 중입니다.",
                        'lock_info': existing_lock
                    }

            # 새 락 획득
            lock_data = {
                'project_code': project_code,
                'user_email': user_email,
                'user_name': user_name,
                'tab_id': tab_id,
                'locked_at': datetime.now().isoformat()
            }

            success = self.redis.acquire_lock(lock_key, timeout=self.lock_timeout_seconds)

            if success:
                self.redis.set(lock_key, json.dumps(lock_data), ex=self.lock_timeout_seconds)
                logger.info(f"락 획득: {project_code} by {user_email}")
                return {
                    'success': True,
                    'message': '잠금을 획득했습니다.',
                    'lock_info': lock_data
                }
            else:
                return {
                    'success': False,
                    'message': '잠금 획득에 실패했습니다.'
                }

    def release_lock(self, project_code: str, user_email: str, tab_id: str) -> Dict:
        """프로젝트 잠금 해제"""
        with self._lock:
            lock_key = f"lock:project:{project_code}"
            existing_lock_data = self.redis.get(lock_key)

            if not existing_lock_data:
                return {
                    'success': True,
                    'message': '잠금이 이미 해제되었습니다.'
                }

            existing_lock = json.loads(existing_lock_data)

            # 소유자만 해제 가능
            if (existing_lock['user_email'] == user_email and
                existing_lock['tab_id'] == tab_id):
                self.redis.release_lock(lock_key)
                logger.info(f"락 해제: {project_code} by {user_email}")
                return {
                    'success': True,
                    'message': '잠금이 해제되었습니다.'
                }
            else:
                return {
                    'success': False,
                    'message': '다른 사용자의 잠금은 해제할 수 없습니다.'
                }
```

---

### Week 3: 테스트 및 배포

#### Day 1-2: 자동화 테스트 작성
```python
# tests/test_redis_integration.py
import pytest
import time
from dashboard.utils.redis_client import RedisClient
from dashboard.utils.project_lock_manager import ProjectLockManager

class TestRedisIntegration:
    """Redis 통합 테스트"""

    def test_redis_connection_failure(self, monkeypatch):
        """Redis 연결 실패 시 서비스 부팅 실패"""
        monkeypatch.setenv('REDIS_HOST', 'invalid_host')

        with pytest.raises(SystemExit):
            RedisClient()

    def test_lock_ttl_expiry(self):
        """락 TTL 만료 확인"""
        lock_manager = ProjectLockManager(lock_timeout_minutes=1)

        # 락 획득
        result = lock_manager.acquire_lock(
            "TEST001",
            "test@example.com",
            "Test User",
            "tab123"
        )
        assert result['success'] == True

        # 61초 대기 (TTL 초과)
        time.sleep(61)

        # 다른 사용자가 획득 가능해야 함
        result2 = lock_manager.acquire_lock(
            "TEST001",
            "other@example.com",
            "Other User",
            "tab456"
        )
        assert result2['success'] == True

    def test_concurrent_lock_acquisition(self):
        """동시 락 획득 경합"""
        from threading import Thread

        lock_manager = ProjectLockManager()
        results = []

        def acquire():
            result = lock_manager.acquire_lock(
                "TEST001",
                f"user{len(results)}@example.com",
                f"User {len(results)}",
                f"tab{len(results)}"
            )
            results.append(result)

        # 10개 스레드 동시 실행
        threads = [Thread(target=acquire) for _ in range(10)]
        for t in threads:
            t.start()
        for t in threads:
            t.join()

        # 정확히 하나만 성공해야 함
        success_count = sum(1 for r in results if r['success'])
        assert success_count == 1
```

#### Day 3-4: Manual QA
- Redis 재시작 중 요청 처리
- 메모리 부족 시나리오
- 동시 편집 충돌 (브라우저 2개)
- Optimistic Lock + Redis Lock 조합

#### Day 5: 배포
```bash
# 1. Redis 서비스 시작 확인
redis-cli ping

# 2. 환경 변수 설정 확인
cat .env

# 3. Gunicorn 시작 (workers=1 유지)
gunicorn -c gunicorn.conf.py dashboard.app:app

# 4. 헬스체크
curl http://localhost:8000/health
```

---

## 테스트 전략

### 1. 자동화 테스트 (Pytest)
```bash
# 설치
pip install pytest pytest-cov redis

# 실행
pytest tests/test_redis_integration.py -v
```

**커버리지:**
- Redis 연결 실패
- 락 TTL 만료
- 동시 락 경합

### 2. Manual QA 체크리스트
- [ ] Redis 재시작 중 요청 → 503 반환 확인
- [ ] 메모리 부족 (maxmemory 초과) → eviction 확인
- [ ] 동시 편집 (브라우저 2개) → 락 충돌 확인
- [ ] Optimistic Lock + Redis Lock → 409 + 락 메시지
- [ ] 캐시 TTL 만료 → 자동 갱신 확인

### 3. 성능 테스트
```bash
# Apache Bench
ab -n 1000 -c 10 http://localhost:8000/api/projects

# 목표:
# - 평균 응답시간 < 200ms
# - 에러율 < 0.1%
```

---

## 배포 및 모니터링

### 배포 절차
```bash
# 1. Redis 설치
choco install redis-64
redis-server --service-install

# 2. 환경 변수 설정
cp .env.example .env
# REDIS_HOST=localhost 확인

# 3. 의존성 설치
pip install redis

# 4. 서비스 재시작
gunicorn -c gunicorn.conf.py dashboard.app:app
```

### 모니터링
```bash
# Redis 상태 확인
redis-cli info memory
redis-cli info stats

# 주요 메트릭:
# - used_memory: 사용 중인 메모리
# - connected_clients: 연결된 클라이언트 수
# - total_commands_processed: 총 명령어 수
```

**알람 설정 (PowerShell 스크립트):**
```powershell
# check_redis.ps1
$ping = redis-cli ping
if ($ping -ne "PONG") {
    Write-Host "Redis is down!"
    # 이메일/Slack 알림
}

$memory = redis-cli info memory | Select-String "used_memory_human"
Write-Host $memory
```

**작업 스케줄러 등록:**
- 5분마다 실행
- 로그 파일 저장

---

## 확장 계획

### Phase 1: 점진적 Workers 증가

**Week 1 배포 후:**
```bash
# workers=1 유지
gunicorn -c gunicorn.conf.py
```

**Week 2: 모니터링**
```bash
# CPU 사용률 확인
# 응답 시간 확인

# 조건: CPU < 50% AND 응답시간 > 200ms
# → workers=2로 증가
```

**Week 3: 추가 모니터링**
```bash
# 조건: CPU < 50% AND 응답시간 > 200ms
# → workers=4로 증가

# 조건: CPU > 80%
# → workers 증가 중단 (컨텍스트 스위칭 비용)
```

### Phase 2: 모바일 환경 대응

**환경변수만 변경:**
```bash
# 로컬 (현재)
REDIS_HOST=localhost
REDIS_PORT=6379

# 클라우드 (미래)
REDIS_HOST=redis-xxxxx.aws.com
REDIS_PORT=6379
REDIS_PASSWORD=xxxxx
```

**코드 변경: 0줄**

**추가 고려사항:**
- HTTPS 인증서
- 방화벽 설정
- Load Balancer 구성
- Redis Sentinel/Cluster (고가용성)

---

## 참고 문서

### Redis 설치
- Windows: https://redis.io/docs/getting-started/installation/install-redis-on-windows/
- Linux: https://redis.io/docs/getting-started/installation/install-redis-on-linux/

### redis-py 문서
- https://redis-py.readthedocs.io/

### 모니터링
- Redis CLI: https://redis.io/docs/ui/cli/
- Redis Insight: https://redis.io/insight/ (GUI 도구)

---

## 작업 이력

### 2025-01-XX: 계획 수립
- Option A vs Redis 비교
- 전문가 피드백 3라운드
- 최종 Redis 선택

### 2025-01-XX: 문서 작성
- 구현 플랜 작성
- 테스트 전략 수립
- 배포 절차 정리

### 2025-01-XX: 구현 시작
- Redis 설치
- redis_client.py 작성
- 테스트 작성
