# Redis 통합 테스트 가이드

## 사전 요구사항

이 테스트를 실행하려면 **Redis 서버가 로컬에서 실행 중이어야 합니다**.

## Redis 설치 및 실행

### Windows 환경

#### 방법 1: Memurai (Redis for Windows)
1. [Memurai 다운로드](https://www.memurai.com/get-memurai) (무료)
2. 설치 후 자동으로 서비스 시작
3. 기본 포트: 6379

#### 방법 2: WSL2 + Redis
```bash
# WSL2에서 실행
sudo apt update
sudo apt install redis-server
sudo service redis-server start
```

#### 방법 3: Docker
```bash
docker run -d -p 6379:6379 redis:7-alpine
```

### 연결 확인
```bash
# Redis CLI로 연결 테스트
redis-cli ping
# 응답: PONG
```

## 테스트 실행

### 전체 테스트 실행
```bash
pytest tests/integration/test_redis_integration.py -v
```

### 특정 테스트만 실행
```bash
# 캐시 테스트만
pytest tests/integration/test_redis_integration.py -k "cache" -v

# 락 테스트만
pytest tests/integration/test_redis_integration.py -k "lock" -v
```

### Fail Fast 테스트 제외 (CI 환경)
```bash
pytest tests/integration/test_redis_integration.py -v \
  -k "not test_redis_connection_failure_on_init"
```

## 테스트 시나리오

### 1. 캐시 매니저 테스트
- ✅ `test_cache_basic_get_set`: 기본 get/set 동작
- ✅ `test_cache_ttl_expiry`: TTL 자동 만료
- ✅ `test_cache_delete`: 캐시 삭제
- ✅ `test_cache_invalidate_by_pattern`: 패턴 기반 무효화
- ✅ `test_cache_race_condition_prevention`: 레이스 컨디션 방지

### 2. 락 매니저 테스트
- ✅ `test_lock_acquire_release`: 락 획득/해제
- ✅ `test_lock_ttl_expiry`: 락 TTL 자동 만료
- ✅ `test_lock_grace_period_recovery`: Grace Period 복구
- ✅ `test_concurrent_lock_acquisition`: 동시 락 획득
- ✅ `test_lock_force_release`: 관리자 강제 해제

### 3. 에러 처리 테스트
- ✅ `test_redis_connection_failure_on_init`: Fail Fast (sys.exit)
- ✅ `test_cache_service_unavailable_handling`: 장애 시 예외 처리
- ✅ `test_lock_service_unavailable_handling`: 락 장애 처리

## 주의사항

### Fail Fast 테스트
`test_redis_connection_failure_on_init`는 실제로 `sys.exit(1)`을 호출하므로:
- CI 환경에서는 skip 권장
- 수동 테스트 시에만 실행

### 테스트 간 Redis 정리
- 각 테스트는 `flushdb()`로 Redis를 정리합니다
- 프로덕션 Redis에서 테스트하지 마세요!
- 테스트용 Redis DB 번호 사용 (예: `REDIS_DB=15`)

### 환경 변수
테스트 전 `.env` 파일 확인:
```bash
REDIS_HOST=localhost
REDIS_PORT=6379
REDIS_DB=15           # 테스트용 DB
REDIS_PASSWORD=       # 필요 시
```

## 트러블슈팅

### Error: "Connection refused"
- Redis 서버가 실행 중인지 확인
- 포트가 올바른지 확인 (기본 6379)

### Error: "NOAUTH Authentication required"
- `.env`에 `REDIS_PASSWORD` 설정

### Error: "SystemExit: 1"
- Fail Fast 동작 (정상)
- Redis 연결 실패 시 의도적으로 프로세스 종료

## 성능 벤치마크

예상 실행 시간:
- 전체 테스트: ~30초 (TTL 대기 포함)
- 캐시 테스트: ~10초
- 락 테스트: ~20초

## CI/CD 통합

GitHub Actions 예시:
```yaml
jobs:
  test:
    runs-on: ubuntu-latest
    services:
      redis:
        image: redis:7-alpine
        ports:
          - 6379:6379
    steps:
      - name: Run Redis tests
        run: pytest tests/integration/test_redis_integration.py -v
```
