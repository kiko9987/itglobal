# 🛡️ 보안 및 성능 업그레이드 가이드

## 📋 개요

여러 영업 사원이 동시에 사용하는 환경을 위한 **데이터 동시성, 권한 관리, 실시간 동기화** 강화 시스템입니다.

---

- **구현 (예시)**: utils/enhanced_permission_manager.py *(현재 저장소에는 포함되지 않음)*

## 🚨 해결된 문제점

### 1. **데이터 동시성 문제** ✅
- **문제**: 여러 사용자 동시 편집 시 데이터 손실
- **해결**: Redis 기반 분산 잠금 시스템
- **구현 (예시)**: utils/distributed_lock_manager.py *(현재 저장소에는 포함되지 않음)*

### 2. **권한 시스템 취약점** ✅
- **문제**: 필드별 세분화된 권한 부재
- **해결**: 역할 기반 접근 제어 (RBAC)
- **구현 (예시)**: utils/enhanced_permission_manager.py *(현재 저장소에는 포함되지 않음)*

### 3. **보안 취약점** ✅
- **문제**: CSRF, XSS, SQL Injection 위험
- **해결**: 포괄적 보안 미들웨어
- **구현**: `utils/security_middleware.py`

### 4. **실시간 동기화 불안정** ✅
- **문제**: 네트워크 끊김 시 작업 손실
- **해결**: 오프라인 지원 및 자동 복구
- **구현**: `static/js/realtime_sync_manager.js`

---

## 🔧 설치 및 설정

### 1. **Redis 설치 (Windows)**

```bash
# Chocolatey 사용
choco install redis-64

# 또는 수동 설치
# https://redis.io/docs/getting-started/installation/install-redis-on-windows/
```

### 2. **Python 패키지 설치**

```bash
pip install redis
```

### 3. **환경 변수 설정**

`.env` 파일에 추가:
```env
# Redis 설정
REDIS_HOST=localhost
REDIS_PORT=6379
REDIS_DB=0

# 보안 설정
SECURITY_ENABLED=true
RATE_LIMIT_ENABLED=true
```

### 4. **Redis 서비스 시작**

```bash
# Windows 서비스로 시작
redis-server

# 또는 백그라운드로 실행
redis-server --daemonize yes
```

---

## 🎯 새로운 기능

### 1. **분산 필드 잠금**

```python
# 사용 예시
from utils.distributed_lock_manager import acquire_field_lock

# 필드 편집 시 자동 잠금
with acquire_field_lock('S2024-KIM', '계약금액', 'user123'):
    # 안전한 데이터 수정
    update_project_field(...)
```

### 2. **세분화된 권한 시스템**

```python
# 사용 예시
from utils.enhanced_permission_manager import require_permission, Permission

@require_permission(Permission.FINANCIAL_EDIT)
def update_financial_data():
    # 금액 수정 권한이 있는 사용자만 접근 가능
    pass
```

### 3. **보안 미들웨어**

```python
# 자동 적용되는 보안 기능
@rate_limit(limit=10, window=300)  # 5분에 10회 제한
@validate_input(
    ('projectCode', 'project_code'),
    ('amount', 'safe_string')
)
def secure_api():
    pass
```

### 4. **실시간 동기화**

```javascript
// 오프라인 작업 지원
window.realtimeSyncManager.queueOfflineOperation({
    type: 'update_project',
    data: { projectCode: 'S2024-KIM', fieldName: '계약금액', newValue: '1200만원' }
});
```

---

## 📊 권한 매트릭스

| 역할 | 기본정보 | 공사정보 | 금액조회 | 금액편집 | 수금편집 | 손익편집 | 관리기능 |
|------|----------|----------|----------|----------|----------|----------|----------|
| **관리자** | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |
| **지역관리자** | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ❌ |
| **수석영업** | ✅ | ✅ | ✅ | ❌ | ❌ | ✅ | ❌ |
| **일반영업** | ✅ | ✅ | ✅ | ❌ | ❌ | ❌ | ❌ |
| **재무담당** | ❌ | ❌ | ✅ | ✅ | ✅ | ✅ | ❌ |
| **조회자** | ✅ | ✅ | ❌ | ❌ | ❌ | ❌ | ❌ |

---

## 🔒 보안 강화 사항

### 1. **입력값 검증**
- XSS, SQL Injection 패턴 자동 차단
- 프로젝트 코드 형식 검증 (`A1234-KIM`)
- 이메일, 전화번호 형식 검증

### 2. **Rate Limiting**
- API 호출 제한 (기본: 1시간 1000회)
- 과도한 요청 시 IP 차단
- 브루트 포스 공격 방지

### 3. **CSRF 보호**
- 모든 상태 변경 요청에 토큰 검증
- 토큰 만료 시간 설정 (기본: 1시간)
- 세션 기반 토큰 생성

---

## 📈 성능 향상

### 1. **스마트 캐싱**
- 구글 시트 데이터: 30초 TTL
- 폴더 매핑 정보: 1일 TTL
- 브라우저 캐시 최적화

### 2. **실시간 업데이트**
- WebSocket 기반 즉시 반영
- 변경분만 전송 (Delta Sync)
- 오프라인 큐 시스템

---

## 🛠️ 모니터링

### 1. **캐시 모니터링**
- 접속: `/admin/cache-monitor`
- 실시간 히트율 확인
- 메모리 사용량 추적

### 2. **보안 이벤트 로그**

```python
# 보안 통계 확인
security_stats = security_middleware.get_security_stats()
print(f"차단된 IP: {security_stats['blocked_ips']}개")
```

### 3. **필드 잠금 현황**

```python
# 현재 잠금 상태 확인
locks = distributed_lock_manager.get_all_locks()
for lock_key, lock_info in locks.items():
    print(f"{lock_key}: {lock_info['user_id']}가 편집 중")
```

---

## 🧪 테스트 시나리오

### 1. **동시성 테스트**
```bash
# 두 명이 동시에 같은 필드 수정 시도
curl -X POST /api/inline-update -d '{"projectCode":"S2024-KIM","fieldName":"계약금액","newValue":"1200"}' &
curl -X POST /api/inline-update -d '{"projectCode":"S2024-KIM","fieldName":"계약금액","newValue":"1300"}' &
```

### 2. **권한 테스트**
```javascript
// 일반 영업사원으로 금액 수정 시도
fetch('/api/inline-update', {
    method: 'POST',
    body: JSON.stringify({
        projectCode: 'S2024-KIM',
        fieldName: '계약금액',
        newValue: '1500만원'
    })
});
// 예상 결과: 403 Forbidden
```

### 3. **오프라인 테스트**
```javascript
// 네트워크 끊기 시뮬레이션
navigator.serviceWorker.postMessage('SIMULATE_OFFLINE');
// 작업 수행 후 온라인 복구 시 자동 동기화 확인
```

---

## 🚀 향후 개선 계획

### Phase 1 (완료) ✅
- [x] 분산 잠금 시스템
- [x] 권한 관리 강화
- [x] 보안 미들웨어
- [x] 실시간 동기화

### Phase 2 (계획)
- [ ] 가상 스크롤링 (대용량 데이터)
- [ ] 증분 동기화
- [ ] 모바일 최적화
- [ ] PWA 지원

### Phase 3 (장기)
- [ ] AI 기반 이상 감지
- [ ] 마이크로서비스 전환
- [ ] 고가용성 구성
- [ ] 자동 확장 시스템

---

## 📞 문제 해결

### 1. **Redis 연결 실패**
```bash
# Redis 서비스 상태 확인
redis-cli ping
# 응답: PONG (정상)

# 포트 확인
netstat -an | findstr :6379
```

### 2. **권한 오류**
```python
# 사용자 권한 확인
user_permission = enhanced_permission_manager.get_user_permission('user123')
print(user_permission.permissions)
```

### 3. **캐시 문제**
```python
# 캐시 상태 확인
cache_stats = smart_cache.get_cache_info()
print(f"히트율: {cache_stats['hit_rate']}%")
```

---

## 📚 참고 자료

- [Redis 공식 문서](https://redis.io/documentation)
- [Flask-SocketIO 가이드](https://flask-socketio.readthedocs.io/)
- [OWASP 보안 가이드](https://owasp.org/www-project-top-ten/)
- [웹 보안 체크리스트](https://github.com/shieldfy/API-Security-Checklist)

---

**⚠️ 주의사항**
- Redis 서버는 반드시 방화벽으로 보호
- 관리자 권한은 최소한의 인원에게만 부여
- 정기적인 보안 감사 실시
- 백업 시스템 구축 필수
