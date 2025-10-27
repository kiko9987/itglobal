# 리팩토링 백로그

프로젝트가 성장하면서 점진적으로 개선할 수 있는 항목들을 정리한 문서입니다.

## 우선순위 기준

- 🔴 **High**: 현재 병목이 되거나 버그 가능성이 높은 부분
- 🟡 **Medium**: 코드 품질 향상에 도움이 되는 부분
- 🟢 **Low**: 이론적으로 더 나은 방법이지만 현재는 문제없는 부분

---

## Phase 1: 검증 시스템 개선 (완료)

### ✅ 완료된 작업
- [x] Marshmallow 스키마 구축
- [x] API 엔드포인트 검증 적용
- [x] Google Sheets 에러 처리 개선
- [x] 사용자 친화적 에러 메시지

---

## Phase 2: 데이터 타입 최적화

### 🟢 Low: Marshmallow Date 필드 전환

**현재 상태:**
```python
# dashboard/schemas/project_schemas.py
start_date = fields.Str(allow_none=True)

@validates('start_date')
def validate_start_date(self, value):
    if value:
        try:
            datetime.strptime(value, '%Y-%m-%d')
        except ValueError:
            raise ValidationError("시작일 형식이 올바르지 않습니다 (YYYY-MM-DD)")
```

**제안된 개선:**
```python
start_date = fields.Date(
    format='%Y-%m-%d',
    allow_none=True,
    error_messages={
        "invalid": "시작일 형식이 올바르지 않습니다 (YYYY-MM-DD)"
    }
)
```

**장점:**
- 더 간결한 코드
- 타입 안정성 향상 (date 객체 반환)
- Marshmallow best practice

**트레이드오프:**
- Google Sheets는 문자열 형식 선호
- 현재 코드베이스 전체가 문자열로 날짜 처리
- API 응답에서 추가 직렬화 필요

**적용 시점:**
- Service Layer 분리 완료 후
- 비즈니스 로직과 데이터 전송 계층이 명확히 구분될 때

**관련 파일:**
- `dashboard/schemas/project_schemas.py`
- `dashboard/blueprints/projects.py`

---

## Phase 3: 재시도 로직 고도화

### 🟢 Low: Tenacity 데코레이터 활용

**현재 상태:**
```python
# dashboard/utils/google_sheets.py
def _execute_with_retry(self, request, operation_name):
    for attempt in range(self.MAX_RETRIES):
        try:
            result = request.execute()
            # ... 상세한 에러 처리 로직 ...
        except HttpError as e:
            # 수동 재시도 로직
```

**제안된 개선:**
```python
from tenacity import retry, stop_after_attempt, wait_exponential_jitter

def _should_retry_google_sheets_error(exception):
    """재시도 가능한 에러만 필터링"""
    if not isinstance(exception, HttpError):
        return False
    return exception.resp.status in [429, 500, 503]

@retry(
    stop=stop_after_attempt(5),
    wait=wait_exponential_jitter(initial=1, max=32),
    retry=retry_if_exception(_should_retry_google_sheets_error),
    before_sleep=before_sleep_log(logger, logging.WARNING),
    after=after_log(logger, logging.INFO)
)
def _execute_with_tenacity(self, request, operation_name):
    try:
        return request.execute()
    except HttpError as e:
        error_code = e.resp.status
        user_msg = self._get_user_friendly_error_message(error_code, e._get_reason())
        raise Exception(f"Google Sheets API 오류: {user_msg}") from e
```

**장점:**
- 더 선언적인 코드
- 재시도 정책 변경이 쉬움
- 표준 라이브러리 활용

**트레이드오프:**
- 현재 코드가 더 명확하고 디버깅 용이
- 세밀한 제어(에러 코드별 처리)가 복잡해질 수 있음
- 팀원들이 tenacity에 익숙하지 않을 수 있음

**적용 시점:**
- 재시도 로직이 더 복잡해질 때
- Circuit Breaker, Timeout 등 추가 정책 필요 시
- 팀원들의 tenacity 이해도가 높아질 때

**관련 파일:**
- `dashboard/utils/google_sheets.py`

---

## Phase 4: 아키텍처 개선 (장기 계획)

### 🟡 Medium: Service Layer 분리

**목표:**
- 비즈니스 로직을 Blueprint에서 분리
- 테스트 가능한 순수 Python 함수로 구성
- Google Sheets 의존성 격리

**구조 예시:**
```
dashboard/
├── blueprints/          # HTTP 요청/응답 처리만
│   └── projects.py
├── services/            # 비즈니스 로직 (NEW)
│   ├── project_service.py
│   ├── validation_service.py
│   └── sheets_sync_service.py
├── repositories/        # 데이터 접근 (NEW)
│   ├── sheets_repository.py
│   └── cache_repository.py
└── schemas/             # 데이터 검증
    └── project_schemas.py
```

**적용 시점:**
- 프로젝트가 10,000 LOC 이상으로 성장할 때
- 단위 테스트 도입이 필요할 때
- 여러 데이터 소스(Sheets + DB) 통합이 필요할 때

**예상 공수:** 2-3주

---

### 🟡 Medium: Task Queue 도입

**목표:**
- Google Sheets 동기화를 비동기로 처리
- 사용자 응답 속도 향상
- Rate Limit 완화

**기술 스택 후보:**
- Celery + Redis
- RQ (Redis Queue) - 더 간단한 대안
- APScheduler (현재 사용 중, 확장 가능)

**적용 시점:**
- Google Sheets API 호출이 병목이 될 때
- 대량 데이터 동기화가 필요할 때
- 백그라운드 작업(리포트 생성 등)이 증가할 때

**예상 공수:** 1-2주

---

### 🟢 Low: 프론트엔드 리팩토링

**목표:**
- React 컴포넌트화
- 수동 DOM 조작 제거
- TypeScript 도입

**적용 시점:**
- 프론트엔드 복잡도가 관리 불가능해질 때
- 새로운 프론트엔드 개발자 합류 시
- SPA(Single Page Application) 전환 필요 시

**예상 공수:** 4-6주

---

## 성능 최적화 백로그

### 측정이 우선

현재는 성능 문제가 없으므로, 아래 항목들은 **실제 병목이 측정된 후** 고려합니다.

#### 1. 데이터베이스 인덱싱
- [ ] 프로젝트 코드 인덱스
- [ ] 날짜 범위 쿼리 최적화

#### 2. 캐시 전략 개선
- [ ] Redis 도입 검토
- [ ] 캐시 무효화 정책 최적화

#### 3. API 응답 최적화
- [ ] Pagination 구현
- [ ] GraphQL 검토 (over-fetching 방지)

---

## 보안 강화 백로그

### 🔴 High: 입력 검증 확대

**현재 상태:**
- ✅ 프로젝트 생성/수정 검증 완료
- ❌ 다른 API 엔드포인트 검증 부족

**TODO:**
- [ ] `/api/projects/cancel` 검증
- [ ] `/api/projects/resume` 검증
- [ ] 모든 POST/PUT 엔드포인트 감사

**적용 시점:** 가능한 빨리

---

### 🟡 Medium: Rate Limiting

**목표:**
- 사용자별 API 호출 제한
- DDoS 방어
- 서비스 안정성 향상

**기술 스택:**
- Flask-Limiter
- Redis 백엔드

**적용 시점:**
- 외부 공개 API 제공 시
- 동시 사용자 수 증가 시

---

## 모니터링 개선 백로그

### 🟡 Medium: 구조화된 로깅

**목표:**
- JSON 로그 포맷
- 중앙 로그 수집 (ELK Stack / CloudWatch)
- 로그 기반 알림

**적용 시점:**
- 서비스 규모가 커질 때
- 운영 이슈 추적이 어려워질 때

---

### 🟢 Low: APM (Application Performance Monitoring)

**목표:**
- 성능 병목 자동 탐지
- 트랜잭션 추적
- 실시간 알림

**기술 스택 후보:**
- New Relic
- Datadog
- Sentry (에러 추적)

**적용 시점:**
- 프로덕션 트래픽 증가 시
- 운영 예산이 확보될 때

---

## 마치며

이 문서는 "언젠가 해야 할 일" 목록이 아닙니다.

**실제로 문제가 될 때, 측정된 데이터를 기반으로 우선순위를 재평가**하고 진행하세요.

> "Premature optimization is the root of all evil" - Donald Knuth

현재 코드는 매우 건강한 상태입니다. 무리한 최적화보다는 **비즈니스 가치 전달**에 집중하고, 병목이 생길 때 이 문서를 참고하세요.

---

**최종 업데이트:** 2025-10-12
**작성자:** Claude Code
