# API 표준 및 규칙

이 문서는 시스템 전체 API의 표준 규칙을 정의합니다.

---

## 📋 HTTP 상태 코드 규칙

### 2xx: 성공

| 코드 | 의미 | 사용 시점 | 예시 |
|------|------|----------|------|
| **200 OK** | 조회/수정 성공 | GET, PUT, PATCH 성공 | 프로젝트 목록 조회, 프로젝트 수정 |
| **201 Created** | 생성 성공 | POST로 리소스 생성 | 프로젝트 생성, 메모 생성 |
| **204 No Content** | 성공 (응답 본문 없음) | DELETE 성공 | 프로젝트 삭제, 메모 삭제 |
| **207 Multi-Status** | 부분 성공 | 배치 작업에서 일부 성공/실패 | 여러 메모 일괄 저장 |

### 4xx: 클라이언트 에러

| 코드 | 의미 | 사용 시점 | 예시 |
|------|------|----------|------|
| **400 Bad Request** | 잘못된 요청 | 입력 검증 실패, 형식 오류 | 필수 필드 누락, 이미 존재하는 리소스 |
| **401 Unauthorized** | 인증 필요 | 로그인 필요 | 세션 만료, 토큰 없음 |
| **403 Forbidden** | 권한 없음 | 인증은 되었지만 권한 부족 | viewer가 수정 시도 |
| **404 Not Found** | 리소스 없음 | 요청한 리소스가 존재하지 않음 | 존재하지 않는 프로젝트 조회 |

### 5xx: 서버 에러

| 코드 | 의미 | 사용 시점 | 예시 |
|------|------|----------|------|
| **500 Internal Server Error** | 내부 서버 오류 | 예상치 못한 서버 오류 | Python Exception, DB 오류 |
| **503 Service Unavailable** | 서비스 불가 | 외부 서비스 장애 | Google Sheets API 장애, Calendar API 장애 |

---

## 📦 표준 응답 구조

### 성공 응답

```json
{
  "success": true,
  "data": {
    // 실제 데이터
  },
  "message": "성공 메시지 (선택)",
  "meta": {
    "timestamp": "2025-10-29T12:34:56",
    "request_id": "a1b2c3d4",
    "total": 100  // 페이지네이션 시
  }
}
```

### 에러 응답

```json
{
  "success": false,
  "error": {
    "code": "VALIDATION_ERROR",
    "message": "사용자 친화적 에러 메시지",
    "details": {
      // 개발자용 상세 정보 (선택)
    }
  },
  "meta": {
    "timestamp": "2025-10-29T12:34:56",
    "request_id": "a1b2c3d4"
  }
}
```

### 부분 성공 응답 (207)

```json
{
  "success": true,
  "data": {
    "results": [
      {"item": 1, "success": true, "message": "저장 완료"},
      {"item": 2, "success": false, "message": "검증 실패"}
    ],
    "summary": {
      "total_count": 2,
      "success_count": 1,
      "failed_count": 1
    }
  },
  "message": "2개 중 1개 성공, 1개 실패",
  "meta": {
    "timestamp": "2025-10-29T12:34:56",
    "request_id": "a1b2c3d4"
  }
}
```

---

## 🔧 사용 방법

### 1. APIResponse 헬퍼 사용 (권장)

```python
from dashboard.api.responses import APIResponse, APIErrorCode

# 성공 응답 (200)
@app.route('/api/projects/list')
def get_projects():
    projects = get_all_projects()
    return APIResponse.success(
        data=projects,
        meta={'total': len(projects)}
    )

# 생성 성공 (201)
@app.route('/api/projects', methods=['POST'])
def create_project():
    project = create_new_project(request.json)
    return APIResponse.created(
        data=project,
        message="프로젝트가 생성되었습니다"
    )

# 검증 오류 (400)
@app.route('/api/projects/<code>')
def get_project(code):
    if not code:
        return APIResponse.validation_error(
            details={'field': 'project_code'},
            message="프로젝트 코드가 필요합니다"
        )

# 리소스 없음 (404)
    project = find_project(code)
    if not project:
        return APIResponse.not_found(
            resource="프로젝트",
            resource_id=code
        )

    return APIResponse.success(data=project)

# 부분 성공 (207)
@app.route('/api/memos/batch', methods=['POST'])
def save_memos_batch():
    results = process_batch(request.json)
    # Note: 부분 성공은 수동으로 구현 (APIResponse에 partial_success 없음)
    return jsonify({
        'success': True,
        'data': {'results': results},
        'meta': {'timestamp': datetime.now().isoformat()}
    }), 207

# 외부 서비스 오류 (503)
@app.route('/api/sync')
def sync_sheets():
    try:
        sync_with_sheets()
    except SheetsAPIError as e:
        return APIResponse.service_unavailable(
            message="Google Sheets 동기화에 실패했습니다"
        )
```

### 2. 수동 응답 (레거시, 점진적 마이그레이션)

```python
# 기존 방식 (허용되지만 점진적으로 ApiResponse로 변경)
return jsonify({
    'success': True,
    'data': projects,
    'meta': {'timestamp': datetime.now().isoformat()}
}), 200
```

---

## 📝 에러 코드 목록

### 클라이언트 에러 (4xx)

| 코드 | HTTP | 설명 | 사용 예시 |
|------|------|------|----------|
| `VALIDATION_ERROR` | 400 | 입력 검증 실패 | Marshmallow 검증 실패 |
| `BAD_REQUEST` | 400 | 잘못된 요청 | 필수 필드 누락, 형식 오류 |
| `MISSING_REQUIRED_FIELD` | 400 | 필수 필드 누락 | project_code 없음 |
| `UNAUTHORIZED` | 401 | 인증 필요 | 로그인 필요 |
| `FORBIDDEN` | 403 | 권한 없음 | viewer가 수정 시도 |
| `RESOURCE_NOT_FOUND` | 404 | 리소스 없음 | 프로젝트 없음 |
| `CONFLICT` | 409 | 리소스 충돌 | 중복 프로젝트 코드 |
| `RATE_LIMITED` | 429 | 요청 제한 | API 호출 초과 |

### 서버 에러 (5xx)

| 코드 | HTTP | 설명 | 사용 예시 |
|------|------|------|----------|
| `INTERNAL_ERROR` | 500 | 내부 서버 오류 | Python Exception |
| `DATABASE_ERROR` | 500 | 데이터베이스 오류 | SQLite 오류 |
| `SERVICE_UNAVAILABLE` | 503 | 서비스 불가 | 외부 서비스 장애 |
| `EXTERNAL_SERVICE_ERROR` | 503 | 외부 서비스 장애 | 일반 외부 API 오류 |
| `TIMEOUT_ERROR` | 504 | 타임아웃 | API 응답 지연 |

### 비즈니스 로직 에러

| 코드 | HTTP | 설명 | 사용 예시 |
|------|------|------|----------|
| `PROJECT_NOT_FOUND` | 404 | 프로젝트 없음 | 존재하지 않는 프로젝트 조회 |
| `PROJECT_ALREADY_EXISTS` | 409 | 프로젝트 중복 | 동일한 코드로 생성 시도 |
| `PROJECT_LOCKED` | 423 | 프로젝트 잠김 | 수정 불가 상태 |
| `INSUFFICIENT_PERMISSIONS` | 403 | 권한 부족 | 담당자 아닌 사용자의 수정 시도 |
| `INVALID_PROJECT_STATUS` | 400 | 잘못된 상태 | 진행중인 프로젝트를 삭제 시도 |

---

## 🎯 모범 사례 (Best Practices)

### 1. 적절한 HTTP 코드 사용

✅ **좋은 예**:
```python
# 프로젝트 생성 → 201
return APIResponse.created(data=project)

# 프로젝트 삭제 → 204
return APIResponse.no_content()

# 검증 실패 → 400
return APIResponse.validation_error(details=errors)

# 리소스 없음 → 404
return APIResponse.not_found(resource="프로젝트", resource_id=code)
```

❌ **나쁜 예**:
```python
# 생성 성공인데 200 (201이 맞음)
return APIResponse.success(data=project)  # status_code=200 (기본값)

# 검증 실패인데 500 (400이 맞음)
return APIResponse.error(APIErrorCode.INTERNAL_ERROR, "검증 실패")
```

### 2. 에러 메시지 작성

✅ **좋은 예**:
```python
# 사용자 친화적 + 개발자 상세 정보
return APIResponse.error(
    APIErrorCode.VALIDATION_ERROR,
    message="프로젝트 코드는 영문, 숫자, 하이픈만 사용 가능합니다",
    details={
        'field': 'project_code',
        'value': 'IT@2024',
        'pattern': '^[A-Z0-9-]+$'
    }
)
```

❌ **나쁜 예**:
```python
# 너무 기술적이거나 불친절
return APIResponse.error(
    APIErrorCode.VALIDATION_ERROR,
    message="Regex match failed: ^[A-Z0-9-]+$"
)
```

### 3. 일관된 메타 정보

✅ **좋은 예**:
```python
# 페이지네이션 정보 (PaginationHelper 사용)
from dashboard.api.responses import PaginationHelper

pagination = PaginationHelper.create_pagination_meta(
    page=1, limit=20, total=100, items=projects
)
return APIResponse.success(data={'items': projects}, pagination=pagination)

# 필터 정보
return APIResponse.success(
    data=filtered_projects,
    total_filtered=30,
    filters={'status': 'active', 'manager': '홍길동'}
)
```

### 4. 헬퍼 메서드 활용

✅ **좋은 예**:
```python
# 생성: created() 사용
return APIResponse.created(data=new_project)

# 검증 오류: validation_error() 사용
return APIResponse.validation_error(details=errors)

# 리소스 없음: not_found() 사용
return APIResponse.not_found(resource="프로젝트", resource_id=code)

# 충돌: conflict() 사용
return APIResponse.conflict(resource="프로젝트")

# 권한 없음: forbidden() 사용
return APIResponse.forbidden(message="담당자만 수정할 수 있습니다")
```

---

## 🔄 마이그레이션 가이드

### 기존 API를 표준으로 전환

**단계**:
1. 신규 API부터 `ApiResponse` 사용
2. 수정이 필요한 기존 API 점진적 전환
3. 레거시 응답 구조는 당분간 유지 (하위 호환성)

### 표준 API 응답 구조

**표준 임포트**:
```python
from dashboard.api.responses import APIResponse, APIErrorCode
```

**주요 클래스**:
- `APIResponse`: 표준 응답 생성기
- `APIErrorCode`: 표준 에러 코드 Enum
- `PaginationHelper`: 페이지네이션 헬퍼
- `APIException`: API 예외 기본 클래스

**마이그레이션 전략**:
- ✅ 기존 코드는 계속 작동 (dashboard.api.responses 사용 중)
- ✅ 신규 API는 헬퍼 메서드 활용 (created, validation_error 등)
- ✅ 레거시 수동 jsonify 응답은 점진적으로 APIResponse로 전환

**예시**:

**Before (레거시)**:
```python
@app.route('/api/projects')
def get_projects():
    try:
        projects = load_projects()
        return jsonify({
            'success': True,
            'data': projects
        }), 200
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500
```

**After (표준)**:
```python
from dashboard.api.responses import APIResponse, APIErrorCode

@app.route('/api/projects')
def get_projects():
    try:
        projects = load_projects()
        return APIResponse.success(
            data=projects,
            total=len(projects)
        )
    except Exception as e:
        logger.error(f"프로젝트 목록 조회 오류: {e}")
        return APIResponse.internal_error(
            message="프로젝트 목록을 불러올 수 없습니다"
        )
```

---

## 🧪 테스트 예시

```python
from dashboard.api.responses import APIResponse, APIErrorCode

def test_api_response_success():
    response, status_code = APIResponse.success(data={'name': 'test'})
    assert status_code == 200
    data = response.get_json()
    assert data['success'] is True
    assert data['data']['name'] == 'test'
    assert 'timestamp' in data['meta']

def test_api_response_created():
    response, status_code = APIResponse.created(data={'id': 123})
    assert status_code == 201
    data = response.get_json()
    assert data['success'] is True

def test_api_response_validation_error():
    response, status_code = APIResponse.validation_error(
        details={'field': 'email'},
        message="이메일 형식이 올바르지 않습니다"
    )
    assert status_code == 400
    data = response.get_json()
    assert data['success'] is False
    assert data['error']['code'] == 'VALIDATION_ERROR'

def test_api_response_not_found():
    response, status_code = APIResponse.not_found(
        resource="프로젝트",
        resource_id="IT-2024-001"
    )
    assert status_code == 404
    data = response.get_json()
    assert data['error']['code'] == 'RESOURCE_NOT_FOUND'
```

---

## 📚 참고 문서

- `dashboard/api/responses.py` - 표준 응답 헬퍼 구현
- `dashboard/tests/test_api_response.py` - API 응답 테스트 예시
- REST API HTTP 상태 코드: https://developer.mozilla.org/en-US/docs/Web/HTTP/Status
- RFC 4918 (WebDAV) - 207 Multi-Status: https://tools.ietf.org/html/rfc4918#section-11.1

---

## 🔄 변경 이력

### 2025-10-30
- 기존 api/responses.py를 표준으로 채택
- 모든 임포트 경로 수정 (utils → api)
- 에러 코드 목록 확장 (비즈니스 로직 에러 추가)
- 헬퍼 메서드 활용 예시 추가

### 2025-10-29
- 초기 문서 작성
- HTTP 상태 코드 규칙 정의
- 표준 응답 구조 및 에러 코드 목록 작성
- 사용 예시 및 모범 사례 추가
