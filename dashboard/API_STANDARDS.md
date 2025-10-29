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

### 1. ApiResponse 헬퍼 사용 (권장)

```python
from dashboard.utils.api_response import ApiResponse, ErrorCode

# 성공 응답 (200)
@app.route('/api/projects/list')
def get_projects():
    projects = get_all_projects()
    return ApiResponse.success(
        data=projects,
        meta={'total': len(projects)}
    )

# 생성 성공 (201)
@app.route('/api/projects', methods=['POST'])
def create_project():
    project = create_new_project(request.json)
    return ApiResponse.success(
        data=project,
        message="프로젝트가 생성되었습니다",
        status_code=201
    )

# 검증 오류 (400)
@app.route('/api/projects/<code>')
def get_project(code):
    if not code:
        return ApiResponse.error(
            ErrorCode.MISSING_REQUIRED_FIELD,
            "프로젝트 코드가 필요합니다",
            details={'field': 'project_code'}
        )

# 리소스 없음 (404)
    project = find_project(code)
    if not project:
        return ApiResponse.error(
            ErrorCode.RESOURCE_NOT_FOUND,
            "프로젝트를 찾을 수 없습니다",
            details={'project_code': code}
        )

    return ApiResponse.success(data=project)

# 부분 성공 (207)
@app.route('/api/memos/batch', methods=['POST'])
def save_memos_batch():
    results = process_batch(request.json)
    return ApiResponse.partial_success(
        results=results,
        success_count=5,
        failed_count=2,
        total_count=7
    )

# 외부 서비스 오류 (503)
@app.route('/api/sync')
def sync_sheets():
    try:
        sync_with_sheets()
    except SheetsAPIError as e:
        return ApiResponse.error(
            ErrorCode.SHEETS_API_ERROR,
            "Google Sheets 동기화에 실패했습니다",
            details={'error': str(e)}
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
| `MISSING_REQUIRED_FIELD` | 400 | 필수 필드 누락 | project_code 없음 |
| `INVALID_FORMAT` | 400 | 형식 오류 | 날짜 형식 틀림 |
| `ALREADY_EXISTS` | 400 | 이미 존재 | 중복 프로젝트 코드 |
| `AUTHENTICATION_REQUIRED` | 401 | 인증 필요 | 로그인 필요 |
| `PERMISSION_DENIED` | 403 | 권한 없음 | viewer가 수정 시도 |
| `RESOURCE_NOT_FOUND` | 404 | 리소스 없음 | 프로젝트 없음 |

### 서버 에러 (5xx)

| 코드 | HTTP | 설명 | 사용 예시 |
|------|------|------|----------|
| `INTERNAL_ERROR` | 500 | 내부 서버 오류 | Python Exception |
| `DATABASE_ERROR` | 500 | 데이터베이스 오류 | SQLite 오류 |
| `EXTERNAL_SERVICE_ERROR` | 503 | 외부 서비스 장애 | 일반 외부 API 오류 |
| `SHEETS_API_ERROR` | 503 | Google Sheets 오류 | Sheets API 장애 |
| `CALENDAR_API_ERROR` | 503 | Calendar 오류 | Calendar API 장애 |

### 특수 (2xx)

| 코드 | HTTP | 설명 | 사용 예시 |
|------|------|------|----------|
| `PARTIAL_SUCCESS` | 207 | 부분 성공 | 배치 작업에서 일부 실패 |

---

## 🎯 모범 사례 (Best Practices)

### 1. 적절한 HTTP 코드 사용

✅ **좋은 예**:
```python
# 프로젝트 생성 → 201
return ApiResponse.success(data=project, status_code=201)

# 프로젝트 삭제 → 204
return ApiResponse.success(status_code=204)

# 검증 실패 → 400
return ApiResponse.error(ErrorCode.VALIDATION_ERROR, "...")

# 리소스 없음 → 404
return ApiResponse.error(ErrorCode.RESOURCE_NOT_FOUND, "...")
```

❌ **나쁜 예**:
```python
# 생성 성공인데 200 (201이 맞음)
return ApiResponse.success(data=project)  # status_code=200 (기본값)

# 검증 실패인데 500 (400이 맞음)
return ApiResponse.error(ErrorCode.INTERNAL_ERROR, "검증 실패")
```

### 2. 에러 메시지 작성

✅ **좋은 예**:
```python
# 사용자 친화적 + 개발자 상세 정보
return ApiResponse.error(
    ErrorCode.VALIDATION_ERROR,
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
return ApiResponse.error(
    ErrorCode.VALIDATION_ERROR,
    message="Regex match failed: ^[A-Z0-9-]+$"
)
```

### 3. 일관된 메타 정보

✅ **좋은 예**:
```python
# 페이지네이션 정보
return ApiResponse.success(
    data=projects,
    meta={
        'total': 100,
        'page': 1,
        'per_page': 20,
        'total_pages': 5
    }
)

# 필터 정보
return ApiResponse.success(
    data=filtered_projects,
    meta={
        'total': 30,
        'filters': {'status': 'active', 'manager': '홍길동'}
    }
)
```

### 4. 부분 성공 처리

✅ **좋은 예**:
```python
# 배치 작업에서 개별 결과 제공
return ApiResponse.partial_success(
    results=[
        {'field': '계약금', 'success': True, 'message': '저장 완료'},
        {'field': '중도금', 'success': False, 'message': '필드명 오류'},
        {'field': '잔금', 'success': True, 'message': '저장 완료'}
    ],
    success_count=2,
    failed_count=1,
    total_count=3
)
```

---

## 🔄 마이그레이션 가이드

### 기존 API를 표준으로 전환

**단계**:
1. 신규 API부터 `ApiResponse` 사용
2. 수정이 필요한 기존 API 점진적 전환
3. 레거시 응답 구조는 당분간 유지 (하위 호환성)

### 하위 호환성 (Backward Compatibility)

기존 코드에서 사용 중인 레거시 이름은 계속 지원됩니다:

**레거시 이름 (계속 작동)**:
```python
from dashboard.utils.api_response import APIResponse, APIErrorCode, api_response

# 레거시 클래스 사용 (여전히 작동)
APIResponse.success(data=projects)
APIErrorCode.VALIDATION_ERROR

# 레거시 함수 사용 (여전히 작동)
api_response(success=True, data=projects)
```

**새 표준 이름 (권장)**:
```python
from dashboard.utils.api_response import ApiResponse, ErrorCode

# 새 표준 클래스 사용 (권장)
ApiResponse.success(data=projects)
ErrorCode.VALIDATION_ERROR
```

**마이그레이션 전략**:
- ✅ 기존 코드는 수정 없이 계속 작동
- ✅ 신규 코드는 `ApiResponse`, `ErrorCode` 사용
- ✅ 기존 코드 수정 시 점진적으로 새 이름으로 전환
- ⚠️ `api_response()` 함수는 제한적 기능만 제공 (에러 코드 자동 매핑 불가)

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
@app.route('/api/projects')
def get_projects():
    try:
        projects = load_projects()
        return ApiResponse.success(
            data=projects,
            meta={'total': len(projects)}
        )
    except Exception as e:
        logger.error(f"프로젝트 목록 조회 오류: {e}")
        return ApiResponse.error(
            ErrorCode.INTERNAL_ERROR,
            "프로젝트 목록을 불러올 수 없습니다",
            details={'exception': str(e)}
        )
```

---

## 🧪 테스트 예시

```python
def test_api_response_success():
    response, status_code = ApiResponse.success(data={'name': 'test'})
    assert status_code == 200
    assert response.json['success'] is True
    assert response.json['data']['name'] == 'test'
    assert 'timestamp' in response.json['meta']

def test_api_response_error():
    response, status_code = ApiResponse.error(
        ErrorCode.VALIDATION_ERROR,
        "테스트 에러"
    )
    assert status_code == 400
    assert response.json['success'] is False
    assert response.json['error']['code'] == 'VALIDATION_ERROR'

def test_api_response_partial_success():
    response, status_code = ApiResponse.partial_success(
        results=[{'item': 1, 'success': True}],
        success_count=1,
        failed_count=0,
        total_count=1
    )
    assert status_code == 207
    assert response.json['data']['summary']['success_count'] == 1
```

---

## 📚 참고 문서

- `dashboard/utils/api_response.py` - 표준 응답 헬퍼 구현
- REST API HTTP 상태 코드: https://developer.mozilla.org/en-US/docs/Web/HTTP/Status
- RFC 4918 (WebDAV) - 207 Multi-Status: https://tools.ietf.org/html/rfc4918#section-11.1

---

## 🔄 변경 이력

### 2025-10-29
- 초기 문서 작성
- HTTP 상태 코드 규칙 정의
- 표준 응답 구조 및 에러 코드 목록 작성
- 사용 예시 및 모범 사례 추가
