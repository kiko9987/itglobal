# API 응답 포맷 표준

## 개요

이 문서는 프로젝트 대시보드 API의 응답 포맷 표준을 정의합니다.

## 표준 포맷 (권장)

### 성공 응답

```python
{
    "success": true,
    "data": { ... },           # 선택적: 응답 데이터
    "message": "작업 완료",    # 선택적: 사용자 메시지
    "timestamp": "2025-01-20T10:30:00"  # 선택적: 타임스탬프
}
```

### 에러 응답

```python
{
    "success": false,
    "error": "에러 메시지",
    "code": "ERROR_CODE",      # 선택적: 에러 코드
    "details": { ... },        # 선택적: 추가 정보
    "timestamp": "2025-01-20T10:30:00"  # 선택적: 타임스탬프
}
```

## 레거시 포맷

일부 API 엔드포인트는 'ok' 필드를 사용합니다:

```python
{
    "ok": true,
    "data": { ... },
    "error": "에러 메시지"     # ok가 false일 때
}
```

**주의:** 새로운 API 엔드포인트는 'success' 포맷을 사용해야 합니다.

## 헬퍼 함수 (향후 추가 예정)

```python
from dashboard.utils.api_response import success_response, error_response

# 성공 응답
@app.route('/api/example')
def example():
    return success_response(
        data={"result": "example"},
        message="성공적으로 완료되었습니다"
    )

# 에러 응답
@app.route('/api/error-example')
def error_example():
    return error_response(
        error="잘못된 요청",
        code="INVALID_REQUEST",
        status_code=400
    )
```

## HTTP 상태 코드 가이드라인

- **200 OK**: 성공적인 GET, PUT, PATCH 요청
- **201 Created**: 성공적인 POST 요청 (리소스 생성)
- **204 No Content**: 성공적인 DELETE 요청 (응답 본문 없음)
- **400 Bad Request**: 잘못된 요청 (유효성 검증 실패)
- **401 Unauthorized**: 인증 실패
- **403 Forbidden**: 권한 부족
- **404 Not Found**: 리소스를 찾을 수 없음
- **409 Conflict**: 리소스 충돌 (중복 생성 등)
- **422 Unprocessable Entity**: 처리 불가능한 엔티티 (비즈니스 로직 오류)
- **429 Too Many Requests**: Rate Limiting
- **500 Internal Server Error**: 서버 내부 오류

## 현재 포맷 사용 현황

### 'success' 포맷 사용 파일 (7개)
- `blueprints/data_management.py`
- `blueprints/monitoring.py`
- `blueprints/projects.py`
- `blueprints/folders.py`
- `blueprints/locks.py`
- `blueprints/users.py`
- `blueprints/stats.py`

### 'ok' 포맷 사용 파일 (2개)
- `blueprints/projects.py` (일부 엔드포인트)
- `blueprints/metadata.py`
- `api/v1/auth.py`

## 마이그레이션 가이드

기존 'ok' 포맷을 'success' 포맷으로 변경하는 방법:

### Before (레거시)
```python
return jsonify({
    'ok': True,
    'token': csrf_token
}), 200
```

### After (표준)
```python
return jsonify({
    'success': True,
    'data': {'token': csrf_token}
}), 200
```

## 검증 규칙

1. **필수 필드**: 모든 응답은 `success` 필드를 포함해야 합니다
2. **데이터 래핑**: 데이터는 `data` 키 안에 래핑합니다
3. **에러 정보**: 에러 시 `error` 필드에 메시지를 포함합니다
4. **일관성**: 같은 엔드포인트는 항상 같은 포맷을 사용합니다

## 예외 사항

다음 엔드포인트는 특수한 포맷을 사용할 수 있습니다:

- **OAuth 콜백**: 외부 서비스 표준에 따름
- **Webhook**: 외부 서비스 요구사항에 따름
- **프록시 엔드포인트**: 원본 응답을 그대로 전달

## 추가 참고 사항

- 응답 시간이 긴 작업은 비동기 패턴 사용을 고려하세요
- 페이지네이션이 필요한 경우 `pagination` 객체를 포함하세요:
  ```python
  {
      "success": true,
      "data": [...],
      "pagination": {
          "page": 1,
          "per_page": 20,
          "total": 100,
          "pages": 5
      }
  }
  ```

## 버전 관리

- **작성일**: 2025-01-20
- **작성자**: Claude Code
- **버전**: 1.0
- **다음 리뷰 예정**: 2025-07-20
