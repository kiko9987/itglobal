# API 응답 스키마 표준

## 배경

현재 코드베이스에 응답 형식 3가지 혼재 (2026-07-08 감사):
- `{'error': '...'}` — 28건
- `{'success': False, 'error': '...'}` — 53건
- `APIResponse.error(...)` — 16건

프론트-백 에러 계약을 통일해서 클라이언트 에러 처리 단순화 필요.

## 표준 형식 (신규 코드는 이 형식 사용)

### 성공 응답
```json
{
  "success": true,
  "data": { ... },        // 실제 응답 데이터
  "message": "선택적 성공 메시지"
}
```

### 실패 응답
```json
{
  "success": false,
  "error": {
    "id": "err_abc123",           // generate_error_id() 결과, 로그 추적용
    "code": "PROJECT_NOT_FOUND",  // 대문자 상수, 프론트 분기용
    "message": "사용자에게 보여줄 한글 메시지"
  }
}
```

HTTP 상태 코드는 별개:
- 400: 클라이언트 요청 오류 (필드 누락·형식 오류)
- 401: 미인증
- 403: 권한 없음
- 404: 리소스 없음
- 409: 충돌 (락 획득 실패 등)
- 500: 서버 내부 오류

## 사용법

### 백엔드 (Python)

```python
from dashboard.api.responses import APIResponse

# 성공
return APIResponse.success(data={'project_code': 'G0001-JW'})

# 실패 — 표준 헬퍼 사용
return APIResponse.error(
    message='프로젝트를 찾을 수 없습니다.',
    code='PROJECT_NOT_FOUND',
    status_code=404
)

# 신규 코드에서 이 방식만 사용. jsonify({'error': ...}) 직접 반환 X.
```

### 프론트엔드 (JavaScript)

```javascript
const res = await fetch('/api/projects/xxx', ...);
const body = await res.json();

if (body.success) {
    // 성공 처리
    const data = body.data;
} else {
    // 실패 처리 — 표준 형식
    const errorMsg = body.error?.message || '알 수 없는 오류';
    const errorCode = body.error?.code;
    showAlert(errorMsg, 'danger');
}
```

## 마이그레이션 계획

### Phase 1 (완료) — 표준 정의 · 신규 코드 강제
- 이 문서 작성 · `APIResponse` 헬퍼 사용 지침
- 신규 API 작성 시 이 형식만 사용

### Phase 2 (백로그, 별도 사이클 2~3시간)
기존 97건 순차 변환. 파일별로:
1. `blueprints/constructors.py` (10건)
2. `blueprints/folders.py` (9건)
3. `blueprints/locks.py` (11건)
4. `blueprints/projects.py` (14건)
5. `blueprints/users.py` (11건)
6. 나머지 (analytics·monitoring·slack_bot·stats·metadata·data_management, 총 42건)

각 파일 변환 후 프론트 호출 지점도 함께 검증 (`response.json()` 후 접근 형태).

### Phase 3 — 프론트 통일
- 모든 fetch 호출 지점을 표준 스키마로 통일
- 공통 헬퍼 `handleAPIResponse(res)` 도입 검토

## 주의

프론트가 지금 두 형식(`.error`, `.error.message`) 다 처리하고 있을 가능성 있음. 백엔드 변환 시 프론트 fallback 확인 필수.

지금 실서비스 개시에는 실질 영향 없음 (기능 정상 작동). 유지보수성·확장성 이슈만.
