# 자동 API 문서 생성 시스템 (auto-docs)

> **참고**: 이 문서는 `/api/*` 엔드포인트를 OpenAPI 3.0 스펙 등으로 자동 문서화하는 **서브시스템** 안내입니다. 프로젝트 전체 개요 및 사용법은 루트 [README.md](../../README.md), 운영 가이드는 [OPERATIONS.md](../../OPERATIONS.md)를 참고하세요.

Flask 라우트와 Marshmallow 스키마를 정적 분석해서 OpenAPI 3.0 스펙 · Swagger UI · Markdown · Postman 컬렉션을 자동 생성.

## 위치

```
dashboard/docs/
├── analyzers/           # 코드 분석기
│   ├── route_analyzer.py    # Flask 라우트 → OpenAPI paths
│   └── schema_analyzer.py   # Marshmallow → OpenAPI schemas
├── generators/          # 문서 생성기
│   └── openapi_generator.py
├── generated/           # 산출물 (JSON/YAML/Markdown/Postman)
├── doc_generator.py     # 중앙 관리자
├── api_endpoints.py     # REST API (`/api/v1/docs/*`)
└── README.md           # 이 파일
```

## 접근

- **Swagger UI**: `/docs`
- **관리자 대시보드**: `/admin/docs` (관리자 권한 필요)
- **REST API**: `/api/v1/docs/status`, `/api/v1/docs/generate`, `/api/v1/docs/download/{format}`

## CLI

```bash
flask generate-docs      # 모든 형식 생성
flask validate-docs      # 코드-문서 일치성 검증
flask docs-stats         # 문서화 통계
```

## 동작 원리

### 1. 라우트 분석 (`route_analyzer.py`)
Flask 앱을 순회하며 URL 패턴 · HTTP 메서드 · docstring · 데코레이터 · 파라미터를 추출.

### 2. 스키마 분석 (`schema_analyzer.py`)
Marshmallow 스키마의 필드 타입 · 검증 규칙 · 중첩 관계를 OpenAPI 스키마로 변환.

### 3. OpenAPI 생성 (`openapi_generator.py`)
분석 결과를 조합해 OpenAPI 3.0 스펙 생성 → JSON/YAML로 내보내고 파생 형식(Markdown, Postman) 생성.

### 4. 웹 관리 (`api_endpoints.py` + `templates/docs_management.html`)
관리자가 웹 UI에서 원클릭 생성 · 미리보기 · 다운로드.

## 환경 변수

```env
AUTO_DOCS_ENABLED=true
DOCS_UPDATE_INTERVAL=300
GENERATE_MARKDOWN=true
GENERATE_POSTMAN=true
VALIDATE_DOCS=true
```

## 확장

새 문서 형식 추가:
1. `generators/`에 생성기 클래스 추가
2. `doc_generator.py`에 등록
3. `api_endpoints.py`에 다운로드 라우트 추가
4. `templates/docs_management.html`에 UI 요소 추가

새 스키마 라이브러리 지원:
1. `schema_analyzer.py`에 타입 매핑 추가
2. 예제 데이터 생성 규칙 추가

## 트러블슈팅

| 증상 | 확인 |
|---|---|
| 문서 생성 실패 | Flask 실행 상태, 관리자 권한, `generated/` 디렉터리 쓰기 권한 |
| 일부 API 누락 | 라우트가 `/api/` prefix 사용 여부, docstring 존재 여부 |
| 스키마 변환 오류 | Marshmallow 버전 호환성, 커스텀 필드 매핑 |

로그: `dashboard/logs/itglobal_dashboard.log`
