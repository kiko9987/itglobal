# IT Global Dashboard API

> **버전**: 1.0.0
> **생성일**: 2025-09-19 16:46:29

## 개요

IT Global 비즈니스 관리 대시보드를 위한 종합 API입니다.

## 주요 기능
- 프로젝트 관리 및 추적
- 매출 및 재무 분석
- 실시간 모니터링 및 메트릭
- 사용자 관리 및 인증
- 캐시 및 시스템 관리

## 인증 방식
이 API는 여러 인증 방법을 지원합니다:
- 세션 기반 인증 (웹 인터페이스)
- Bearer 토큰 인증 (API 접근)
- Google OAuth 2.0 (웹 인터페이스)

## 표준 응답 형식
모든 API 응답은 다음과 같은 표준 형식을 따릅니다:
```json
{
  "success": true|false,
  "data": any,
  "error": {...},
  "meta": {...}
}
```

## 버전 관리
API는 URL 경로 기반 버전 관리를 사용합니다 (/api/v1/, /api/v2/ 등).

## 서버 정보

- **Development server**: `http://localhost:5000`
- **Production server**: `https://dashboard.itglobal.com`

## 인증

### bearerAuth
- **타입**: http
- **스키마**: bearer

### sessionAuth
- **타입**: apiKey
- **위치**: cookie


## Projects

프로젝트 관리 관련 API

### GET /api/projects/list

**프로젝트 목록 API**

/api/projects/list 경로에서 데이터를 조회합니다.

**응답:**

- **200**: 성공
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---

### POST /api/projects/auto

**신규 프로젝트 자동 코드 생성 및 추가**

/api/projects/auto 경로에서 새로운 데이터를 생성합니다.

**응답:**

- **200**: 성공
- **201**: 생성됨
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---

### POST /api/projects

**새 프로젝트 생성 API**

/api/projects 경로에서 새로운 데이터를 생성합니다.

**요청 본문:**

```json
// Content-Type: application/json
{
  // 요청 데이터 구조
}
```

**응답:**

- **200**: 성공
- **201**: 생성됨
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---

### GET /api/projects/<project_code>

**프로젝트 상세 정보 API**

/api/projects/<project_code> 경로에서 데이터를 조회합니다.

**파라미터:**

- `project_code` (path) (필수): project_code 파라미터

**응답:**

- **200**: 성공
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---

### PUT /api/projects/<project_code>

**프로젝트 수정 API**

/api/projects/<project_code> 경로에서 데이터를 업데이트합니다.

**파라미터:**

- `project_code` (path) (필수): project_code 파라미터

**응답:**

- **200**: 성공
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---

### GET /api/projects/<project_code>/folder_id

**프로젝트의 Google Drive 폴더 ID 반환**

/api/projects/<project_code>/folder_id 경로에서 데이터를 조회합니다.

**파라미터:**

- `project_code` (path) (필수): project_code 파라미터

**응답:**

- **200**: 성공
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---


## General

기타 API

### GET /api/next-project-code

**다음 프로젝트 코드 생성 API**

/api/next-project-code 경로에서 데이터를 조회합니다.

**응답:**

- **200**: 성공
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---

### GET /api/health

**간단한 헬스체크 API**

/api/health 경로에서 데이터를 조회합니다.

**응답:**

- **200**: 성공
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류

---

### GET /api/docs/

**The /apidocs**

/api/docs/ 경로에서 데이터를 조회합니다.

**응답:**

- **200**: 성공
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---

### GET /api/docs

**API Documentation Landing Page**

/api/docs 경로에서 데이터를 조회합니다.

**응답:**

- **200**: 성공
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---

### GET /api/summary

**요약 통계 API**

/api/summary 경로에서 데이터를 조회합니다.

**응답:**

- **200**: 성공
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---

### GET /api/monthly-sales

**월별 매출 API**

/api/monthly-sales 경로에서 데이터를 조회합니다.

**응답:**

- **200**: 성공
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---

### GET /api/regional-analysis

**지역별 분석 API**

/api/regional-analysis 경로에서 데이터를 조회합니다.

**응답:**

- **200**: 성공
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---

### GET /api/outstanding-analysis

**미수금 분석 API**

/api/outstanding-analysis 경로에서 데이터를 조회합니다.

**응답:**

- **200**: 성공
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---

### GET /api/cache-stats

**캐시 상태 모니터링 API**

/api/cache-stats 경로에서 데이터를 조회합니다.

**응답:**

- **200**: 성공
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---

### POST /api/cache-clear

**캐시 전체 삭제 API (관리자 전용)**

/api/cache-clear 경로에서 새로운 데이터를 생성합니다.

**응답:**

- **200**: 성공
- **201**: 생성됨
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---

### GET /api/missing-data

**누락 데이터 분석 API**

/api/missing-data 경로에서 데이터를 조회합니다.

**응답:**

- **200**: 성공
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---

### GET /api/brand-analysis

**브랜드별 분석 API**

/api/brand-analysis 경로에서 데이터를 조회합니다.

**응답:**

- **200**: 성공
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---

### GET /api/preview-project-code

**프로젝트 코드 미리보기 생성**

/api/preview-project-code 경로에서 데이터를 조회합니다.

**응답:**

- **200**: 성공
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---

### GET /api/meta/options

**드롭다운용 옵션 API (사업자, 담당자 목록)**

/api/meta/options 경로에서 데이터를 조회합니다.

**응답:**

- **200**: 성공
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---

### GET /api/refresh-data

**데이터 새로고침 API (강제 새로고침)**

/api/refresh-data 경로에서 데이터를 조회합니다.

**응답:**

- **200**: 성공
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---

### GET /api/quick-refresh

**빠른 데이터 새로고침 API (캐시 우선 사용)**

/api/quick-refresh 경로에서 데이터를 조회합니다.

**응답:**

- **200**: 성공
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---

### GET /api/receivables

**미수금 관리용 데이터 API**

/api/receivables 경로에서 데이터를 조회합니다.

**응답:**

- **200**: 성공
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---

### POST /api/card-lock/acquire

**카드 잠금 획득**

/api/card-lock/acquire 경로에서 새로운 데이터를 생성합니다.

**응답:**

- **200**: 성공
- **201**: 생성됨
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---

### POST /api/card-lock/release

**카드 잠금 해제**

/api/card-lock/release 경로에서 새로운 데이터를 생성합니다.

**응답:**

- **200**: 성공
- **201**: 생성됨
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---

### GET /api/field-lock/status/<project_code>

**프로젝트의 잠금 상태 조회**

/api/field-lock/status/<project_code> 경로에서 데이터를 조회합니다.

**파라미터:**

- `project_code` (path) (필수): project_code 파라미터

**응답:**

- **200**: 성공
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---

### GET /api/field-lock/status/all

**모든 프로젝트의 잠금 상태 한 번에 조회 (성능 최적화)**

/api/field-lock/status/all 경로에서 데이터를 조회합니다.

**응답:**

- **200**: 성공
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---

### POST /api/release-all-user-locks

**특정 사용자의 모든 잠금 강제 해제 (페이지 새로고침/종료 시)**

/api/release-all-user-locks 경로에서 새로운 데이터를 생성합니다.

**응답:**

- **200**: 성공
- **201**: 생성됨
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---

### POST /api/inline-update

**간단한 인라인 업데이트 API (직접 구현)**

/api/inline-update 경로에서 새로운 데이터를 생성합니다.

**응답:**

- **200**: 성공
- **201**: 생성됨
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---

### GET /api/audit-logs

**감사 로그 조회 API (페이지네이션 지원)**

/api/audit-logs 경로에서 데이터를 조회합니다.

**응답:**

- **200**: 성공
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---

### GET /api/user/role

**사용자 역할 정보 반환**

/api/user/role 경로에서 데이터를 조회합니다.

**응답:**

- **200**: 성공
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---

### POST /api/folder/convert-paths-to-ids

**기존 프로젝트들의 폴더 경로를 Google Drive ID로 일괄 변환**

/api/folder/convert-paths-to-ids 경로에서 새로운 데이터를 생성합니다.

**응답:**

- **200**: 성공
- **201**: 생성됨
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---

### GET /api/folder/name/<project_code>

**프로젝트 폴더명 가져오기**

/api/folder/name/<project_code> 경로에서 데이터를 조회합니다.

**파라미터:**

- `project_code` (path) (필수): project_code 파라미터

**응답:**

- **200**: 성공
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---

### GET /api/folder/open/<project_code>

**프로젝트 폴더를 윈도우 탐색기로 열기**

/api/folder/open/<project_code> 경로에서 데이터를 조회합니다.

**파라미터:**

- `project_code` (path) (필수): project_code 파라미터

**응답:**

- **200**: 성공
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---


## Test

테스트 및 데모용 API

### GET /api/v1/test

**테스트용 API 엔드포인트**

/api/v1/test 경로에서 데이터를 조회합니다.

**응답:**

- **200**: 성공
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---

### GET /api/v1/test/health

**헬스 체크 엔드포인트**

/api/v1/test/health 경로에서 데이터를 조회합니다.

**응답:**

- **200**: 성공
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류

---

### POST /api/v1/test/echo

**에코 테스트 엔드포인트**

/api/v1/test/echo 경로에서 새로운 데이터를 생성합니다.

**응답:**

- **200**: 성공
- **201**: 생성됨
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류

---


## Projects v1

Projects v1 관련 API

### GET /api/v1/projects

**프로젝트 목록 조회 (페이지네이션 지원)**

summary: 프로젝트 목록 조회
description: 페이지네이션과 필터링을 지원하는 프로젝트 목록 조회
parameters:
in: query
type: integer
minimum: 1
default: 1
description: 페이지 번호
in: query
type: integer
minimum: 1
maximum: 100
default: 20
description: 페이지당 항목 수
in: query
type: string
default: created_at
description: 정렬 필드
in: query
type: string
enum: [asc, desc]
default: desc
description: 정렬 순서
in: query
type: string
enum: [active, completed, cancelled, on_hold]
description: 상태별 필터링
in: query
type: string
description: 프로젝트명 검색
responses:
200:
description: 프로젝트 목록 조회 성공
schema:
$ref: '#/definitions/StandardResponse'

**응답:**

- **200**: 성공
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---

### POST /api/v1/projects

**새 프로젝트 생성**

summary: 프로젝트 생성
description: 새로운 프로젝트를 생성합니다
parameters:
in: body
required: true
schema:
type: object
properties:
name:
type: string
description: 프로젝트명 (필수)
example: "웹사이트 리뉴얼"
description:
type: string
description: 프로젝트 설명
example: "회사 홈페이지 전면 리뉴얼 프로젝트"
start_date:
type: string
format: date
description: 시작일 (필수)
example: "2025-01-01"
end_date:
type: string
format: date
description: 종료일
example: "2025-06-30"
budget:
type: number
description: 예산
example: 50000000
required:
responses:
201:
description: 프로젝트 생성 성공
400:
description: 입력 데이터 유효성 검사 실패
409:
description: 중복된 프로젝트명

**요청 본문:**

```json
// Content-Type: application/json
{
  // 요청 데이터 구조
}
```

**응답:**

- **200**: 성공
- **201**: 생성됨
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---

### GET /api/v1/projects/<project_code>

**특정 프로젝트 조회**

summary: 프로젝트 상세 조회
description: 프로젝트 코드로 특정 프로젝트의 상세 정보를 조회합니다
parameters:
in: path
required: true
type: string
description: 프로젝트 코드
example: "PRJ001"
responses:
200:
description: 프로젝트 조회 성공
404:
description: 프로젝트를 찾을 수 없음

**파라미터:**

- `project_code` (path) (필수): project_code 파라미터

**응답:**

- **200**: 성공
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---

### PATCH /api/v1/projects/<project_code>

**프로젝트 부분 업데이트**

summary: 프로젝트 업데이트
description: 프로젝트의 특정 필드만 업데이트합니다
parameters:
in: path
required: true
type: string
description: 프로젝트 코드
in: body
required: true
schema:
type: object
properties:
name:
type: string
description: 프로젝트명
description:
type: string
description: 프로젝트 설명
status:
type: string
enum: [active, completed, cancelled, on_hold]
description: 프로젝트 상태
end_date:
type: string
format: date
description: 종료일
budget:
type: number
description: 예산
progress:
type: integer
minimum: 0
maximum: 100
description: 진행률
responses:
200:
description: 프로젝트 업데이트 성공
404:
description: 프로젝트를 찾을 수 없음
400:
description: 입력 데이터 유효성 검사 실패

**파라미터:**

- `project_code` (path) (필수): project_code 파라미터

**응답:**

- **200**: 성공
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---

### PUT /api/v1/projects/<project_code>/status

**프로젝트 상태 변경**

summary: 프로젝트 상태 변경
description: 프로젝트의 상태를 변경합니다
parameters:
in: path
required: true
type: string
description: 프로젝트 코드
in: body
required: true
schema:
type: object
properties:
status:
type: string
enum: [active, completed, cancelled, on_hold]
description: 새로운 상태
reason:
type: string
description: 상태 변경 사유
required:
responses:
200:
description: 상태 변경 성공
404:
description: 프로젝트를 찾을 수 없음
400:
description: 잘못된 상태 변경 요청

**파라미터:**

- `project_code` (path) (필수): project_code 파라미터

**응답:**

- **200**: 성공
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---

### GET /api/v1/projects/summary

**프로젝트 요약 통계**

summary: 프로젝트 요약 통계
description: 전체 프로젝트의 요약 통계 정보를 조회합니다
responses:
200:
description: 요약 통계 조회 성공

**응답:**

- **200**: 성공
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---


## name: page

name: page 관련 API

### GET /api/v1/projects

**프로젝트 목록 조회 (페이지네이션 지원)**

summary: 프로젝트 목록 조회
description: 페이지네이션과 필터링을 지원하는 프로젝트 목록 조회
parameters:
in: query
type: integer
minimum: 1
default: 1
description: 페이지 번호
in: query
type: integer
minimum: 1
maximum: 100
default: 20
description: 페이지당 항목 수
in: query
type: string
default: created_at
description: 정렬 필드
in: query
type: string
enum: [asc, desc]
default: desc
description: 정렬 순서
in: query
type: string
enum: [active, completed, cancelled, on_hold]
description: 상태별 필터링
in: query
type: string
description: 프로젝트명 검색
responses:
200:
description: 프로젝트 목록 조회 성공
schema:
$ref: '#/definitions/StandardResponse'

**응답:**

- **200**: 성공
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---


## name: limit

name: limit 관련 API

### GET /api/v1/projects

**프로젝트 목록 조회 (페이지네이션 지원)**

summary: 프로젝트 목록 조회
description: 페이지네이션과 필터링을 지원하는 프로젝트 목록 조회
parameters:
in: query
type: integer
minimum: 1
default: 1
description: 페이지 번호
in: query
type: integer
minimum: 1
maximum: 100
default: 20
description: 페이지당 항목 수
in: query
type: string
default: created_at
description: 정렬 필드
in: query
type: string
enum: [asc, desc]
default: desc
description: 정렬 순서
in: query
type: string
enum: [active, completed, cancelled, on_hold]
description: 상태별 필터링
in: query
type: string
description: 프로젝트명 검색
responses:
200:
description: 프로젝트 목록 조회 성공
schema:
$ref: '#/definitions/StandardResponse'

**응답:**

- **200**: 성공
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---


## name: sort

name: sort 관련 API

### GET /api/v1/projects

**프로젝트 목록 조회 (페이지네이션 지원)**

summary: 프로젝트 목록 조회
description: 페이지네이션과 필터링을 지원하는 프로젝트 목록 조회
parameters:
in: query
type: integer
minimum: 1
default: 1
description: 페이지 번호
in: query
type: integer
minimum: 1
maximum: 100
default: 20
description: 페이지당 항목 수
in: query
type: string
default: created_at
description: 정렬 필드
in: query
type: string
enum: [asc, desc]
default: desc
description: 정렬 순서
in: query
type: string
enum: [active, completed, cancelled, on_hold]
description: 상태별 필터링
in: query
type: string
description: 프로젝트명 검색
responses:
200:
description: 프로젝트 목록 조회 성공
schema:
$ref: '#/definitions/StandardResponse'

**응답:**

- **200**: 성공
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---


## name: order

name: order 관련 API

### GET /api/v1/projects

**프로젝트 목록 조회 (페이지네이션 지원)**

summary: 프로젝트 목록 조회
description: 페이지네이션과 필터링을 지원하는 프로젝트 목록 조회
parameters:
in: query
type: integer
minimum: 1
default: 1
description: 페이지 번호
in: query
type: integer
minimum: 1
maximum: 100
default: 20
description: 페이지당 항목 수
in: query
type: string
default: created_at
description: 정렬 필드
in: query
type: string
enum: [asc, desc]
default: desc
description: 정렬 순서
in: query
type: string
enum: [active, completed, cancelled, on_hold]
description: 상태별 필터링
in: query
type: string
description: 프로젝트명 검색
responses:
200:
description: 프로젝트 목록 조회 성공
schema:
$ref: '#/definitions/StandardResponse'

**응답:**

- **200**: 성공
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---


## name: status

name: status 관련 API

### GET /api/v1/projects

**프로젝트 목록 조회 (페이지네이션 지원)**

summary: 프로젝트 목록 조회
description: 페이지네이션과 필터링을 지원하는 프로젝트 목록 조회
parameters:
in: query
type: integer
minimum: 1
default: 1
description: 페이지 번호
in: query
type: integer
minimum: 1
maximum: 100
default: 20
description: 페이지당 항목 수
in: query
type: string
default: created_at
description: 정렬 필드
in: query
type: string
enum: [asc, desc]
default: desc
description: 정렬 순서
in: query
type: string
enum: [active, completed, cancelled, on_hold]
description: 상태별 필터링
in: query
type: string
description: 프로젝트명 검색
responses:
200:
description: 프로젝트 목록 조회 성공
schema:
$ref: '#/definitions/StandardResponse'

**응답:**

- **200**: 성공
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---


## name: search

name: search 관련 API

### GET /api/v1/projects

**프로젝트 목록 조회 (페이지네이션 지원)**

summary: 프로젝트 목록 조회
description: 페이지네이션과 필터링을 지원하는 프로젝트 목록 조회
parameters:
in: query
type: integer
minimum: 1
default: 1
description: 페이지 번호
in: query
type: integer
minimum: 1
maximum: 100
default: 20
description: 페이지당 항목 수
in: query
type: string
default: created_at
description: 정렬 필드
in: query
type: string
enum: [asc, desc]
default: desc
description: 정렬 순서
in: query
type: string
enum: [active, completed, cancelled, on_hold]
description: 상태별 필터링
in: query
type: string
description: 프로젝트명 검색
responses:
200:
description: 프로젝트 목록 조회 성공
schema:
$ref: '#/definitions/StandardResponse'

**응답:**

- **200**: 성공
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---


## name: body

name: body 관련 API

### POST /api/v1/projects

**새 프로젝트 생성**

summary: 프로젝트 생성
description: 새로운 프로젝트를 생성합니다
parameters:
in: body
required: true
schema:
type: object
properties:
name:
type: string
description: 프로젝트명 (필수)
example: "웹사이트 리뉴얼"
description:
type: string
description: 프로젝트 설명
example: "회사 홈페이지 전면 리뉴얼 프로젝트"
start_date:
type: string
format: date
description: 시작일 (필수)
example: "2025-01-01"
end_date:
type: string
format: date
description: 종료일
example: "2025-06-30"
budget:
type: number
description: 예산
example: 50000000
required:
responses:
201:
description: 프로젝트 생성 성공
400:
description: 입력 데이터 유효성 검사 실패
409:
description: 중복된 프로젝트명

**요청 본문:**

```json
// Content-Type: application/json
{
  // 요청 데이터 구조
}
```

**응답:**

- **200**: 성공
- **201**: 생성됨
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---

### PATCH /api/v1/projects/<project_code>

**프로젝트 부분 업데이트**

summary: 프로젝트 업데이트
description: 프로젝트의 특정 필드만 업데이트합니다
parameters:
in: path
required: true
type: string
description: 프로젝트 코드
in: body
required: true
schema:
type: object
properties:
name:
type: string
description: 프로젝트명
description:
type: string
description: 프로젝트 설명
status:
type: string
enum: [active, completed, cancelled, on_hold]
description: 프로젝트 상태
end_date:
type: string
format: date
description: 종료일
budget:
type: number
description: 예산
progress:
type: integer
minimum: 0
maximum: 100
description: 진행률
responses:
200:
description: 프로젝트 업데이트 성공
404:
description: 프로젝트를 찾을 수 없음
400:
description: 입력 데이터 유효성 검사 실패

**파라미터:**

- `project_code` (path) (필수): project_code 파라미터

**응답:**

- **200**: 성공
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---

### PUT /api/v1/projects/<project_code>/status

**프로젝트 상태 변경**

summary: 프로젝트 상태 변경
description: 프로젝트의 상태를 변경합니다
parameters:
in: path
required: true
type: string
description: 프로젝트 코드
in: body
required: true
schema:
type: object
properties:
status:
type: string
enum: [active, completed, cancelled, on_hold]
description: 새로운 상태
reason:
type: string
description: 상태 변경 사유
required:
responses:
200:
description: 상태 변경 성공
404:
description: 프로젝트를 찾을 수 없음
400:
description: 잘못된 상태 변경 요청

**파라미터:**

- `project_code` (path) (필수): project_code 파라미터

**응답:**

- **200**: 성공
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---


## name

name 관련 API

### POST /api/v1/projects

**새 프로젝트 생성**

summary: 프로젝트 생성
description: 새로운 프로젝트를 생성합니다
parameters:
in: body
required: true
schema:
type: object
properties:
name:
type: string
description: 프로젝트명 (필수)
example: "웹사이트 리뉴얼"
description:
type: string
description: 프로젝트 설명
example: "회사 홈페이지 전면 리뉴얼 프로젝트"
start_date:
type: string
format: date
description: 시작일 (필수)
example: "2025-01-01"
end_date:
type: string
format: date
description: 종료일
example: "2025-06-30"
budget:
type: number
description: 예산
example: 50000000
required:
responses:
201:
description: 프로젝트 생성 성공
400:
description: 입력 데이터 유효성 검사 실패
409:
description: 중복된 프로젝트명

**요청 본문:**

```json
// Content-Type: application/json
{
  // 요청 데이터 구조
}
```

**응답:**

- **200**: 성공
- **201**: 생성됨
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---


## start_date

start_date 관련 API

### POST /api/v1/projects

**새 프로젝트 생성**

summary: 프로젝트 생성
description: 새로운 프로젝트를 생성합니다
parameters:
in: body
required: true
schema:
type: object
properties:
name:
type: string
description: 프로젝트명 (필수)
example: "웹사이트 리뉴얼"
description:
type: string
description: 프로젝트 설명
example: "회사 홈페이지 전면 리뉴얼 프로젝트"
start_date:
type: string
format: date
description: 시작일 (필수)
example: "2025-01-01"
end_date:
type: string
format: date
description: 종료일
example: "2025-06-30"
budget:
type: number
description: 예산
example: 50000000
required:
responses:
201:
description: 프로젝트 생성 성공
400:
description: 입력 데이터 유효성 검사 실패
409:
description: 중복된 프로젝트명

**요청 본문:**

```json
// Content-Type: application/json
{
  // 요청 데이터 구조
}
```

**응답:**

- **200**: 성공
- **201**: 생성됨
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---


## name: project_code

name: project_code 관련 API

### GET /api/v1/projects/<project_code>

**특정 프로젝트 조회**

summary: 프로젝트 상세 조회
description: 프로젝트 코드로 특정 프로젝트의 상세 정보를 조회합니다
parameters:
in: path
required: true
type: string
description: 프로젝트 코드
example: "PRJ001"
responses:
200:
description: 프로젝트 조회 성공
404:
description: 프로젝트를 찾을 수 없음

**파라미터:**

- `project_code` (path) (필수): project_code 파라미터

**응답:**

- **200**: 성공
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---

### PATCH /api/v1/projects/<project_code>

**프로젝트 부분 업데이트**

summary: 프로젝트 업데이트
description: 프로젝트의 특정 필드만 업데이트합니다
parameters:
in: path
required: true
type: string
description: 프로젝트 코드
in: body
required: true
schema:
type: object
properties:
name:
type: string
description: 프로젝트명
description:
type: string
description: 프로젝트 설명
status:
type: string
enum: [active, completed, cancelled, on_hold]
description: 프로젝트 상태
end_date:
type: string
format: date
description: 종료일
budget:
type: number
description: 예산
progress:
type: integer
minimum: 0
maximum: 100
description: 진행률
responses:
200:
description: 프로젝트 업데이트 성공
404:
description: 프로젝트를 찾을 수 없음
400:
description: 입력 데이터 유효성 검사 실패

**파라미터:**

- `project_code` (path) (필수): project_code 파라미터

**응답:**

- **200**: 성공
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---

### PUT /api/v1/projects/<project_code>/status

**프로젝트 상태 변경**

summary: 프로젝트 상태 변경
description: 프로젝트의 상태를 변경합니다
parameters:
in: path
required: true
type: string
description: 프로젝트 코드
in: body
required: true
schema:
type: object
properties:
status:
type: string
enum: [active, completed, cancelled, on_hold]
description: 새로운 상태
reason:
type: string
description: 상태 변경 사유
required:
responses:
200:
description: 상태 변경 성공
404:
description: 프로젝트를 찾을 수 없음
400:
description: 잘못된 상태 변경 요청

**파라미터:**

- `project_code` (path) (필수): project_code 파라미터

**응답:**

- **200**: 성공
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---


## status

status 관련 API

### PUT /api/v1/projects/<project_code>/status

**프로젝트 상태 변경**

summary: 프로젝트 상태 변경
description: 프로젝트의 상태를 변경합니다
parameters:
in: path
required: true
type: string
description: 프로젝트 코드
in: body
required: true
schema:
type: object
properties:
status:
type: string
enum: [active, completed, cancelled, on_hold]
description: 새로운 상태
reason:
type: string
description: 상태 변경 사유
required:
responses:
200:
description: 상태 변경 성공
404:
description: 프로젝트를 찾을 수 없음
400:
description: 잘못된 상태 변경 요청

**파라미터:**

- `project_code` (path) (필수): project_code 파라미터

**응답:**

- **200**: 성공
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---


## Documentation

Documentation 관련 API

### POST /api/v1/docs/generate

**API 문서 생성**

summary: API 문서 자동 생성
description: 현재 애플리케이션 상태를 분석하여 모든 형태의 API 문서를 생성합니다
requestBody:
content:
application/json:
schema:
type: object
properties:
formats:
type: array
items:
type: string
enum: [json, yaml, markdown, postman, guide]
description: 생성할 문서 형식 (미지정시 모든 형식)
force:
type: boolean
description: 강제 재생성 여부
default: false
responses:
200:
description: 문서 생성 성공
500:
description: 문서 생성 실패

**응답:**

- **200**: 성공
- **201**: 생성됨
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---

### GET /api/v1/docs/status

**문서 생성 상태 조회**

summary: 문서 생성 시스템 상태 확인
description: 현재 문서 생성 시스템의 상태와 통계 정보를 반환합니다
responses:
200:
description: 상태 조회 성공

**응답:**

- **200**: 성공
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---

### POST /api/v1/docs/validate

**문서 유효성 검증**

summary: 생성된 문서의 유효성 검증
description: 현재 애플리케이션 상태와 생성된 문서의 일치성을 검증합니다
responses:
200:
description: 검증 완료

**응답:**

- **200**: 성공
- **201**: 생성됨
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---

### GET /api/v1/docs/download/<file_type>

**문서 파일 다운로드**

summary: 생성된 문서 파일 다운로드
description: 지정된 형식의 문서 파일을 다운로드합니다
parameters:
in: path
required: true
schema:
type: string
enum: [json, yaml, markdown, postman, guide, all]
responses:
200:
description: 파일 다운로드 성공
404:
description: 파일을 찾을 수 없음

**파라미터:**

- `file_type` (path) (필수): file_type 파라미터

**응답:**

- **200**: 성공
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---

### GET /api/v1/docs/preview/<file_type>

**문서 미리보기**

summary: 생성된 문서 파일 미리보기
description: 지정된 형식의 문서 내용을 JSON으로 반환합니다
parameters:
in: path
required: true
schema:
type: string
enum: [json, yaml, markdown, postman, guide]
responses:
200:
description: 미리보기 성공
404:
description: 파일을 찾을 수 없음

**파라미터:**

- `file_type` (path) (필수): file_type 파라미터

**응답:**

- **200**: 성공
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---

### GET /api/v1/docs/analytics

**문서 분석 정보 조회**

summary: API 문서 분석 및 통계 정보
description: 현재 API의 구조 분석 결과와 문서화 통계를 반환합니다
responses:
200:
description: 분석 정보 조회 성공

**응답:**

- **200**: 성공
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---

### PUT /api/v1/docs/config

**문서 시스템 설정 관리**

summary: 문서 생성 시스템 설정 조회/변경
description: 자동 문서 생성 시스템의 설정을 조회하거나 변경합니다

**응답:**

- **200**: 성공
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---

### GET /api/v1/docs/config

**문서 시스템 설정 관리**

summary: 문서 생성 시스템 설정 조회/변경
description: 자동 문서 생성 시스템의 설정을 조회하거나 변경합니다

**응답:**

- **200**: 성공
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---


## name: file_type

name: file_type 관련 API

### GET /api/v1/docs/download/<file_type>

**문서 파일 다운로드**

summary: 생성된 문서 파일 다운로드
description: 지정된 형식의 문서 파일을 다운로드합니다
parameters:
in: path
required: true
schema:
type: string
enum: [json, yaml, markdown, postman, guide, all]
responses:
200:
description: 파일 다운로드 성공
404:
description: 파일을 찾을 수 없음

**파라미터:**

- `file_type` (path) (필수): file_type 파라미터

**응답:**

- **200**: 성공
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---

### GET /api/v1/docs/preview/<file_type>

**문서 미리보기**

summary: 생성된 문서 파일 미리보기
description: 지정된 형식의 문서 내용을 JSON으로 반환합니다
parameters:
in: path
required: true
schema:
type: string
enum: [json, yaml, markdown, postman, guide]
responses:
200:
description: 미리보기 성공
404:
description: 파일을 찾을 수 없음

**파라미터:**

- `file_type` (path) (필수): file_type 파라미터

**응답:**

- **200**: 성공
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---


## API Versioning

API Versioning 관련 API

### GET /api/versions

**모든 API 버전 목록 조회**

summary: List all API versions
description: Get information about all available API versions
responses:
200:
description: List of API versions

**응답:**

- **200**: 성공
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류

---

### GET /api/version/current

**현재 요청의 API 버전 정보**

summary: Get current API version
description: Get information about the API version being used for this request
responses:
200:
description: Current API version information

**응답:**

- **200**: 성공
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---


## Monitoring

시스템 모니터링 관련 API

### GET /api/monitoring/usage

**API 사용량 통계 조회 (관리자 전용)**

/api/monitoring/usage 경로에서 데이터를 조회합니다.

**응답:**

- **200**: 성공
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---

### GET /api/monitoring/prefetch

**백그라운드 프리패치 통계 조회 (관리자 전용)**

/api/monitoring/prefetch 경로에서 데이터를 조회합니다.

**응답:**

- **200**: 성공
- **400**: 잘못된 요청
- **404**: 찾을 수 없음
- **500**: 서버 오류
- **401**: 인증 필요
- **403**: 권한 없음

---


## 데이터 스키마

### StandardResponse

**속성:**

- `success` (boolean) (필수): 요청 성공 여부
- `data` (unknown) (필수): 응답 데이터
- `error` (unknown) (필수): 
- `meta` (unknown) (필수): 

### ErrorResponse


### ErrorDetails

**속성:**

- `code` (string) (필수): 에러 코드
- `message` (string) (필수): 에러 메시지
- `details` (object): 추가 에러 정보

### ResponseMeta

**속성:**

- `timestamp` (string) (필수): 응답 시간
- `version` (string) (필수): API 버전
- `request_id` (string) (필수): 요청 ID

