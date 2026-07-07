# 리팩토링 백로그 (2026-07-08 코드 리뷰)

Agent D 리포트에서 발견된 중복·개선 여지. 실서비스 개시 임박 시점이라 즉시 리팩토링은 회귀 리스크. 안정화 이후 순차 처리.

## P2: 중복 헬퍼 통합

### 슬랙 봇 헬퍼 중복
- **파일**: `dashboard/blueprints/slack_bot.py`, `dashboard/blueprints/slack_helpers.py`
- **중복 지점**:
  - 모달 state 추출 로직 (여러 handler에서 유사 코드)
  - 시간 포맷팅 (KST 변환 · relative time)
  - 이니셜 매핑 (`_to_initial` vs `_slack_user_to_initial` 부분 겹침)
- **조치안**: `slack_helpers.py`에 헬퍼 통합 → slack_bot.py는 그 헬퍼만 호출
- **소요**: 1~2시간

### get_project_records 반복 호출
- **파일**: 5개 (`admin.py`, `leads.py`, `projects.py`, `calendar_sync_scheduler.py`, `test_project_service.py`)
- **총 8건 호출**
- **패턴**: 각 endpoint에서 `records = get_project_records()` → filter → 응답
- **조치안**: `project_query_helpers.py` 신규 — `get_by_code`, `filter_by_status` 등 helper.  
  각 호출 지점 단순화.
- **소요**: 30~45분

### 에러 응답 스키마 통일 (97건)
- Task 33에서 문서화 완료 ([api-response-schema.md](api-response-schema.md))
- 실제 순차 변환은 이 백로그의 별도 대사이클
- **소요**: 2~3시간 (파일별 프론트 검증 병행)

## P3: 코드 청소

### 로거 초기화 통일
- **현상**: `logger = logging.getLogger(__name__)` (표준) vs `logger = get_logger(__name__)` (커스텀) 혼용
- **조치안**: `get_logger()`로 통일. `logging_config.py` 의존.
- **소요**: 20분 (sed 자동화 가능, 검증 함께)

### docstring 스타일 통일
- **현상**: NumPy 스타일과 짧은 한 줄 요약 혼재
- **조치안**: `dashboard/services/`, `dashboard/utils/` 위주로 NumPy 스타일 통일
- **소요**: 1시간 (파일별 순차)

### TODO/FIXME 정리
- **총 6건 발견** (slack_bot.py 3, lead_helpers.py 2, channeltalk.py 1)
- **조치안**: 각 TODO 내용 확인 → GitHub issue 변환 or 즉시 해결
- **소요**: 30분 조사 + 개별 fix 시간

### CSP `unsafe-inline` 제거
- **파일**: `dashboard/utils/security_middleware.py:319`
- **현재**: Content-Security-Policy에 `'unsafe-inline'` 포함 (개발 편의)
- **조치안**: 모든 인라인 script/style을 별도 파일로 → nonce 기반 CSP
- **소요**: 반나절 (프론트 리팩토링 필요)

## P3: 기타

### API 명세 문서
- **문서화 대상**: `/api/*` 모든 endpoint, request/response 스키마
- **조치안**: `docs/api-reference.md` 신규 or Swagger 자동 생성
- **소요**: 1~2시간 (자동 생성 사용 시), 반나절 (수기)

### 스케줄러 job 문서
- **문서화 대상**: APScheduler 각 job (당근·홈페이지·payment 등) 주기·목적·실패 시 재시도
- **조치안**: `docs/scheduler-jobs.md` 신규
- **소요**: 1시간

### 부하 테스트 실행
- 이미 [`docs/deployment/load-test.md`](load-test.md) + `locustfile.py` 준비됨
- 실서비스 매니저 온보딩 며칠 전에 실측 검증
- **소요**: 실행 30분 + 결과 분석 30분

## 우선순위 요약

**실서비스 첫 2주 안 처리 권장**:
1. 슬랙 봇 헬퍼 중복 통합 (안정성·유지보수)
2. 에러 응답 스키마 통일 (프론트-백 계약 명확화)
3. `/api/monitoring/scheduler-status`는 이미 완료 ✓

**첫 달 안 처리**:
4. get_project_records helper 추출
5. 로거 초기화 통일
6. TODO 정리

**분기별**:
7. CSP unsafe-inline 제거
8. API 명세 문서 자동화
9. docstring 스타일 통일

## 참고

- 각 리팩토링 전에 `test_core_flows.py` 실행으로 회귀 검증
- 큰 변경은 별도 브랜치 → PR 리뷰 (혼자 개발이라도)
- 실서비스 트래픽 있는 시점엔 배포 직후 관측 강화
