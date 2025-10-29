# 캐시 계약 문서 (Cache Contract)

이 문서는 시스템 전체의 캐시 키, TTL 정책, 무효화 규칙을 정의합니다.

---

## 📋 캐시 전략 (Cache Strategy)

### CacheStrategy Enum

| 전략 | TTL | 용도 | 예시 |
|------|-----|------|------|
| `CRITICAL_DATA` | 300초 (5분) | 핵심 비즈니스 데이터 | 프로젝트 목록, 프로젝트 상세 |
| `METADATA` | 600초 (10분) | 메타데이터 | 담당자 목록, 사업자 목록, 거래처 목록 |
| `STATIC_CONFIG` | 3600초 (1시간) | 설정 데이터 | 환경 설정, 권한 정보 |
| `UI_STATE` | 300초 (5분) | UI 상태 | 필터 설정, 정렬 옵션 |
| `TEMPORARY` | 60초 (1분) | 임시 데이터 | 셀 노트, 리드 목록 |
| `FOLDER_MAPPING` | 86400초 (24시간) | 폴더 매핑 | Google Drive 폴더 ID |

---

## 🔑 캐시 키 네이밍 규칙

### 1. 프로젝트 관련

```python
# 전체 프로젝트 목록 (DataFrame)
"current_sheet_data"

# 개별 프로젝트 상세 (추후 구현 시)
"project_{project_code}"  # 예: "project_IT-2024-001"

# 프로젝트 필터링 결과 (추후 구현 시)
"projects_filter_{hash}"  # 예: "projects_filter_abc123"
```

### 2. 셀 노트 관련

```python
# 전체 셀 노트 (Google Sheets 기준)
"cell_notes_{sheet_id}"  # 예: "cell_notes_1a2b3c4d5e"

# 특정 행의 노트 모음 (임시)
"notes_row_{sheet_id}_{row_num}"  # 예: "notes_row_1a2b3c_10"
```

### 3. 메타데이터 관련

```python
# 담당자/사업자/거래처 메타데이터
"metadata_{type}"  # 예: "metadata_담당자", "metadata_사업자"

# 사용자 목록
"user_list"

# 권한 정보
"permissions_{user_email}"
```

### 4. 리드 관련

```python
# 리드 목록
"leads_list"

# 특정 리드 상세
"lead_{lead_no}"
```

### 5. 폴더 관련

```python
# Google Drive 폴더 ID
"folder_id_{folder_type}"  # 예: "folder_id_projects"
```

---

## ♻️ 캐시 무효화 규칙

### 1. 프로젝트 생성/수정/삭제 시

**무효화 대상**:
- `current_sheet_data` (전체 목록)
- `cell_notes_{sheet_id}` (셀 노트)
- `metadata_*` (메타데이터, 담당자/사업자 변경 시)

**호출 방법**:
```python
from dashboard.services.project_service import invalidate_project_cache

# 전체 캐시 무효화 + 백그라운드 갱신 트리거
invalidate_project_cache()

# 특정 프로젝트만 무효화 (추후 구현)
invalidate_project_cache(project_code="IT-2024-001")

# 무효화만 하고 갱신 트리거 안함
invalidate_project_cache(trigger_refresh=False)
```

### 2. 메모 저장 시

**무효화 대상**:
- `cell_notes_{sheet_id}` (셀 노트만)
- `current_sheet_data`는 무효화하지 않음 (메모는 별도 저장)

**호출 방법**:
```python
from dashboard.utils.cache_invalidation import invalidate_cache

invalidate_cache(f"cell_notes_{sheet_id}")
```

### 3. 메타데이터 변경 시

**무효화 대상**:
- `metadata_{type}` (해당 메타데이터만)

**호출 방법**:
```python
invalidate_cache(f"metadata_{metadata_type}")
```

### 4. 리드 상태 변경 시

**무효화 대상**:
- `leads_list` (전체 리드 목록)
- `lead_{lead_no}` (특정 리드)

---

## 📖 사용 예시

### 1. 데이터 조회 (smart_get)

```python
from dashboard.utils.smart_cache_manager import smart_get, CacheStrategy

# 프로젝트 목록 조회 (캐시 우선)
df = smart_get("current_sheet_data", CacheStrategy.CRITICAL_DATA)

if df is None:
    # 캐시 없으면 Google Sheets에서 로드
    df = load_data_from_sheets()
    smart_set("current_sheet_data", df, CacheStrategy.CRITICAL_DATA)
```

### 2. 데이터 저장 (smart_set)

```python
from dashboard.utils.smart_cache_manager import smart_set, CacheStrategy

# 프로젝트 목록 캐싱 (5분 TTL)
smart_set("current_sheet_data", df, CacheStrategy.CRITICAL_DATA)

# 메타데이터 캐싱 (10분 TTL)
smart_set("metadata_담당자", managers_list, CacheStrategy.METADATA)

# 임시 데이터 캐싱 (1분 TTL)
smart_set("notes_row_123", notes, CacheStrategy.TEMPORARY)
```

### 3. 캐시 무효화 (invalidate_cache)

```python
from dashboard.utils.cache_invalidation import invalidate_cache

# 단일 키 무효화
invalidate_cache("current_sheet_data")

# 패턴 무효화 (와일드카드)
invalidate_cache("cell_notes_*")

# 여러 키 무효화
invalidate_cache(["current_sheet_data", "metadata_담당자"])
```

### 4. 프로젝트 캐시 무효화 (통합 헬퍼)

```python
from dashboard.services.project_service import invalidate_project_cache

# 프로젝트 생성 후
def add_project_auto():
    # ... 프로젝트 생성 로직 ...
    invalidate_project_cache()  # 전체 캐시 무효화 + 백그라운드 갱신

# 프로젝트 수정 후
def update_project():
    # ... 프로젝트 수정 로직 ...
    invalidate_project_cache(project_code)  # 특정 프로젝트 무효화

# 메모만 수정 후
def save_field_memo():
    # ... 메모 저장 로직 ...
    invalidate_cache(f"cell_notes_{sheet_id}")  # 셀 노트만 무효화
```

---

## ⚠️ 주의사항

### 1. 캐시 키 하드코딩 금지

❌ **나쁜 예**:
```python
df = smart_get("current_sheet_data", CacheStrategy.CRITICAL_DATA)
```

✅ **좋은 예** (추후 개선):
```python
from dashboard.utils.cache_keys import CACHE_KEYS

df = smart_get(CACHE_KEYS.PROJECT_LIST, CacheStrategy.CRITICAL_DATA)
```

### 2. 무효화 누락 방지

프로젝트 데이터를 변경하는 모든 API는 반드시 `invalidate_project_cache()`를 호출해야 합니다.

**체크리스트**:
- ✅ 프로젝트 생성 (add_project_auto)
- ✅ 프로젝트 수정 (update_project, update_project_inline)
- ✅ 프로젝트 삭제 (delete_project)
- ✅ 프로젝트 취소/재개 (cancel_project_api, resume_project_api)
- ✅ 필드 배치 수정 (save_field_memos_batch - 메모만 해당)

### 3. TTL과 백그라운드 갱신 조화

- **CRITICAL_DATA (5분 TTL)**: 백그라운드 프리로더가 4분마다 갱신
- **METADATA (10분 TTL)**: 수동 무효화 의존
- **TEMPORARY (1분 TTL)**: 짧은 수명, 자주 갱신

### 4. 캐시 미스 대응

캐시가 없을 때는 항상 원본 소스(Google Sheets, DB)에서 로드하고 캐싱:

```python
df = smart_get("current_sheet_data", CacheStrategy.CRITICAL_DATA)

if df is None:
    # Fallback: 원본 소스에서 로드
    df = load_data_from_sheets()
    smart_set("current_sheet_data", df, CacheStrategy.CRITICAL_DATA)
```

---

## 📊 캐시 성능 지표

### 현재 성능

- **API 호출 감소**: 75% (백그라운드 프리로더 효과)
- **평균 응답 시간**: 캐시 히트 시 50ms 이내
- **캐시 히트율**: 85-90% (CRITICAL_DATA 기준)

### 모니터링 포인트

1. **캐시 히트율**: `smart_cache_manager`에서 자동 추적
2. **무효화 빈도**: 너무 빈번하면 TTL 조정 고려
3. **메모리 사용량**: Redis 메모리 모니터링

---

## 🔄 변경 이력

### 2025-10-29
- 초기 문서 작성
- 캐시 전략, 키 네이밍 규칙, 무효화 규칙 정의
- 사용 예시 및 주의사항 추가

---

## 📚 참고 문서

- `dashboard/utils/smart_cache_manager.py` - 캐시 매니저 구현
- `dashboard/utils/cache_invalidation.py` - 캐시 무효화 서비스
- `dashboard/services/project_service.py` - 프로젝트 캐시 무효화 헬퍼
- `dashboard/utils/background_prefetch.py` - 백그라운드 프리로더
