# 모던 시스템 클래스명 추출 및 레거시 매핑 분석

## 📋 모던 시스템에서 추출된 클래스명들

### HTML 템플릿에서 추출 (modern_project_list.html)
```css
/* 레이아웃 관련 */
.container-fluid
.content-container
.row, .mb-2, .col-12
.d-flex, .align-items-center, .justify-content-between

/* 프로젝트 헤더 */
.project-header
.header-alert-container
.cache-status-container
.cache-status
.cache-refresh-btn

/* 버튼 */
.new-project-btn
.btn-soft-primary
.btn-outline-secondary

/* 테이블 */
.table-section
.table-wrapper
.projects-table
.table, .table-hover

/* 모바일 */
.mobile-card-container

/* 모달 */
.modal, .fade
.modal-dialog, .modal-lg
.modal-content, .modal-header, .modal-body, .modal-footer
.modal-title
.btn-close
```

### JS 컴포넌트에서 추출
```css
/* 액션 버튼 */
.btn-group, .btn-group-${size}
.btn-more-actions
.dropdown-toggle, .dropdown-menu, .dropdown-item

/* 상태 관련 */
.loading
.collection-status-badge
.status-badge-icon, .status-badge-text

/* 인라인 편집 */
.inline-editor, .inline-edit-input

/* 모바일 카드 */
.mobile-card-view

/* 아코디언 */
.project-accordion-container

/* 가상화 */
.virtual-scroll-area, .virtual-scroll-visible
.loading-placeholder

/* 토스트 */
.toast-container

/* 필드 락 */
.lock-indicator-icon
.lockingIndicatorClass, .lockIndicatorClass
.lockedByOtherClass

/* 알림 */
.alert, .alert-dismissible, .fade, .show
.custom-tooltiptext
```

## 🔍 레거시 vs 모던 매핑 분석

### ✅ 올바르게 매핑된 클래스들
| 기능 | 레거시 클래스 | 모던 클래스 | 상태 |
|------|-------------|------------|------|
| 테이블 래퍼 | `.table-wrapper` | `.table-wrapper` | ✅ 일치 |
| 모바일 카드 | `.mobile-card-container` | `.mobile-card-container` | ✅ 일치 |
| 모달 시스템 | `.modal`, `.modal-*` | `.modal`, `.modal-*` | ✅ 일치 |
| 버튼 그룹 | `.btn-group` | `.btn-group` | ✅ 일치 |

### ❌ 누락된 레거시 클래스들
| 레거시 클래스 | 기능 | 모던에서 찾음? | 대체 방안 |
|-------------|------|------------|---------|
| `.btn-soft-*` | 소프트 버튼 스타일 | ❌ | `design-system/buttons.css`에 추가 필요 |
| `.user-menu-*` | 사용자 메뉴 | ❌ | 새 모듈 `user-dropdown.css` 필요 |
| `.filter-*` | 필터 시스템 | 부분적 | `filter-system.css` 보강 필요 |
| `.accordion-*` | 아코디언 시스템 | 부분적 | 새 모듈 `accordion-system.css` 필요 |
| `.badge-category-*` | 구분 뱃지 | ❌ | `badges.css` 보강 필요 |
| `.amount-*` | 금액 표시 | ❌ | 새 모듈 필요 |
| `.datatable-*` | DataTables 통합 | ❌ | 새 모듈 필요 |

### ⚠️ 불완전한 매핑들
| 기능 | 문제점 | 해결 방안 |
|------|-------|---------|
| 상태 뱃지 | 일부 상태만 구현 | `badges.css` 확장 |
| 애니메이션 | 트랜지션 효과 미구현 | `animations.css` 보강 |
| 반응형 디자인 | 브레이크포인트 불일치 | `responsive.css` 재검토 |

## 🚨 중요한 발견사항들

### 1. CSS 변수 불일치
**레거시:**
```css
:root {
  --primary-color: #FF6B35;
  --success-color: #10b981;
  --warning-color: #f59e0b;
}
```

**모던 확인 필요:**
- `design-system/tokens.css`의 변수들과 비교 필요

### 2. 누락된 글로벌 스타일들
```css
/* 레거시에만 있는 중요한 스타일들 */
* {
  user-select: text;
  caret-color: transparent !important;
}

.inline-edit-input:focus {
  caret-color: auto !important;
}
```

### 3. 사용자 정의 클래스들 (모던에서 누락)
- `.btn-soft-primary`, `.btn-soft-secondary` 등
- `.user-dropdown-*` 관련 모든 클래스
- `.cache-status-*` (새로 추가된 기능)
- `.project-accordion-*` (부분적으로만 구현)

## 📋 우선순위별 액션 아이템

### 🚨 Critical (즉시 구현 필요)
1. **글로벌 커서 설정** - `global-cursor.css` 생성
2. **소프트 버튼** - `buttons.css`에 `.btn-soft-*` 추가
3. **사용자 드롭다운** - `user-dropdown.css` 생성
4. **CSS 변수 동기화** - `tokens.css` 업데이트

### ⚠️ High (조기 구현 권장)
1. **아코디언 시스템** - `accordion-system.css` 생성
2. **구분 뱃지** - `badges.css` 확장
3. **DataTables 통합** - `datatable-integration.css` 생성
4. **금액 포맷** - `currency-formatting.css` 생성

### 📝 Medium (점진적 개선)
1. **애니메이션 효과** - `animations.css` 보강
2. **반응형 개선** - `responsive.css` 재검토
3. **성능 최적화** - CSS 모듈 최적화

## 🔧 다음 단계
1. 누락된 모듈들을 순차적으로 생성
2. 기존 모듈들의 불완전한 부분 보강
3. 클래스명 일관성 검증 스크립트 작성
4. 시각적 회귀 테스트 구현