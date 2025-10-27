# 레거시 CSS 전체 분석 보고서

## 📊 개요
- **파일**: `dashboard/static/legacy/css/project-list.css`
- **크기**: 2,738줄
- **특징**: 모든 스타일이 하나의 파일에 집중된 모놀리식 구조

## 🎯 주요 기능 섹션 분석

### 1. 디자인 토큰 및 변수 (1-77줄)
```css
:root {
    /* Brand Colors */
    --primary-color: #FF6B35;
    --primary-dark: #E55A2B;
    --primary-light: #FF8A65;

    /* Dark Theme Base Colors */
    --dark-primary: #1a1a1a;
    --dark-secondary: #2d2d2d;
    --dark-accent: #3a3a3a;

    /* Status Colors */
    --success-color: #10b981;
    --warning-color: #f59e0b;
    --danger-color: #ef4444;
    --info-color: #06b6d4;

    /* Typography, Spacing, Shadows */
}
```

### 2. 글로벌 스타일 (78-100줄)
- 기본 body 스타일
- 컨테이너 여백 설정
- 텍스트 선택 및 커서 설정

### 3. 헤더 시스템 (96-168줄)
- 헤더 배경 그라데이션
- 장식 요소들

### 4. 버튼 시스템 (169-249줄)
- Soft Button 스타일
- 다양한 버튼 변형들

### 5. 사용자 메뉴 & 드롭다운 (250-369줄)
- 사용자 메뉴 버튼
- 드롭다운 스타일링
- z-index 관리

### 6. 콘텐츠 컨테이너 (370-444줄)
- 메인 콘텐츠 영역
- 구분 뱃지 스타일들

### 7. 필터 시스템 (445-533줄)
- 필터 섹션 레이아웃
- 필터 버튼들

### 8. 폼 엘리먼트 (534-567줄)
- 입력 필드 스타일
- 필터 선택 강조

### 9. 테이블 시스템 (568-876줄)
- 테이블 레이아웃
- DataTables 통합
- 헤더 정렬
- 셀 스타일링

### 10. 상태 뱃지 (877-925줄)
- 상태별 색상 시스템

### 11. 금액 표시 (926-940줄)
- 통화 포맷

### 12. 액션 버튼 (941-961줄)
- 테이블 내 액션 버튼들

### 13. 아코디언 시스템 (962-1124줄)
- 행 세부사항 표시
- 아코디언 헤더
- 정보 카드들

### 14. 모달 시스템 (1125-계속...)
- 모달 스타일링
- 폼 체크 라벨들

## 🔍 모던 시스템과의 비교 분석

### ✅ 이미 모던 모듈에 존재하는 기능
| 레거시 기능 | 모던 모듈 위치 | 상태 |
|------------|---------------|------|
| 디자인 토큰 | `design-system/tokens.css` | ✅ 부분적 |
| 버튼 시스템 | `design-system/buttons.css` | ✅ 부분적 |
| 애니메이션 | `design-system/animations.css` | ✅ 부분적 |
| 뱃지 시스템 | `design-system/badges.css` | ✅ 부분적 |
| 모바일 카드 | `components/mobile-card-view.css` | ✅ 완료 |
| 테이블 | `components/project-table.css` | ✅ 부분적 |
| 모달 | `components/modal-system.css` | ✅ 부분적 |
| 알림 | `components/notification-system.css` | ✅ 부분적 |

### ❌ 누락된 중요 기능들
| 기능 | 레거시 위치 | 모던 모듈 | 상태 |
|------|------------|----------|------|
| 사용자 메뉴 드롭다운 | 250-369줄 | 없음 | ❌ 누락 |
| 아코디언 시스템 | 962-1124줄 | 없음 | ❌ 누락 |
| 구분 뱃지 (4가지 색상) | 379-443줄 | 부분적 | ⚠️ 불완전 |
| 필터 시스템 강조 | 545-552줄 | 부분적 | ⚠️ 불완전 |
| DataTables 통합 | 609-876줄 | 없음 | ❌ 누락 |
| 금액 표시 포맷 | 926-940줄 | 없음 | ❌ 누락 |
| 글로벌 커서 설정 | 3-24줄 | 없음 | ❌ 누락 |

## 🚨 발견된 문제점들

### 1. 모놀리식 구조
- 모든 스타일이 하나의 파일에 집중
- 유지보수의 어려움
- 선택적 로딩 불가능

### 2. 하드코딩된 값들
- Magic number 남용
- 반응형 대응 어려움

### 3. CSS 변수 불일치
- 레거시: `--primary-color: #FF6B35`
- 모던: 확인 필요

## 📋 다음 단계 액션 아이템

1. **즉시 필요한 누락 모듈들**:
   - `user-dropdown.css`
   - `accordion-system.css`
   - `datatable-integration.css`
   - `currency-formatting.css`
   - `global-cursor.css`

2. **불완전한 모듈 보강**:
   - `filter-system.css` 강화
   - `badges.css` 구분 뱃지 추가
   - `tokens.css` 변수 동기화

3. **검증이 필요한 모듈들**:
   - 모든 기존 모듈의 완전성 확인
   - 클래스명 매핑 테이블 생성