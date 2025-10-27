# ITGlobalApp API 사용 가이드

## 개요

`window.ITGlobalApp`는 프로젝트 관리 시스템의 통합 API 네임스페이스입니다.
기존의 전역 함수들(`refreshProjectData`, `openNewProjectModal` 등)을 대체하며, 더 안전하고 체계적인 API 구조를 제공합니다.

**도입 배경**:
- 전역 네임스페이스 오염 방지
- API 구조 명확화 및 유지보수성 향상
- TypeScript 친화적인 구조 제공
- 레거시 코드와의 호환성 유지 (deprecated 경고와 함께)

---

## 빠른 시작

### 권장 사용법 (New)

```javascript
// 데이터 새로고침
ITGlobalApp.refreshData(force, showMessage);

// 새 프로젝트 모달 열기
ITGlobalApp.openNewProjectModal();

// 감사 로그 모달 열기
ITGlobalApp.openAuditLogs();

// 현재 데이터 가져오기
const data = ITGlobalApp.getCurrentData();

// 캐시 지우기
ITGlobalApp.clearCache();
```

### 레거시 방식 (Deprecated)

```javascript
// ⚠️ DEPRECATED: 첫 호출 시 경고 메시지가 표시됩니다
refreshProjectData(force, showMessage);  // → ITGlobalApp.refreshData 사용 권장
openNewProjectModal();                    // → ITGlobalApp.openNewProjectModal 사용 권장
loadAuditLogs();                          // → ITGlobalApp.openAuditLogs 사용 권장
```

**Throttling**: 각 레거시 함수는 **한 번만** deprecated 경고를 출력하므로 콘솔이 도배되지 않습니다.

---

## API 레퍼런스

### 1. 데이터 관련 API

| 함수 | 설명 | 파라미터 | 반환값 |
|------|------|----------|--------|
| `ITGlobalApp.refreshData(force, showMessage)` | 프로젝트 데이터 새로고침 | `force`: 강제 새로고침 (boolean)<br>`showMessage`: 성공 메시지 표시 (boolean) | Promise |
| `ITGlobalApp.getCurrentData()` | 현재 프로젝트 데이터 가져오기 | - | Array |
| `ITGlobalApp.api.getFilteredData()` | 필터링된 데이터 가져오기 | - | Array |
| `ITGlobalApp.clearCache()` | 로컬 캐시 지우기 | - | void |
| `ITGlobalApp.api.getCachedData()` | 캐시된 데이터 가져오기 | - | Object |

### 2. 모달 관련 API

| 함수 | 설명 | 파라미터 | 반환값 |
|------|------|----------|--------|
| `ITGlobalApp.openNewProjectModal()` | 신규 프로젝트 모달 열기 | - | void |
| `ITGlobalApp.openAuditLogs()` | 감사 로그 모달 열기 | - | void |
| `ITGlobalApp.showProjectDetails(projectCode)` | 프로젝트 상세 정보 표시 | `projectCode`: 프로젝트 코드 (string) | void |

### 3. UI 관련 API

| 함수 | 설명 | 파라미터 | 반환값 |
|------|------|----------|--------|
| `ITGlobalApp.api.showSuccess(message)` | 성공 메시지 표시 | `message`: 메시지 내용 (string) | void |
| `ITGlobalApp.api.showError(message)` | 에러 메시지 표시 | `message`: 메시지 내용 (string) | void |

### 4. 시스템 정보 API

| 함수 | 설명 | 파라미터 | 반환값 |
|------|------|----------|--------|
| `ITGlobalApp.api.getComponents()` | 컴포넌트 인스턴스 가져오기 | - | Object |
| `ITGlobalApp.api.getPermissions()` | 사용자 권한 정보 가져오기 | - | Object |
| `ITGlobalApp.api.getSocket()` | 소켓 인스턴스 가져오기 | - | SocketManager |

---

## 사용 예시

### 예시 1: 데이터 새로고침 후 특정 작업 수행

```javascript
// 강제 새로고침 + 성공 메시지 표시
ITGlobalApp.refreshData(true, true)
  .then(() => {
    console.log('데이터 새로고침 완료');
    // 추가 작업 수행
  })
  .catch(error => {
    console.error('새로고침 실패:', error);
  });
```

### 예시 2: 버튼 클릭 시 모달 열기

```html
<!-- 권장 방식 -->
<button onclick="ITGlobalApp.openNewProjectModal()">
  새 프로젝트 추가
</button>

<!-- 레거시 방식 (deprecated 경고 발생) -->
<button onclick="openNewProjectModal()">
  새 프로젝트 추가
</button>
```

### 예시 3: 프로젝트 데이터 필터링

```javascript
// 현재 데이터 가져오기
const allProjects = ITGlobalApp.getCurrentData();

// 필터링된 데이터 가져오기 (검색/필터 적용 후)
const filteredProjects = ITGlobalApp.api.getFilteredData();

console.log(`전체: ${allProjects.length}개, 필터링: ${filteredProjects.length}개`);
```

### 예시 4: 사용자 권한에 따른 UI 제어

```javascript
const permissions = ITGlobalApp.api.getPermissions();

if (permissions.canEdit) {
  document.getElementById('edit-btn').style.display = 'block';
} else {
  document.getElementById('edit-btn').style.display = 'none';
}
```

### 예시 5: 컴포넌트 직접 접근 (고급)

```javascript
const components = ITGlobalApp.api.getComponents();

// ModernProjectModal 컴포넌트에 직접 접근
if (components.modernModal) {
  components.modernModal.open();
}

// AuditLogModal 컴포넌트에 직접 접근
if (components.auditLogModal) {
  components.auditLogModal.open();
}
```

---

## 마이그레이션 가이드

### 기존 코드 → 새 코드

| 기존 (Deprecated) | 새 코드 (Recommended) |
|-------------------|------------------------|
| `refreshProjectData(true, true)` | `ITGlobalApp.refreshData(true, true)` |
| `getCurrentProjectData()` | `ITGlobalApp.getCurrentData()` |
| `clearProjectCache()` | `ITGlobalApp.clearCache()` |
| `openNewProjectModal()` | `ITGlobalApp.openNewProjectModal()` |
| `loadAuditLogs()` | `ITGlobalApp.openAuditLogs()` |
| `openAuditLogs()` | `ITGlobalApp.openAuditLogs()` |
| `showProjectDetails(code)` | `ITGlobalApp.showProjectDetails(code)` |

### 마이그레이션 전략

1. **점진적 마이그레이션**: 레거시 함수는 당분간 유지되므로 급하게 변경할 필요 없음
2. **신규 코드 작성 시**: 무조건 `ITGlobalApp` 사용
3. **기존 코드 수정 시**: 기회가 될 때마다 새 API로 변경
4. **콘솔 경고 확인**: Deprecated 경고가 발생하면 해당 부분 리팩토링

---

## 구조 상세

### ITGlobalApp 네임스페이스 구조

```javascript
window.ITGlobalApp = {
  // 편의 함수 (직접 접근 가능)
  refreshData: Function,
  getCurrentData: Function,
  clearCache: Function,
  openNewProjectModal: Function,
  openAuditLogs: Function,
  showProjectDetails: Function,

  // API 객체 (더 많은 기능)
  api: {
    refreshData: Function,
    getCurrentData: Function,
    getFilteredData: Function,
    clearCache: Function,
    getCachedData: Function,
    showSuccess: Function,
    showError: Function,
    getComponents: Function,
    getPermissions: Function,
    getSocket: Function
  },

  // 모달 객체
  modals: {
    openNewProjectModal: Function,
    openAuditLogs: Function,
    showProjectDetails: Function
  }
};
```

### TypeScript 타입 정의 (참고용)

```typescript
interface ITGlobalApp {
  // 편의 함수
  refreshData(force?: boolean, showMessage?: boolean): Promise<void>;
  getCurrentData(): Project[];
  clearCache(): void;
  openNewProjectModal(): void;
  openAuditLogs(): void;
  showProjectDetails(projectCode: string): void;

  // API 객체
  api: {
    refreshData(force?: boolean, internal?: boolean, showMessage?: boolean): Promise<void>;
    getCurrentData(): Project[];
    getFilteredData(): Project[];
    clearCache(): void;
    getCachedData(): CachedData | null;
    showSuccess(message: string): void;
    showError(message: string): void;
    getComponents(): Components;
    getPermissions(): UserPermissions;
    getSocket(): SocketManager;
  };

  // 모달 객체
  modals: {
    openNewProjectModal(): void;
    openAuditLogs(): void;
    showProjectDetails(projectCode: string): void;
  };
}

declare global {
  interface Window {
    ITGlobalApp: ITGlobalApp;
  }
}
```

---

## 주의사항

### 1. 전역 변수 오염 방지

❌ **하지 말 것**:
```javascript
window.myFunction = () => { /* ... */ };
window.myData = [];
```

✅ **권장**:
```javascript
// 기존 ITGlobalApp에 속성 추가 (확장)
ITGlobalApp.myCustomFunction = () => { /* ... */ };
```

### 2. Deprecated 함수 사용

- 레거시 함수는 **첫 호출 시 한 번만** 경고 출력
- 콘솔에 `[DEPRECATED]` 메시지가 보이면 해당 코드 리팩토링 권장
- 긴급하지 않으므로 점진적으로 개선

### 3. API 안정성

- `ITGlobalApp.api.*` 함수들은 안정적인 API로 간주
- Breaking change 발생 시 최소 1개월 전에 deprecation 공지
- 하위 호환성 유지 원칙

---

## 문제 해결

### Q1. `ITGlobalApp is not defined` 에러 발생

**원인**: `ProjectListAPI.js`가 로드되기 전에 호출
**해결**: DOM 로드 완료 후 호출

```javascript
// ❌ 잘못된 예
<script>ITGlobalApp.refreshData();</script>

// ✅ 올바른 예
<script>
  document.addEventListener('DOMContentLoaded', () => {
    ITGlobalApp.refreshData();
  });
</script>
```

### Q2. Deprecated 경고가 계속 나타남

**원인**: 레거시 함수 호출 중
**해결**: 새 API로 변경

```javascript
// 경고 발생 코드
openNewProjectModal();

// 해결
ITGlobalApp.openNewProjectModal();
```

### Q3. 컴포넌트가 undefined

**원인**: 컴포넌트 초기화 전에 접근
**해결**: 존재 여부 확인 후 사용

```javascript
const components = ITGlobalApp.api.getComponents();
if (components.modernModal) {
  components.modernModal.open();
} else {
  console.error('ModernModal이 아직 로드되지 않았습니다.');
}
```

---

## 추가 자료

- **소스 코드**: `dashboard/src/js/api/ProjectListAPI.js`
- **관련 문서**: `docs/보고서_완료작업_및_개선사항.md`
- **커밋 히스토리**: `c675d3b` (ITGlobalApp 네임스페이스 도입)

---

**최종 업데이트**: 2025-10-21
**작성자**: Claude Code
**버전**: 1.0.0
