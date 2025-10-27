# 공사 취소 기능 구현 가이드

**작성일**: 2025-01-15
**상태**: 백엔드 완료, 프론트엔드 수동 작업 필요

---

## 1. 완료된 작업

### ✅ 백엔드 API (inline_update.py)

두 개의 새 엔드포인트 추가 완료:

#### `/api/project/cancel` - 공사 취소
```python
- 공사상태를 '취소'로 변경
- 수금완료를 FALSE로 변경
- 공사확정일을 빈 값으로 초기화
- 구글 시트 행 배경색을 진한 회색(#808080)으로 변경
```

#### `/api/project/resume` - 공사 재개
```python
- 공사상태를 '공사진행'으로 변경
- 구글 시트 행 배경색을 흰색으로 복원
```

---

## 2. 수동 작업이 필요한 파일들

### 📝 A. ProjectRowAccordion.js (5,645줄)

파일 위치: `dashboard/src/js/components/ProjectRowAccordion.js`

#### 작업 1: generateCancelResumeButton() 함수 수정 (Line 952-972)

**현재 코드:**
```javascript
generateCancelResumeButton(projectCode, rowData) {
  const status = rowData['공사상태'] || '';

  if (status === '공사중단') {
    return `
      <button type="button" class="btn btn-success btn-sm resume-project-btn"
              data-project-code="${projectCode}" title="공사 재개">
        <i class="fas fa-play me-1"></i>재개
      </button>
    `;
  } else if (status === '공사진행' || status === '공사대기') {
    return `
      <button type="button" class="btn btn-warning btn-sm cancel-project-btn"
              data-project-code="${projectCode}" title="공사 중단">
        <i class="fas fa-pause me-1"></i>중단
      </button>
    `;
  }

  return '';
}
```

**변경할 코드:**
```javascript
generateCancelResumeButton(projectCode, rowData) {
  const status = rowData['공사상태'] || '';

  // 공사 취소 상태면 재개 버튼 표시
  if (status === '취소') {
    return `
      <button type="button" class="btn btn-success btn-sm resume-construction-btn"
              data-project-code="${projectCode}" title="공사 재개">
        <i class="fas fa-play me-1"></i>재개
      </button>
    `;
  }

  // 공사중단 상태면 재개 버튼 (기존 로직)
  if (status === '공사중단') {
    return `
      <button type="button" class="btn btn-success btn-sm resume-project-btn"
              data-project-code="${projectCode}" title="공사 재개">
        <i class="fas fa-play me-1"></i>재개
      </button>
    `;
  }

  // 진행중이거나 대기중이면 취소 버튼 표시
  if (status === '공사진행' || status === '공사대기' || status === '공사완료' || status === '수금필요') {
    return `
      <button type="button" class="btn btn-danger btn-sm cancel-construction-btn"
              data-project-code="${projectCode}" title="공사 취소">
        <i class="fas fa-ban me-1"></i>취소
      </button>
    `;
  }

  return '';
}
```

---

#### 작업 2: 이벤트 리스너 등록 (setupEventListeners 함수 내부 추가)

파일 내에서 `setupEventListeners()` 또는 `bindEvents()` 함수를 찾아 다음 코드를 추가:

```javascript
// 공사 취소 버튼 이벤트
document.addEventListener('click', (e) => {
  if (e.target.closest('.cancel-construction-btn')) {
    const button = e.target.closest('.cancel-construction-btn');
    const projectCode = button.dataset.projectCode;
    this.cancelConstruction(projectCode);
  }
});

// 공사 재개 버튼 이벤트
document.addEventListener('click', (e) => {
  if (e.target.closest('.resume-construction-btn')) {
    const button = e.target.closest('.resume-construction-btn');
    const projectCode = button.dataset.projectCode;
    this.resumeConstruction(projectCode);
  }
});
```

---

#### 작업 3: cancelConstruction() 함수 추가 (파일 끝부분에 추가)

```javascript
/**
 * 공사 취소
 */
async cancelConstruction(projectCode) {
  if (!confirm(`정말로 프로젝트 ${projectCode}의 공사를 취소하시겠습니까?\n\n취소하면:\n- 모든 편집 기능이 비활성화됩니다\n- 수금완료가 해제됩니다\n- 공사확정일이 초기화됩니다`)) {
    return;
  }

  try {
    this.showMessage('공사를 취소하고 있습니다...', 'info');

    const response = await fetch('/api/project/cancel', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        projectCode: projectCode
      })
    });

    const result = await response.json();

    if (result.ok) {
      console.log('[공사취소] 서버 응답 성공');

      // UI 스타일 적용
      this.applyCancelledProjectStyles(projectCode);

      // 버튼 업데이트 (취소 → 재개)
      this.updateCancelResumeButton(projectCode, '취소');

      // 부분 업데이트 이벤트 발송
      window.dispatchEvent(new CustomEvent('projectUpdated', {
        detail: {
          projectCode: projectCode,
          action: 'cancel',
          partialUpdate: true
        }
      }));

      this.showMessage('공사가 취소되었습니다.', 'success');
    } else {
      throw new Error(result.error || '공사 취소 처리 중 오류가 발생했습니다.');
    }
  } catch (error) {
    console.error('[공사취소] 오류:', error);
    this.showMessage('공사 취소 처리 중 오류가 발생했습니다: ' + error.message, 'error');
  }
}

/**
 * 공사 재개
 */
async resumeConstruction(projectCode) {
  if (!confirm(`프로젝트 ${projectCode}의 공사를 재개하시겠습니까?`)) {
    return;
  }

  try {
    this.showMessage('공사를 재개하고 있습니다...', 'info');

    const response = await fetch('/api/project/resume', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        projectCode: projectCode
      })
    });

    const result = await response.json();

    if (result.ok) {
      console.log('[공사재개] 서버 응답 성공');

      // UI 스타일 제거
      this.removeCancelledProjectStyles(projectCode);

      // 버튼 업데이트 (재개 → 취소)
      this.updateCancelResumeButton(projectCode, '공사진행');

      // 부분 업데이트 이벤트 발송
      window.dispatchEvent(new CustomEvent('projectUpdated', {
        detail: {
          projectCode: projectCode,
          action: 'resume',
          partialUpdate: true
        }
      }));

      this.showMessage('공사가 재개되었습니다.', 'success');
    } else {
      throw new Error(result.error || '공사 재개 처리 중 오류가 발생했습니다.');
    }
  } catch (error) {
    console.error('[공사재개] 오류:', error);
    this.showMessage('공사 재개 처리 중 오류가 발생했습니다: ' + error.message, 'error');
  }
}

/**
 * 취소된 프로젝트 스타일 적용
 */
applyCancelledProjectStyles(projectCode) {
  // 테이블 행에 빨간 취소선 추가
  const tableRows = document.querySelectorAll(`tr[data-project-code="${projectCode}"]`);
  tableRows.forEach(row => {
    row.classList.add('project-cancelled');
    row.style.position = 'relative';

    // 빨간 취소선 추가 (가로로 긋기)
    if (!row.querySelector('.cancelled-line')) {
      const line = document.createElement('div');
      line.className = 'cancelled-line';
      line.style.cssText = `
        position: absolute;
        top: 50%;
        left: 0;
        right: 0;
        height: 3px;
        background-color: #dc3545;
        z-index: 5;
        pointer-events: none;
      `;
      row.appendChild(line);
    }
  });

  // 아코디언 스타일 적용
  const rowDetails = this.accordionContainer;
  if (rowDetails) {
    rowDetails.classList.add('project-cancelled');
    rowDetails.style.filter = 'grayscale(50%)';
    rowDetails.style.opacity = '0.8';
    rowDetails.style.position = 'relative';

    // 모든 편집 버튼 비활성화
    const editButtons = rowDetails.querySelectorAll('.unified-edit-btn, .unified-save-btn, .unified-cancel-btn, .edit-btn, .save-btn, .cancel-btn');
    editButtons.forEach(btn => {
      btn.disabled = true;
      btn.style.opacity = '0.5';
      btn.style.pointerEvents = 'none';
    });

    // "취소된 공사" 워터마크 추가
    if (!rowDetails.querySelector('.cancelled-watermark')) {
      const watermark = document.createElement('div');
      watermark.className = 'cancelled-watermark';
      watermark.textContent = '취소된 공사';
      watermark.style.cssText = `
        position: absolute;
        top: 50%;
        left: 50%;
        transform: translate(-50%, -50%) rotate(-15deg);
        font-size: 3rem;
        font-weight: bold;
        color: rgba(220, 53, 69, 0.2);
        pointer-events: none;
        z-index: 1000;
        white-space: nowrap;
      `;
      rowDetails.appendChild(watermark);
    }
  }
}

/**
 * 취소된 프로젝트 스타일 제거
 */
removeCancelledProjectStyles(projectCode) {
  // 테이블 행 스타일 제거
  const tableRows = document.querySelectorAll(`tr[data-project-code="${projectCode}"]`);
  tableRows.forEach(row => {
    row.classList.remove('project-cancelled');
    row.style.position = '';

    // 취소선 제거
    const line = row.querySelector('.cancelled-line');
    if (line) line.remove();
  });

  // 아코디언 스타일 제거
  const rowDetails = this.accordionContainer;
  if (rowDetails) {
    rowDetails.classList.remove('project-cancelled');
    rowDetails.style.filter = '';
    rowDetails.style.opacity = '';

    // 편집 버튼 활성화
    const editButtons = rowDetails.querySelectorAll('.unified-edit-btn, .unified-save-btn, .unified-cancel-btn, .edit-btn, .save-btn, .cancel-btn');
    editButtons.forEach(btn => {
      btn.disabled = false;
      btn.style.opacity = '';
      btn.style.pointerEvents = '';
    });

    // 워터마크 제거
    const watermark = rowDetails.querySelector('.cancelled-watermark');
    if (watermark) watermark.remove();
  }
}
```

---

#### 작업 4: updateCancelResumeButton() 함수 수정 (Line 2453-2456)

**기존 코드:**
```javascript
updateCancelResumeButton(projectCode, newStatus) {
  const buttonContainer = document.querySelector(`.project-title-section[data-project-code="${projectCode}"] > div > div:last-child`);
  buttonContainer.innerHTML = this.generateCancelResumeButton(projectCode, { '공사상태': newStatus });
}
```

**변경 없음** (이미 적절함) - 단, generateCancelResumeButton이 수정되면 자동으로 적용됨

---

### 📝 B. ModernProjectFilters.js (수금 관리 모드 필터링)

파일 위치: `dashboard/src/js/components/ModernProjectFilters.js`

#### 작업: applyFilters() 함수에서 수금 모드 필터링 추가

`applyFilters()` 함수 내부에서 다음 로직 추가:

```javascript
// 수금 관리 모드에서는 취소된 공사 제외
if (this.isReceivablesMode) {
  filteredData = filteredData.filter(item => {
    const status = item['공사상태'] || '';
    return status !== '취소';  // 취소된 공사 제외
  });
}
```

---

### 📝 C. UnifiedBadgeSystem.js (상태 뱃지 추가)

파일 위치: `dashboard/src/js/components/UnifiedBadgeSystem.js`

#### 작업: createStatusBadge() 함수에 '취소' 상태 추가

```javascript
createStatusBadge(status) {
  const badgeMap = {
    '공사대기': { class: 'status-waiting', icon: 'clock', text: '대기' },
    '공사진행': { class: 'status-progress', icon: 'spinner', text: '진행중' },
    '공사완료': { class: 'status-complete', icon: 'check-circle', text: '완료' },
    '수금필요': { class: 'status-collection', icon: 'money-bill-wave', text: '수금필요' },
    '공사중단': { class: 'status-paused', icon: 'pause-circle', text: '중단' },
    '취소': { class: 'status-cancelled', icon: 'ban', text: '취소'},  // ← 추가
    '확인필요': { class: 'status-review', icon: 'exclamation-circle', text: '확인필요' }
  };

  const config = badgeMap[status] || { class: 'status-waiting', icon: 'question-circle', text: status || '대기' };

  return `
    <span class="badge ${config.class}">
      <i class="fas fa-${config.icon} me-1"></i>${config.text}
    </span>
  `;
}
```

---

### 📝 D. CSS 스타일 추가

파일 위치: `dashboard/src/css/components/table.css` 또는 `main.css`

```css
/* 공사 취소 상태 스타일 */
.project-cancelled {
  position: relative;
}

.cancelled-line {
  position: absolute;
  top: 50%;
  left: 0;
  right: 0;
  height: 3px;
  background-color: #dc3545;
  z-index: 5;
  pointer-events: none;
}

.cancelled-watermark {
  position: absolute;
  top: 50%;
  left: 50%;
  transform: translate(-50%, -50%) rotate(-15deg);
  font-size: 3rem;
  font-weight: bold;
  color: rgba(220, 53, 69, 0.2);
  pointer-events: none;
  z-index: 1000;
  white-space: nowrap;
  user-select: none;
}

/* 취소 상태 뱃지 */
.status-cancelled {
  background-color: #6c757d !important;
  color: white !important;
  font-weight: 600;
}

.status-cancelled i {
  animation: none;
}
```

---

## 3. 구현 체크리스트

### 백엔드
- [x] `/api/project/cancel` API 엔드포인트 추가
- [x] `/api/project/resume` API 엔드포인트 추가
- [x] 구글 시트 색상 변경 로직
- [x] 수금완료 및 공사확정일 초기화 로직

### 프론트엔드
- [ ] ProjectRowAccordion.js - generateCancelResumeButton() 수정
- [ ] ProjectRowAccordion.js - 이벤트 리스너 추가
- [ ] ProjectRowAccordion.js - cancelConstruction() 함수 추가
- [ ] ProjectRowAccordion.js - resumeConstruction() 함수 추가
- [ ] ProjectRowAccordion.js - applyCancelledProjectStyles() 함수 추가
- [ ] ProjectRowAccordion.js - removeCancelledProjectStyles() 함수 추가
- [ ] ModernProjectFilters.js - 수금 모드 필터링 추가
- [ ] UnifiedBadgeSystem.js - 취소 상태 뱃지 추가
- [ ] CSS 스타일 추가

### 테스트
- [ ] 공사 취소 버튼 클릭 테스트
- [ ] 공사 재개 버튼 클릭 테스트
- [ ] 테이블 행에 빨간 취소선 표시 확인
- [ ] 아코디언 그레이스케일 및 워터마크 확인
- [ ] 수금 모드에서 취소된 공사 필터링 확인
- [ ] 구글 시트 색상 변경 확인

---

## 4. 주의사항

### 구글 시트 컬럼 확인 필요
`inline_update.py`의 다음 부분은 실제 구글 시트 구조에 맞게 수정 필요:

```python
# Line 182-192
updates = [
    {
        'range': f'공사 현황!G{row_number}',  # G열: 공사상태 (추정 - 확인 필요)
        'values': [['취소']]
    },
    {
        'range': f'공사 현황!AG{row_number}',  # AG열: 수금완료 (추정 - 확인 필요)
        'values': [['FALSE']]
    },
    {
        'range': f'공사 현황!I{row_number}',  # I열: 공사확정일 (추정 - 확인 필요)
        'values': [['']]
    }
]
```

**확인 방법:**
1. 구글 시트를 열어서 각 컬럼의 위치 확인
2. 해당 컬럼명이 정확히 어느 열인지 확인 (A=1, B=2, ...)
3. 코드 수정

---

## 5. 구현 후 테스트 시나리오

1. **공사 취소 테스트**
   ```
   1. 프로젝트 행 클릭하여 아코디언 열기
   2. "공사 취소" 버튼 클릭
   3. 확인 대화상자에서 "확인" 클릭
   4. 확인 사항:
      - 테이블 행에 빨간 취소선 표시
      - 아코디언이 그레이스케일로 변경
      - "취소된 공사" 워터마크 표시
      - 모든 편집 버튼 비활성화
      - 버튼이 "재개"로 변경
   ```

2. **공사 재개 테스트**
   ```
   1. 취소된 프로젝트 행 클릭하여 아코디언 열기
   2. "재개" 버튼 클릭
   3. 확인 대화상자에서 "확인" 클릭
   4. 확인 사항:
      - 빨간 취소선 제거
      - 아코디언 그레이스케일 제거
      - 워터마크 제거
      - 편집 버튼 활성화
      - 버튼이 "취소"로 변경
   ```

3. **수금 모드 필터링 테스트**
   ```
   1. 수금 관리 모드 활성화
   2. 확인 사항:
      - 취소된 공사가 목록에 나타나지 않음
   ```

4. **구글 시트 확인**
   ```
   1. 구글 시트에서 취소한 프로젝트 행 확인
   2. 확인 사항:
      - 행 배경색이 진한 회색(#808080)
      - 공사상태가 "취소"
      - 수금완료가 FALSE
      - 공사확정일이 비어있음
   ```

---

**작성자**: Claude Code
**최종 수정**: 2025-01-15
