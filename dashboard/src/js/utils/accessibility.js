/**
 * 접근성(A11y) 유틸리티
 * ARIA 레이블, 키보드 네비게이션, 포커스 관리
 */

/**
 * 버튼에 ARIA 속성 추가
 * @param {HTMLElement} button - 버튼 요소
 * @param {string} label - ARIA 레이블
 * @param {string} description - ARIA 설명 (선택적)
 */
export function enhanceButtonAccessibility(button, label, description = null) {
  if (!button) return;

  // ARIA 레이블
  if (!button.hasAttribute('aria-label')) {
    button.setAttribute('aria-label', label);
  }

  // ARIA 설명
  if (description) {
    button.setAttribute('aria-describedby', `desc-${Math.random().toString(36).substr(2, 9)}`);
    const descEl = document.createElement('span');
    descEl.id = button.getAttribute('aria-describedby');
    descEl.className = 'sr-only';  // 스크린 리더 전용
    descEl.textContent = description;
    button.appendChild(descEl);
  }

  // 키보드 포커스 가능하도록 보장
  if (!button.hasAttribute('tabindex')) {
    button.setAttribute('tabindex', '0');
  }

  // Enter/Space 키 지원
  if (!button._keyboardHandlerAttached) {
    button.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' || e.key === ' ') {
        e.preventDefault();
        button.click();
      }
    });
    button._keyboardHandlerAttached = true;
  }
}

/**
 * 입력 필드에 ARIA 레이블 추가
 * @param {HTMLInputElement} input - 입력 요소
 * @param {string} label - 레이블 텍스트
 * @param {boolean} required - 필수 입력 여부
 */
export function enhanceInputAccessibility(input, label, required = false) {
  if (!input) return;

  // ARIA 레이블
  input.setAttribute('aria-label', label);

  // 필수 입력 표시
  if (required) {
    input.setAttribute('aria-required', 'true');
    input.setAttribute('required', '');
  }

  // 에러 상태 지원
  if (input.classList.contains('is-invalid')) {
    input.setAttribute('aria-invalid', 'true');
  }
}

/**
 * 모달 접근성 개선
 * @param {HTMLElement} modal - 모달 요소
 * @param {string} title - 모달 제목
 */
export function enhanceModalAccessibility(modal, title) {
  if (!modal) return;

  // ARIA 역할
  modal.setAttribute('role', 'dialog');
  modal.setAttribute('aria-modal', 'true');

  // 레이블
  if (title) {
    const titleId = `modal-title-${Math.random().toString(36).substr(2, 9)}`;
    modal.setAttribute('aria-labelledby', titleId);

    const titleEl = modal.querySelector('h1, h2, h3, h4, h5, h6');
    if (titleEl) {
      titleEl.id = titleId;
    }
  }

  // 포커스 트랩 (모달 열릴 때 포커스 가능한 요소들만 탐색)
  const focusableElements = modal.querySelectorAll(
    'button, [href], input, select, textarea, [tabindex]:not([tabindex="-1"])'
  );

  if (focusableElements.length > 0) {
    // 첫 번째 요소에 포커스
    focusableElements[0].focus();

    // Tab 키로 마지막 요소 다음에 첫 요소로 순환
    const firstFocusable = focusableElements[0];
    const lastFocusable = focusableElements[focusableElements.length - 1];

    modal.addEventListener('keydown', (e) => {
      if (e.key === 'Tab') {
        if (e.shiftKey) {
          // Shift + Tab
          if (document.activeElement === firstFocusable) {
            e.preventDefault();
            lastFocusable.focus();
          }
        } else {
          // Tab
          if (document.activeElement === lastFocusable) {
            e.preventDefault();
            firstFocusable.focus();
          }
        }
      }

      // ESC 키로 모달 닫기
      if (e.key === 'Escape') {
        const closeBtn = modal.querySelector('[data-bs-dismiss="modal"], .close, .btn-close');
        if (closeBtn) {
          closeBtn.click();
        }
      }
    });
  }
}

/**
 * 테이블 접근성 개선
 * @param {HTMLTableElement} table - 테이블 요소
 */
export function enhanceTableAccessibility(table) {
  if (!table) return;

  // ARIA 역할
  table.setAttribute('role', 'table');

  // 테이블 헤더에 scope 속성 추가
  const headers = table.querySelectorAll('thead th');
  headers.forEach(th => {
    if (!th.hasAttribute('scope')) {
      th.setAttribute('scope', 'col');
    }
  });

  // 행 헤더에 scope 속성 추가
  const rowHeaders = table.querySelectorAll('tbody th');
  rowHeaders.forEach(th => {
    if (!th.hasAttribute('scope')) {
      th.setAttribute('scope', 'row');
    }
  });

  // 정렬 가능한 헤더에 ARIA 속성
  const sortableHeaders = table.querySelectorAll('th[data-sortable="true"], th.sorting');
  sortableHeaders.forEach(th => {
    th.setAttribute('role', 'button');
    th.setAttribute('tabindex', '0');
    th.setAttribute('aria-label', `${th.textContent} (정렬 가능)`);

    // 키보드로 정렬
    if (!th._sortKeyboardHandlerAttached) {
      th.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' || e.key === ' ') {
          e.preventDefault();
          th.click();
        }
      });
      th._sortKeyboardHandlerAttached = true;
    }
  });
}

/**
 * 스킵 링크 추가 (메인 콘텐츠로 건너뛰기)
 */
export function addSkipLink() {
  // 이미 존재하면 건너뜀
  if (document.querySelector('.skip-link')) return;

  const skipLink = document.createElement('a');
  skipLink.href = '#main-content';
  skipLink.className = 'skip-link sr-only sr-only-focusable';
  skipLink.textContent = '메인 콘텐츠로 건너뛰기';
  skipLink.style.cssText = `
    position: absolute;
    top: 0;
    left: 0;
    z-index: 10000;
    padding: 10px 20px;
    background: #000;
    color: #fff;
    text-decoration: none;
    transform: translateY(-100%);
    transition: transform 0.2s;
  `;

  skipLink.addEventListener('focus', () => {
    skipLink.style.transform = 'translateY(0)';
  });

  skipLink.addEventListener('blur', () => {
    skipLink.style.transform = 'translateY(-100%)';
  });

  document.body.insertBefore(skipLink, document.body.firstChild);

  // 메인 콘텐츠에 ID 추가
  const mainContent = document.querySelector('main, .main-content, #main');
  if (mainContent && !mainContent.id) {
    mainContent.id = 'main-content';
  }
}

/**
 * 라이브 리전 알림 (스크린 리더용)
 * @param {string} message - 알림 메시지
 * @param {string} priority - 'polite' 또는 'assertive'
 */
export function announceToScreenReader(message, priority = 'polite') {
  let liveRegion = document.getElementById('sr-live-region');

  if (!liveRegion) {
    liveRegion = document.createElement('div');
    liveRegion.id = 'sr-live-region';
    liveRegion.className = 'sr-only';
    liveRegion.setAttribute('aria-live', priority);
    liveRegion.setAttribute('aria-atomic', 'true');
    document.body.appendChild(liveRegion);
  }

  // 메시지 업데이트
  liveRegion.textContent = message;

  // 3초 후 메시지 제거
  setTimeout(() => {
    liveRegion.textContent = '';
  }, 3000);
}

/**
 * 포커스 가능한 요소 찾기
 * @param {HTMLElement} container - 컨테이너 요소
 * @returns {NodeList} - 포커스 가능한 요소들
 */
export function getFocusableElements(container) {
  return container.querySelectorAll(
    'a[href], button:not([disabled]), textarea:not([disabled]), ' +
    'input:not([disabled]), select:not([disabled]), ' +
    '[tabindex]:not([tabindex="-1"])'
  );
}

/**
 * 포커스 트랩 (컨테이너 내에서만 포커스 순환)
 * @param {HTMLElement} container - 컨테이너 요소
 * @returns {Function} - 트랩 해제 함수
 */
export function trapFocus(container) {
  const focusableElements = getFocusableElements(container);
  const firstFocusable = focusableElements[0];
  const lastFocusable = focusableElements[focusableElements.length - 1];

  const handleTabKey = (e) => {
    if (e.key !== 'Tab') return;

    if (e.shiftKey) {
      if (document.activeElement === firstFocusable) {
        e.preventDefault();
        lastFocusable.focus();
      }
    } else {
      if (document.activeElement === lastFocusable) {
        e.preventDefault();
        firstFocusable.focus();
      }
    }
  };

  container.addEventListener('keydown', handleTabKey);

  // 첫 번째 요소에 포커스
  if (firstFocusable) {
    firstFocusable.focus();
  }

  // 트랩 해제 함수 반환
  return () => {
    container.removeEventListener('keydown', handleTabKey);
  };
}

export default {
  enhanceButtonAccessibility,
  enhanceInputAccessibility,
  enhanceModalAccessibility,
  enhanceTableAccessibility,
  addSkipLink,
  announceToScreenReader,
  getFocusableElements,
  trapFocus
};
