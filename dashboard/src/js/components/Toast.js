/**
 * Toast 알림 컴포넌트
 * 사용자에게 피드백 메시지를 표시
 */
export default class Toast {
  constructor() {
    this.container = this.createContainer();
  }

  /**
   * Toast 컨테이너 생성
   */
  createContainer() {
    let container = document.getElementById('toast-container');

    if (!container) {
      container = document.createElement('div');
      container.id = 'toast-container';
      container.className = 'toast-container';
      container.style.cssText = `
        position: fixed;
        top: 20px;
        right: 20px;
        z-index: 9999;
        max-width: 350px;
      `;
      document.body.appendChild(container);
    }

    return container;
  }

  /**
   * Toast 메시지 표시
   * @param {string} message
   * @param {string} type
   * @param {number} duration
   */
  show(message, type = 'info', duration = 4000) {
    const toast = this.createToast(message, type);
    this.container.appendChild(toast);

    // 애니메이션
    requestAnimationFrame(() => {
      toast.classList.add('show');
    });

    // 자동 제거
    setTimeout(() => {
      this.hide(toast);
    }, duration);

    return toast;
  }

  /**
   * Toast 엘리먼트 생성
   */
  createToast(message, type) {
    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;

    const icon = this.getIcon(type);

    toast.innerHTML = `
      <div class="toast-content">
        <i class="${icon}"></i>
        <span class="toast-message">${message}</span>
        <button class="toast-close" aria-label="닫기">
          <i class="fas fa-times"></i>
        </button>
      </div>
    `;

    // 스타일 적용
    toast.style.cssText = `
      background: ${this.getBackgroundColor(type)};
      color: white;
      padding: 12px 16px;
      border-radius: 8px;
      margin-bottom: 8px;
      box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
      transform: translateX(100%);
      transition: transform 0.3s ease-in-out;
      display: flex;
      align-items: center;
      font-size: 14px;
      min-width: 300px;
    `;

    // 닫기 버튼 이벤트
    const closeBtn = toast.querySelector('.toast-close');
    closeBtn.addEventListener('click', () => this.hide(toast));

    return toast;
  }

  /**
   * Toast 숨기기
   */
  hide(toast) {
    toast.style.transform = 'translateX(100%)';
    setTimeout(() => {
      if (toast.parentNode) {
        toast.parentNode.removeChild(toast);
      }
    }, 300);
  }

  /**
   * 타입별 아이콘 반환
   */
  getIcon(type) {
    const icons = {
      success: 'fas fa-check-circle',
      error: 'fas fa-exclamation-circle',
      warning: 'fas fa-exclamation-triangle',
      info: 'fas fa-info-circle'
    };
    return icons[type] || icons.info;
  }

  /**
   * 타입별 배경색 반환
   */
  getBackgroundColor(type) {
    const colors = {
      success: '#10b981',
      error: '#ef4444',
      warning: '#f59e0b',
      info: '#06b6d4'
    };
    return colors[type] || colors.info;
  }
}

// CSS 스타일 추가
const style = document.createElement('style');
style.textContent = `
  .toast.show {
    transform: translateX(0) !important;
  }

  .toast-content {
    display: flex;
    align-items: center;
    gap: 8px;
    width: 100%;
  }

  .toast-message {
    flex: 1;
    font-weight: 500;
  }

  .toast-close {
    background: none;
    border: none;
    color: inherit;
    cursor: pointer;
    padding: 4px;
    border-radius: 4px;
    transition: background-color 0.2s;
  }

  .toast-close:hover {
    background-color: rgba(255, 255, 255, 0.2);
  }

  @media (max-width: 768px) {
    .toast-container {
      right: 10px !important;
      left: 10px !important;
      max-width: none !important;
    }

    .toast {
      min-width: auto !important;
    }
  }
`;

if (!document.getElementById('toast-styles')) {
  style.id = 'toast-styles';
  document.head.appendChild(style);
}