/**
 * Toast 알림 컴포넌트
 *
 * 사이트 헤더 상단 제목 옆 슬롯(`#systemAlertContainer`, modern_base.html)에
 * 컴팩트 pill 형태로 표시. 시공자 관리 모달의 showConstructorAlert() 스타일 계승.
 *
 * 이전 구현은 화면 우측 상단 fixed 컨테이너였음 (bd6299d 시점까지). 매니저 피드백:
 * 모든 시스템 알림을 사이트 제목 옆으로 통일하고 싶다는 요청 반영해 재작성.
 * 슬롯이 없는 페이지(로그인 등)에서는 화면 우측 상단 fallback으로 대체.
 */
export default class Toast {
  constructor() {
    // 헤더 슬롯 우선, 없으면 fallback 컨테이너 생성
    this.headerSlot = document.getElementById('systemAlertContainer');
    if (!this.headerSlot) {
      this.headerSlot = this._createFallbackContainer();
    }
  }

  /**
   * 헤더 슬롯 없는 페이지용 fallback (화면 우측 상단)
   */
  _createFallbackContainer() {
    let container = document.getElementById('toast-fallback-container');
    if (!container) {
      container = document.createElement('div');
      container.id = 'toast-fallback-container';
      container.style.cssText = `
        position: fixed;
        top: 20px;
        right: 20px;
        z-index: 9999;
        max-width: 350px;
        display: flex;
        flex-direction: column;
        gap: 8px;
      `;
      document.body.appendChild(container);
    }
    return container;
  }

  /**
   * Toast 메시지 표시
   * @param {string} message 표시할 메시지
   * @param {string} type success | error | warning | info
   * @param {number} duration 자동 사라지는 시간 (ms). 기본 4000. 0이면 수동 닫기만.
   */
  show(message, type = 'info', duration = 4000) {
    // 이전 알림 제거 (헤더 슬롯은 최근 하나만 유지)
    const isHeaderSlot = this.headerSlot?.id === 'systemAlertContainer';
    if (isHeaderSlot) {
      this.headerSlot.innerHTML = '';
    }

    const alertEl = this._createAlert(message, type);
    this.headerSlot.appendChild(alertEl);

    // fade-in
    requestAnimationFrame(() => {
      alertEl.style.opacity = '1';
    });

    // 자동 fade-out + 제거
    if (duration > 0) {
      setTimeout(() => this.hide(alertEl), duration);
    }

    return alertEl;
  }

  /**
   * pill 형태의 alert 엘리먼트 생성 (showConstructorAlert 스타일)
   */
  _createAlert(message, type) {
    const styleMap = {
      success: { bg: '#198754', color: '#fff', icon: 'fa-check-circle' },
      error:   { bg: '#dc3545', color: '#fff', icon: 'fa-exclamation-circle' },
      danger:  { bg: '#dc3545', color: '#fff', icon: 'fa-exclamation-circle' },
      warning: { bg: '#ffc107', color: '#212529', icon: 'fa-exclamation-triangle' },
      info:    { bg: '#0d6efd', color: '#fff', icon: 'fa-info-circle' },
    };
    const s = styleMap[type] || styleMap.info;

    const el = document.createElement('div');
    el.setAttribute('role', type === 'error' || type === 'danger' ? 'alert' : 'status');
    el.style.cssText = `
      display: inline-flex;
      align-items: center;
      gap: 0.4rem;
      background-color: ${s.bg};
      color: ${s.color};
      border-radius: 0.375rem;
      padding: 0.35rem 0.7rem;
      font-size: 0.8125rem;
      font-weight: 500;
      white-space: nowrap;
      box-shadow: 0 1px 3px rgba(0, 0, 0, 0.08);
      opacity: 0;
      transition: opacity 0.2s ease-in-out;
      max-width: 100%;
    `;
    el.innerHTML = `
      <i class="fas ${s.icon}" style="font-size: 0.85rem;" aria-hidden="true"></i>
      <span style="overflow: hidden; text-overflow: ellipsis;">${this._escape(message)}</span>
    `;
    return el;
  }

  /**
   * fade-out 후 DOM 제거
   */
  hide(el) {
    if (!el || !el.parentNode) return;
    el.style.opacity = '0';
    setTimeout(() => {
      if (el.parentNode) el.parentNode.removeChild(el);
    }, 250);
  }

  /**
   * XSS 방지: 메시지 HTML 이스케이프
   */
  _escape(str) {
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }
}
