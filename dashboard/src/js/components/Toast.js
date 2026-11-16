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
    // 그룹웨어 틀에선 #systemAlertContainer 가 위쪽 탭 줄 끝(시공사례 탭 옆)이라 알림이 화면 제목과 동떨어졌음
    //   (부장님 지적 11-16) → 화면 제목 옆 자리(#headerAlertContainer, 프로젝트·리드)가 있으면 거기에.
    //   예전 틀에선 #systemAlertContainer 가 사이트 제목 옆이라 그대로.
    const pageSlot = document.body?.classList.contains('gw-shell')
      ? document.getElementById('headerAlertContainer') : null;
    this.headerSlot = pageSlot || document.getElementById('systemAlertContainer');
    if (pageSlot) {
      // 제목 옆 자리는 높이가 2rem 로 고정돼 있어 두 줄 알림이 아래 줄을 덮음 → 최소 높이로
      pageSlot.style.height = 'auto';
      pageSlot.style.minHeight = '2rem';
      if (!pageSlot.hasAttribute('aria-live')) pageSlot.setAttribute('aria-live', 'polite');
    }
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
    const isHeaderSlot = ['systemAlertContainer', 'headerAlertContainer'].includes(this.headerSlot?.id);
    if (isHeaderSlot) {
      this.headerSlot.innerHTML = '';
    }

    const alertEl = this._createAlert(message, type);
    // 저절로 안 사라지는 알림(오류 등)은 닫기 단추를 붙인다(PC 점검 11-16)
    if (!(duration > 0)) {
      const close = document.createElement('button');
      close.type = 'button';
      close.setAttribute('aria-label', '알림 닫기');
      close.innerHTML = '<i class="fas fa-times" aria-hidden="true"></i>';
      close.style.cssText = 'background:none;border:0;color:inherit;padding:0 0 0 .35rem;margin:0;cursor:pointer;line-height:1;min-width:24px;min-height:24px;';
      close.addEventListener('click', () => this.hide(alertEl));
      alertEl.appendChild(close);
    }
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
      white-space: normal;
      line-height: 1.35;
      box-shadow: 0 1px 3px rgba(0, 0, 0, 0.08);
      opacity: 0;
      transition: opacity 0.2s ease-in-out;
      max-width: min(560px, 100%);
    `;
    // 긴 안내가 '…'로 잘려 뒷부분(해야 할 일)을 못 읽던 것 → 줄바꿈 허용(PC 점검 11-16)
    el.innerHTML = `
      <i class="fas ${s.icon}" style="font-size: 0.85rem;" aria-hidden="true"></i>
      <span style="overflow-wrap: anywhere;">${this._escape(message)}</span>
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
