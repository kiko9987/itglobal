/**
 * 버튼 상태 관리 유틸리티
 * 로딩, 성공, 실패 상태를 시각적으로 표현
 *
 * 최소 스피너 노출 시간 400ms 보장:
 *   write-behind 로 응답이 100ms 이내로 빨라져 스피너가 회전하기도 전에
 *   사라지는 문제(정지된 것처럼 보임) 방지.
 */
const MIN_LOADING_MS = 400;

export class ButtonStateManager {
  constructor() {
    this.originalStates = new Map();
    this.loadingStartTs = new WeakMap(); // 버튼별 setLoading 시각
  }

  /**
   * 버튼 원본 상태 저장
   * @param {HTMLElement} button
   */
  saveOriginalState(button) {
    if (!button || this.originalStates.has(button)) return;

    this.originalStates.set(button, {
      innerHTML: button.innerHTML,
      disabled: button.disabled,
      className: button.className
    });
  }

  /**
   * 버튼을 로딩 상태로 변경
   * @param {HTMLElement} button
   * @param {string} text
   */
  setLoading(button, text = '저장중...') {
    if (!button) return;

    this.saveOriginalState(button);
    this.loadingStartTs.set(button, Date.now());

    button.disabled = true;
    // fa-circle-notch: 비대칭 원호라 짧은 회전도 명확히 보임 (fa-spinner 대칭 8-radial 은 회전 안 보임)
    button.innerHTML = `<i class="fas fa-circle-notch fa-spin me-1"></i>${text}`;
    button.classList.remove('btn-success', 'btn-danger', 'btn-warning');
    button.classList.add('btn-secondary');

    // 클릭 효과
    this.animateButton(button, 'scale(0.95)', 100);
  }

  /**
   * 최소 로딩 시간 남은 만큼 대기 (setLoading 후 400ms 이하면 남은 시간 sleep)
   * @param {HTMLElement} button
   * @returns {Promise<void>}
   */
  _waitMinLoading(button) {
    const start = this.loadingStartTs.get(button);
    if (!start) return Promise.resolve();
    const elapsed = Date.now() - start;
    const remain = MIN_LOADING_MS - elapsed;
    if (remain <= 0) return Promise.resolve();
    return new Promise(r => setTimeout(r, remain));
  }

  /**
   * 버튼을 성공 상태로 변경
   * @param {HTMLElement} button
   * @param {string} text
   * @param {Function} callback
   */
  async setSuccess(button, text = '저장', callback = null) {
    if (!button) return;

    // 스피너 최소 400ms 노출 보장 — 사용자에게 '요청 처리됨' 시각 신호 확보
    await this._waitMinLoading(button);

    button.disabled = false;
    button.innerHTML = `<i class="fas fa-check me-1"></i>${text}`;
    button.classList.remove('btn-secondary', 'btn-danger', 'btn-warning');
    button.classList.add('btn-success');

    // 성공 애니메이션
    this.animateButton(button, 'scale(1.05)', 200, callback);
  }

  /**
   * 버튼을 실패 상태로 변경
   * @param {HTMLElement} button
   * @param {string} text
   * @param {string} originalText
   */
  async setError(button, text = '실패', originalText = '저장') {
    if (!button) return;

    // 실패도 최소 400ms 로딩 노출 (즉시 실패 시 스피너 안 보이는 문제 방지)
    await this._waitMinLoading(button);

    button.disabled = false;
    button.innerHTML = `<i class="fas fa-exclamation-triangle me-1"></i>${text}`;
    button.classList.remove('btn-secondary', 'btn-success');
    button.classList.add('btn-danger');

    // 실패 애니메이션 (흔들림)
    button.style.animation = 'shake 0.5s';

    // 1.5초 후 원래 상태로 복원
    setTimeout(() => {
      button.style.animation = '';
      this.reset(button, originalText);
    }, 1500);
  }

  /**
   * 버튼을 기본 상태로 복원
   * @param {HTMLElement} button
   * @param {string} text
   */
  reset(button, text = null) {
    if (!button) return;

    const originalState = this.originalStates.get(button);

    if (originalState && !text) {
      button.innerHTML = originalState.innerHTML;
      button.disabled = originalState.disabled;
      button.className = originalState.className;
    } else {
      button.disabled = false;
      button.innerHTML = text || `<i class="fas fa-check me-1"></i>저장`;
      button.classList.remove('btn-secondary', 'btn-danger', 'btn-warning');
      button.classList.add('btn-success');
    }

    button.style.transform = 'scale(1)';
    button.style.animation = '';
  }

  /**
   * 버튼 애니메이션 헬퍼
   * @param {HTMLElement} button
   * @param {string} transform
   * @param {number} duration
   * @param {Function} callback
   */
  animateButton(button, transform, duration = 200, callback = null) {
    button.style.transform = transform;
    setTimeout(() => {
      button.style.transform = 'scale(1)';
      if (callback) callback();
    }, duration);
  }

  /**
   * 모든 저장된 상태 정리
   */
  cleanup() {
    this.originalStates.clear();
  }
}

// 전역 인스턴스 생성
export const buttonStateManager = new ButtonStateManager();

// 레거시 호환성을 위한 글로벌 객체
window.ButtonStateManager = {
  setLoading: (button, text) => buttonStateManager.setLoading(button, text),
  setSuccess: (button, text, callback) => buttonStateManager.setSuccess(button, text, callback),
  setError: (button, text, originalText) => buttonStateManager.setError(button, text, originalText),
  reset: (button, text) => buttonStateManager.reset(button, text)
};