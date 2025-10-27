import logger from '../utils/logger.js';

/**
 * CacheStatusManager - 서버 푸시 기반 캐시 상태 관리
 * 전문가 리뷰: "캐시 갱신·락 상태는 서버 푸시나 명시적 user action에 연동하도록 바꿔야
 * 네트워크 낭비와 race condition을 줄일 수 있습니다"
 */

export default class CacheStatusManager {
  constructor(socket = null) {
    this.socket = socket;
    this.currentStatus = null;
    this.statusListeners = [];
    this.isMonitoring = false;

    // 기존 setInterval 방식을 단계적으로 제거하기 위한 플래그
    this.useServerPush = true;
    this.fallbackInterval = null;
  }

  /**
   * 캐시 상태 모니터링 시작 (서버 푸시 기반)
   */
  startMonitoring() {
    if (this.isMonitoring) return;

    logger.debug('[CacheStatusManager] 서버 푸시 기반 캐시 상태 모니터링 시작');

    // 초기 캐시 상태 확인 (한 번만)
    this.checkCacheStatusOnce();

    if (this.socket && this.useServerPush) {
      this.setupServerPushListeners();
    } else {
      logger.warn('[CacheStatusManager] 소켓이 없어 폴백 모드로 전환');
      this.setupFallbackMode();
    }

    this.isMonitoring = true;
  }

  /**
   * 서버 푸시 리스너 설정
   */
  setupServerPushListeners() {
    if (!this.socket) return;

    // 캐시 상태 변경 이벤트 리스너
    this.socket.on('cache_status_updated', (status) => {
      logger.debug('[CacheStatusManager] 서버 푸시 캐시 상태 업데이트:', status);
      this.updateCacheStatus(status);
    });

    // 캐시 새로고침 완료 이벤트
    this.socket.on('cache_refresh_completed', (result) => {
      logger.debug('[CacheStatusManager] 캐시 새로고침 완료:', result);
      if (result.success) {
        this.checkCacheStatusOnce(); // 최신 상태 다시 확인
      }
    });

    // 프리패치 상태 변경 이벤트
    this.socket.on('prefetch_status_changed', (status) => {
      logger.debug('[CacheStatusManager] 프리패치 상태 변경:', status);
      // 현재 상태 업데이트
      if (this.currentStatus) {
        this.currentStatus.prefetch_running = status.running;
        this.notifyStatusListeners(this.currentStatus);
      }
    });

    logger.debug('[CacheStatusManager] 서버 푸시 리스너 설정 완료');
  }

  /**
   * 폴백 모드 (최소한의 polling)
   */
  setupFallbackMode() {
    // 서버 푸시가 불가능한 경우만 최소한의 polling 사용
    // 기존 3분에서 10분으로 간격 대폭 증가
    this.fallbackInterval = setInterval(() => {
      logger.debug('[CacheStatusManager] 폴백 모드 상태 확인');
      this.checkCacheStatusOnce();
    }, 600000); // 10분
  }

  /**
   * 한 번만 캐시 상태 확인 (명시적 액션)
   */
  async checkCacheStatusOnce() {
    try {
      const response = await fetch('/api/cache/status');
      const data = await response.json();

      this.updateCacheStatus(data);
    } catch (error) {
      logger.error('[CacheStatusManager] 캐시 상태 확인 실패:', error);
      this.updateCacheStatus({
        cache_health: 'error',
        prefetch_running: false,
        last_successful_update: null
      });
    }
  }

  /**
   * 캐시 상태 업데이트
   */
  updateCacheStatus(status) {
    this.currentStatus = status;
    this.updateCacheStatusUI(status);
    this.notifyStatusListeners(status);
  }

  /**
   * 캐시 상태 UI 업데이트
   */
  updateCacheStatusUI(status) {
    const iconElement = document.getElementById('cacheStatusIcon');
    const textElement = document.getElementById('cacheStatusText');
    const refreshButton = document.getElementById('manualRefreshBtn');

    if (!iconElement || !textElement || !refreshButton) return;

    const { cache_health, prefetch_running, last_successful_update } = status;

    let iconClass, iconColor, statusText, showRefreshBtn = false;

    switch (cache_health) {
      case 'healthy':
        iconClass = 'fas fa-check-circle';
        iconColor = 'text-success';
        statusText = '캐시 정상';
        break;
      case 'warning':
        iconClass = 'fas fa-exclamation-triangle';
        iconColor = 'text-warning';
        statusText = '캐시 경고';
        showRefreshBtn = true;
        break;
      case 'error':
        iconClass = 'fas fa-times-circle';
        iconColor = 'text-danger';
        statusText = '캐시 오류';
        showRefreshBtn = true;
        break;
      default:
        iconClass = 'fas fa-sync-alt fa-spin';
        iconColor = 'text-muted';
        statusText = '상태 확인 중';
    }

    // 프리패치가 실행되지 않는 경우 경고 표시
    if (!prefetch_running) {
      iconClass = 'fas fa-exclamation-triangle';
      iconColor = 'text-warning';
      statusText = '자동 갱신 중지됨';
      showRefreshBtn = true;
    }

    // UI 업데이트
    iconElement.innerHTML = `<i class="${iconClass} ${iconColor}" title="${statusText}"></i>`;
    textElement.textContent = statusText;
    textElement.className = iconColor.replace('text-', 'text-');

    // 새로고침 버튼 표시/숨김
    refreshButton.style.display = showRefreshBtn ? 'inline-block' : 'none';

    // 마지막 업데이트 시간을 툴팁으로 표시
    if (last_successful_update) {
      const lastUpdate = new Date(last_successful_update);
      const timeAgo = this.getTimeAgo(lastUpdate);
      iconElement.setAttribute('title', `${statusText} (마지막 갱신: ${timeAgo})`);
    }
  }

  /**
   * 수동 캐시 새로고침 (명시적 user action)
   */
  async manualRefreshCache() {
    const refreshButton = document.getElementById('manualRefreshBtn');
    const iconElement = document.getElementById('cacheStatusIcon');
    const textElement = document.getElementById('cacheStatusText');

    if (!refreshButton || !iconElement || !textElement) return;

    // 버튼 비활성화 및 로딩 표시
    refreshButton.disabled = true;
    iconElement.innerHTML = '<i class="fas fa-sync-alt fa-spin text-primary"></i>';
    textElement.textContent = '새로고침 중...';
    textElement.className = 'text-primary';

    try {
      const response = await fetch('/api/cache/refresh', { method: 'POST' });
      const result = await response.json();

      if (result.success) {
        // 성공 시 UI 업데이트
        iconElement.innerHTML = '<i class="fas fa-check-circle text-success"></i>';
        textElement.textContent = '새로고침 완료';
        textElement.className = 'text-success';

        // 서버 푸시를 기다리거나 1초 후 상태 재확인
        setTimeout(() => {
          this.checkCacheStatusOnce();
        }, 1000);

        logger.debug('[CacheStatusManager] 수동 캐시 새로고침 성공:', result);

        // 새로고침 성공 이벤트 발생
        this.notifyRefreshComplete(true, result);

      } else {
        throw new Error(result.message || '새로고침 실패');
      }
    } catch (error) {
      logger.error('[CacheStatusManager] 수동 캐시 새로고침 실패:', error);

      // 오류 시 UI 업데이트
      iconElement.innerHTML = '<i class="fas fa-times-circle text-danger"></i>';
      textElement.textContent = '새로고침 실패';
      textElement.className = 'text-danger';

      // 2초 후 원래 상태로 복원
      setTimeout(() => this.checkCacheStatusOnce(), 2000);

      // 새로고침 실패 이벤트 발생
      this.notifyRefreshComplete(false, error);

    } finally {
      // 버튼 다시 활성화
      setTimeout(() => {
        refreshButton.disabled = false;
      }, 2000);
    }
  }

  /**
   * 상태 변경 리스너 등록
   */
  onStatusChange(callback) {
    this.statusListeners.push(callback);
  }

  /**
   * 상태 리스너들에게 알림
   */
  notifyStatusListeners(status) {
    this.statusListeners.forEach(callback => {
      try {
        callback(status);
      } catch (error) {
        logger.error('[CacheStatusManager] 상태 리스너 오류:', error);
      }
    });
  }

  /**
   * 새로고침 완료 이벤트 발생
   */
  notifyRefreshComplete(success, result) {
    const event = new CustomEvent('cacheRefreshComplete', {
      detail: { success, result }
    });
    window.dispatchEvent(event);
  }

  /**
   * 시간차이를 사람이 읽기 쉬운 형태로 변환
   */
  getTimeAgo(date) {
    const now = new Date();
    const diffMs = now - date;
    const diffMins = Math.floor(diffMs / 60000);
    const diffHours = Math.floor(diffMins / 60);
    const diffDays = Math.floor(diffHours / 24);

    if (diffMins < 1) return '방금 전';
    if (diffMins < 60) return `${diffMins}분 전`;
    if (diffHours < 24) return `${diffHours}시간 전`;
    return `${diffDays}일 전`;
  }

  /**
   * 모니터링 중지
   */
  stopMonitoring() {
    if (!this.isMonitoring) return;

    logger.debug('[CacheStatusManager] 캐시 상태 모니터링 중지');

    // 소켓 리스너 제거
    if (this.socket) {
      this.socket.off('cache_status_updated');
      this.socket.off('cache_refresh_completed');
      this.socket.off('prefetch_status_changed');
    }

    // 폴백 인터벌 정리
    if (this.fallbackInterval) {
      clearInterval(this.fallbackInterval);
      this.fallbackInterval = null;
    }

    this.isMonitoring = false;
  }

  /**
   * 정리 작업
   */
  cleanup() {
    this.stopMonitoring();
    this.statusListeners = [];
    this.currentStatus = null;
    logger.debug('[CacheStatusManager] 정리 작업 완료');
  }

  /**
   * 현재 상태 반환
   */
  getCurrentStatus() {
    return this.currentStatus;
  }

  /**
   * 글로벌 함수 등록 (레거시 호환성)
   */
  exposeGlobalFunction() {
    window.manualRefreshCache = () => this.manualRefreshCache();
  }
}