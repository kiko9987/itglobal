import logger from '../utils/logger.js';
import { buttonStateManager } from '../utils/ButtonStateManager.js';

/**
 * Project List API - 글로벌 API 함수들을 노출하는 모듈
 */

/**
 * 글로벌 API 함수들을 window 객체에 노출
 * @param {Object} app - ProjectListApp 인스턴스
 * @returns {Object} API 객체
 */
export function exposeGlobalAPI(app) {
  // API 객체 구성
  const api = {
    // 데이터 관련
    refreshData: (force = false, showMessage = false) => app.refreshData(force, true, showMessage),

    // 완전한 동기화 (서버 캐시 삭제 + 클라이언트 갱신)
    fullRefresh: async (showMessage = true) => {
      logger.debug('[ProjectListAPI] fullRefresh 호출 - 서버 캐시 삭제 + 클라이언트 갱신');

      const btn = document.getElementById('fullRefreshBtn');

      try {
        // 버튼을 로딩 상태로 변경
        if (btn) {
          buttonStateManager.setLoading(btn, '새로고침 중...');
        }

        // 1단계: 서버 캐시 삭제 (DOM 요소 없이 직접 API 호출)
        logger.debug('[ProjectListAPI] 서버 캐시 삭제 중...');
        const response = await fetch('/api/cache/refresh', {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json'
          },
          body: JSON.stringify({})
        });
        const result = await response.json();

        if (!result.success) {
          logger.warn('[ProjectListAPI] 서버 캐시 삭제 실패:', result.message);
          throw new Error(result.message || '서버 캐시 삭제 실패');
        }

        logger.info('[ProjectListAPI] 서버 캐시 삭제 성공');

        // 2단계: 클라이언트 데이터 갱신
        api.refreshData(true, showMessage);

        // 성공 상태로 변경
        if (btn) {
          buttonStateManager.setSuccess(btn, '완료', () => {
            // 1.5초 후 원래 상태로 복원
            setTimeout(() => {
              buttonStateManager.reset(btn);
            }, 1500);
          });
        }

      } catch (error) {
        logger.error('[ProjectListAPI] 서버 캐시 삭제 오류:', error);

        // 오류 시 에러 상태로 변경
        if (btn) {
          buttonStateManager.setError(btn, '실패', '새로고침');
        }

        // 새로고침 실패 = 시스템 파이프라인 에러 → 사이트 최상단 헤더
        if (app.showSystemAlert) {
          app.showSystemAlert('새로고침에 실패했습니다.', 'error');
        } else if (app.showErrorMessage) {
          app.showErrorMessage('새로고침에 실패했습니다.');
        }
      }
    },

    getCurrentData: () => {
      if (Array.isArray(app?.stateManager?.currentData) && app.stateManager.currentData.length) {
        return app.stateManager.currentData;
      }
      if (Array.isArray(app?.components?.filters?.currentData) && app.components.filters.currentData.length) {
        return app.components.filters.currentData;
      }
      return app.currentData;
    },
    getFilteredData: () => app.filteredData,

    // 캐시 관련
    clearCache: () => app.clearLocalCache(),
    getCachedData: () => app.getCachedData(),

    // UI 관련 (스코프 명시 헬퍼도 함께 노출)
    showSuccess: (message) => app.showSuccessMessage(message),  // 기본: 페이지
    showError: (message) => app.showErrorMessage(message),      // 기본: 시스템
    showSystemAlert: (message, type) => app.showSystemAlert(message, type),
    showPageAlert: (message, type) => app.showPageAlert(message, type),

    // 컴포넌트 접근
    getComponents: () => app.components,

    // 권한 관련
    getPermissions: () => app.userPermissions,

    // 소켓 관련
    getSocket: () => app.socket
  };

  // 모달 함수들
  const modals = {
    openNewProjectModal: () => {
      logger.debug('[ProjectListAPI] openNewProjectModal 호출됨');
      const components = app.components;
      logger.debug('[ProjectListAPI] app.components:', components);
      logger.debug('[ProjectListAPI] modernModal 존재:', !!components.modernModal);

      if (components.modernModal) {
        logger.debug('[ProjectListAPI] modernModal.open() 호출');
        components.modernModal.open();
      } else {
        logger.error('[ProjectListAPI] ModernProjectModal 컴포넌트가 초기화되지 않았습니다.');
        logger.error('[ProjectListAPI] 사용 가능한 컴포넌트:', Object.keys(components));
      }
    },

    openAuditLogs: () => {
      const components = app.components;
      if (components.auditLogModal) {
        components.auditLogModal.open();
      } else {
        logger.error('[AUDIT_LOGS] AuditLogModal이 아직 로드되지 않았습니다.');
      }
    },

    showProjectDetails: (projectCode) => {
      const detailsRow = document.getElementById('details-' + projectCode);
      if (detailsRow) {
        detailsRow.classList.remove('d-none');
      }
    },

    syncCalendar: async () => {
      logger.debug('[ProjectListAPI] syncCalendar 호출됨');
      const btn = document.getElementById('calendarSyncBtn');

      try {
        // 버튼을 로딩 상태로 변경
        if (btn) {
          buttonStateManager.setLoading(btn, '동기화 중...');
        }

        // API 호출
        const response = await fetch('/api/projects/calendar/sync', {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json'
          },
          body: JSON.stringify({})
        });

        const data = await response.json();

        if (!data.success) {
          throw new Error(data.error || '동기화에 실패했습니다.');
        }

        logger.info('[ProjectListAPI] 캘린더 동기화 성공');

        // 데이터 새로고침
        if (api.refreshData) {
          api.refreshData(true, false);
        }

        // 성공 상태로 변경
        if (btn) {
          buttonStateManager.setSuccess(btn, '완료', () => {
            // 1.5초 후 원래 상태로 복원
            setTimeout(() => {
              buttonStateManager.reset(btn);
            }, 1500);
          });
        }

      } catch (error) {
        logger.error('[ProjectListAPI] 캘린더 동기화 실패:', error);

        // 오류 시 에러 상태로 변경
        if (btn) {
          buttonStateManager.setError(btn, '실패', '캘린더 동기화');
        }
      }
    }
  };

  // 통합 네임스페이스: window.ITGlobalApp
  window.ITGlobalApp = {
    // API 객체
    api: api,

    // 모달 함수들
    modals: modals,

    // 편의 함수 (직접 접근)
    refreshData: api.refreshData,
    fullRefresh: api.fullRefresh,
    getCurrentData: api.getCurrentData,
    clearCache: api.clearCache,
    openNewProjectModal: modals.openNewProjectModal,
    openAuditLogs: modals.openAuditLogs,
    showProjectDetails: modals.showProjectDetails,
    syncCalendar: modals.syncCalendar
  };

  // 레거시 호환성: 기존 전역 변수들 유지 (deprecated 경고)
  // throttling: 각 함수당 한 번만 경고 (콘솔 도배 방지)
  const createDeprecatedWrapper = (newPath, fn) => {
    let warningShown = false;
    return function(...args) {
      if (!warningShown) {
        logger.warn(`[DEPRECATED] 이 함수는 더 이상 사용되지 않습니다. 대신 ${newPath}를 사용하세요.`);
        warningShown = true;
      }
      return fn(...args);
    };
  };

  window.projectAPI = api; // 유지 (많이 사용됨)
  window.refreshProjectData = createDeprecatedWrapper('ITGlobalApp.refreshData', api.refreshData);
  window.getCurrentProjectData = createDeprecatedWrapper('ITGlobalApp.getCurrentData', api.getCurrentData);
  window.clearProjectCache = createDeprecatedWrapper('ITGlobalApp.clearCache', api.clearCache);
  window.openNewProjectModal = createDeprecatedWrapper('ITGlobalApp.openNewProjectModal', modals.openNewProjectModal);
  window.loadAuditLogs = createDeprecatedWrapper('ITGlobalApp.openAuditLogs', modals.openAuditLogs);
  window.openAuditLogs = createDeprecatedWrapper('ITGlobalApp.openAuditLogs', modals.openAuditLogs);
  window.showProjectDetails = createDeprecatedWrapper('ITGlobalApp.showProjectDetails', modals.showProjectDetails);
  window.syncCalendar = modals.syncCalendar; // 새로운 함수는 deprecated 경고 없이 직접 노출
  window.refreshData = api.refreshData; // 새로고침 버튼용 직접 노출
  window.fullRefresh = api.fullRefresh; // 완전 동기화 (서버 캐시 + 클라이언트) 직접 노출

  logger.info('[ProjectListAPI] 전역 API가 window.ITGlobalApp 네임스페이스에 노출되었습니다.');
  logger.info('[ProjectListAPI] 레거시 전역 함수들은 deprecated 경고와 함께 유지됩니다.');

  return api;
}

export default exposeGlobalAPI;
