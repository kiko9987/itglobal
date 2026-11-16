import logger from '../utils/logger.js';
import { buttonStateManager } from '../utils/ButtonStateManager.js';
import { showFriendlyError } from '../utils/userFriendlyErrors.js';

// 새로고침 진행 중 플래그 (연타 방지) + 진행 중 AbortController
let _refreshInFlight = false;
let _refreshAbortCtrl = null;

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
      // 연타 방지 — 진행 중이면 조용히 return
      if (_refreshInFlight) {
        logger.debug('[ProjectListAPI] 새로고침 이미 진행 중, skip');
        return;
      }
      // 펼침에서 편집 중이면 서버 캐시를 지우기 전에 먼저 묻는다 — 이어서 편집하기로 하면 아무것도 안 함 (11-16)
      if (typeof app.confirmLeaveEditForRefresh === 'function' && !app.confirmLeaveEditForRefresh()) {
        return;
      }
      _refreshInFlight = true;

      // 진행 중 요청은 새 요청 시작 전 취소 (AbortController)
      if (_refreshAbortCtrl) {
        try { _refreshAbortCtrl.abort(); } catch (_) {}
      }
      _refreshAbortCtrl = new AbortController();
      // 서버 응답 없으면 60초 후 자동 취소 (구글 시트 full fetch ≈ 10~15초라 여유)
      const timeoutId = setTimeout(() => {
        try { _refreshAbortCtrl?.abort(); } catch (_) {}
      }, 60000);

      logger.debug('[ProjectListAPI] fullRefresh 호출 - 서버 캐시 삭제 + 클라이언트 갱신');

      const btn = document.getElementById('fullRefreshBtn');

      // 3초 경과 시 사용자에게 상세 진행 상태 안내 (Google Sheets full fetch 10~15초 케이스)
      const _progressTimerId = setTimeout(() => {
        if (btn && _refreshInFlight) {
          try {
            const icon = btn.querySelector('i');
            const iconHtml = icon ? icon.outerHTML : '<i class="fas fa-circle-notch fa-spin me-1"></i>';
            btn.innerHTML = `${iconHtml}Google Sheets 재로드 중...`;
          } catch (_) {}
        }
      }, 3000);

      try {
        // 버튼을 로딩 상태로 변경
        if (btn) {
          buttonStateManager.setLoading(btn, '새로고침 중...');
        }

        // 1단계: 서버 캐시 삭제 (DOM 요소 없이 직접 API 호출)
        logger.debug('[ProjectListAPI] 서버 캐시 삭제 중...');
        const response = await fetch('/api/cache/refresh', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          credentials: 'same-origin',
          body: JSON.stringify({}),
          signal: _refreshAbortCtrl.signal,
        });

        // 응답 body 안전 파싱 — 서버 재시작 중이면 body 잘림
        let result;
        try {
          const text = await response.text();
          result = text ? JSON.parse(text) : {};
        } catch (parseErr) {
          throw new Error(response.ok
            ? '서버 응답 처리 중 문제가 발생했습니다. 잠시 후 다시 시도해 주세요.'
            : `서버 오류 (HTTP ${response.status})`);
        }

        if (!result.success) {
          logger.warn('[ProjectListAPI] 서버 캐시 삭제 실패:', result.message || result.error);
          const err = new Error(result.message || result.error || '서버 캐시 삭제 실패');
          if (result.error_id) err.error_id = result.error_id;
          throw err;
        }

        logger.info('[ProjectListAPI] 서버 캐시 삭제 성공');

        // 2단계: 클라이언트 데이터 갱신 — refreshData 가 돌려주는 결과로 판정.
        //   예전엔 refreshData 가 실패를 삼켜(던지지 않음) 목록을 못 불러와도 '완료'가 떴고,
        //   다른 새로고침이 도는 중이라 바로 돌아온 경우도 '완료'였음 (PC 점검 11-16).
        //   알림은 여기서 한 번만 띄우도록 refreshData 쪽 알림(showMessage)은 끈다.
        let refreshStatus = 'failed';
        try {
          refreshStatus = await app.refreshData(true, true, false, { userInitiated: true });
          // 다른 새로고침(소켓 등)이 도는 중 — 끝나길 기다렸다가 지운 캐시 기준으로 다시 (최대 3번)
          for (let i = 0; i < 3 && refreshStatus === 'busy'; i++) {
            try { await app._refreshPromise; } catch (_) { /* 앞 새로고침 결과와 상관없이 다시 시도 */ }
            refreshStatus = await app.refreshData(true, true, false, { userInitiated: true });
          }
        } catch (refreshErr) {
          refreshStatus = 'failed';
          logger.error('[ProjectListAPI] 클라이언트 데이터 갱신 실패:', refreshErr);
        }

        // 성공 상태로 변경 (버튼 매니저가 최소 스피너 노출 시간 자동 보장)
        if (refreshStatus === 'ok') {
          if (btn) {
            buttonStateManager.setSuccess(btn, '완료', () => {
              setTimeout(() => buttonStateManager.reset(btn), 1500);
            });
          }
        } else if (refreshStatus === 'cancelled') {
          // 기다리는 사이 다시 편집을 시작해 '이어서 편집'을 고름 — 실패가 아니므로 단추만 원래대로
          if (btn) buttonStateManager.reset(btn);
        } else {
          if (btn) buttonStateManager.setError(btn, '실패', '새로고침');
          if (app.showSystemAlert) {
            app.showSystemAlert(
              '목록을 다시 불러오지 못했습니다. 잠시 뒤 [새로고침]을 다시 누르거나 F5를 눌러 주세요.',
              'error',
            );
          }
        }

      } catch (error) {
        // AbortError = 새 요청이 이전 요청 취소한 것. 사용자 액션이 아니라 무시.
        if (error.name === 'AbortError') {
          logger.debug('[ProjectListAPI] 이전 새로고침 요청 취소됨 (새 요청이 대신 진행)');
          return;
        }
        logger.error('[ProjectListAPI] 서버 캐시 삭제 오류:', error);

        // 오류 시 에러 상태로 변경
        if (btn) {
          buttonStateManager.setError(btn, '실패', '새로고침');
        }

        // 사용자 친화적 알림 (error_id 있으면 표시)
        showFriendlyError(error, '새로고침', { duration: 8000 });

      } finally {
        clearTimeout(timeoutId);
        clearTimeout(_progressTimerId);
        _refreshInFlight = false;
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

        // 오류 시 에러 상태로 변경 + 서버가 준 이유를 보여 준다(예전엔 '실패'만 떠서 왜인지 몰랐음 — PC 점검 2차 11-16)
        if (btn) {
          buttonStateManager.setError(btn, '실패', '캘린더 동기화');
        }
        try {
          const { default: Toast } = await import('../components/Toast.js');
          new Toast().show(error?.message || '캘린더 동기화에 실패했습니다.', 'warning', 7000);
        } catch (_) { /* 알림 실패는 무시 */ }
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
