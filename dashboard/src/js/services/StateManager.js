import logger from '../utils/logger.js';

/**
 * StateManager - 애플리케이션 상태 관리 전담
 * 전문가 리뷰: "상태 저장을 각각 별도 클래스로 나눠 app.init()이 orchestration만 담당하게"
 */

export default class StateManager {
  constructor() {
    this.currentData = [];
    this.filteredData = [];
    this.userPermissions = {};
    this.isInitialized = false;
    this.components = {};

    // 상태 변경 리스너들
    this.listeners = {
      dataChange: [],
      filterChange: [],
      permissionChange: []
    };
  }

  /**
   * 초기 상태 설정
   */
  initializeState() {
    // 사용자 권한 확인
    const userRole = window.userRole || 'viewer';
    this.userPermissions = {
      canEdit: userRole === 'admin' || userRole === 'editor',
      canDelete: userRole === 'admin',
      canCreate: userRole === 'admin' || userRole === 'editor'
    };

    this.isInitialized = true;

    // 권한 변경 알림
    this.notifyListeners('permissionChange', this.userPermissions);
  }

  /**
   * 사용자 권한 설정 (호환성 메서드)
   */
  setUserPermissions(permissions) {
    this.userPermissions = { ...this.userPermissions, ...permissions };

    // 권한 변경 알림
    this.notifyListeners('permissionChange', this.userPermissions);
  }

  /**
   * 컴포넌트 등록
   */
  registerComponents(components) {
    this.components = components;
  }

  /**
   * 현재 데이터 업데이트
   * @param {Array} data - 새로운 데이터
   * @param {boolean} skipFilterApply - 필터 재적용 건너뛰기 (기본값: false)
   */
  setCurrentData(data, skipFilterApply = false) {
    this.currentData = data;

    // 전역 호환성 (레거시) - 읽기 전용으로 보호
    Object.defineProperty(window, 'projectsData', {
      value: data,
      writable: false,
      configurable: true  // 다음 업데이트를 위해 재설정 가능하도록
    });

    // 데이터 변경 알림
    this.notifyListeners('dataChange', data);

    // 필터 재적용 (조건부)
    if (!skipFilterApply) {
      this.applyCurrentFilters();
    } else {
    }
  }

  /**
   * 필터링된 데이터 업데이트
   */
  setFilteredData(data) {
    this.filteredData = data;

    // 필터 변경 알림
    this.notifyListeners('filterChange', data);
  }

  /**
   * 현재 필터 적용
   */
  applyCurrentFilters() {
    if (this.components.filters && this.currentData.length > 0) {
      const filterResult = this.components.filters.applyFilters(this.currentData);
      const filteredData = filterResult !== null ? filterResult : this.currentData;
      this.setFilteredData(filteredData);
    } else {
      this.setFilteredData(this.currentData);
    }
  }

  /**
   * 데이터 새로고침 (필터 상태 유지)
   */
  refreshWithNewData(newData) {
    this.setCurrentData(newData);
  }

  /**
   * 권한에 따른 UI 적용
   */
  applyPermissions() {
    const editButtons = document.querySelectorAll('.btn-edit');
    const deleteButtons = document.querySelectorAll('.btn-delete');
    const createButton = document.getElementById('createProjectBtn');

    if (!this.userPermissions.canEdit) {
      editButtons.forEach(btn => btn.style.display = 'none');
    }

    if (!this.userPermissions.canDelete) {
      deleteButtons.forEach(btn => btn.style.display = 'none');
    }

    if (!this.userPermissions.canCreate && createButton) {
      createButton.style.display = 'none';
    }

  }

  /**
   * 상태 변경 리스너 등록
   */
  onDataChange(callback) {
    this.listeners.dataChange.push(callback);
  }

  onFilterChange(callback) {
    this.listeners.filterChange.push(callback);
  }

  onPermissionChange(callback) {
    this.listeners.permissionChange.push(callback);
  }

  /**
   * 리스너들에게 알림
   */
  notifyListeners(event, data) {
    if (this.listeners[event]) {
      this.listeners[event].forEach(callback => {
        try {
          callback(data);
        } catch (error) {
          logger.error(`[StateManager] ${event} 리스너 오류:`, error);
        }
      });
    }
  }

  /**
   * 현재 상태 반환
   */
  getCurrentState() {
    return {
      currentData: this.currentData,
      filteredData: this.filteredData,
      userPermissions: this.userPermissions,
      isInitialized: this.isInitialized,
      componentCount: Object.keys(this.components).length
    };
  }

  /**
   * 특정 프로젝트 데이터 찾기
   */
  findProject(projectCode) {
    return this.currentData.find(project => project['프로젝트 코드'] === projectCode);
  }

  /**
   * 필터된 데이터에서 특정 프로젝트 찾기
   */
  findFilteredProject(projectCode) {
    return this.filteredData.find(project => project['프로젝트 코드'] === projectCode);
  }

  /**
   * 특정 프로젝트 데이터 업데이트 (부분 업데이트)
   * @param {string} projectCode - 프로젝트 코드 (새 프로젝트 코드)
   * @param {Object} updatedProject - 업데이트된 프로젝트 데이터
   * @param {string} oldProjectCode - 이전 프로젝트 코드 (프로젝트 코드 변경 시)
   * @returns {boolean} 업데이트 성공 여부
   */
  updateSingleProject(projectCode, updatedProject, oldProjectCode = null, options = {}) {
    try {
      // options.silent: true이면 dataChange 알림 + 필터 재적용 skip
      // (2026-07-07): 취소/재개처럼 호출자가 이미 로컬 DOM을 부분 업데이트한 경우,
      // 리스트 전체 리렌더는 아코디언 상태·스크롤 위치를 리셋시켜서 UX 저해.
      // 이 경우 currentData만 조용히 교체하고 렌더링 트리거는 건너뜀.
      const silent = options.silent === true;

      const searchCode = oldProjectCode || projectCode;

      const index = this.currentData.findIndex(p => p['프로젝트 코드'] === searchCode);

      if (index === -1) {
        logger.warn(`[StateManager] 프로젝트 ${searchCode}를 찾을 수 없습니다`);
        logger.warn(`[StateManager] currentData 크기: ${this.currentData.length}`);
        return false;
      }

      this.currentData[index] = updatedProject;

      Object.defineProperty(window, 'projectsData', {
        value: this.currentData,
        writable: false,
        configurable: true
      });

      if (oldProjectCode && oldProjectCode !== projectCode) {
        logger.debug(`[StateManager] 프로젝트 코드 변경: ${oldProjectCode} → ${projectCode}${silent ? ' (silent)' : ''}`);
      } else {
        logger.debug(`[StateManager] 프로젝트 ${projectCode} 업데이트 완료${silent ? ' (silent)' : ''}`);
      }

      if (!silent) {
        this.notifyListeners('dataChange', this.currentData);
        this.applyCurrentFilters();
      }

      return true;
    } catch (error) {
      logger.error('[StateManager] 프로젝트 업데이트 실패:', error);
      return false;
    }
  }

  /**
   * 상태 디버그 정보
   */
  getDebugInfo() {
    return {
      currentDataCount: this.currentData.length,
      filteredDataCount: this.filteredData.length,
      permissions: this.userPermissions,
      isInitialized: this.isInitialized,
      listenerCounts: {
        dataChange: this.listeners.dataChange.length,
        filterChange: this.listeners.filterChange.length,
        permissionChange: this.listeners.permissionChange.length
      }
    };
  }

  /**
   * 정리 작업
   */
  cleanup() {
    this.listeners.dataChange = [];
    this.listeners.filterChange = [];
    this.listeners.permissionChange = [];
    this.components = {};
  }
}