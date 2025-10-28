/**
 * ModeManager - 중앙 집중식 모드 상태 관리 클래스
 *
 * 시스템의 모든 모드 상태를 한 곳에서 관리하고,
 * 모드 전환 시 검증 및 이벤트 발생을 담당합니다.
 *
 * 특징:
 * - Single Source of Truth: 모드 상태의 유일한 출처
 * - 안전한 상태 변경: 검증 로직을 통한 안전한 변경
 * - 이벤트 기반: 모드 변경 시 리스너에게 자동 알림
 * - 중앙 집중식 로깅: 모든 모드 변경 추적
 */

import {
  TABLE_MODE,
  ACCORDION_MODE,
  getCurrentModeCombination,
  getModeDescription
} from '../constants/ViewModes.js';
import logger from './logger.js';

export default class ModeManager {
  // Private fields (ES2022)
  #tableMode = TABLE_MODE.NORMAL;
  #accordionMode = ACCORDION_MODE.VIEW;
  #listeners = [];
  #changeHistory = []; // 모드 변경 이력
  #maxHistorySize = 50; // 최대 이력 보관 개수

  constructor(options = {}) {
    // 초기 모드 설정 (옵션)
    if (options.initialTableMode) {
      this.#tableMode = options.initialTableMode;
    }
    if (options.initialAccordionMode) {
      this.#accordionMode = options.initialAccordionMode;
    }

    // 디버그 모드
    this.debugMode = options.debugMode || false;

    // 초기 상태 로깅
    this.#log('info', `[ModeManager] 초기화 완료: ${this.getDescription()}`);
  }

  // ============================================================================
  // 상태 조회 (Getters)
  // ============================================================================

  /**
   * 현재 테이블 모드 조회
   * @returns {string} TABLE_MODE.NORMAL | TABLE_MODE.RECEIVABLES
   */
  getTableMode() {
    return this.#tableMode;
  }

  /**
   * 현재 아코디언 모드 조회
   * @returns {string} ACCORDION_MODE.VIEW | ACCORDION_MODE.EDIT
   */
  getAccordionMode() {
    return this.#accordionMode;
  }

  /**
   * 현재 모드 조합 조회
   * @returns {object} { table: string, accordion: string }
   */
  getCurrentCombination() {
    return getCurrentModeCombination(
      this.#tableMode === TABLE_MODE.RECEIVABLES,
      this.#accordionMode === ACCORDION_MODE.EDIT
    );
  }

  /**
   * 현재 모드 설명 조회
   * @returns {string} "일반 관리 + 뷰 모드"
   */
  getDescription() {
    return getModeDescription(
      this.#tableMode === TABLE_MODE.RECEIVABLES,
      this.#accordionMode === ACCORDION_MODE.EDIT
    );
  }

  /**
   * 기존 코드 호환성을 위한 boolean 조회
   */
  isReceivablesMode() {
    return this.#tableMode === TABLE_MODE.RECEIVABLES;
  }

  isEditMode() {
    return this.#accordionMode === ACCORDION_MODE.EDIT;
  }

  // ============================================================================
  // 상태 변경 (Setters with Validation)
  // ============================================================================

  /**
   * 테이블 모드 변경
   * @param {string} newMode - TABLE_MODE.NORMAL | TABLE_MODE.RECEIVABLES
   * @param {object} options - { force: boolean, skipValidation: boolean }
   * @returns {boolean} 변경 성공 여부
   */
  setTableMode(newMode, options = {}) {
    // 1. 유효성 검증
    if (!Object.values(TABLE_MODE).includes(newMode)) {
      this.#log('error', `[ModeManager] 유효하지 않은 테이블 모드: ${newMode}`);
      return false;
    }

    // 2. 같은 모드면 스킵
    if (this.#tableMode === newMode) {
      this.#log('debug', `[ModeManager] 이미 ${newMode} 상태입니다`);
      return true;
    }

    // 3. 검증 (force 옵션으로 스킵 가능)
    if (!options.skipValidation && !options.force) {
      const validationResult = this.#validateTableModeChange(newMode);
      if (!validationResult.valid) {
        this.#log('warn', `[ModeManager] 테이블 모드 변경 차단: ${validationResult.reason}`);
        return false;
      }
    }

    // 4. 이전 상태 저장
    const oldMode = this.#tableMode;

    // 5. 상태 변경
    this.#tableMode = newMode;

    // 6. 이력 저장
    this.#addHistory({
      type: 'tableMode',
      oldValue: oldMode,
      newValue: newMode,
      timestamp: Date.now()
    });

    // 7. 이벤트 발생
    this.#emit('tableModeChanged', {
      oldMode,
      newMode,
      currentCombination: this.getCurrentCombination()
    });

    // 8. 로깅 (운영 환경에서는 debug 레벨로 억제)
    this.#log('debug', `[ModeManager] 테이블 모드 변경: ${oldMode} → ${newMode}`);

    return true;
  }

  /**
   * 아코디언 모드 변경
   * @param {string} newMode - ACCORDION_MODE.VIEW | ACCORDION_MODE.EDIT
   * @param {object} options - { force: boolean, skipValidation: boolean }
   * @returns {boolean} 변경 성공 여부
   */
  setAccordionMode(newMode, options = {}) {
    // 1. 유효성 검증
    if (!Object.values(ACCORDION_MODE).includes(newMode)) {
      this.#log('error', `[ModeManager] 유효하지 않은 아코디언 모드: ${newMode}`);
      return false;
    }

    // 2. 같은 모드면 스킵
    if (this.#accordionMode === newMode) {
      this.#log('debug', `[ModeManager] 이미 ${newMode} 상태입니다`);
      return true;
    }

    // 3. 검증 (force 옵션으로 스킵 가능)
    if (!options.skipValidation && !options.force) {
      const validationResult = this.#validateAccordionModeChange(newMode);
      if (!validationResult.valid) {
        this.#log('warn', `[ModeManager] 아코디언 모드 변경 차단: ${validationResult.reason}`);
        return false;
      }
    }

    // 4. 이전 상태 저장
    const oldMode = this.#accordionMode;

    // 5. 상태 변경
    this.#accordionMode = newMode;

    // 6. 이력 저장
    this.#addHistory({
      type: 'accordionMode',
      oldValue: oldMode,
      newValue: newMode,
      timestamp: Date.now()
    });

    // 7. 이벤트 발생
    this.#emit('accordionModeChanged', {
      oldMode,
      newMode,
      currentCombination: this.getCurrentCombination()
    });

    // 8. 로깅 (운영 환경에서는 debug 레벨로 억제)
    this.#log('debug', `[ModeManager] 아코디언 모드 변경: ${oldMode} → ${newMode}`);

    return true;
  }

  // ============================================================================
  // 검증 로직 (Private)
  // ============================================================================

  /**
   * 테이블 모드 변경 검증
   * @private
   */
  #validateTableModeChange(newMode) {
    // 규칙 1: 편집 모드 활성 시 테이블 모드 변경 차단
    if (this.#accordionMode === ACCORDION_MODE.EDIT) {
      return {
        valid: false,
        reason: '편집 모드가 활성화되어 있습니다. 편집을 종료한 후 테이블 모드를 변경해주세요.'
      };
    }

    // 모든 검증 통과
    return { valid: true };
  }

  /**
   * 아코디언 모드 변경 검증
   * @private
   */
  #validateAccordionModeChange(newMode) {
    // 현재는 아코디언 모드 변경에 특별한 제약 없음
    // 필요시 여기에 검증 로직 추가 가능
    return { valid: true };
  }

  // ============================================================================
  // 이벤트 시스템
  // ============================================================================

  /**
   * 이벤트 리스너 등록
   * @param {string} event - 이벤트명 ('tableModeChanged' | 'accordionModeChanged' | 'modeChanged')
   * @param {function} callback - 콜백 함수
   */
  on(event, callback) {
    if (typeof callback !== 'function') {
      this.#log('error', '[ModeManager] 콜백은 함수여야 합니다');
      return;
    }

    this.#listeners.push({ event, callback });
    this.#log('debug', `[ModeManager] 이벤트 리스너 등록: ${event}`);
  }

  /**
   * 이벤트 리스너 제거
   * @param {string} event - 이벤트명
   * @param {function} callback - 콜백 함수
   */
  off(event, callback) {
    const index = this.#listeners.findIndex(
      l => l.event === event && l.callback === callback
    );

    if (index !== -1) {
      this.#listeners.splice(index, 1);
      this.#log('debug', `[ModeManager] 이벤트 리스너 제거: ${event}`);
    }
  }

  /**
   * 이벤트 발생
   * @private
   */
  #emit(event, data) {
    // 특정 이벤트 리스너 실행
    this.#listeners
      .filter(l => l.event === event)
      .forEach(l => {
        try {
          l.callback(data);
        } catch (error) {
          this.#log('error', `[ModeManager] 이벤트 핸들러 오류 (${event}):`, error);
        }
      });

    // 'modeChanged' 이벤트는 모든 모드 변경 시 발생
    if (event !== 'modeChanged') {
      this.#listeners
        .filter(l => l.event === 'modeChanged')
        .forEach(l => {
          try {
            l.callback({
              eventType: event,
              ...data
            });
          } catch (error) {
            this.#log('error', '[ModeManager] modeChanged 핸들러 오류:', error);
          }
        });
    }
  }

  // ============================================================================
  // 이력 관리
  // ============================================================================

  /**
   * 모드 변경 이력 추가
   * @private
   */
  #addHistory(entry) {
    this.#changeHistory.push(entry);

    // 최대 크기 초과 시 오래된 이력 제거
    if (this.#changeHistory.length > this.#maxHistorySize) {
      this.#changeHistory.shift();
    }
  }

  /**
   * 모드 변경 이력 조회
   * @param {number} limit - 조회할 이력 개수 (기본: 10)
   * @returns {array} 이력 배열
   */
  getHistory(limit = 10) {
    return this.#changeHistory.slice(-limit);
  }

  /**
   * 이력 초기화
   */
  clearHistory() {
    this.#changeHistory = [];
    this.#log('info', '[ModeManager] 이력 초기화 완료');
  }

  // ============================================================================
  // 로깅
  // ============================================================================

  /**
   * 로그 출력
   * @private
   */
  #log(level, ...args) {
    // 디버그 모드가 아니면 debug 레벨 로그는 출력하지 않음
    if (level === 'debug' && !this.debugMode) {
      return;
    }

    const logMethod = {
      debug: console.log,
      info: console.log,
      warn: console.warn,
      error: console.error
    }[level] || console.log;

    logMethod(...args);
  }

  // ============================================================================
  // 유틸리티
  // ============================================================================

  /**
   * 현재 상태 출력 (디버깅용)
   */
  printState() {
    logger.group('[ModeManager] 현재 상태');
    logger.debug('테이블 모드:', this.#tableMode);
    logger.debug('아코디언 모드:', this.#accordionMode);
    logger.debug('모드 조합:', this.getCurrentCombination());
    logger.debug('모드 설명:', this.getDescription());
    logger.debug('활성 리스너:', this.#listeners.length);
    logger.debug('이력 개수:', this.#changeHistory.length);
    logger.groupEnd();
  }

  /**
   * 상태 초기화
   */
  reset() {
    this.#tableMode = TABLE_MODE.NORMAL;
    this.#accordionMode = ACCORDION_MODE.VIEW;
    this.#changeHistory = [];
    this.#log('info', '[ModeManager] 상태 초기화 완료');

    // 이벤트 발생
    this.#emit('modeReset', {
      tableMode: this.#tableMode,
      accordionMode: this.#accordionMode
    });
  }

  /**
   * 리스너 정리
   */
  destroy() {
    this.#listeners = [];
    this.#changeHistory = [];
    this.#log('info', '[ModeManager] 정리 완료');
  }
}
