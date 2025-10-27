/**
 * 편집 상태 관리 클래스 (Single Source of Truth)
 *
 * 문제: 기존에는 5개 소스에서 데이터를 읽음
 *   1. DOM input.value
 *   2. this.currentProject
 *   3. DataTable
 *   4. this.originalMemos
 *   5. Multi-select checkboxes
 *
 * 해결: 모든 편집 데이터를 EditState 하나에서 관리
 */
import { normalizeFieldValue } from './valueNormalizer.js';

import logger from '../utils/logger.js';
// 🆕 디버그 모드 플래그 (프로덕션에서는 false로 설정)
const DEBUG_EDIT_STATE = false;

export default class EditState {
  constructor() {
    this.originalData = {};   // 편집 시작 시점 스냅샷 (불변)
    this.currentData = {};    // 현재 편집 중인 값 (유일한 진실)
    this.dirtyFields = new Set();  // 변경된 필드 추적
    this.isActive = false;    // 편집 모드 활성 여부
  }

  /**
   * 편집 모드 시작 - 초기 데이터 저장
   * @param {Object} projectData - 프로젝트 전체 데이터
   */
  initialize(projectData) {
    // 깊은 복사로 원본 보존
    this.originalData = JSON.parse(JSON.stringify(projectData));
    this.currentData = JSON.parse(JSON.stringify(projectData));
    this.dirtyFields.clear();
    this.isActive = true;

    if (DEBUG_EDIT_STATE) {
      logger.debug('[EditState] 편집 모드 초기화:', {
        projectCode: projectData['프로젝트 코드'],
        fieldCount: Object.keys(projectData).length
      });
    }
  }

  /**
   * 필드 타입 추론 (필드명 기반)
   * @param {string} fieldName - 필드 이름
   * @returns {string} 필드 타입 ('money', 'checkbox', 'text')
   */
  getFieldType(fieldName) {
    // 금액 필드
    const moneyFields = ['총액', '제품대', '도급비', '자재비', '기타비', '계약금', '중도금', '잔금', '미수금'];
    if (moneyFields.some(f => fieldName.includes(f))) {
      return 'money';
    }

    // 체크박스 필드
    if (fieldName === '부가세' || fieldName === '수금 확인') {
      return 'checkbox';
    }

    // 나머지는 텍스트 (날짜, 드롭다운, 일반 텍스트 포함)
    return 'text';
  }

  /**
   * 필드 값 업데이트 (모든 변경의 단일 진입점)
   * @param {string} fieldName - 필드 이름
   * @param {any} newValue - 새 값
   * @param {boolean} markDirty - dirty 플래그 설정 여부 (기본: true)
   */
  updateField(fieldName, newValue, markDirty = true) {
    if (!this.isActive) {
      if (DEBUG_EDIT_STATE) {
        logger.warn('[EditState] 편집 모드가 활성화되지 않았습니다.');
      }
      return;
    }

    const oldValue = this.currentData[fieldName];

    // 🆕 필드 타입별 자동 정규화
    const fieldType = this.getFieldType(fieldName);
    const normalizedValue = normalizeFieldValue(newValue, fieldType);

    this.currentData[fieldName] = normalizedValue;

    if (markDirty) {
      this.dirtyFields.add(fieldName);
    }

    if (DEBUG_EDIT_STATE) {
      logger.debug(`[EditState] 필드 업데이트: ${fieldName}`, {
        oldValue,
        newValue,
        normalizedValue,
        fieldType,
        isDirty: this.dirtyFields.has(fieldName)
      });
    }
  }

  /**
   * 필드 값 조회
   * @param {string} fieldName - 필드 이름
   * @returns {any} 현재 값
   */
  getField(fieldName) {
    return this.currentData[fieldName];
  }

  /**
   * 원본 값 조회
   * @param {string} fieldName - 필드 이름
   * @returns {any} 편집 시작 시점 값
   */
  getOriginalField(fieldName) {
    return this.originalData[fieldName];
  }

  /**
   * 필드 변경 여부 확인
   * @param {string} fieldName - 필드 이름
   * @returns {boolean} 변경되었는지 여부
   */
  isFieldDirty(fieldName) {
    return this.dirtyFields.has(fieldName);
  }

  /**
   * 변경사항 수집 (단순화된 로직!)
   * @returns {Object} 변경된 필드와 정규화된 값
   */
  collectChanges() {
    const changes = {};

    this.dirtyFields.forEach(fieldName => {
      const currentValue = this.currentData[fieldName];
      const originalValue = this.originalData[fieldName];

      // 필드 타입별 정규화 (currentData는 이미 정규화되어 있지만, originalData는 아닐 수 있음)
      const fieldType = this.getFieldType(fieldName);
      const normalizedCurrent = normalizeFieldValue(currentValue, fieldType);
      const normalizedOriginal = normalizeFieldValue(originalValue, fieldType);

      // 실제로 변경되었는지 확인
      if (normalizedCurrent !== normalizedOriginal) {
        // 🆕 정규화된 값 반환 (서버/UI 타입 정합성 보장)
        changes[fieldName] = normalizedCurrent;
      }
    });

    if (DEBUG_EDIT_STATE) {
      logger.debug('[EditState] 변경사항 수집:', {
        totalDirtyFields: this.dirtyFields.size,
        actualChanges: Object.keys(changes).length,
        changes
      });
    }

    return changes;
  }

  /**
   * 값 정규화 (비교용)
   * @param {any} value - 정규화할 값
   * @returns {string} 정규화된 값
   */
  normalizeValue(value) {
    if (value === null || value === undefined || value === 'null' || value === 'undefined' || value === '-') {
      return '';
    }
    return String(value);
  }

  /**
   * 특정 필드 초기화 (변경 취소)
   * @param {string} fieldName - 필드 이름
   */
  resetField(fieldName) {
    this.currentData[fieldName] = this.originalData[fieldName];
    this.dirtyFields.delete(fieldName);

    if (DEBUG_EDIT_STATE) {
      logger.debug(`[EditState] 필드 초기화: ${fieldName} = ${this.originalData[fieldName]}`);
    }
  }

  /**
   * 전체 초기화 (편집 취소)
   */
  reset() {
    this.currentData = JSON.parse(JSON.stringify(this.originalData));
    this.dirtyFields.clear();

    if (DEBUG_EDIT_STATE) {
      logger.debug('[EditState] 전체 초기화 완료');
    }
  }

  /**
   * 편집 모드 종료
   */
  destroy() {
    this.originalData = {};
    this.currentData = {};
    this.dirtyFields.clear();
    this.isActive = false;

    if (DEBUG_EDIT_STATE) {
      logger.debug('[EditState] 편집 모드 종료');
    }
  }

  /**
   * 디버그 정보 출력
   */
  debug() {
    if (DEBUG_EDIT_STATE) {
      console.group('[EditState] 디버그 정보');
      logger.debug('활성 상태:', this.isActive);
      logger.debug('변경된 필드:', Array.from(this.dirtyFields));
      logger.debug('현재 데이터:', this.currentData);
      logger.debug('원본 데이터:', this.originalData);
      console.groupEnd();
    }
  }
}
