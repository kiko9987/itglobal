/**
 * 뷰 모드 상수 정의
 *
 * 시스템에는 2가지 독립적인 모드 축이 존재:
 * 1. 테이블 레벨 모드 (전체 프로젝트 목록 관리 방식)
 * 2. 아코디언 레벨 모드 (개별 프로젝트 상세 편집 방식)
 */

/**
 * 테이블 레벨 모드
 * - 전체 프로젝트 목록의 관리 방식을 결정
 * - ProjectTable 컴포넌트에서 관리
 */
export const TABLE_MODE = {
  /** 프로젝트 일반 관리 모드 - 모든 프로젝트 표시, 일반 관리 기능 */
  NORMAL: 'normal',

  /** 프로젝트 수금 관리 모드 - 미수금 프로젝트 중심, 수금 관련 기능 강조 */
  RECEIVABLES: 'receivables'
};

/**
 * 아코디언 레벨 모드
 * - 개별 프로젝트의 상세 정보 편집 방식을 결정
 * - ProjectRowAccordion 컴포넌트에서 관리
 */
export const ACCORDION_MODE = {
  /** 아코디언 일반 모드 - 정보 조회만 가능, 편집 불가 */
  VIEW: 'view',

  /** 아코디언 편집 모드 - 필드 편집 가능, EditState 활성화 */
  EDIT: 'edit'
};

/**
 * 모드 조합 상태
 * - 실제 시스템에서 가능한 4가지 조합 상태
 */
export const MODE_COMBINATION = {
  /** 일반 관리 + 뷰 모드 */
  NORMAL_VIEW: {
    table: TABLE_MODE.NORMAL,
    accordion: ACCORDION_MODE.VIEW
  },

  /** 일반 관리 + 편집 모드 */
  NORMAL_EDIT: {
    table: TABLE_MODE.NORMAL,
    accordion: ACCORDION_MODE.EDIT
  },

  /** 수금 관리 + 뷰 모드 */
  RECEIVABLES_VIEW: {
    table: TABLE_MODE.RECEIVABLES,
    accordion: ACCORDION_MODE.VIEW
  },

  /** 수금 관리 + 편집 모드 */
  RECEIVABLES_EDIT: {
    table: TABLE_MODE.RECEIVABLES,
    accordion: ACCORDION_MODE.EDIT
  }
};

/**
 * 테이블 모드 변환 헬퍼
 * boolean 플래그를 명확한 모드 상수로 변환
 */
export function getTableMode(isReceivablesMode) {
  return isReceivablesMode ? TABLE_MODE.RECEIVABLES : TABLE_MODE.NORMAL;
}

/**
 * 아코디언 모드 변환 헬퍼
 * boolean 플래그를 명확한 모드 상수로 변환
 */
export function getAccordionMode(isUnifiedEditMode) {
  return isUnifiedEditMode ? ACCORDION_MODE.EDIT : ACCORDION_MODE.VIEW;
}

/**
 * 현재 모드 조합 조회 헬퍼
 * 두 플래그 값으로 현재 시스템 상태 조회
 */
export function getCurrentModeCombination(isReceivablesMode, isUnifiedEditMode) {
  const tableMode = getTableMode(isReceivablesMode);
  const accordionMode = getAccordionMode(isUnifiedEditMode);

  return {
    table: tableMode,
    accordion: accordionMode
  };
}

/**
 * 모드 설명 조회 헬퍼
 * 디버깅/로깅용
 */
export function getModeDescription(isReceivablesMode, isUnifiedEditMode) {
  const combo = getCurrentModeCombination(isReceivablesMode, isUnifiedEditMode);

  const tableDesc = combo.table === TABLE_MODE.NORMAL ? '일반 관리' : '수금 관리';
  const accordionDesc = combo.accordion === ACCORDION_MODE.VIEW ? '뷰 모드' : '편집 모드';

  return `${tableDesc} + ${accordionDesc}`;
}
